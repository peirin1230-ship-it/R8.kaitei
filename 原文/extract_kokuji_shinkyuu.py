#!/usr/bin/env python3
"""
告示PDF新旧対照表 抽出スクリプト

R8告示PDF（令和８年厚生労働省告示第69号）とR6告示PDF（令和６年告示第57号）を比較し、
difflib で文字レベルの差分を検出して XLSX 形式の新旧対照表を生成する。
差分部分にはアンダーライン書式を適用する。
"""

import fitz  # PyMuPDF
import re
import sys
import os
import difflib
import xlsxwriter

# marker_utils は親ディレクトリにあるため、パスを追加
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from marker_utils import strip_sequence_marker, normalize_width, normalize_for_content_compare

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# ============================================================
# 定数
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
R8_PDF_PATH = os.path.join(SCRIPT_DIR,
    "診療報酬の算定方法の一部を改正する件（令和８年厚生労働省告示第69号）医科点数表.pdf")
R6_PDF_PATH = os.path.join(SCRIPT_DIR, "..", "令和6年度診療報酬", "原本",
    "診療報酬の算定方法の一部を改正する告示 令和６年 厚生労働省告示第57号 別表第一 （医科点数表）.pdf")
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, "output",
    "R8年度医科点数表(告示)_新旧対照表.xlsx")

MIN_FONT_SIZE = 5.0
RUBY_MAX_SIZE = 6.0
Y_GROUP_TOLERANCE = 3.0
# 告示PDFは単一カラム。見出し判定の最大x座標
HEADING_MAX_X = 250.0
# 目次ページ数（R8: 2ページ、R6: 2ページ）
R8_TOC_PAGES = 2
R6_TOC_PAGES = 2

# ============================================================
# 正規表現パターン（extract_shinkyuu.py から再利用）
# ============================================================
RE_CHAPTER = re.compile(r'^第([１２３４５６７８９０\d]+)章\s+(.*)')
RE_PART = re.compile(r'^第([１２３４５６７８９０\d]+)\s*部\s+(.*)')
RE_SECTION = re.compile(r'^第([１２３４５６７８９０\d]+)節\s+(.*)')
RE_SUBSECTION = re.compile(r'^第([１２３４５６７８９０\d]+)\s*款\s+(.*)')
RE_TSUSOKU = re.compile(r'^通則')
RE_KUBUN = re.compile(
    r'^([ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰａ-ｚA-Z]'
    r'[０-９0-9]{3}(?:－[０-９0-9]+(?:－[０-９0-9]+)?)?)\s+(.*)')
RE_NOTE = re.compile(r'^注([１２３４５６７８９0-9]*)\s')
RE_NOTE_NUM_ONLY = re.compile(r'^([０-９0-9１２３４５６７８９]+)\s')
RE_PAGE_NUM = re.compile(r'^\d{1,4}$')
# 見出しのみの行を検出（「第N部 名前」「第N節 名前」「区分」「削除」等）
RE_HEADING_ONLY = re.compile(
    r'^(第[１２３４５６７８９０\d]+\s*[章部節款]\s+.*|区分|削除\s*|通則)$')
# カナ項目検出
RE_KANA_ITEM = re.compile(r'^([アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン])\s')
# サブアイテム検出パターン（行頭）
RE_SUBITEM_PAREN = re.compile(r'^([（(][０-９0-9]+[）)])[\s\u3000]')
RE_SUBITEM_KANA = re.compile(
    r'^([アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン])[\s\u3000]')
RE_SUBITEM_ALPHA = re.compile(r'^([ａｂｃｄｅｆｇｈｉｊｋｌｍ])[\s\u3000]')
# paren型サブアイテムの名称抽出（noteフィールド用）
RE_PAREN_NAME = re.compile(
    r'^[（(][０-９0-9]+[）)]\s*(.+?)\s*(?:[０-９0-9,，]+点|[①②③④⑤]|$)')
RE_IS_ITEM_NAME = re.compile(r'(?:料|加算[０-９0-9]*)$')
# 番号付きサブ項目の検出（B001「１ ウイルス疾患指導料」等）
# 全角数字 + 短い名称（医療行為名で終わるもの）をブロック境界として認識
# 注の条件文（「場合は」「した」等を含む長文）は除外する
RE_NAMED_SUBITEM = re.compile(
    r'^([１２３４５６７８９０0-9]+)\s+'
    r'(.{2,25}?(?:料|指導|削除)(?:\s*[（(].+?[）)])?)\s*$')
RE_NAMED_SUBITEM_WITH_POINTS = re.compile(
    r'^([１２３４５６７８９０0-9]+)\s+'
    r'(.{2,25}?(?:料|指導|削除)(?:\s*[（(].+?[）)])?)'
    r'\s+[\d０-９,，]+点')
# 条件文パターン（サブ項目名ではない）
RE_NOT_SUBITEM_NAME = re.compile(r'場合|した|する|について|ものと|である|おいて|により|から|まで')


# ============================================================
# PDFテキスト抽出
# ============================================================

def extract_page_lines_single_column(page):
    """単一カラムの告示PDFからテキスト行を抽出する。

    Returns:
        [(text, x0), ...] のリスト（y座標順）
    """
    td = page.get_text('dict')

    spans_data = []
    for block in td['blocks']:
        if 'lines' not in block:
            continue
        for line in block['lines']:
            for span in line['spans']:
                if span['size'] < MIN_FONT_SIZE:
                    continue
                # ルビ（ひらがなのみの小さいスパン）を除外、それ以外は保持
                if span['size'] < RUBY_MAX_SIZE and re.match(r'^[ぁ-ん]+$', span['text'].strip()):
                    continue
                text = span['text']
                if not text.strip():
                    continue
                bbox = span['bbox']
                spans_data.append({
                    'x0': bbox[0],
                    'y0': bbox[1],
                    'x1': bbox[2],
                    'y1': bbox[3],
                    'text': text,
                })

    if not spans_data:
        return []

    spans_data.sort(key=lambda s: (s['y0'], s['x0']))

    # y座標でグループ化して行を構成
    y_groups = []
    for s in spans_data:
        placed = False
        for g in y_groups:
            if abs(g['y'] - s['y0']) < Y_GROUP_TOLERANCE:
                g['spans'].append(s)
                placed = True
                break
        if not placed:
            y_groups.append({'y': s['y0'], 'spans': [s]})

    y_groups.sort(key=lambda g: g['y'])

    lines = []
    for g in y_groups:
        sorted_spans = sorted(g['spans'], key=lambda s: s['x0'])
        text = ''.join(s['text'] for s in sorted_spans).strip()
        # ダッシュ文字を全角マイナス（－ U+FF0D）に正規化
        text = text.replace('\u2015', '－').replace('\u2014', '－').replace('\u2013', '－')
        x0 = sorted_spans[0]['x0']

        # ページ番号行を除外
        if RE_PAGE_NUM.match(text):
            continue
        if not text:
            continue

        lines.append((text, x0))

    return lines


# ============================================================
# 階層追跡（告示PDF用）
# ============================================================

class HierarchyTracker:
    """章/部/節/款/通則/区分番号/注/サブ項目 の階層を追跡する"""

    def __init__(self):
        self.chapter = ""
        self.part = ""
        self.section = ""
        self.subsection = ""
        self.item_code = ""
        self.item_name = ""
        self.sub_item = ""  # B001「１ ウイルス疾患指導料」等の番号付きサブ項目
        self.note = ""
        self.last_note_x = 0
        self.last_item_x = 0
        self.name_incomplete = False

    def _check_named_subitem(self, text):
        """番号付きサブ項目（B001「１ ウイルス疾患指導料」等）を検出する。
        点数付き行からも名称を抽出する。
        条件文（注の本文）は除外する。"""
        for pat in (RE_NAMED_SUBITEM, RE_NAMED_SUBITEM_WITH_POINTS):
            m = pat.match(text)
            if m:
                name = m.group(2).strip()
                # 条件文パターンを含む場合は注の本文であり、サブ項目名ではない
                if RE_NOT_SUBITEM_NAME.search(name):
                    continue
                return m.group(1), name
        return None, None

    def update(self, text, x_pos):
        if not text:
            return False

        # 項目名が括弧未完結の場合、次行のテキストを連結
        if self.name_incomplete:
            if RE_CHAPTER.match(text) or RE_PART.match(text) or \
               RE_SECTION.match(text) or RE_SUBSECTION.match(text) or \
               RE_KUBUN.match(text) or RE_NOTE.match(text) or \
               RE_TSUSOKU.match(text):
                self.name_incomplete = False
            else:
                addition = text.strip()
                addition = re.sub(r'\s+[\d０-９,，]+点.*$', '', addition)
                self.item_name += addition.strip()
                op = self.item_name.count('（') + self.item_name.count('(')
                cp = self.item_name.count('）') + self.item_name.count(')')
                self.name_incomplete = (op > cp)
                return True

        m = RE_CHAPTER.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.chapter = f"第{m.group(1)}章 {m.group(2)}".strip()
            self.part = ""
            self.section = ""
            self.subsection = ""
            self.item_code = ""
            self.item_name = ""
            self.sub_item = ""
            self.note = ""
            return True

        m = RE_PART.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.part = f"第{m.group(1).strip()}部 {m.group(2)}".strip()
            self.section = ""
            self.subsection = ""
            self.item_code = ""
            self.item_name = ""
            self.sub_item = ""
            self.note = ""
            return True

        m = RE_SECTION.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.section = f"第{m.group(1).strip()}節 {m.group(2)}".strip()
            self.subsection = ""
            self.item_code = ""
            self.item_name = ""
            self.sub_item = ""
            self.note = ""
            return True

        m = RE_SUBSECTION.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.subsection = f"第{m.group(1).strip()}款 {m.group(2)}".strip()
            self.item_code = ""
            self.item_name = ""
            self.sub_item = ""
            self.note = ""
            return True

        m = RE_TSUSOKU.match(text)
        if m:
            self.item_code = "通則"
            self.item_name = ""
            self.sub_item = ""
            self.note = ""
            self.last_item_x = x_pos
            return True

        m = RE_KUBUN.match(text)
        if m:
            self.item_code = m.group(1)
            raw_name = m.group(2).strip()
            raw_name = re.sub(r'\s+[\d０-９,，]+点.*$', '', raw_name)
            self.item_name = raw_name.strip()
            self.sub_item = ""
            self.note = ""
            self.last_item_x = x_pos
            op = self.item_name.count('（') + self.item_name.count('(')
            cp = self.item_name.count('）') + self.item_name.count(')')
            self.name_incomplete = (op > cp)
            return True

        # 番号付きサブ項目の検出（注の検出より前に行う）
        sub_num, sub_name = self._check_named_subitem(text)
        if sub_num is not None:
            self.sub_item = f"{sub_num} {sub_name}"
            self.note = ""
            return True

        m = RE_NOTE.match(text)
        if m:
            note_num = m.group(1)
            self.note = f"注{note_num}" if note_num else "注"
            self.last_note_x = x_pos
            return True

        m = RE_NOTE_NUM_ONLY.match(text)
        if m:
            num_str = m.group(1)
            if self.note and abs(x_pos - self.last_note_x) < 15:
                self.note = f"注{num_str}"
                return True
            if self.item_code == "通則" and abs(x_pos - self.last_item_x) < 15:
                self.note = f"注{num_str}"
                return True

        return False

    def snapshot(self):
        return {
            'chapter': self.chapter,
            'part': self.part,
            'section': self.section,
            'subsection': self.subsection,
            'item_code': self.item_code,
            'item_name': self.item_name,
            'sub_item': self.sub_item,
            'note': self.note,
        }


# ============================================================
# ブロック分割
# ============================================================

def is_block_boundary(text, x0):
    """テキストが新しいブロック境界（区分番号変更、注変更、通則）かを判定"""
    if RE_KUBUN.match(text):
        return True
    if RE_TSUSOKU.match(text):
        return True
    if RE_CHAPTER.match(text) and x0 < HEADING_MAX_X:
        return True
    if RE_PART.match(text) and x0 < HEADING_MAX_X:
        return True
    if RE_SECTION.match(text) and x0 < HEADING_MAX_X:
        return True
    if RE_SUBSECTION.match(text) and x0 < HEADING_MAX_X:
        return True
    # 番号付きサブ項目（B001「１ ウイルス疾患指導料」等）
    if RE_NAMED_SUBITEM.match(text) or RE_NAMED_SUBITEM_WITH_POINTS.match(text):
        return True
    # 注の開始
    m = RE_NOTE.match(text)
    if m:
        return True
    return False


def is_heading_only_block(text):
    """ブロックが見出しのみ（構造情報のみ）かを判定。
    見出しのみのブロックはマッチング対象外にする。
    """
    lines = text.strip().split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if not RE_HEADING_ONLY.match(line):
            return False
    return True


def split_blocks_at_subitems(blocks):
    """ブロックをサブアイテム（(1)、ア、ａ等）の境界で分割する。

    通則ブロックは内部構造が複雑なため分割しない。
    同一タイプのマーカーが2つ以上ある場合のみ分割する。
    """
    result = []
    for block in blocks:
        if block['item_code'] == '通則':
            result.append(block)
            continue
        # 再帰的に分割（paren→kana→alphaの順で階層分割）
        sub_blocks = _split_single_block(block)
        for sb in sub_blocks:
            if sb.get('_from_split'):
                deeper = _split_single_block(sb)
                for db in deeper:
                    if db.get('_from_split') and db is not sb:
                        even_deeper = _split_single_block(db)
                        result.extend(even_deeper)
                    else:
                        result.append(db)
            else:
                result.append(sb)
    return result


def _split_single_block(block):
    """単一ブロックをサブアイテム境界で分割する。"""
    lines = block['text'].split('\n')

    # 各行のサブアイテムマーカーを検出
    markers = []  # (line_idx, type, marker_text)
    for i, line in enumerate(lines):
        stripped = line.strip()
        m = RE_SUBITEM_PAREN.match(stripped)
        if m:
            markers.append((i, 'paren', m.group(1)))
            continue
        m = RE_SUBITEM_KANA.match(stripped)
        if m:
            markers.append((i, 'kana', m.group(1)))
            continue
        m = RE_SUBITEM_ALPHA.match(stripped)
        if m:
            markers.append((i, 'alpha', m.group(1)))
            continue

    if len(markers) < 2:
        return [block]

    # 同一タイプのマーカーが2つ以上あるタイプのみ有効
    # 階層構造を保持するため1タイプずつ分割
    # 最初に出現するマーカーのタイプを上位とする（動的優先度）
    type_counts = {}
    for _, mtype, _ in markers:
        type_counts[mtype] = type_counts.get(mtype, 0) + 1

    valid_types = None
    for _, mtype, _ in markers:
        if type_counts.get(mtype, 0) >= 2:
            valid_types = {mtype}
            break
    if not valid_types:
        return [block]

    markers = [(i, t, m) for i, t, m in markers if t in valid_types]
    if len(markers) < 2:
        return [block]

    result = []
    note_base = block['note']

    # プリアンブル（最初のマーカーより前のテキスト）
    first_line = markers[0][0]
    if first_line > 0:
        preamble = '\n'.join(lines[:first_line])
        if preamble.strip():
            result.append({**block, 'text': preamble})

    # サブアイテムブロック
    for idx, (start, mtype, marker) in enumerate(markers):
        end = markers[idx + 1][0] if idx + 1 < len(markers) else len(lines)
        sub_text = '\n'.join(lines[start:end])
        effective_marker = marker
        if mtype == 'paren':
            first_line_text = lines[start].strip()
            m_name = RE_PAREN_NAME.match(first_line_text)
            if m_name:
                candidate = m_name.group(1).strip()
                if RE_IS_ITEM_NAME.search(candidate):
                    effective_marker = f"{marker}{candidate}"
        sub_note = f"{note_base}\u3000{effective_marker}" if note_base else effective_marker
        result.append({**block, 'text': sub_text, 'note': sub_note,
                       '_from_split': True})

    return result


def extract_blocks_from_pdf(pdf_path, toc_pages):
    """告示PDFからテキストブロックを抽出する。

    各ブロックは同一の階層コンテキスト内の連続テキスト。
    区分番号・注・通則の変更でブロックを分割する。

    Returns:
        [{chapter, part, section, subsection, item_code, item_name, note, text, page}, ...]
    """
    print(f"  PDFを読み込み中: {os.path.basename(pdf_path)}", file=sys.stderr)
    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    print(f"  総ページ数: {total_pages}", file=sys.stderr)

    tracker = HierarchyTracker()
    blocks = []
    current_lines = []
    current_ctx = None
    current_page = 0

    def save_current_block():
        nonlocal current_lines, current_ctx, current_page
        if current_lines and current_ctx:
            block_text = '\n'.join(current_lines)
            blocks.append({
                'chapter': current_ctx['chapter'],
                'part': current_ctx['part'],
                'section': current_ctx['section'],
                'subsection': current_ctx['subsection'],
                'item_code': current_ctx['item_code'],
                'item_name': current_ctx['item_name'],
                'sub_item': current_ctx.get('sub_item', ''),
                'note': current_ctx['note'],
                'text': block_text,
                'page': current_page,
            })

    for pg_idx in range(total_pages):
        if (pg_idx + 1) % 100 == 0:
            print(f"    処理中: {pg_idx+1}/{total_pages} ページ...", file=sys.stderr)

        page = doc[pg_idx]
        page_num = pg_idx + 1
        is_toc = (pg_idx < toc_pages)

        lines = extract_page_lines_single_column(page)

        for text, x0 in lines:
            # ブロック境界の判定（tracker更新前に行う）
            boundary = is_block_boundary(text, x0)

            # 裸の数字行も注の境界として検出（注１の後の２,３,...,16等）
            if not boundary:
                m = RE_NOTE_NUM_ONLY.match(text)
                if m:
                    if tracker.note and abs(x0 - tracker.last_note_x) < 15:
                        boundary = True
                    elif tracker.item_code == "通則" and abs(x0 - tracker.last_item_x) < 15:
                        boundary = True

            # 階層の更新
            tracker.update(text, x0)

            if is_toc:
                continue

            ctx = tracker.snapshot()

            if boundary or current_ctx is None:
                save_current_block()
                current_lines = [text]
                current_ctx = ctx
                current_page = page_num
            else:
                current_lines.append(text)

    # 最後のブロック
    save_current_block()

    doc.close()
    print(f"  抽出ブロック数: {len(blocks)}", file=sys.stderr)
    return blocks


def make_block_key(block):
    """ブロックのマッチングキーを生成。
    sub_itemを含めて、同一item_code内の異なるサブ項目を区別する。
    通則ブロック内の注は正規化する（R6/R8で分割が違うことがある）。
    """
    note = block['note']
    if block['item_code'] == '通則':
        note = ''
    return (block['chapter'], block['part'], block['section'],
            block['subsection'], block['item_code'],
            block.get('sub_item', ''), note)


def make_block_key_short(block):
    """フォールバック用の短いキー（sub_item含む）"""
    note = block['note']
    if block['item_code'] == '通則':
        note = ''
    return (block['item_code'], block.get('sub_item', ''), note)


# ============================================================
# 差分検出
# ============================================================

def normalize_text_for_compare(text):
    """比較用にテキストを正規化する。
    全角半角の違い、空白・改行の違いを吸収する。
    """
    # 全角英数字・記号・スペースを半角に統一（NFKC正規化）
    t = normalize_width(text)
    # 全ての空白文字（改行含む）を半角スペースに統一し、連続を1つに
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def compute_diff_segments(r8_text, r6_text):
    """2つのテキストの文字レベル差分をセグメントリストに変換する。

    Returns:
        (r8_segments, r6_segments): 各セグメントは [(text, is_changed), ...]
        テキストが同一の場合は (None, None) を返す。
    """
    # 正規化して比較（改行・空白の差異を無視）
    r8_norm = normalize_text_for_compare(r8_text)
    r6_norm = normalize_text_for_compare(r6_text)

    if r8_norm == r6_norm:
        return None, None

    # 先頭マーカー・文中注番号参照の変更のみなら同一扱い
    if normalize_for_content_compare(r8_text, normalize_text_for_compare) == \
       normalize_for_content_compare(r6_text, normalize_text_for_compare):
        return None, None

    # スペースを除去した全文が同一なら改行位置の違いのみ（同一扱い）
    if re.sub(r'\s', '', r8_norm) == re.sub(r'\s', '', r6_norm):
        return None, None

    # autojunk=False で正確な文字レベル差分を取得
    # (autojunk=True だとK番号リスト等で過剰に広いreplaceブロックが生成される)
    sm = difflib.SequenceMatcher(None, r6_norm, r8_norm, autojunk=False)
    ratio = sm.ratio()

    # 類似度が低すぎる場合は全体をアンダーライン
    if ratio < 0.3:
        return [(r8_text, True)], [(r6_text, True)]

    r8_segments = []
    r6_segments = []

    # opcodeを一括取得して文字シフトパターンを検出
    opcodes = sm.get_opcodes()

    for oi, (op, i1, i2, j1, j2) in enumerate(opcodes):
        r6_part = r6_norm[i1:i2]
        r8_part = r8_norm[j1:j2]
        if op == 'equal':
            r8_segments.append((r8_part, False))
            r6_segments.append((r6_part, False))
        elif op == 'replace':
            # スペースを全て除去して同一なら改行位置の違いのみ
            if (not r6_part.strip() and not r8_part.strip()) or \
               re.sub(r'\s', '', r6_part) == re.sub(r'\s', '', r8_part):
                r6_segments.append((r6_part, False))
                r8_segments.append((r8_part, False))
            else:
                # 大きなreplaceブロック内の改行差分を分離するため再分析
                sub_sm = difflib.SequenceMatcher(
                    None, r6_part, r8_part, autojunk=False)
                for sub_op, si1, si2, sj1, sj2 in sub_sm.get_opcodes():
                    sr6 = r6_part[si1:si2]
                    sr8 = r8_part[sj1:sj2]
                    if sub_op == 'equal':
                        r6_segments.append((sr6, False))
                        r8_segments.append((sr8, False))
                    elif sub_op == 'replace':
                        if re.sub(r'\s', '', sr6) == re.sub(r'\s', '', sr8):
                            r6_segments.append((sr6, False))
                            r8_segments.append((sr8, False))
                        else:
                            r6_segments.append((sr6, True))
                            r8_segments.append((sr8, True))
                    elif sub_op == 'insert':
                        r8_segments.append((sr8, bool(sr8.strip())))
                    elif sub_op == 'delete':
                        r6_segments.append((sr6, bool(sr6.strip())))
        elif op == 'insert':
            # 改行位置の違いによる文字シフトを検出:
            # insert(X) + equal(' ') + delete(X) パターン
            is_shift = False
            if r8_part.strip() and len(r8_part.strip()) <= 2:
                if oi + 2 < len(opcodes):
                    next_op = opcodes[oi + 1]
                    next2_op = opcodes[oi + 2]
                    if (next_op[0] == 'equal' and
                        r6_norm[next_op[1]:next_op[2]].strip() == '' and
                        next2_op[0] == 'delete' and
                        r6_norm[next2_op[1]:next2_op[2]] == r8_part):
                        is_shift = True
            r8_segments.append((r8_part, not is_shift and bool(r8_part.strip())))
        elif op == 'delete':
            # delete(X) + equal(' ') + insert(X) パターン
            is_shift = False
            if r6_part.strip() and len(r6_part.strip()) <= 2:
                if oi + 2 < len(opcodes):
                    next_op = opcodes[oi + 1]
                    next2_op = opcodes[oi + 2]
                    if (next_op[0] == 'equal' and
                        r8_norm[next_op[3]:next_op[4]].strip() == '' and
                        next2_op[0] == 'insert' and
                        r8_norm[next2_op[3]:next2_op[4]] == r6_part):
                        is_shift = True
            r6_segments.append((r6_part, not is_shift and bool(r6_part.strip())))

    # 変更セグメントが1つもなければ同一テキスト扱い
    r8_has_change = any(is_changed for _, is_changed in r8_segments)
    r6_has_change = any(is_changed for _, is_changed in r6_segments)
    if not r8_has_change and not r6_has_change:
        return None, None

    return r8_segments, r6_segments


def text_similarity(text1, text2):
    """2つのテキストの類似度を返す（0.0〜1.0）"""
    t1 = normalize_text_for_compare(text1)
    t2 = normalize_text_for_compare(text2)
    if not t1 and not t2:
        return 1.0
    if not t1 or not t2:
        return 0.0
    return difflib.SequenceMatcher(None, t1, t2).ratio()


# ============================================================
# XLSX出力
# ============================================================

def write_rich_cell(ws, row_idx, col_idx, segments, normal_fmt, ul_fmt):
    """セグメントリストからxlsxwriterのリッチテキストセルを書き込む。"""
    if not segments:
        ws.write_string(row_idx, col_idx, '', normal_fmt)
        return

    if len(segments) == 1 and not segments[0][1]:
        ws.write_string(row_idx, col_idx, segments[0][0], normal_fmt)
        return

    parts = []
    for text, is_changed in segments:
        if not text:
            continue
        fmt = ul_fmt if is_changed else normal_fmt
        parts.extend([fmt, text])

    if not parts:
        ws.write_string(row_idx, col_idx, '', normal_fmt)
        return

    all_same = all(parts[i] is parts[0] for i in range(0, len(parts), 2))
    if all_same:
        text = ''.join(parts[i] for i in range(1, len(parts), 2))
        ws.write_string(row_idx, col_idx, text, parts[0])
    else:
        ws.write_rich_string(row_idx, col_idx, *parts, normal_fmt)


# ============================================================
# メイン処理
# ============================================================

def main():
    print("=" * 60, file=sys.stderr)
    print("告示PDF新旧対照表 抽出開始", file=sys.stderr)
    print("=" * 60, file=sys.stderr)

    # Phase 1: 両方のPDFからブロックを抽出
    print("\n[Phase 1] PDFからブロック抽出", file=sys.stderr)
    r8_blocks_raw = extract_blocks_from_pdf(R8_PDF_PATH, R8_TOC_PAGES)
    r6_blocks_raw = extract_blocks_from_pdf(R6_PDF_PATH, R6_TOC_PAGES)

    # 見出しのみのブロックを除外
    r8_blocks = [b for b in r8_blocks_raw if not is_heading_only_block(b['text'])]
    r6_blocks = [b for b in r6_blocks_raw if not is_heading_only_block(b['text'])]
    print(f"  R8: {len(r8_blocks_raw)} → {len(r8_blocks)} ブロック（見出し除外後）",
          file=sys.stderr)
    print(f"  R6: {len(r6_blocks_raw)} → {len(r6_blocks)} ブロック（見出し除外後）",
          file=sys.stderr)

    # 通則ブロックの結合：同一キーの通則ブロックをまとめる
    # （R6/R8で分割が異なることがあるため）
    def merge_tsusoku_blocks(blocks):
        merged = []
        i = 0
        while i < len(blocks):
            b = blocks[i]
            if b['item_code'] == '通則':
                # 同じ階層の通則ブロックを連続して結合
                combined_text = b['text']
                combined_page = b['page']
                j = i + 1
                while j < len(blocks) and blocks[j]['item_code'] == '通則' and \
                      blocks[j]['chapter'] == b['chapter'] and \
                      blocks[j]['part'] == b['part'] and \
                      blocks[j]['section'] == b['section']:
                    combined_text += '\n' + blocks[j]['text']
                    j += 1
                merged.append({
                    'chapter': b['chapter'],
                    'part': b['part'],
                    'section': b['section'],
                    'subsection': b['subsection'],
                    'item_code': b['item_code'],
                    'item_name': b['item_name'],
                    'note': '',
                    'text': combined_text,
                    'page': combined_page,
                })
                i = j
            else:
                merged.append(b)
                i += 1
        return merged

    r8_blocks = merge_tsusoku_blocks(r8_blocks)
    r6_blocks = merge_tsusoku_blocks(r6_blocks)
    print(f"  R8: {len(r8_blocks)} ブロック（通則結合後）", file=sys.stderr)
    print(f"  R6: {len(r6_blocks)} ブロック（通則結合後）", file=sys.stderr)

    # サブアイテム分割: (1)、ア、ａ等の境界でブロックを細分化
    r8_before = len(r8_blocks)
    r6_before = len(r6_blocks)
    r8_blocks = split_blocks_at_subitems(r8_blocks)
    r6_blocks = split_blocks_at_subitems(r6_blocks)
    print(f"  R8: {r8_before} → {len(r8_blocks)} ブロック（サブアイテム分割後）",
          file=sys.stderr)
    print(f"  R6: {r6_before} → {len(r6_blocks)} ブロック（サブアイテム分割後）",
          file=sys.stderr)

    # Phase 2: ブロックマッチング
    print("\n[Phase 2] ブロックマッチング", file=sys.stderr)

    # R6ブロックをキーでインデックス化
    r6_by_key = {}
    for b in r6_blocks:
        key = make_block_key(b)
        r6_by_key.setdefault(key, []).append(b)

    r6_by_short_key = {}
    for b in r6_blocks:
        key = make_block_key_short(b)
        r6_by_short_key.setdefault(key, []).append(b)

    # サブアイテム分割ブロック用: item_codeのみのインデックス
    # 注番号がR8/R6間でずれた場合のフォールバック
    r6_by_code = {}
    for b in r6_blocks:
        r6_by_code.setdefault(b['item_code'], []).append(b)

    matched_r6_ids = set()
    # R6ブロックindex → output_rows内の挿入位置を記録（削除項目の位置計算用）
    r6_id_to_idx = {id(b): i for i, b in enumerate(r6_blocks)}
    r6_idx_to_output_pos = {}
    output_rows = []

    # R6ブロックをitem_code+sub_itemでインデックス化（同一サブ項目内マッチ用）
    r6_by_code_sub = {}
    for b in r6_blocks:
        cs_key = (b['item_code'], b.get('sub_item', ''))
        r6_by_code_sub.setdefault(cs_key, []).append(b)

    # R6ブロックを注の親部分でインデックス化（番号ずれ対策）
    # 例: "注10　ヤ" → 親キー (item_code, sub_item, "注10")
    def _note_parent(note):
        """注フィールドからイロハ/カナ等のサブマーカーを除去して親部分を返す。"""
        if not note:
            return note
        # "注10　ヤ" → "注10", "注５　(1)" → "注５" 等
        m = re.match(r'^(注[０-９0-9]*)', note)
        if m:
            return m.group(1)
        return note

    r6_by_note_parent = {}
    for b in r6_blocks:
        parent = _note_parent(b['note'])
        np_key = (b['item_code'], b.get('sub_item', ''), parent)
        r6_by_note_parent.setdefault(np_key, []).append(b)

    for r8b in r8_blocks:
        key = make_block_key(r8b)
        r6_match = None

        # 完全キーマッチ（複数候補があればテキスト類似度で選択）
        if key in r6_by_key:
            candidates = [c for c in r6_by_key[key] if id(c) not in matched_r6_ids]
            if len(candidates) == 1:
                r6_match = candidates[0]
            elif len(candidates) > 1:
                best = max(candidates,
                           key=lambda c: text_similarity(r8b['text'], c['text']))
                r6_match = best

        # 完全キーマッチ後の類似度チェック
        # キーが同じでも内容が大きく異なる場合は拒否（番号ずれ対策）
        skip_fallback = False
        if r6_match:
            sim = text_similarity(r8b['text'], r6_match['text'])
            if sim < 0.5:
                r6_match = None

        # フォールバック1: 同じ親注内で内容ベースマッチング（番号ずれ対策）
        # 例: 注10のイロハが入れ替わった場合、同じ注10内で最も類似する候補を探す
        if r6_match is None and not skip_fallback:
            parent = _note_parent(r8b['note'])
            np_key = (r8b['item_code'], r8b.get('sub_item', ''), parent)
            if np_key in r6_by_note_parent:
                candidates = [c for c in r6_by_note_parent[np_key]
                              if id(c) not in matched_r6_ids]
                if candidates:
                    best = max(candidates,
                               key=lambda c: text_similarity(r8b['text'], c['text']))
                    sim = text_similarity(r8b['text'], best['text'])
                    if sim > 0.5:
                        r6_match = best

        # フォールバック2: 短いキーでマッチ（類似度で選択）
        if r6_match is None and not skip_fallback:
            short_key = make_block_key_short(r8b)
            if short_key in r6_by_short_key:
                candidates = [c for c in r6_by_short_key[short_key]
                              if id(c) not in matched_r6_ids]
                if candidates:
                    best = max(candidates,
                               key=lambda c: text_similarity(r8b['text'], c['text']))
                    if text_similarity(r8b['text'], best['text']) > 0.3:
                        r6_match = best

        # フォールバック3: item_code+sub_itemでマッチ（注番号ズレ対策）
        if r6_match is None and not skip_fallback:
            cs_key = (r8b['item_code'], r8b.get('sub_item', ''))
            if cs_key in r6_by_code_sub and r8b['item_code'] != '通則':
                candidates = [c for c in r6_by_code_sub[cs_key]
                              if id(c) not in matched_r6_ids]
                if candidates:
                    best = max(candidates,
                               key=lambda c: text_similarity(r8b['text'], c['text']))
                    sim = text_similarity(r8b['text'], best['text'])
                    if sim > 0.5:
                        r6_match = best

        # フォールバック4: item_codeのみでマッチ（sub_itemが異なる場合）
        # より厳しい類似度閾値で誤マッチを防止
        if r6_match is None and not skip_fallback:
            code = r8b['item_code']
            if code in r6_by_code and code != '通則':
                candidates = [c for c in r6_by_code[code]
                              if id(c) not in matched_r6_ids]
                if candidates:
                    best = max(candidates,
                               key=lambda c: text_similarity(r8b['text'], c['text']))
                    sim = text_similarity(r8b['text'], best['text'])
                    if sim > 0.7:
                        r6_match = best

        if r6_match:
            matched_r6_ids.add(id(r6_match))
            r6_idx = r6_id_to_idx[id(r6_match)]

        r8_text = r8b['text']
        r6_text = r6_match['text'] if r6_match else ""

        if r6_match is None:
            # R8のみ → 新設
            r8_segments = [(r8_text, True)]
            r6_segments = [("（新設）", False)]
        else:
            r8_segments, r6_segments = compute_diff_segments(r8_text, r6_text)

            # --- マーカー変更フィルタリング ---
            if r8_segments is not None:
                r8_marker, r8_body = strip_sequence_marker(r8_text)
                r6_marker, r6_body = strip_sequence_marker(r6_text)
                if r8_marker and r6_marker and r8_marker != r6_marker:
                    body_r8_segs, body_r6_segs = compute_diff_segments(
                        r8_body, r6_body)
                    if body_r8_segs is None:
                        r8_segments = None
                        r6_segments = None
                    else:
                        r8_segments = [(r8_marker + ' ', False)] + body_r8_segs
                        r6_segments = [(r6_marker + ' ', False)] + body_r6_segs

        if r8_segments is None:
            # テキスト同一 → スキップ
            if r6_match:
                r6_idx_to_output_pos[r6_idx] = len(output_rows) - 1 if output_rows else -1
            continue

        if r6_match:
            r6_idx_to_output_pos[r6_idx] = len(output_rows)

        output_rows.append({
            'chapter': r8b['chapter'],
            'part': r8b['part'],
            'section': r8b['section'],
            'subsection': r8b['subsection'],
            'item_code': r8b['item_code'],
            'item_name': r8b['item_name'],
            'sub_item': r8b.get('sub_item', ''),
            'note': r8b['note'],
            'r8_segments': r8_segments,
            'r6_segments': r6_segments,
            'page': r8b['page'],
            '_from_split': r8b.get('_from_split', False),
        })

    # R6のみ（削除）のブロックをR6での出現順に基づく正しい位置に挿入
    # 各削除項目について、R6ブロックリスト内で直前のマッチ済みブロックを探し、
    # そのブロックに対応するoutput_rows内の位置の直後に挿入する
    deleted_items = []
    for r6_idx, r6b in enumerate(r6_blocks):
        if id(r6b) not in matched_r6_ids:
            # R6ブロックリスト内で直前のマッチ済みブロックを探す
            insert_pos = 0  # デフォルト: 先頭
            for prev_idx in range(r6_idx - 1, -1, -1):
                if prev_idx in r6_idx_to_output_pos:
                    insert_pos = r6_idx_to_output_pos[prev_idx] + 1
                    break
            r6_text = r6b['text']
            deleted_items.append((insert_pos, r6_idx, {
                'chapter': r6b['chapter'],
                'part': r6b['part'],
                'section': r6b['section'],
                'subsection': r6b['subsection'],
                'item_code': r6b['item_code'],
                'item_name': r6b['item_name'],
                'sub_item': r6b.get('sub_item', ''),
                'note': r6b['note'],
                'r8_segments': [("（削除）", False)],
                'r6_segments': [(r6_text, True)],
                'page': 0,
            }))

    # 後方から挿入してインデックスずれを回避
    # (insert_pos, r6_idx) を降順ソート: 同一位置ではR6の後方から挿入し相対順序を維持
    deleted_items.sort(key=lambda x: (x[0], x[1]), reverse=True)
    for insert_pos, _r6_idx, row in deleted_items:
        if insert_pos >= len(output_rows):
            output_rows.append(row)
        else:
            output_rows.insert(insert_pos, row)

    print(f"  差分ブロック数: {len(output_rows)}", file=sys.stderr)

    # Phase 2.5: カナ項目を注フィールドに付与
    # ブロックテキスト内で最初の変更箇所の直前にあるカナマーカーを検出し、
    # 注フィールドを「注N　ア」形式に更新する
    print("\n[Phase 2.5] カナ項目を注に付与", file=sys.stderr)
    re_kana_in_text = re.compile(
        r'(?:^|\s)([アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン])\s')
    kana_count = 0
    for row in output_rows:
        note = row.get('note', '')
        if not note:
            continue
        # サブアイテム分割済みの行はスキップ
        if '\u3000' in note or row.get('_from_split'):
            continue
        for segs in [row['r8_segments'], row['r6_segments']]:
            if not segs:
                continue
            # セグメントを結合してテキスト内の位置を追跡
            # 最初の変更箇所の位置を特定し、その直前のカナマーカーを探す
            pos = 0
            first_change_pos = None
            for text, is_changed in segs:
                if is_changed and text.strip() and first_change_pos is None:
                    first_change_pos = pos
                    break
                pos += len(text)
            if first_change_pos is None:
                continue
            # 変更箇所より前のテキストを結合
            preceding = ''.join(t for t, _ in segs)[:first_change_pos]
            # 直前のカナマーカーを検索（最後に出現するもの）
            matches = list(re_kana_in_text.finditer(preceding))
            if matches:
                found_kana = matches[-1].group(1)
                row['note'] = f"{note}\u3000{found_kana}"
                kana_count += 1
                break
    print(f"  カナ付与: {kana_count}件", file=sys.stderr)

    # B-2: カナ単体noteを直前行の親noteで補完（出力行に対して適用）
    RE_KANA_ONLY = re.compile(
        r'^[アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン]$')
    kana_fix_count = 0
    for i, row in enumerate(output_rows):
        if not row.get('_from_split'):
            continue
        if not RE_KANA_ONLY.match(row['note']):
            continue
        for j in range(i - 1, -1, -1):
            if output_rows[j]['item_code'] != row['item_code']:
                break
            prev_note = output_rows[j]['note']
            if prev_note and not RE_KANA_ONLY.match(prev_note):
                # prev_noteからカナ単体パーツを除去してベースnoteを抽出
                base_parts = [p for p in prev_note.split('\u3000')
                              if not RE_KANA_ONLY.match(p)]
                base_note = '\u3000'.join(base_parts)
                if base_note:
                    row['note'] = f"{base_note}\u3000{row['note']}"
                kana_fix_count += 1
                break
    print(f"  カナnote補完: {kana_fix_count}件", file=sys.stderr)

    # Phase 3: XLSX出力
    print(f"\n[Phase 3] XLSX出力: {OUTPUT_XLSX}", file=sys.stderr)

    os.makedirs(os.path.dirname(OUTPUT_XLSX), exist_ok=True)
    wb = xlsxwriter.Workbook(OUTPUT_XLSX)
    ws = wb.add_worksheet('新旧対照表')

    header_fmt = wb.add_format({
        'bold': True, 'font_name': '游ゴシック', 'font_size': 11,
        'text_wrap': True, 'valign': 'top',
        'border': 1, 'bg_color': '#D9E1F2',
    })
    normal_fmt = wb.add_format({
        'font_name': '游ゴシック', 'font_size': 11,
        'text_wrap': True, 'valign': 'top',
    })
    ul_fmt = wb.add_format({
        'font_name': '游ゴシック', 'font_size': 11,
        'underline': True,
        'text_wrap': True, 'valign': 'top',
    })

    headers = ['章', '部', '節', '款', '通則/項目コード', '注',
               '改正後（R8）', '改正前（R6）', 'ページ']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 12)
    ws.set_column(3, 3, 12)
    ws.set_column(4, 4, 25)
    ws.set_column(5, 5, 8)
    ws.set_column(6, 6, 55)
    ws.set_column(7, 7, 55)
    ws.set_column(8, 8, 6)

    for idx, row in enumerate(output_rows):
        r = idx + 1
        item_display = row['item_code']
        if row.get('item_name'):
            item_display = f"{row['item_code']} {row['item_name']}"
        if row.get('sub_item'):
            item_display = f"{item_display} {row['sub_item']}"

        ws.write_string(r, 0, row.get('chapter', ''), normal_fmt)
        ws.write_string(r, 1, row.get('part', ''), normal_fmt)
        ws.write_string(r, 2, row.get('section', ''), normal_fmt)
        ws.write_string(r, 3, row.get('subsection', ''), normal_fmt)
        ws.write_string(r, 4, item_display, normal_fmt)
        ws.write_string(r, 5, row.get('note', ''), normal_fmt)
        write_rich_cell(ws, r, 6, row['r8_segments'], normal_fmt, ul_fmt)
        write_rich_cell(ws, r, 7, row['r6_segments'], normal_fmt, ul_fmt)
        if row['page'] > 0:
            ws.write_number(r, 8, row['page'], normal_fmt)
        else:
            ws.write_string(r, 8, '', normal_fmt)

    wb.close()
    print(f"\n完了！ {len(output_rows)} 件の差分を出力しました。", file=sys.stderr)


if __name__ == '__main__':
    main()
