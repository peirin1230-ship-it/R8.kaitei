#!/usr/bin/env python3
"""
令和8年度 医科診療報酬点数表 新旧対照表 抽出スクリプト

PDFの左側（改正後）と右側（改正前）を比較し、
アンダーライン（傍線）が引かれた変更箇所のみを抽出して
XLSX形式の新旧対照表を生成する。アンダーライン部分には下線書式を適用。
"""

import fitz  # PyMuPDF
import re
import sys
import os
import copy
import xlsxwriter

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# ============================================================
# 定数
# ============================================================
PDF_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "総－２別紙１－１医科診療報酬点数表.pdf")
OUTPUT_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "R8年度診療報酬改定_新旧対照表.xlsx")

LEFT_X_MAX = 420.0        # 左カラム（改正後）の右端 x
UL_HEIGHT_MAX = 2.0       # アンダーライン矩形の最大高さ
UL_WIDTH_MIN = 5.0        # アンダーライン矩形の最小幅
UL_MATCH_TOLERANCE = 4.0  # テキスト bbox 下端とアンダーライン y のマッチ許容差
Y_GROUP_TOLERANCE = 3.0   # 同一行とみなす y 座標の許容差
MIN_FONT_SIZE = 6.0       # 抽出するテキストの最小フォントサイズ(ルビ除外用)
SKIP_PAGES_BEFORE = 0     # 全ページを走査（目次も階層コンテキスト取得に使用）
# 目次ページ終了位置（ここまではアンダーラインの変更を出力しない）
TOC_END_PAGE = 4

# 罫線の x 座標（これらは除外する）
BORDER_XS = [93.86, 420.41, 747.09]
BORDER_TOLERANCE = 5.0

# ============================================================
# 正規表現パターン
# ============================================================
RE_CHAPTER = re.compile(r'^第([１２３４５６７８９０\d]+)章\s+(.*)')
RE_PART = re.compile(r'^第([１２３４５６７８９０\d]+)部\s+(.*)')
RE_SECTION = re.compile(r'^第([１２３４５６７８９０\d]+)節\s+(.*)')
RE_SUBSECTION = re.compile(r'^第([１２３４５６７８９０\d]+)款\s+(.*)')
# 見出しの最大 x 座標（本文中の参照を除外するため）
HEADING_MAX_X = 180.0
RE_TSUSOKU = re.compile(r'^通則')
RE_KUBUN = re.compile(
    r'^([ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰａ-ｚA-Z]'
    r'[０-９0-9]{3}(?:－[０-９0-9]+(?:－[０-９0-9]+)?)?)\s+(.*)')
RE_NOTE = re.compile(r'^注([１２３４５６７８９0-9]*)\s')
RE_NOTE_NUM_ONLY = re.compile(r'^([０-９0-9１２３４５６７８９]+)\s')

# ブロック分割用：行頭の項目開始パターン
# （削る）（新設）、注N、数字、イロハ項目
RE_BLOCK_START = re.compile(
    r'^(（削る）|（新設）|注[１２３４５６７８９0-9]*\s|'
    r'[０-９0-9１２３４５６７８９]+\s|'
    r'[ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰａ-ｚA-Z][０-９0-9]{3}|'
    r'[イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス]\s)')


# ============================================================
# ヘルパー関数
# ============================================================

def get_underlines(page):
    """ページからアンダーライン矩形を抽出する。

    同一y座標上で近接するアンダーライン矩形は自動的にマージし、
    文字間のギャップによる検出漏れを防ぐ。
    """
    raw = []
    for d in page.get_drawings():
        rect = d['rect']
        x0 = rect.x0
        # 罫線を除外
        if any(abs(x0 - bx) < BORDER_TOLERANCE for bx in BORDER_XS):
            continue
        if rect.height < UL_HEIGHT_MAX and rect.width > UL_WIDTH_MIN:
            raw.append(rect)

    # --- 同一y座標の近接矩形をマージ ---
    # PDFでは「９」「８」のように個別の文字にそれぞれ短いUL矩形が
    # 割り当てられることがある。間のスペースが検出漏れするのを防ぐため、
    # 同じ行（y座標が近い）でx方向のギャップが小さい矩形を統合する。
    UL_MERGE_Y_TOL = 1.0    # 同一行とみなすy座標の差
    UL_MERGE_GAP = 3.0      # マージするx方向の最大ギャップ（PDF生成上の微小ギャップのみ対象）

    if not raw:
        return raw

    # y座標でグループ化
    y_groups = []
    for u in sorted(raw, key=lambda r: (r.y0, r.x0)):
        placed = False
        for g in y_groups:
            if abs(g['y'] - u.y0) < UL_MERGE_Y_TOL:
                g['rects'].append(u)
                placed = True
                break
        if not placed:
            y_groups.append({'y': u.y0, 'rects': [u]})

    merged = []
    for g in y_groups:
        rects = sorted(g['rects'], key=lambda r: r.x0)
        cur_x0, cur_y0, cur_x1, cur_y1 = (
            rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1)
        for r in rects[1:]:
            if r.x0 - cur_x1 < UL_MERGE_GAP:
                cur_x1 = max(cur_x1, r.x1)
                cur_y0 = min(cur_y0, r.y0)
                cur_y1 = max(cur_y1, r.y1)
            else:
                merged.append(fitz.Rect(cur_x0, cur_y0, cur_x1, cur_y1))
                cur_x0, cur_y0, cur_x1, cur_y1 = r.x0, r.y0, r.x1, r.y1
        merged.append(fitz.Rect(cur_x0, cur_y0, cur_x1, cur_y1))

    return merged


def _char_is_underlined(char_bbox, underlines):
    """文字がアンダーラインを持っているかを判定（文字中心 x で判定）"""
    x0, y0, x1, y1 = char_bbox
    cx = (x0 + x1) / 2
    for u in underlines:
        if abs(u.y0 - y1) < UL_MATCH_TOLERANCE:
            if u.x0 <= cx <= u.x1:
                return True
    return False


def _build_segments(chars):
    """連続する同一UL状態の文字をセグメントにグループ化する。

    Args:
        chars: [{'char': str, 'underlined': bool, ...}, ...] のリスト（x順ソート済み）

    Returns:
        [(text, is_underlined), ...] のリスト
    """
    if not chars:
        return []
    segments = []
    current_text = chars[0]['char']
    current_ul = chars[0]['underlined']
    for c in chars[1:]:
        if c['underlined'] == current_ul:
            current_text += c['char']
        else:
            segments.append((current_text, current_ul))
            current_text = c['char']
            current_ul = c['underlined']
    segments.append((current_text, current_ul))
    return segments


def _strip_segments(segments):
    """セグメントリストの先頭・末尾の空白を除去する。"""
    if not segments:
        return []
    result = list(segments)
    # 先頭セグメントの左空白を除去
    text, ul = result[0]
    text = text.lstrip()
    while not text and len(result) > 1:
        result.pop(0)
        text, ul = result[0]
        text = text.lstrip()
    result[0] = (text, ul)
    # 末尾セグメントの右空白を除去
    text, ul = result[-1]
    text = text.rstrip()
    while not text and len(result) > 1:
        result.pop()
        text, ul = result[-1]
        text = text.rstrip()
    result[-1] = (text, ul)
    # 空セグメントを除去
    result = [(t, u) for t, u in result if t]
    return result


def _segments_text(segments):
    """セグメントリストからテキストを結合して返す。"""
    return ''.join(t for t, _ in segments)


def extract_page_lines(page, underlines):
    """ページからテキスト行を抽出し、左右カラムに分ける。

    rawdict を使い文字単位でアンダーライン判定を行うことで、
    行内の部分的なアンダーラインを正確に反映する。
    """
    td = page.get_text('rawdict')

    chars_data = []
    for block in td['blocks']:
        if 'lines' not in block:
            continue
        for line in block['lines']:
            for span in line['spans']:
                if span['size'] < MIN_FONT_SIZE:
                    continue
                for char_info in span.get('chars', []):
                    c = char_info['c']
                    bbox = char_info['bbox']
                    origin = char_info['origin']
                    is_ul = _char_is_underlined(bbox, underlines)
                    side = 'left' if bbox[0] < LEFT_X_MAX else 'right'
                    chars_data.append({
                        'x': origin[0],
                        'y': origin[1],
                        'char': c,
                        'bbox': bbox,
                        'underlined': is_ul,
                        'side': side,
                    })

    chars_data.sort(key=lambda c: (c['y'], c['x']))

    # y 座標でグループ化
    y_groups = []
    for c in chars_data:
        placed = False
        for g in y_groups:
            if abs(g['y'] - c['y']) < Y_GROUP_TOLERANCE:
                g['chars'].append(c)
                placed = True
                break
        if not placed:
            y_groups.append({'y': c['y'], 'chars': [c]})

    y_groups.sort(key=lambda g: g['y'])

    lines = []
    for g in y_groups:
        left_chars = sorted([c for c in g['chars'] if c['side'] == 'left'],
                            key=lambda c: c['x'])
        right_chars = sorted([c for c in g['chars'] if c['side'] == 'right'],
                             key=lambda c: c['x'])

        # セグメント構築（文字単位のUL状態を反映）
        left_segments_raw = _build_segments(left_chars)
        right_segments_raw = _build_segments(right_chars)

        left_text = _segments_text(left_segments_raw).strip()
        right_text = _segments_text(right_segments_raw).strip()

        left_ul = any(c['underlined'] for c in left_chars) if left_chars else False
        right_ul = any(c['underlined'] for c in right_chars) if right_chars else False

        left_x = min((c['x'] for c in left_chars), default=999.0)
        right_x = min((c['x'] for c in right_chars), default=999.0)

        # ページ番号行の除外
        combined = (left_text + right_text).strip()
        if combined.isdigit() and len(combined) <= 4:
            continue
        if not combined:
            continue

        # セグメントの空白除去
        left_segments = _strip_segments(left_segments_raw)
        right_segments = _strip_segments(right_segments_raw)

        lines.append({
            'y': g['y'],
            'left_text': left_text,
            'right_text': right_text,
            'left_x': left_x,
            'right_x': right_x,
            'left_underlined': left_ul,
            'right_underlined': right_ul,
            'left_segments': left_segments,
            'right_segments': right_segments,
        })

    return lines


def line_has_change(line):
    """行に変更があるかを判定"""
    return (
        line['left_underlined'] or
        line['right_underlined'] or
        line['left_text'] == '（削る）' or
        line['right_text'] == '（新設）' or
        line['left_text'] == '（新設）' or
        line['right_text'] == '（削る）'
    )


def should_split_block(line, prev_line):
    """新しい変更ブロックを開始すべきかを判定。

    以下の場合に新ブロックを開始する:
    1. (削る) が左に出現
    2. (新設) が右に出現（ただし前の行も新設の場合は除く）
    3. アンダーラインのある側が切り替わった場合
    """
    lt = line['left_text']
    rt = line['right_text']

    # （削る）/（新設）は常に新ブロック（左右どちらに出現しても）
    if lt == '（削る）' or lt == '（新設）':
        return True
    if rt == '（新設）' or rt == '（削る）':
        return True

    # アンダーラインの側が切り替わり、かつ新しい項目が始まる場合のみ分割
    # （同じ変更箇所の中で行ごとにアンダーライン側が変わることがあるため、
    #  側の切り替わりだけでは分割しない）
    if prev_line:
        prev_left_active = prev_line['left_underlined'] or prev_line['left_text'] == '（削る）'
        prev_right_active = prev_line['right_underlined'] or prev_line['right_text'] == '（新設）'
        curr_left_active = line['left_underlined']
        curr_right_active = line['right_underlined']

        side_changed = False
        if prev_left_active and not prev_right_active:
            if curr_right_active and not curr_left_active:
                side_changed = True
        if prev_right_active and not prev_left_active:
            if curr_left_active and not curr_right_active:
                side_changed = True

        # 側が切り替わった場合は、新しい項目の開始パターンがある場合のみ分割
        if side_changed and (RE_BLOCK_START.match(lt) or RE_BLOCK_START.match(rt)):
            return True

    # 両側にアンダーラインがあり、新しい項目が始まる場合
    if line['left_underlined'] or line['right_underlined']:
        # 両方のテキストが項目開始パターンにマッチ
        if (RE_BLOCK_START.match(lt) and RE_BLOCK_START.match(rt)):
            return True

    return False


# ============================================================
# 階層構造の追跡
# ============================================================

class HierarchyTracker:
    """章/部/節/款/通則/区分番号/注 の階層を追跡する"""

    def __init__(self):
        self.chapter = ""
        self.part = ""
        self.section = ""
        self.subsection = ""
        self.item_code = ""
        self.item_name = ""
        self.note = ""
        self.last_note_x = 0  # 注のインデント位置
        self.last_item_x = 0  # 項目のインデント位置
        self.name_incomplete = False  # 項目名の括弧が未完結

    def update(self, text, x_pos):
        """テキストと x 座標から階層情報を更新する"""
        if not text or text in ('（削る）', '（新設）'):
            return False

        # 項目名が括弧未完結の場合、次行のテキストを連結する
        if self.name_incomplete:
            # 新しい見出しや項目が始まった場合は連結を中止
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
            self.note = ""
            return True

        m = RE_PART.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.part = f"第{m.group(1)}部 {m.group(2)}".strip()
            self.section = ""
            self.subsection = ""
            self.item_code = ""
            self.item_name = ""
            self.note = ""
            return True

        m = RE_SECTION.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.section = f"第{m.group(1)}節 {m.group(2)}".strip()
            self.subsection = ""
            self.item_code = ""
            self.item_name = ""
            self.note = ""
            return True

        m = RE_SUBSECTION.match(text)
        if m and x_pos < HEADING_MAX_X:
            self.subsection = f"第{m.group(1)}款 {m.group(2)}".strip()
            self.item_code = ""
            self.item_name = ""
            self.note = ""
            return True

        m = RE_TSUSOKU.match(text)
        if m:
            self.item_code = "通則"
            self.item_name = ""
            self.note = ""
            self.last_item_x = x_pos
            return True

        m = RE_KUBUN.match(text)
        if m:
            self.item_code = m.group(1)
            raw_name = m.group(2).strip()
            # 点数表記（例: "291点"、"1,000点"）と余分な空白を除去
            raw_name = re.sub(r'\s+[\d０-９,，]+点.*$', '', raw_name)
            self.item_name = raw_name.strip()
            self.note = ""
            self.last_item_x = x_pos
            # 括弧が閉じていない場合は次行で継続
            op = self.item_name.count('（') + self.item_name.count('(')
            cp = self.item_name.count('）') + self.item_name.count(')')
            self.name_incomplete = (op > cp)
            return True

        m = RE_NOTE.match(text)
        if m:
            note_num = m.group(1)
            self.note = f"注{note_num}" if note_num else "注"
            self.last_note_x = x_pos
            return True

        # 数字のみで始まる行は文脈に応じて解釈
        m = RE_NOTE_NUM_ONLY.match(text)
        if m:
            num_str = m.group(1)
            # 注の文脈内かつ注の x 位置に近い場合 → 注N
            if self.note and abs(x_pos - self.last_note_x) < 15:
                self.note = f"注{num_str}"
                return True
            # 通則の文脈で、通則の番号付き項目の場合
            if self.item_code == "通則" and abs(x_pos - self.last_item_x) < 15:
                self.note = f"注{num_str}"
                return True

        return False

    def snapshot(self):
        """現在の状態のコピーを返す"""
        return {
            'chapter': self.chapter,
            'part': self.part,
            'section': self.section,
            'subsection': self.subsection,
            'item_code': self.item_code,
            'item_name': self.item_name,
            'note': self.note,
            'name_incomplete': self.name_incomplete,
        }


# ============================================================
# PDF行結合
# ============================================================

def join_pdf_lines(line_segments_list):
    """PDFの行（セグメントリストのリスト）を改行区切りで結合する。

    Args:
        line_segments_list: [[(text, is_underlined), ...], ...] のリスト
            外側リスト = 行、内側リスト = 行内セグメント
    """
    if not line_segments_list:
        return ""
    return '\n'.join(_segments_text(segs) for segs in line_segments_list)


# ============================================================
# 注の番号を検出
# ============================================================

def detect_note_label(text):
    """テキストから注のラベルを検出する。
    明示的に「注N ...」で始まる場合のみ検出。
    数字のみの行は文脈（tracker）に任せる。
    """
    if not text or text in ('（削る）', '（新設）'):
        return ""
    m = RE_NOTE.match(text)
    if m:
        num = m.group(1)
        return f"注{num}" if num else "注"
    return ""


# ============================================================
# メイン処理
# ============================================================

def is_sentence_start(text):
    """行が文の始まりかどうかを推定する。
    注番号、項目番号、区分番号、イロハ項目の開始、
    ただし書きの開始など。
    """
    if not text:
        return False
    if text in ('（削る）', '（新設）'):
        return True
    if RE_NOTE.match(text):
        return True
    if RE_NOTE_NUM_ONLY.match(text):
        return True
    if RE_KUBUN.match(text):
        return True
    if RE_CHAPTER.match(text) or RE_PART.match(text):
        return True
    if RE_SECTION.match(text) or RE_SUBSECTION.match(text):
        return True
    if RE_TSUSOKU.match(text):
        return True
    # ただし書き（法令文で独立した節として扱われる）
    if text.startswith('ただし、') or text.startswith('ただし，'):
        return True
    return False


def text_ends_sentence(text):
    """テキストが文末（。）で終わっているかを判定。

    括弧が閉じていない場合（例:「（...限る。」）は文途中と判定。
    日本語の法令文では「（○○に限る。）」のように括弧内に「。」が
    入ることが多いため、括弧のバランスをチェックする。
    """
    if not text:
        return False
    t = text.rstrip()
    if not (t.endswith('。') or t.endswith('。）') or t.endswith('。」')):
        return False
    # 括弧バランスチェック: 開き括弧が閉じ括弧より多い場合は
    # 「。」が括弧の中にある（＝文末ではない）
    open_count = t.count('（') + t.count('(')
    close_count = t.count('）') + t.count(')')
    if open_count > close_count:
        return False
    return True


def column_complete(text):
    """カラムのテキストが完結しているかを判定。

    空/特殊マーカーの場合、または文末（。）で終わる場合は完結と判定。
    これにより、左右両カラムの完結状態を独立に評価できる。
    """
    if not text or text.strip() == '' or text in ('（削る）', '（新設）'):
        return True
    return text_ends_sentence(text)


def extend_block_context(lines, ul_start, ul_end):
    """アンダーライン領域を前後に拡張して完全な文にする。

    Args:
        lines: ページの全行リスト
        ul_start: アンダーライン開始行のインデックス
        ul_end: アンダーライン終了行のインデックス（含む）

    Returns:
        (start_idx, end_idx): 拡張後の開始・終了インデックス
    """
    start = ul_start
    end = ul_end

    # 前方拡張: 文の始まりまで遡る
    # 最初のアンダーライン行が文の途中から始まっている場合、
    # 前の行まで遡って文頭を見つける
    first_left = lines[ul_start]['left_text']
    first_right = lines[ul_start]['right_text']

    # (削る)(新設) の場合は拡張しない
    if first_left not in ('（削る）', '（新設）', '') or \
       first_right not in ('（削る）', '（新設）', ''):

        # 前方に遡る（最大20行）
        for i in range(ul_start - 1, max(ul_start - 21, -1), -1):
            # 別の変更ブロック（アンダーライン行）は越えない
            if line_has_change(lines[i]):
                start = i + 1
                break
            prev_left = lines[i]['left_text']
            prev_right = lines[i]['right_text']

            # 前の行が文の始まりパターンなら、そこから開始
            if is_sentence_start(prev_left) or is_sentence_start(prev_right):
                start = i
                break
            # 前の行が文末で、左右両方とも完結していれば次の行から開始
            if column_complete(prev_left) and column_complete(prev_right):
                start = i + 1
                break
            start = i

    # 後方拡張: 左右両カラムが文末（。）に達するまで進む
    last_left = lines[ul_end]['left_text']
    last_right = lines[ul_end]['right_text']

    # 左右いずれかのカラムが文途中であれば拡張する
    if not column_complete(last_left) or not column_complete(last_right):
        for i in range(ul_end + 1, min(ul_end + 21, len(lines))):
            cur_left = lines[i]['left_text']
            cur_right = lines[i]['right_text']
            # 次の項目開始 → 手前で止める
            if is_sentence_start(cur_left) or \
               is_sentence_start(cur_right):
                end = i - 1
                break
            end = i
            # 左右両カラムとも完結したら停止
            if column_complete(cur_left) and column_complete(cur_right):
                break
            # 別の変更箇所に達したら手前で止める
            if line_has_change(lines[i]):
                end = i - 1
                break

    return start, end


def main():
    print("PDFを読み込んでいます...", file=sys.stderr)
    doc = fitz.open(PDF_PATH)
    total_pages = doc.page_count
    print(f"総ページ数: {total_pages}", file=sys.stderr)

    tracker = HierarchyTracker()
    all_rows = []
    # 完成した項目名の辞書（ページをまたぐ項目名の補完に使用）
    completed_names = {}

    # 前ページから継続する情報
    carry_block = None
    carry_prev_change_line = None

    for pg_idx in range(SKIP_PAGES_BEFORE, total_pages):
        if (pg_idx + 1) % 100 == 0:
            print(f"  処理中: {pg_idx+1}/{total_pages} ページ...",
                  file=sys.stderr)

        page = doc[pg_idx]
        page_num = pg_idx + 1
        is_toc = (pg_idx < TOC_END_PAGE)

        underlines = get_underlines(page)
        lines = extract_page_lines(page, underlines)

        # 階層コンテキスト更新（全行）＋行ごとのスナップショットを保存
        # 各行処理後の tracker 状態を記録し、ブロック作成時に
        # その行での注情報を正確に取得できるようにする
        line_tracker_states = []
        for line in lines:
            text_for_ctx = line['left_text'] if line['left_text'] else ""
            if text_for_ctx and text_for_ctx not in ('（削る）', '（新設）'):
                tracker.update(text_for_ctx, line['left_x'])
            if not text_for_ctx or text_for_ctx in ('（削る）', '（新設）'):
                rt = line['right_text']
                if rt and rt not in ('（削る）', '（新設）'):
                    tracker.update(rt, line['right_x'])
            line_tracker_states.append(tracker.snapshot())
            # 項目名が完成したら辞書に記録（ページまたぎの補完用）
            if tracker.item_code and not tracker.name_incomplete:
                completed_names[tracker.item_code] = tracker.item_name

        if is_toc:
            continue

        # ===== 変更ブロックの検出と文脈拡張 =====
        # まず、アンダーライン領域を特定
        i = 0
        current_block = carry_block
        prev_change_line = carry_prev_change_line
        carry_block = None
        carry_prev_change_line = None

        while i < len(lines):
            line = lines[i]
            is_change = line_has_change(line)

            if is_change:
                start_new = (current_block is None) or \
                    should_split_block(line, prev_change_line)

                if start_new and current_block is not None:
                    all_rows.append(current_block)
                    current_block = None

                if current_block is None:
                    # アンダーライン領域の開始
                    ul_start = i

                    # アンダーライン領域の終了を探す
                    ul_end = i
                    j = i + 1
                    while j < len(lines):
                        if line_has_change(lines[j]):
                            if should_split_block(lines[j], lines[ul_end]):
                                break
                            ul_end = j
                        else:
                            break
                        j += 1

                    # 文脈拡張
                    ext_start, ext_end = extend_block_context(
                        lines, ul_start, ul_end)

                    # ブロック開始行での tracker 状態を使用
                    # （ページ最終状態ではなく、その行での状態を取得）
                    # コピーして使用（元のスナップショットを変更しない）
                    ctx = dict(line_tracker_states[ul_start])

                    # tracker が検出した注情報を保存（後処理で活用）
                    tracker_note = ctx.get('note', '')

                    # 変更行からコンテキスト情報を取得
                    left_active = ""
                    right_active = ""
                    first_line = lines[ul_start]
                    if first_line['left_text'] and \
                       first_line['left_text'] not in ('（削る）', '（新設）'):
                        left_active = first_line['left_text']
                    if first_line['right_text'] and \
                       first_line['right_text'] not in ('（削る）', '（新設）'):
                        right_active = first_line['right_text']
                    active_text = left_active or right_active

                    # 注フィールドは後処理で最終判定（tracker_note はバックアップ）
                    ctx['tracker_note'] = tracker_note
                    ctx['note'] = ""

                    if active_text:
                        m = RE_KUBUN.match(active_text)
                        if m:
                            ctx['item_code'] = m.group(1)
                            raw_name = m.group(2).strip()
                            raw_name = re.sub(
                                r'\s+[\d０-９,，]+点.*$', '', raw_name)
                            ctx['item_name'] = raw_name.strip()
                            ctx['note'] = ""

                    # 項目名が複数行にまたがる場合（括弧が閉じていない場合）、
                    # 後続行の tracker 状態から完成した名前を取得する
                    # （例: Ｋ０４６－２ の名前が2行目で括弧が閉じる）
                    item_name = ctx.get('item_name', '')
                    op = item_name.count('（') + item_name.count('(')
                    cp = item_name.count('）') + item_name.count(')')
                    if op > cp:
                        for ahead in range(ul_start + 1,
                                           len(line_tracker_states)):
                            s = line_tracker_states[ahead]
                            if s['item_code'] == ctx['item_code'] and \
                               not s.get('name_incomplete'):
                                ctx['item_name'] = s['item_name']
                                break

                    # 拡張された範囲からテキストを収集（セグメント単位のUL情報付き）
                    left_lines = []
                    right_lines = []
                    for k in range(ext_start, ext_end + 1):
                        if k < len(lines):
                            if lines[k]['left_text']:
                                left_lines.append(lines[k]['left_segments'])
                            if lines[k]['right_text']:
                                right_lines.append(lines[k]['right_segments'])

                    current_block = {
                        'page': page_num,
                        'context': ctx,
                        'left_lines': left_lines,
                        'right_lines': right_lines,
                    }

                    prev_change_line = lines[ul_end]
                    i = ext_end + 1
                    continue
                else:
                    # 既存ブロックに追加（carry_block 継続時）
                    if line['left_text']:
                        current_block['left_lines'].append(
                            line['left_segments'])
                    if line['right_text']:
                        current_block['right_lines'].append(
                            line['right_segments'])
                    prev_change_line = line
            else:
                if current_block is not None:
                    # 後方拡張: 左右両カラムが文末に達するまで追加
                    lt = line['left_text']
                    rt = line['right_text']
                    # 左右いずれかのカラムが文途中であれば拡張
                    last_left = _segments_text(
                        current_block['left_lines'][-1]) \
                        if current_block['left_lines'] else ''
                    last_right = _segments_text(
                        current_block['right_lines'][-1]) \
                        if current_block['right_lines'] else ''
                    need_extend = not column_complete(last_left) or \
                        not column_complete(last_right)
                    if need_extend and not is_sentence_start(lt) and \
                       not is_sentence_start(rt):
                        if lt:
                            current_block['left_lines'].append(
                                line['left_segments'])
                        if rt:
                            current_block['right_lines'].append(
                                line['right_segments'])
                        # 左右両方が完結したら確定
                        new_last_left = _segments_text(
                            current_block['left_lines'][-1]) \
                            if current_block['left_lines'] else ''
                        new_last_right = _segments_text(
                            current_block['right_lines'][-1]) \
                            if current_block['right_lines'] else ''
                        if column_complete(new_last_left) and \
                           column_complete(new_last_right):
                            all_rows.append(current_block)
                            current_block = None
                            prev_change_line = None
                        i += 1
                        continue
                    else:
                        all_rows.append(current_block)
                        current_block = None
                        prev_change_line = None

            i += 1

        # ページ末尾で未完了のブロック
        if current_block is not None:
            carry_block = current_block
            carry_prev_change_line = prev_change_line

    # 最後のブロック
    if carry_block is not None:
        all_rows.append(carry_block)

    doc.close()

    # ============================================================
    # 項目名の後処理（ページまたぎの補完）
    # ============================================================
    # ページ境界をまたいで項目名が続く場合（例: Ｋ０７６－３ の名前が
    # 717ページ末尾で始まり718ページ先頭で閉じる）、ブロック作成時には
    # 不完全な名前しか取得できない。全ページ処理後に完成した名前で補完する。
    for row in all_rows:
        ctx = row['context']
        item_name = ctx.get('item_name', '')
        item_code = ctx.get('item_code', '')
        if item_name and item_code:
            op = item_name.count('（') + item_name.count('(')
            cp = item_name.count('）') + item_name.count(')')
            if op > cp and item_code in completed_names:
                ctx['item_name'] = completed_names[item_code]

    # ============================================================
    # 注フィールドの後処理
    # ============================================================
    # 出力順序に基づいて注を判定する。
    # - テキストが「注N ...」で始まる → 注N
    # - テキストが裸の数字で始まり、同じ項目コード内の直前ブロックが
    #   注を持っていた場合 → 注N（省略された注番号）
    # - それ以外 → 空
    last_note = ""
    last_item_code = ""
    for row in all_rows:
        ctx = row['context']
        item_code = ctx.get('item_code', '')

        # 項目コードが変わったら注のコンテキストをリセット
        if item_code != last_item_code:
            last_note = ""
            last_item_code = item_code

        # 変更テキストの先頭行を取得（セグメントリストからテキスト結合）
        active_texts = []
        for lines_list in [row['left_lines'], row['right_lines']]:
            if lines_list:
                first_text = _segments_text(lines_list[0])
                if first_text and first_text not in ('（削る）', '（新設）'):
                    active_texts.append(first_text)

        # 明示的な注（「注N ...」パターン）
        note = ""
        for txt in active_texts:
            note = detect_note_label(txt)
            if note:
                break

        # 省略された注番号（前のブロックまたは tracker が注を検出済みの場合、
        # 裸の数字も注番号とみなす）
        # tracker_note: ブロック作成時の tracker の注状態。PDF上で注の
        # 文脈内にある行では、tracker が正しく注番号を追跡している。
        # last_note: 同じ項目コード内の直前ブロックで検出された注。
        if not note:
            tracker_note = ctx.get('tracker_note', '')
            if last_note or tracker_note:
                for txt in active_texts:
                    m = RE_NOTE_NUM_ONLY.match(txt)
                    if m:
                        note = f"注{m.group(1)}"
                        break

        # tracker_note フォールバック:
        # テキストが「注N」にも裸の数字にも該当しない場合でも、
        # tracker が注を検出済みで、テキストが区分番号や見出しでなければ
        # tracker_note を使用する（例: イ 時間外対応体制加算 等のサブ項目）
        if not note:
            tracker_note = ctx.get('tracker_note', '')
            if tracker_note:
                is_kubun_or_heading = False
                for txt in active_texts:
                    if RE_KUBUN.match(txt) or RE_CHAPTER.match(txt) or \
                       RE_PART.match(txt) or RE_SECTION.match(txt) or \
                       RE_SUBSECTION.match(txt) or RE_TSUSOKU.match(txt):
                        is_kubun_or_heading = True
                        break
                if not is_kubun_or_heading:
                    note = tracker_note

        ctx['note'] = note
        if note:
            last_note = note

    # ============================================================
    # （新設）（削る）マーカーのクリーンアップ
    # ============================================================
    # マーカーが含まれる側はマーカーのみを残し、混入テキストを除去する
    for row in all_rows:
        for side_key in ['left_lines', 'right_lines']:
            lines_list = row[side_key]
            has_marker = any(
                _segments_text(segs) in ('（新設）', '（削る）')
                for segs in lines_list)
            if has_marker:
                markers = [segs for segs in lines_list
                           if _segments_text(segs) in ('（新設）', '（削る）')]
                row[side_key] = markers

    # ============================================================
    # XLSX 出力（xlsxwriter でアンダーライン付きリッチテキスト）
    # ============================================================
    # openpyxl の CellRichText は Excel で「修復」が必要になり
    # リッチテキスト書式が壊れるため、xlsxwriter を使用する。
    print(f"XLSXを出力しています: {OUTPUT_XLSX}", file=sys.stderr)
    print(f"抽出された変更ブロック数: {len(all_rows)}", file=sys.stderr)

    wb = xlsxwriter.Workbook(OUTPUT_XLSX)
    ws = wb.add_worksheet('新旧対照表')

    # 書式定義
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

    # ヘッダー行
    headers = ['章', '部', '節', '款', '通則/項目コード', '注',
               '改正後', '改正前', 'ページ']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    # 列幅設定
    ws.set_column(0, 0, 12)   # 章
    ws.set_column(1, 1, 12)   # 部
    ws.set_column(2, 2, 12)   # 節
    ws.set_column(3, 3, 12)   # 款
    ws.set_column(4, 4, 25)   # 通則/項目コード
    ws.set_column(5, 5, 8)    # 注
    ws.set_column(6, 6, 55)   # 改正後
    ws.set_column(7, 7, 55)   # 改正前
    ws.set_column(8, 8, 6)    # ページ

    def write_rich_cell(ws, row_idx, col_idx, line_segments_list):
        """セグメントリストから xlsxwriter のリッチテキストセルを書き込む。

        line_segments_list: [[(text, is_underlined), ...], ...]
            外側リスト=行、内側リスト=行内セグメント
        """
        if not line_segments_list:
            ws.write_string(row_idx, col_idx, '', normal_fmt)
            return

        # リッチテキストの引数リストを構築
        # xlsxwriter.write_rich_string(row, col, *args)
        # args は [format, string, format, string, ...] の形式
        parts = []
        for i, segs in enumerate(line_segments_list):
            if i > 0:
                parts.extend([normal_fmt, '\n'])
            for text, is_ul in segs:
                fmt = ul_fmt if is_ul else normal_fmt
                parts.extend([fmt, text])

        if not parts:
            ws.write_string(row_idx, col_idx, '', normal_fmt)
            return

        # すべて同一書式なら write_string、混在なら write_rich_string
        all_same = all(parts[i] is parts[0] for i in range(0, len(parts), 2))
        if all_same:
            text = ''.join(parts[i] for i in range(1, len(parts), 2))
            ws.write_string(row_idx, col_idx, text, parts[0])
        else:
            ws.write_rich_string(row_idx, col_idx, *parts, normal_fmt)

    for idx, row in enumerate(all_rows):
        r = idx + 1  # 行番号（0はヘッダー）
        ctx = row['context']
        item_display = ctx['item_code']
        if ctx.get('item_name'):
            item_display = f"{ctx['item_code']} {ctx['item_name']}"

        ws.write_string(r, 0, ctx.get('chapter', ''), normal_fmt)
        ws.write_string(r, 1, ctx.get('part', ''), normal_fmt)
        ws.write_string(r, 2, ctx.get('section', ''), normal_fmt)
        ws.write_string(r, 3, ctx.get('subsection', ''), normal_fmt)
        ws.write_string(r, 4, item_display, normal_fmt)
        ws.write_string(r, 5, ctx.get('note', ''), normal_fmt)
        write_rich_cell(ws, r, 6, row['left_lines'])
        write_rich_cell(ws, r, 7, row['right_lines'])
        ws.write_number(r, 8, row['page'], normal_fmt)

    wb.close()
    print(f"完了！ {len(all_rows)} 件の変更箇所を出力しました。",
          file=sys.stderr)


if __name__ == '__main__':
    main()
