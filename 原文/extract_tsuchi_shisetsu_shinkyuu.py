#!/usr/bin/env python3
"""
施設基準 通知PDF 新旧対照表 抽出スクリプト

R8通知PDFとR6通知PDFを比較し、difflib で文字レベルの差分を検出して
XLSX 形式の新旧対照表を生成する。差分部分にはアンダーライン書式を適用する。

基本診療料・特掲診療料の両方を処理する。
"""

import fitz  # PyMuPDF
import re
import sys
import os
import difflib
import xlsxwriter

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# ============================================================
# 定数
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIGS = [
    {
        'label': '基本診療料',
        'r8_pdf': os.path.join(SCRIPT_DIR,
            "基本診療料の施設基準等及びその届出に関する手続きの取扱いについて（令和８年３月５日保医発0305第７号）.pdf"),
        'r6_pdf': os.path.join(SCRIPT_DIR,
            "基本診療料の施設基準等及びその届出に関する手続きの取扱いについて（通知）令和６年３月５日 保医発0305第５号.pdf"),
        'output': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(基本・通知)_新旧対照表.xlsx"),
    },
    {
        'label': '特掲診療料',
        'r8_pdf': os.path.join(SCRIPT_DIR,
            "特掲診療料の施設基準等及びその届出に関する手続きの取扱いについて（令和８年３月５日保医発0305第８号）.pdf"),
        'r6_pdf': os.path.join(SCRIPT_DIR,
            "特掲診療料の施設基準等及びその届出に関する手続きの取扱いについて（通知）令和６年３月５日 保医発0305第６号.pdf"),
        'output': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(特掲・通知)_新旧対照表.xlsx"),
    },
]

MIN_FONT_SIZE = 5.0
RUBY_MAX_SIZE = 6.0
Y_GROUP_TOLERANCE = 3.0

# ============================================================
# 正規表現パターン
# ============================================================
# ページ番号パターン（- 1 - 形式）
RE_PAGE_NUM = re.compile(
    r'^[\-\u2015\u2014\u2013－]?\s*\d{1,4}\s*[\-\u2015\u2014\u2013－]?$')

# 別添パターン
RE_BETTEN = re.compile(r'^別添\s*([１２３４５６７８９0-9]+(?:の[１２３４５６７８９0-9]+)?)\s*$')

# 第N / 第NのM パターン（セクション見出し）
# ※「第57号」等の告示番号を誤検出しないよう、数字は3桁以内に制限
RE_SECTION = re.compile(
    r'^第\s*([１２３４５６７８９０0-9]{1,3}(?:\s*の\s*[１２３４５６７８９０0-9]{1,3})*)\s+(.+)')

# 番号付き項目（1, 2, 3, ... や１, ２, ３, ...）
RE_NUMBERED_ITEM = re.compile(r'^([１２３４５６７８９０0-9]+)\s')
# 番号直後が法令参照・日付の継続の場合は項番ではない
# 例: "23 年政令..." "69 号）..." "18 対１..." "31 日まで..."
RE_FALSE_ITEM_AFTER = re.compile(
    r'^(号|条|項|対[１２３４５６７８９０0-9]|月[１２３４５６７８９０0-9末]'
    r'|年[^間度以]|日[^本常間])')

# 見出しのみの行
RE_HEADING_ONLY = re.compile(
    r'^(別添\s*[１２３４５６７８９0-9]+(?:の[１２３４５６７８９0-9]+)?|'
    r'第\s*[１２３４５６７８９０0-9]+(?:\s*の\s*[１２３４５６７８９０0-9]+)*\s+.*|'
    r'削除\s*)$')

# サブアイテム検出パターン
RE_SUBITEM_PAREN = re.compile(r'^([（(][１２３４５６７８９０0-9]+[）)])[\s\u3000]')
RE_SUBITEM_KANA = re.compile(
    r'^([アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホ'
    r'マミムメモヤユヨラリルレロワヲン])[\s\u3000]')
RE_SUBITEM_ALPHA = re.compile(r'^([ａｂｃｄｅｆｇｈｉｊｋｌｍ])[\s\u3000]')


# ============================================================
# PDFテキスト抽出
# ============================================================

def extract_page_lines(page):
    """横書き単一カラムPDFからテキスト行を抽出する。

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
                if span['size'] < RUBY_MAX_SIZE and re.match(
                        r'^[ぁ-ん]+$', span['text'].strip()):
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
        text = text.replace('\u2015', '－').replace('\u2014', '－').replace(
            '\u2013', '－')
        x0 = sorted_spans[0]['x0']

        if RE_PAGE_NUM.match(text):
            continue
        if not text:
            continue

        lines.append((text, x0))

    return lines


# ============================================================
# 階層追跡（施設基準 通知PDF用）
# ============================================================

class HierarchyTracker:
    """別添/第N/番号 の階層を追跡する"""

    def __init__(self):
        self.betten = ""         # 別添N
        self.section = ""        # 第NのM
        self.section_name = ""   # セクション名
        self.item_num = ""       # 番号付き項目 (1, 2, ...)
        self.last_item_x = 0
        self.pre_betten = True   # 別添の前の部分（第1~第4等）

    def update(self, text, x_pos):
        if not text:
            return False

        # 別添パターン
        m = RE_BETTEN.match(text)
        if m:
            self.betten = f"別添{m.group(1)}"
            self.section = ""
            self.section_name = ""
            self.item_num = ""
            self.pre_betten = False
            return True

        # 第Nパターン（セクション見出し）
        # 「第57号」「第１項」等の法令参照を誤検出しないよう、
        # タイトル部分が「号」「条」「項」で始まるものを除外
        m = RE_SECTION.match(text)
        if m and x_pos < 100:
            sec_title = m.group(2).strip()
            if not re.match(r'^[号条項）)]', sec_title):
                sec_num = m.group(1).replace(' ', '')
                self.section = f"第{sec_num}"
                self.section_name = sec_title
                self.item_num = ""
                return True

        # 番号付き項目（法令参照・日付の継続は除外）
        m = RE_NUMBERED_ITEM.match(text)
        if m and x_pos < 100:
            after = text[m.end():].strip()
            if not RE_FALSE_ITEM_AFTER.match(after):
                self.item_num = m.group(1)
                self.last_item_x = x_pos
                return True

        return False

    def snapshot(self):
        return {
            'betten': self.betten,
            'section': self.section,
            'section_name': self.section_name,
            'item_num': self.item_num,
            'pre_betten': self.pre_betten,
        }


# ============================================================
# ブロック分割
# ============================================================

def is_block_boundary(text, x0):
    """テキストが新しいブロック境界かを判定"""
    if RE_BETTEN.match(text):
        return True
    m = RE_SECTION.match(text)
    if m and x0 < 100:
        title = m.group(2).strip()
        if not re.match(r'^[号条項）)]', title):
            return True
    m = RE_NUMBERED_ITEM.match(text)
    if m and x0 < 100:
        after = text[m.end():].strip()
        if not RE_FALSE_ITEM_AFTER.match(after):
            return True
    return False


def is_heading_only_block(text):
    """ブロックが見出しのみかを判定"""
    lines = text.strip().split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if not RE_HEADING_ONLY.match(line):
            return False
    return True


def split_blocks_at_subitems(blocks):
    """ブロックをサブアイテム境界で分割する"""
    result = []
    for block in blocks:
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
    """単一ブロックをサブアイテム境界で分割する"""
    lines = block['text'].split('\n')

    markers = []
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
    note_base = block['item_num']

    first_line = markers[0][0]
    if first_line > 0:
        preamble = '\n'.join(lines[:first_line])
        if preamble.strip():
            result.append({**block, 'text': preamble})

    for idx, (start, mtype, marker) in enumerate(markers):
        end = markers[idx + 1][0] if idx + 1 < len(markers) else len(lines)
        sub_text = '\n'.join(lines[start:end])
        sub_note = f"{note_base}\u3000{marker}" if note_base else marker
        result.append({**block, 'text': sub_text, 'item_num': sub_note,
                       '_from_split': True})

    return result


def extract_blocks_from_pdf(pdf_path):
    """通知PDFからテキストブロックを抽出する。"""
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
                'betten': current_ctx['betten'],
                'section': current_ctx['section'],
                'section_name': current_ctx['section_name'],
                'item_num': current_ctx['item_num'],
                'text': block_text,
                'page': current_page,
            })

    for pg_idx in range(total_pages):
        if (pg_idx + 1) % 100 == 0:
            print(f"    処理中: {pg_idx+1}/{total_pages} ページ...",
                  file=sys.stderr)

        page = doc[pg_idx]
        page_num = pg_idx + 1

        lines = extract_page_lines(page)

        for text, x0 in lines:
            boundary = is_block_boundary(text, x0)

            tracker.update(text, x0)

            ctx = tracker.snapshot()

            if boundary or current_ctx is None:
                save_current_block()
                current_lines = [text]
                current_ctx = ctx
                current_page = page_num
            else:
                current_lines.append(text)

    save_current_block()
    doc.close()
    print(f"  抽出ブロック数: {len(blocks)}", file=sys.stderr)
    return blocks


# ============================================================
# マッチングキー
# ============================================================

def make_block_key(block):
    """ブロックのマッチングキー"""
    return (block['betten'], block['section'], block['item_num'])


def make_block_key_short(block):
    """フォールバック用短キー"""
    return (block['section'], block['item_num'])


# ============================================================
# 差分検出
# ============================================================

def normalize_text_for_compare(text):
    """比較用にテキストを正規化する"""
    t = re.sub(r'\s+', ' ', text).strip()
    t = t.replace('\u3000', ' ')
    t = re.sub(r' +', ' ', t)
    return t


def compute_diff_segments(r8_text, r6_text):
    """2つのテキストの文字レベル差分をセグメントリストに変換する"""
    r8_norm = normalize_text_for_compare(r8_text)
    r6_norm = normalize_text_for_compare(r6_text)

    if r8_norm == r6_norm:
        return None, None

    sm = difflib.SequenceMatcher(None, r6_norm, r8_norm, autojunk=False)
    ratio = sm.ratio()

    if ratio < 0.3:
        return [(r8_text, True)], [(r6_text, True)]

    r8_segments = []
    r6_segments = []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        r6_part = r6_norm[i1:i2]
        r8_part = r8_norm[j1:j2]
        if op == 'equal':
            r8_segments.append((r8_part, False))
            r6_segments.append((r6_part, False))
        elif op == 'replace':
            if not r6_part.strip() and not r8_part.strip():
                r6_segments.append((r6_part, False))
                r8_segments.append((r8_part, False))
            else:
                r6_segments.append((r6_part, True))
                r8_segments.append((r8_part, True))
        elif op == 'insert':
            if not r8_part.strip():
                r8_segments.append((r8_part, False))
            else:
                r8_segments.append((r8_part, True))
        elif op == 'delete':
            if not r6_part.strip():
                r6_segments.append((r6_part, False))
            else:
                r6_segments.append((r6_part, True))

    # 変更セグメントが1つもなければ同一テキスト扱い
    # （PDFの改行位置の違いによるスペース差分のみの場合）
    r8_has_change = any(is_changed for _, is_changed in r8_segments)
    r6_has_change = any(is_changed for _, is_changed in r6_segments)
    if not r8_has_change and not r6_has_change:
        return None, None

    return r8_segments, r6_segments


def text_similarity(text1, text2):
    """2つのテキストの類似度を返す"""
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
    """セグメントリストからリッチテキストセルを書き込む"""
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

def process_pair(config):
    """1組のPDFペアを処理して新旧対照表を生成する"""
    label = config['label']
    r8_pdf = config['r8_pdf']
    r6_pdf = config['r6_pdf']
    output_xlsx = config['output']

    print(f"\n{'=' * 60}", file=sys.stderr)
    print(f"施設基準（{label}・通知）新旧対照表 抽出開始", file=sys.stderr)
    print(f"{'=' * 60}", file=sys.stderr)

    # Phase 1: ブロック抽出
    print("\n[Phase 1] PDFからブロック抽出", file=sys.stderr)
    r8_blocks_raw = extract_blocks_from_pdf(r8_pdf)
    r6_blocks_raw = extract_blocks_from_pdf(r6_pdf)

    # 見出しのみのブロックを除外、前文（別添・セクション未設定）も除外
    r8_blocks = [b for b in r8_blocks_raw
                 if not is_heading_only_block(b['text'])
                 and (b['betten'] or b['section'])]
    r6_blocks = [b for b in r6_blocks_raw
                 if not is_heading_only_block(b['text'])
                 and (b['betten'] or b['section'])]
    print(f"  R8: {len(r8_blocks_raw)} → {len(r8_blocks)} ブロック（見出し除外後）",
          file=sys.stderr)
    print(f"  R6: {len(r6_blocks_raw)} → {len(r6_blocks)} ブロック（見出し除外後）",
          file=sys.stderr)

    # サブアイテム分割
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

    r6_by_key = {}
    for b in r6_blocks:
        key = make_block_key(b)
        r6_by_key.setdefault(key, []).append(b)

    r6_by_short_key = {}
    for b in r6_blocks:
        key = make_block_key_short(b)
        r6_by_short_key.setdefault(key, []).append(b)

    matched_r6_ids = set()
    r6_id_to_idx = {id(b): i for i, b in enumerate(r6_blocks)}
    r6_idx_to_output_pos = {}
    output_rows = []

    for r8b in r8_blocks:
        key = make_block_key(r8b)
        r6_match = None

        if key in r6_by_key:
            candidates = [c for c in r6_by_key[key]
                          if id(c) not in matched_r6_ids]
            if len(candidates) == 1:
                r6_match = candidates[0]
            elif len(candidates) > 1:
                best = max(candidates,
                           key=lambda c: text_similarity(r8b['text'],
                                                         c['text']))
                r6_match = best

        skip_fallback = False
        if r6_match and r8b.get('_from_split') and r6_match.get(
                '_from_split'):
            if text_similarity(r8b['text'], r6_match['text']) < 0.3:
                r6_match = None
                skip_fallback = True

        if r6_match is None and not skip_fallback:
            short_key = make_block_key_short(r8b)
            if short_key in r6_by_short_key:
                candidates = [c for c in r6_by_short_key[short_key]
                              if id(c) not in matched_r6_ids]
                if candidates:
                    best = max(candidates,
                               key=lambda c: text_similarity(r8b['text'],
                                                             c['text']))
                    if text_similarity(r8b['text'], best['text']) > 0.2:
                        r6_match = best

        if r6_match:
            matched_r6_ids.add(id(r6_match))
            r6_idx = r6_id_to_idx[id(r6_match)]

        r8_text = r8b['text']
        r6_text = r6_match['text'] if r6_match else ""

        if r6_match is None:
            r8_segments = [(r8_text, True)]
            r6_segments = [("（新設）", False)]
        else:
            r8_segments, r6_segments = compute_diff_segments(r8_text, r6_text)

        if r8_segments is None:
            if r6_match:
                r6_idx_to_output_pos[r6_idx] = (
                    len(output_rows) - 1 if output_rows else -1)
            continue

        if r6_match:
            r6_idx_to_output_pos[r6_idx] = len(output_rows)

        output_rows.append({
            'betten': r8b['betten'],
            'section': r8b['section'],
            'section_name': r8b.get('section_name', ''),
            'item_num': r8b['item_num'],
            'r8_segments': r8_segments,
            'r6_segments': r6_segments,
            'page': r8b['page'],
            '_from_split': r8b.get('_from_split', False),
        })

    # 削除項目の挿入
    deleted_items = []
    for r6_idx, r6b in enumerate(r6_blocks):
        if id(r6b) not in matched_r6_ids:
            insert_pos = 0
            for prev_idx in range(r6_idx - 1, -1, -1):
                if prev_idx in r6_idx_to_output_pos:
                    insert_pos = r6_idx_to_output_pos[prev_idx] + 1
                    break
            r6_text = r6b['text']
            deleted_items.append((insert_pos, r6_idx, {
                'betten': r6b['betten'],
                'section': r6b['section'],
                'section_name': r6b.get('section_name', ''),
                'item_num': r6b['item_num'],
                'r8_segments': [("（削除）", False)],
                'r6_segments': [(r6_text, True)],
                'page': 0,
            }))

    deleted_items.sort(key=lambda x: (x[0], x[1]), reverse=True)
    for insert_pos, _r6_idx, row in deleted_items:
        if insert_pos >= len(output_rows):
            output_rows.append(row)
        else:
            output_rows.insert(insert_pos, row)

    print(f"  差分ブロック数: {len(output_rows)}", file=sys.stderr)

    # Phase 3: XLSX出力
    print(f"\n[Phase 3] XLSX出力: {output_xlsx}", file=sys.stderr)

    os.makedirs(os.path.dirname(output_xlsx), exist_ok=True)
    wb = xlsxwriter.Workbook(output_xlsx)
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
        'bold': True, 'underline': True,
        'text_wrap': True, 'valign': 'top',
    })

    headers = ['別添', '項目', '項番', '改正後（R8）', '改正前（R6）', 'ページ']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    ws.set_column(0, 0, 12)   # 別添
    ws.set_column(1, 1, 30)   # 項目（第NのM + 名称）
    ws.set_column(2, 2, 10)   # 項番
    ws.set_column(3, 3, 60)   # 改正後
    ws.set_column(4, 4, 60)   # 改正前
    ws.set_column(5, 5, 6)    # ページ

    for idx, row in enumerate(output_rows):
        r = idx + 1
        section_display = row['section']
        if row.get('section_name'):
            section_display = f"{row['section']} {row['section_name']}"

        ws.write_string(r, 0, row.get('betten', ''), normal_fmt)
        ws.write_string(r, 1, section_display, normal_fmt)
        ws.write_string(r, 2, row.get('item_num', ''), normal_fmt)
        write_rich_cell(ws, r, 3, row['r8_segments'], normal_fmt, ul_fmt)
        write_rich_cell(ws, r, 4, row['r6_segments'], normal_fmt, ul_fmt)
        if row['page'] > 0:
            ws.write_number(r, 5, row['page'], normal_fmt)
        else:
            ws.write_string(r, 5, '', normal_fmt)

    wb.close()
    print(f"\n完了！ {len(output_rows)} 件の差分を出力しました。", file=sys.stderr)


def main():
    for config in CONFIGS:
        process_pair(config)


if __name__ == '__main__':
    main()
