#!/usr/bin/env python3
"""
DPC通知PDF新旧対照表 抽出スクリプト

R8 DPC通知PDF（令和８年保医発0318第４号）とR6 DPC通知PDF（令和６年保医発0321第６号）を比較し、
difflib で文字レベルの差分を検出して XLSX 形式の新旧対照表を生成する。
差分部分にはアンダーライン書式を適用する。
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
R8_PDF_PATH = os.path.join(SCRIPT_DIR,
    "厚生労働大臣が指定する病院の病棟における療養に要する費用の額の算定方法の一部改正等に伴う実施上の留意事項について（通知）（令和８年３月18日保医発0318第４号）.pdf")
R6_PDF_PATH = os.path.join(SCRIPT_DIR,
    "厚生労働大臣が指定する病院の病棟における療養に要する費用の額の算定方法の一部改正等に伴う実施上の留意事項について（通知）令和６年３月21日.pdf")
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, "output",
    "R8年度DPC通知_新旧対照表.xlsx")

MIN_FONT_SIZE = 5.0
RUBY_MAX_SIZE = 6.0
Y_GROUP_TOLERANCE = 3.0

R8_TOC_PAGES = 0
R6_TOC_PAGES = 0

# x座標の閾値（DPC通知PDF固有）
# 56.6: 左マージン（第N、本文継続）
# 67.2: 項番（１, ２, ３）
# 77.7: サブ項番 (１),(２)
# 98.7: カナ（ア, イ, ウ）
ITEM_NUM_MAX_X = 72.0    # 項番の最大x座標
SUB_ITEM_MAX_X = 82.0    # サブ項番の最大x座標
SECTION_MAX_X = 62.0     # セクション（第N）の最大x座標

# ============================================================
# 正規表現パターン（DPC通知用）
# ============================================================
RE_DPC_SECTION = re.compile(r'^第([１２３４５６７８９０\d]+)\s+(.*)')
RE_DPC_ITEM_NUM = re.compile(r'^([１２３４５６７８９０\d]+)\s')
RE_DPC_SUB_ITEM = re.compile(r'^[（(]([１２３４５６７８９０\d]+)[）)]\s')
RE_KANA_ITEM = re.compile(r'^([アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン])\s')
RE_CIRCLED_NUM = re.compile(r'^([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])\s')
RE_PAGE_NUM = re.compile(r'^[\-\u2015\u2014\u2013－]?\s*\d{1,4}\s*[\-\u2015\u2014\u2013－]?$')
RE_HEADING_ONLY = re.compile(r'^(第[１２３４５６７８９０\d]+\s+.+|記)$')

# サブアイテム検出パターン（split_blocks_at_subitems用）
RE_SUBITEM_PAREN = re.compile(r'^([（(][０-９0-9]+[）)])[\s\u3000]')
RE_SUBITEM_KANA = re.compile(
    r'^([アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン])[\s\u3000]')
RE_SUBITEM_ALPHA = re.compile(r'^([ａｂｃｄｅｆｇｈｉｊｋｌｍ])[\s\u3000]')


# ============================================================
# PDFテキスト抽出
# ============================================================

def extract_page_lines_single_column(page):
    """単一カラムの通知PDFからテキスト行を抽出する。

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
        text = text.replace('\u2015', '－').replace('\u2014', '－').replace('\u2013', '－')
        x0 = sorted_spans[0]['x0']

        if RE_PAGE_NUM.match(text):
            continue
        if not text:
            continue

        lines.append((text, x0))

    return lines


# ============================================================
# 階層追跡（DPC通知用）
# ============================================================

class DpcHierarchyTracker:
    """第N/項番/サブ項番/カナ の階層を追跡する（DPC通知用）"""

    def __init__(self):
        self.section = ""       # 第１, 第２, 第３, 第４
        self.item_num = ""      # １, ２, ３
        self.sub_item = ""      # (１), (２)
        self.sub_sub = ""       # ア, イ, ①, ②
        self.in_preamble = True  # 前文フラグ（第１が出現するまで）

    def update(self, text, x_pos):
        if not text:
            return False

        m = RE_DPC_SECTION.match(text)
        if m and x_pos < SECTION_MAX_X:
            self.section = f"第{m.group(1)} {m.group(2)}".strip()
            self.item_num = ""
            self.sub_item = ""
            self.sub_sub = ""
            self.in_preamble = False
            return True

        if self.in_preamble:
            return False

        m = RE_DPC_ITEM_NUM.match(text)
        if m and x_pos < ITEM_NUM_MAX_X:
            self.item_num = m.group(1)
            self.sub_item = ""
            self.sub_sub = ""
            return True

        m = RE_DPC_SUB_ITEM.match(text)
        if m and x_pos < SUB_ITEM_MAX_X:
            self.sub_item = f"({m.group(1)})"
            self.sub_sub = ""
            return True

        m = RE_KANA_ITEM.match(text)
        if m and x_pos > SUB_ITEM_MAX_X:
            self.sub_sub = m.group(1)
            return True

        m = RE_CIRCLED_NUM.match(text)
        if m:
            self.sub_sub = m.group(1)
            return True

        return False

    def snapshot(self):
        return {
            'section': self.section,
            'item_num': self.item_num,
            'sub_item': self.sub_item,
            'sub_sub': self.sub_sub,
        }


# ============================================================
# ブロック分割
# ============================================================

def is_block_boundary(text, x0, tracker):
    """テキストが新しいブロック境界かを判定"""
    if RE_DPC_SECTION.match(text) and x0 < SECTION_MAX_X:
        return True
    if RE_DPC_ITEM_NUM.match(text) and x0 < ITEM_NUM_MAX_X:
        return True
    if RE_DPC_SUB_ITEM.match(text) and x0 < SUB_ITEM_MAX_X:
        return True
    return False


def is_heading_only_block(text):
    """ブロックが見出しのみ（構造情報のみ）かを判定。"""
    lines = text.strip().split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if not RE_HEADING_ONLY.match(line):
            return False
    return True


def split_blocks_at_subitems(blocks):
    """ブロックをサブアイテム（ア、イ、ウ等）の境界で分割する。

    同一タイプのマーカーが2つ以上ある場合のみ分割する。
    """
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
    """単一ブロックをサブアイテム境界で分割する。"""
    lines = block['text'].split('\n')

    markers = []
    for i, line in enumerate(lines):
        stripped = line.strip()
        m = RE_SUBITEM_KANA.match(stripped)
        if m:
            markers.append((i, 'kana', m.group(1)))
            continue
        m = RE_SUBITEM_PAREN.match(stripped)
        if m:
            markers.append((i, 'paren', m.group(1)))
            continue
        m = RE_CIRCLED_NUM.match(stripped)
        if m:
            markers.append((i, 'circled', m.group(1)))
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
    sub_sub_base = block.get('sub_sub', '')

    first_line = markers[0][0]
    if first_line > 0:
        preamble = '\n'.join(lines[:first_line])
        if preamble.strip():
            result.append({**block, 'text': preamble})

    for idx, (start, mtype, marker) in enumerate(markers):
        end = markers[idx + 1][0] if idx + 1 < len(markers) else len(lines)
        sub_text = '\n'.join(lines[start:end])
        sub_sub = f"{sub_sub_base}\u3000{marker}" if sub_sub_base else marker
        result.append({**block, 'text': sub_text, 'sub_sub': sub_sub,
                       '_from_split': True})

    return result


def extract_blocks_from_pdf(pdf_path, toc_pages):
    """DPC通知PDFからテキストブロックを抽出する。

    Returns:
        [{section, item_num, sub_item, sub_sub, text, page}, ...]
    """
    print(f"  PDFを読み込み中: {os.path.basename(pdf_path)}", file=sys.stderr)
    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    print(f"  総ページ数: {total_pages}", file=sys.stderr)

    tracker = DpcHierarchyTracker()
    blocks = []
    current_lines = []
    current_ctx = None
    current_page = 0

    def save_current_block():
        nonlocal current_lines, current_ctx, current_page
        if current_lines and current_ctx:
            block_text = '\n'.join(current_lines)
            blocks.append({
                'section': current_ctx['section'],
                'item_num': current_ctx['item_num'],
                'sub_item': current_ctx['sub_item'],
                'sub_sub': current_ctx['sub_sub'],
                'text': block_text,
                'page': current_page,
            })

    for pg_idx in range(total_pages):
        page = doc[pg_idx]
        page_num = pg_idx + 1
        is_toc = (pg_idx < toc_pages)

        lines = extract_page_lines_single_column(page)

        for text, x0 in lines:
            boundary = is_block_boundary(text, x0, tracker)
            tracker.update(text, x0)

            if is_toc or tracker.in_preamble:
                continue

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


def make_block_key(block):
    """ブロックのマッチングキーを生成。"""
    return (block['section'], block['item_num'], block['sub_item'],
            block.get('sub_sub', ''))


def make_block_key_short(block):
    """フォールバック用の短いキー"""
    return (block['item_num'], block['sub_item'])


# ============================================================
# 差分検出
# ============================================================

def normalize_text_for_compare(text):
    """比較用にテキストを正規化する。"""
    t = re.sub(r'\s+', ' ', text).strip()
    t = t.replace('\u3000', ' ')
    t = re.sub(r' +', ' ', t)
    return t


def compute_diff_segments(r8_text, r6_text):
    """2つのテキストの文字レベル差分をセグメントリストに変換する。

    Returns:
        (r8_segments, r6_segments): 各セグメントは [(text, is_changed), ...]
        テキストが同一の場合は (None, None) を返す。
    """
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
    print("DPC通知PDF新旧対照表 抽出開始", file=sys.stderr)
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
            candidates = [c for c in r6_by_key[key] if id(c) not in matched_r6_ids]
            if len(candidates) == 1:
                r6_match = candidates[0]
            elif len(candidates) > 1:
                best = max(candidates,
                           key=lambda c: text_similarity(r8b['text'], c['text']))
                r6_match = best

        # _from_splitブロック同士のマッチで類似度が低い場合は拒否
        skip_fallback = False
        if r6_match and r8b.get('_from_split') and r6_match.get('_from_split'):
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
                               key=lambda c: text_similarity(r8b['text'], c['text']))
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
                r6_idx_to_output_pos[r6_idx] = len(output_rows) - 1 if output_rows else -1
            continue

        if r6_match:
            r6_idx_to_output_pos[r6_idx] = len(output_rows)

        output_rows.append({
            'section': r8b['section'],
            'item_num': r8b['item_num'],
            'sub_item': r8b.get('sub_item', ''),
            'sub_sub': r8b.get('sub_sub', ''),
            'r8_segments': r8_segments,
            'r6_segments': r6_segments,
            'page': r8b['page'],
        })

    # R6のみ（削除）のブロック
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
                'section': r6b['section'],
                'item_num': r6b['item_num'],
                'sub_item': r6b.get('sub_item', ''),
                'sub_sub': r6b.get('sub_sub', ''),
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

    headers = ['セクション', '項番', 'サブ項番', '改正後（R8）', '改正前（R6）', 'ページ']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    ws.set_column(0, 0, 15)
    ws.set_column(1, 1, 8)
    ws.set_column(2, 2, 12)
    ws.set_column(3, 3, 55)
    ws.set_column(4, 4, 55)
    ws.set_column(5, 5, 6)

    for idx, row in enumerate(output_rows):
        r = idx + 1
        sub_display = row.get('sub_sub', '')
        if row.get('sub_item'):
            sub_display = row['sub_item']
            if row.get('sub_sub'):
                sub_display = f"{row['sub_item']} {row['sub_sub']}"

        ws.write_string(r, 0, row.get('section', ''), normal_fmt)
        ws.write_string(r, 1, row.get('item_num', ''), normal_fmt)
        ws.write_string(r, 2, sub_display, normal_fmt)
        write_rich_cell(ws, r, 3, row['r8_segments'], normal_fmt, ul_fmt)
        write_rich_cell(ws, r, 4, row['r6_segments'], normal_fmt, ul_fmt)
        if row['page'] > 0:
            ws.write_number(r, 5, row['page'], normal_fmt)
        else:
            ws.write_string(r, 5, '', normal_fmt)

    wb.close()
    print(f"\n完了！ {len(output_rows)} 件の差分を出力しました。", file=sys.stderr)


if __name__ == '__main__':
    main()
