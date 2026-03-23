#!/usr/bin/env python3
"""
施設基準 告示PDF 新旧対照表 抽出スクリプト

R8告示PDF（縦書き）とR6告示PDF（縦書き）を比較し、
difflib で文字レベルの差分を検出して XLSX 形式の新旧対照表を生成する。
差分部分にはアンダーライン書式を適用する。

基本診療料・特掲診療料の両方を処理する。
"""

import fitz  # PyMuPDF
import re
import sys
import os
import difflib
import xlsxwriter
from collections import defaultdict

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
            "基本診療料の施設基準等の一部を改正する件（令和８年厚生労働省告示第70号）.pdf"),
        'r6_pdf': os.path.join(SCRIPT_DIR,
            "基本診療料の施設基準等の一部を改正する告示　令和６年 厚生労働省告示第58号.pdf"),
        'output': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(基本・告示)_新旧対照表.xlsx"),
        'r8_toc_pages': 0,
        'r6_toc_pages': 0,
    },
    {
        'label': '特掲診療料',
        'r8_pdf': os.path.join(SCRIPT_DIR,
            "特掲診療料の施設基準等の一部を改正する件（令和８年厚生労働省告示第71号）.pdf"),
        'r6_pdf': os.path.join(SCRIPT_DIR,
            "特掲診療料の施設基準等の一部を改正する件　令和６年 厚生労働省告示第59号.pdf"),
        'output': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(特掲・告示)_新旧対照表.xlsx"),
        'r8_toc_pages': 0,
        'r6_toc_pages': 0,
    },
]

MIN_FONT_SIZE = 5.0
RUBY_MAX_SIZE = 8.0  # フリガナ（ルビ）の最大フォントサイズ
# 縦書きPDFの列グループ化許容差（x座標の差）
X_GROUP_TOLERANCE = 5.0
# ページ番号の除外y閾値（ページ下部のページ番号を除去）
PAGE_NUM_Y_THRESHOLD = 790.0

# ============================================================
# 正規表現パターン（告示PDF用）
# ============================================================
# 第一 ~ 第十 等のセクション見出し（漢数字）
RE_DAI_KANJI = re.compile(
    r'^第([一二三四五六七八九十百]+(?:の[一二三四五六七八九十百]+)*)\s+(.*)')

# 漢数字の番号付き項目（一, 二, 三, ...）
RE_KANJI_NUM = re.compile(
    r'^([一二三四五六七八九十]+(?:の[一二三四五六七八九十]+)*)\s')

# (1), (2), ... パターン
RE_PAREN_NUM = re.compile(r'^[\(（](\d+|[０１２３４５６７８９]+)[\)）]\s')

# イ, ロ, ハ, ニ, ... パターン
RE_IROHA = re.compile(r'^([イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス])\s')

# ページ番号パターン
RE_PAGE_NUM = re.compile(r'^\d{1,4}\s*$')

# 見出しのみの行
RE_HEADING_ONLY = re.compile(
    r'^(第[一二三四五六七八九十百]+(?:の[一二三四五六七八九十百]+)*\s+.*|'
    r'削除\s*)$')


# ============================================================
# 縦書きPDFテキスト抽出
# ============================================================

def extract_page_lines_vertical(page):
    """縦書きPDFからテキスト行を抽出する。

    縦書き = 各列が右→左に進み、列内は上→下に文字が並ぶ。
    PyMuPDFでは各文字が個別に配置されるため、x座標でグループ化して列を形成し、
    右から左の順で読む。

    Returns:
        [(text, x0), ...] のリスト（右→左の列順）
    """
    rd = page.get_text('rawdict')

    chars = []
    for block in rd['blocks']:
        if block['type'] != 0:
            continue
        for line in block['lines']:
            for span in line['spans']:
                if span['size'] < MIN_FONT_SIZE:
                    continue
                # フリガナ（ルビ）を除外：小さいフォントのひらがな・カタカナ
                if span['size'] < RUBY_MAX_SIZE:
                    span_text = ''.join(
                        ch['c'] for ch in span.get('chars', []))
                    if re.match(r'^[ぁ-んァ-ヶー]+$', span_text.strip()):
                        continue
                for ch in span.get('chars', []):
                    c = ch['c']
                    if not c.strip():
                        continue
                    y = ch['bbox'][1]
                    # ページ下部のページ番号を除外
                    if y > PAGE_NUM_Y_THRESHOLD:
                        continue
                    x = ch['bbox'][0]
                    h = ch['bbox'][3] - ch['bbox'][1]
                    chars.append((x, y, c, h))

    if not chars:
        return []

    # x座標でグループ化（同一列の文字をまとめる）
    chars.sort(key=lambda c: c[0])
    x_groups = []
    for x, y, c, h in chars:
        placed = False
        for g in x_groups:
            if abs(g['x'] - x) < X_GROUP_TOLERANCE:
                g['chars'].append((y, c, h))
                placed = True
                break
        if not placed:
            x_groups.append({'x': x, 'chars': [(y, c, h)]})

    # 右から左の順にソート（縦書きの読み順）
    x_groups.sort(key=lambda g: g['x'], reverse=True)

    lines = []
    for g in x_groups:
        sorted_chars = sorted(g['chars'], key=lambda c: c[0])
        # 文字間のギャップを検出してスペースを挿入
        parts = []
        for i, (y, c, h) in enumerate(sorted_chars):
            if i > 0:
                prev_y, _, prev_h = sorted_chars[i - 1]
                gap = y - (prev_y + prev_h)
                # ギャップが文字高さの半分以上ならスペース挿入
                if gap > h * 0.5:
                    parts.append(' ')
            parts.append(c)
        text = ''.join(parts).strip()

        # ダッシュ正規化
        text = text.replace('\u2015', '－').replace('\u2014', '－').replace(
            '\u2013', '－')

        # ページ番号を除外
        if RE_PAGE_NUM.match(text):
            continue
        if not text:
            continue

        lines.append((text, g['x']))

    return lines


# ============================================================
# 階層追跡（告示PDF用）
# ============================================================

class HierarchyTracker:
    """第一~第十/漢数字項目/番号の階層を追跡する"""

    def __init__(self):
        self.section = ""        # 第一, 第二, etc.
        self.section_name = ""   # セクション名
        self.item_num = ""       # 漢数字項目 (一, 二, ...)
        self.sub_item = ""       # (1), (2) or イ, ロ, ...

    def update(self, text):
        if not text:
            return False

        # 第N パターン
        m = RE_DAI_KANJI.match(text)
        if m:
            self.section = f"第{m.group(1)}"
            self.section_name = m.group(2).strip()
            self.item_num = ""
            self.sub_item = ""
            return True

        # 漢数字番号
        m = RE_KANJI_NUM.match(text)
        if m:
            self.item_num = m.group(1)
            self.sub_item = ""
            return True

        # (N) パターン
        m = RE_PAREN_NUM.match(text)
        if m:
            self.sub_item = f"({m.group(1)})"
            return True

        # イロハ パターン
        m = RE_IROHA.match(text)
        if m:
            self.sub_item = m.group(1)
            return True

        return False

    def snapshot(self):
        return {
            'section': self.section,
            'section_name': self.section_name,
            'item_num': self.item_num,
            'sub_item': self.sub_item,
        }


# ============================================================
# ブロック分割
# ============================================================

def is_block_boundary(text):
    """テキストが新しいブロック境界かを判定"""
    if RE_DAI_KANJI.match(text):
        return True
    if RE_KANJI_NUM.match(text):
        return True
    return False


def is_heading_only_block(text):
    """ブロックが見出しのみかを判定"""
    lines_text = text.strip().split('\n')
    for line in lines_text:
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
                result.extend(deeper)
            else:
                result.append(sb)
    return result


def _split_single_block(block):
    """単一ブロックをサブアイテム境界で分割する"""
    lines = block['text'].split('\n')

    markers = []
    for i, line in enumerate(lines):
        stripped = line.strip()
        m = RE_PAREN_NUM.match(stripped)
        if m:
            markers.append((i, 'paren', f"({m.group(1)})"))
            continue
        m = RE_IROHA.match(stripped)
        if m:
            markers.append((i, 'iroha', m.group(1)))
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


def extract_blocks_from_pdf(pdf_path, toc_pages):
    """縦書き告示PDFからテキストブロックを抽出する。"""
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
                'section': current_ctx['section'],
                'section_name': current_ctx['section_name'],
                'item_num': current_ctx['item_num'],
                'text': block_text,
                'page': current_page,
            })

    for pg_idx in range(total_pages):
        if (pg_idx + 1) % 50 == 0:
            print(f"    処理中: {pg_idx+1}/{total_pages} ページ...",
                  file=sys.stderr)

        if pg_idx < toc_pages:
            continue

        page = doc[pg_idx]
        page_num = pg_idx + 1

        lines = extract_page_lines_vertical(page)

        for text, x0 in lines:
            boundary = is_block_boundary(text)
            tracker.update(text)

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
    return (block['section'], block['item_num'])


def make_block_key_short(block):
    """フォールバック用短キー"""
    return (block['item_num'],)


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
    r8_toc = config['r8_toc_pages']
    r6_toc = config['r6_toc_pages']

    print(f"\n{'=' * 60}", file=sys.stderr)
    print(f"施設基準（{label}・告示）新旧対照表 抽出開始", file=sys.stderr)
    print(f"{'=' * 60}", file=sys.stderr)

    # Phase 1: ブロック抽出
    print("\n[Phase 1] PDFからブロック抽出", file=sys.stderr)
    r8_blocks_raw = extract_blocks_from_pdf(r8_pdf, r8_toc)
    r6_blocks_raw = extract_blocks_from_pdf(r6_pdf, r6_toc)

    # 見出しのみのブロックと前文（セクション未設定）を除外
    r8_blocks = [b for b in r8_blocks_raw
                 if not is_heading_only_block(b['text'])
                 and b['section']]
    r6_blocks = [b for b in r6_blocks_raw
                 if not is_heading_only_block(b['text'])
                 and b['section']]
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
            'section': r8b['section'],
            'section_name': r8b.get('section_name', ''),
            'item_num': r8b['item_num'],
            'r8_segments': r8_segments,
            'r6_segments': r6_segments,
            'page': r8b['page'],
            '_from_split': r8b.get('_from_split', False),
        })

    # 削除項目
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

    headers = ['項目', '項番', '改正後（R8）', '改正前（R6）', 'ページ']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    ws.set_column(0, 0, 30)   # 項目（第N + 名称）
    ws.set_column(1, 1, 12)   # 項番
    ws.set_column(2, 2, 60)   # 改正後
    ws.set_column(3, 3, 60)   # 改正前
    ws.set_column(4, 4, 6)    # ページ

    for idx, row in enumerate(output_rows):
        r = idx + 1
        section_display = row['section']
        if row.get('section_name'):
            section_display = f"{row['section']} {row['section_name']}"

        ws.write_string(r, 0, section_display, normal_fmt)
        ws.write_string(r, 1, row.get('item_num', ''), normal_fmt)
        write_rich_cell(ws, r, 2, row['r8_segments'], normal_fmt, ul_fmt)
        write_rich_cell(ws, r, 3, row['r6_segments'], normal_fmt, ul_fmt)
        if row['page'] > 0:
            ws.write_number(r, 4, row['page'], normal_fmt)
        else:
            ws.write_string(r, 4, '', normal_fmt)

    wb.close()
    print(f"\n完了！ {len(output_rows)} 件の差分を出力しました。", file=sys.stderr)


def main():
    for config in CONFIGS:
        process_pair(config)


if __name__ == '__main__':
    main()
