#!/usr/bin/env python3
"""
施設基準 新旧対照表 統合スクリプト

告示Excel + 通知Excel → 1つの統合Excelに結合する。
リッチテキスト（アンダーライン書式）を保持して転記する。

基本診療料・特掲診療料それぞれについて統合を行う。
"""

import sys
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import xlsxwriter

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIGS = [
    {
        'label': '基本診療料',
        'kokuji': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(基本・告示)_新旧対照表.xlsx"),
        'tsuchi': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(基本・通知)_新旧対照表.xlsx"),
        'output': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(基本)_新旧対照表_統合.xlsx"),
    },
    {
        'label': '特掲診療料',
        'kokuji': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(特掲・告示)_新旧対照表.xlsx"),
        'tsuchi': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(特掲・通知)_新旧対照表.xlsx"),
        'output': os.path.join(SCRIPT_DIR, "output",
            "R8年度施設基準(特掲)_新旧対照表_統合.xlsx"),
    },
]

NS = {
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


def parse_shared_strings(xlsx_path):
    """sharedStrings.xmlを解析してリッチテキスト情報を含む文字列リストを返す。"""
    strings = []
    with zipfile.ZipFile(xlsx_path) as z:
        try:
            with z.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
        except KeyError:
            return strings

    root = tree.getroot()
    for si in root.findall('ss:si', NS):
        runs = si.findall('ss:r', NS)
        if runs:
            segments = []
            for r in runs:
                rpr = r.find('ss:rPr', NS)
                is_ul = False
                if rpr is not None:
                    u_elem = rpr.find('ss:u', NS)
                    if u_elem is not None:
                        is_ul = True
                t_elem = r.find('ss:t', NS)
                text = t_elem.text if t_elem is not None and t_elem.text else ''
                segments.append((text, is_ul))
            full_text = ''.join(t for t, _ in segments)
            strings.append({'type': 'rich', 'text': full_text,
                            'segments': segments})
        else:
            t_elem = si.find('ss:t', NS)
            text = t_elem.text if t_elem is not None and t_elem.text else ''
            strings.append({'type': 'plain', 'text': text,
                            'segments': [(text, False)]})

    return strings


def _ref_to_col(ref):
    """Excelのセル参照から0始まりの列インデックスを返す"""
    col = 0
    for ch in ref:
        if ch.isalpha():
            col = col * 26 + (ord(ch.upper()) - ord('A') + 1)
        else:
            break
    return col - 1


def parse_worksheet(xlsx_path, shared_strings):
    """sheet1.xmlを解析してセルデータを返す。"""
    with zipfile.ZipFile(xlsx_path) as z:
        with z.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)

    root = tree.getroot()
    sheet_data = root.find('ss:sheetData', NS)
    if sheet_data is None:
        return [], []

    all_rows = []
    for row_elem in sheet_data.findall('ss:row', NS):
        row_data = {}
        for cell_elem in row_elem.findall('ss:c', NS):
            ref = cell_elem.get('r')
            col_idx = _ref_to_col(ref)
            cell_type = cell_elem.get('t', '')
            v_elem = cell_elem.find('ss:v', NS)
            value = v_elem.text if v_elem is not None and v_elem.text else ''

            if cell_type == 's':
                ss_idx = int(value)
                if ss_idx < len(shared_strings):
                    row_data[col_idx] = shared_strings[ss_idx]
                else:
                    row_data[col_idx] = {'type': 'plain', 'text': '',
                                         'segments': [('', False)]}
            elif cell_type == 'n' or (cell_type == '' and value):
                try:
                    num_val = float(value)
                    if num_val == int(num_val):
                        num_val = int(num_val)
                    row_data[col_idx] = {'type': 'number', 'text': str(num_val),
                                         'segments': [], 'value': num_val}
                except ValueError:
                    row_data[col_idx] = {'type': 'plain', 'text': value,
                                         'segments': [(value, False)]}
            else:
                row_data[col_idx] = {'type': 'plain', 'text': value,
                                     'segments': [(value, False)]}

        all_rows.append(row_data)

    if not all_rows:
        return [], []

    header_row = all_rows[0]
    max_col = max(header_row.keys()) if header_row else 0
    headers = []
    for c in range(max_col + 1):
        cell = header_row.get(c, {'text': ''})
        headers.append(cell.get('text', ''))

    rows = []
    for row_dict in all_rows[1:]:
        row_cells = []
        for c in range(max_col + 1):
            cell = row_dict.get(c, {'type': 'plain', 'text': '',
                                    'segments': [('', False)]})
            row_cells.append(cell)
        rows.append(row_cells)

    return headers, rows


def read_xlsx_with_richtext(path):
    """XLSXファイルからリッチテキスト情報を含むデータを読み込む。"""
    print(f"  読み込み中: {os.path.basename(path)}", file=sys.stderr)
    shared_strings = parse_shared_strings(path)
    headers, rows = parse_worksheet(path, shared_strings)
    print(f"    {len(rows)} 行, shared strings: {len(shared_strings)}",
          file=sys.stderr)
    return headers, rows


def write_cell(ws, row_idx, col_idx, cell_data, normal_fmt, ul_fmt):
    """セルデータをxlsxwriterで書き込む。"""
    if cell_data.get('type') == 'number':
        ws.write_number(row_idx, col_idx, cell_data['value'], normal_fmt)
        return

    segments = cell_data.get('segments', [])
    if not segments:
        ws.write_string(row_idx, col_idx, '', normal_fmt)
        return

    has_ul = any(is_ul for _, is_ul in segments)
    if not has_ul:
        text = ''.join(t for t, _ in segments)
        ws.write_string(row_idx, col_idx, text, normal_fmt)
        return

    parts = []
    for text, is_ul in segments:
        if not text:
            continue
        fmt = ul_fmt if is_ul else normal_fmt
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


def process_merge(config):
    """告示と通知の統合処理"""
    label = config['label']
    kokuji_path = config['kokuji']
    tsuchi_path = config['tsuchi']
    output_path = config['output']

    print(f"\n{'=' * 60}", file=sys.stderr)
    print(f"施設基準（{label}）統合開始", file=sys.stderr)
    print(f"{'=' * 60}", file=sys.stderr)

    k_headers, k_rows = read_xlsx_with_richtext(kokuji_path)
    t_headers, t_rows = read_xlsx_with_richtext(tsuchi_path)

    # 統合Excel出力
    # 告示ヘッダー: 項目, 項番, 改正後（R8）, 改正前（R6）, ページ
    # 通知ヘッダー: 別添, 項目, 項番, 改正後（R8）, 改正前（R6）, ページ
    # 統合ヘッダー: 種別, 別添, 項目, 項番, 改正後（R8）, 改正前（R6）, ページ

    print(f"\n  XLSX出力: {os.path.basename(output_path)}", file=sys.stderr)

    wb = xlsxwriter.Workbook(output_path)
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

    out_headers = ['種別', '別添', '項目', '項番',
                   '改正後（R8）', '改正前（R6）', 'ページ']
    for col, h in enumerate(out_headers):
        ws.write(0, col, h, header_fmt)

    ws.set_column(0, 0, 8)    # 種別
    ws.set_column(1, 1, 12)   # 別添
    ws.set_column(2, 2, 30)   # 項目
    ws.set_column(3, 3, 12)   # 項番
    ws.set_column(4, 4, 60)   # 改正後
    ws.set_column(5, 5, 60)   # 改正前
    ws.set_column(6, 6, 6)    # ページ

    empty_cell = {'type': 'plain', 'text': '', 'segments': [('', False)]}
    row_num = 1
    total_rows = 0

    # 告示行を出力
    for row in k_rows:
        ws.write_string(row_num, 0, '告示', normal_fmt)
        ws.write_string(row_num, 1, '', normal_fmt)  # 別添（告示にはない）
        # 告示: col0=項目, col1=項番, col2=改正後, col3=改正前, col4=ページ
        write_cell(ws, row_num, 2, row[0] if len(row) > 0 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 3, row[1] if len(row) > 1 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 4, row[2] if len(row) > 2 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 5, row[3] if len(row) > 3 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 6, row[4] if len(row) > 4 else empty_cell,
                   normal_fmt, ul_fmt)
        row_num += 1
        total_rows += 1

    # 通知行を出力
    for row in t_rows:
        ws.write_string(row_num, 0, '通知', normal_fmt)
        # 通知: col0=別添, col1=項目, col2=項番, col3=改正後, col4=改正前, col5=ページ
        write_cell(ws, row_num, 1, row[0] if len(row) > 0 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 2, row[1] if len(row) > 1 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 3, row[2] if len(row) > 2 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 4, row[3] if len(row) > 3 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 5, row[4] if len(row) > 4 else empty_cell,
                   normal_fmt, ul_fmt)
        write_cell(ws, row_num, 6, row[5] if len(row) > 5 else empty_cell,
                   normal_fmt, ul_fmt)
        row_num += 1
        total_rows += 1

    wb.close()
    print(f"  完了！ {total_rows} 行を出力しました。", file=sys.stderr)


def main():
    for config in CONFIGS:
        process_merge(config)


if __name__ == '__main__':
    main()
