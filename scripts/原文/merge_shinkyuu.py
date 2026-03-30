#!/usr/bin/env python3
"""
新旧対照表 統合スクリプト

告示Excel + 通知Excel → 1つの統合Excelに結合する。
告示の出現順を基準に、各項目（部+区分番号）ごとに告示→通知の順で並べて出力。
リッチテキスト（アンダーライン書式）を保持して転記する。

xlsxwriterで生成されたXLSXのリッチテキストはopenpyxlで正しく読めないため、
XLSXのXML（sharedStrings.xml + sheet1.xml）を直接解析する。
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
REPO_ROOT = os.path.join(SCRIPT_DIR, "..", "..")
KOKUJI_XLSX = os.path.join(REPO_ROOT, "output",
    "R8年度医科点数表(告示)_新旧対照表.xlsx")
TSUCHI_XLSX = os.path.join(REPO_ROOT, "output",
    "R8年度医科点数表(通知)_新旧対照表.xlsx")
OUTPUT_XLSX = os.path.join(REPO_ROOT, "output",
    "R8年度医科点数表_新旧対照表_統合.xlsx")

# Excel XML名前空間
NS = {
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


def parse_shared_strings(xlsx_path):
    """sharedStrings.xmlを解析してリッチテキスト情報を含む文字列リストを返す。

    Returns:
        [{type: 'plain'|'rich', text: str, segments: [(text, is_underline), ...]}]
    """
    strings = []
    with zipfile.ZipFile(xlsx_path) as z:
        try:
            with z.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
        except KeyError:
            return strings

    root = tree.getroot()
    for si in root.findall('ss:si', NS):
        # <si> の中に <t> (plain) か <r> (rich run) がある
        runs = si.findall('ss:r', NS)
        if runs:
            # リッチテキスト
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
            strings.append({'type': 'rich', 'text': full_text, 'segments': segments})
        else:
            t_elem = si.find('ss:t', NS)
            text = t_elem.text if t_elem is not None and t_elem.text else ''
            strings.append({'type': 'plain', 'text': text, 'segments': [(text, False)]})

    return strings


def parse_worksheet(xlsx_path, shared_strings):
    """sheet1.xmlを解析してセルデータを返す。

    Returns:
        headers: [str, ...]
        rows: [[cell_data, ...], ...]
            cell_data = {type: 'plain'|'rich'|'number', text: str, segments: [...], value: float}
    """
    with zipfile.ZipFile(xlsx_path) as z:
        # sheet1.xmlのパスを取得
        sheet_path = 'xl/worksheets/sheet1.xml'
        with z.open(sheet_path) as f:
            tree = ET.parse(f)

    root = tree.getroot()
    sheet_data = root.find('ss:sheetData', NS)
    if sheet_data is None:
        return [], []

    all_rows = []
    for row_elem in sheet_data.findall('ss:row', NS):
        row_data = {}
        for cell_elem in row_elem.findall('ss:c', NS):
            ref = cell_elem.get('r')  # e.g., "A1", "B2"
            col_idx = _ref_to_col(ref)
            cell_type = cell_elem.get('t', '')
            v_elem = cell_elem.find('ss:v', NS)
            value = v_elem.text if v_elem is not None and v_elem.text else ''

            if cell_type == 's':
                # Shared string reference
                ss_idx = int(value)
                if ss_idx < len(shared_strings):
                    ss = shared_strings[ss_idx]
                    row_data[col_idx] = ss
                else:
                    row_data[col_idx] = {'type': 'plain', 'text': '',
                                         'segments': [('', False)]}
            elif cell_type == 'n' or (cell_type == '' and value):
                # Number
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

    # ヘッダー行（最初の行）
    header_row = all_rows[0]
    max_col = max(header_row.keys()) if header_row else 0
    headers = []
    for c in range(max_col + 1):
        cell = header_row.get(c, {'text': ''})
        headers.append(cell.get('text', ''))

    # データ行
    rows = []
    for row_dict in all_rows[1:]:
        row_cells = []
        for c in range(max_col + 1):
            cell = row_dict.get(c, {'type': 'plain', 'text': '',
                                    'segments': [('', False)]})
            row_cells.append(cell)
        rows.append(row_cells)

    return headers, rows


def _ref_to_col(ref):
    """Excelのセル参照 (e.g., "A1", "AB3") から0始まりの列インデックスを返す"""
    col = 0
    for ch in ref:
        if ch.isalpha():
            col = col * 26 + (ord(ch.upper()) - ord('A') + 1)
        else:
            break
    return col - 1


def read_xlsx_with_richtext(path):
    """XLSXファイルからリッチテキスト情報を含むデータを読み込む。"""
    print(f"  読み込み中: {os.path.basename(path)}", file=sys.stderr)
    shared_strings = parse_shared_strings(path)
    headers, rows = parse_worksheet(path, shared_strings)
    print(f"    {len(rows)} 行, shared strings: {len(shared_strings)}", file=sys.stderr)
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


def _normalize_text(text):
    """比較用にスペースを除去したテキストを返す。"""
    return re.sub(r'\s+', '', text)


def _segments_text(segs):
    """セグメントリストからテキスト全文を結合して返す。"""
    return ''.join(t for t, _ in segs)


def _segments_text_no_ul(segs):
    """セグメントリストからUL以外のテキストを結合して返す。"""
    return ''.join(t for t, u in segs if not u)


def strip_false_underlines(rows):
    """偽陽性のアンダーラインを除去し、偽陽性行のインデックスを返す。

    3段階で処理:
    1. スペースのみのULセグメントからULフラグを除去
    2. 改正後と改正前の全文がスペース差異のみで同一の場合、全ULを除去
       （PDFの行折り返し位置変更による偽陽性）
    3. 片方のみULがある場合、UL部分を除いたテキストが反対側と一致すれば除去
       （前ページの見出しや項目名がUL付きで混入するケース）

    Returns:
        set: 偽陽性として処理された行のインデックス
    """
    # 処理前にULを持つ行を記録
    originally_had_ul = set()
    for i, row in enumerate(rows):
        if len(row) < 8:
            continue
        go, zen = row[6], row[7]
        if (any(u for _, u in go.get('segments', []))
                or any(u for _, u in zen.get('segments', []))):
            originally_had_ul.add(i)

    # Pass 1: スペースのみのULセグメントを除去
    for row in rows:
        for cell in row:
            segs = cell.get('segments', [])
            if not segs or not any(u for _, u in segs):
                continue
            cell['segments'] = [
                (text, False) if is_ul and not text.strip() else (text, is_ul)
                for text, is_ul in segs
            ]

    # Pass 2 & 3
    for row in rows:
        if len(row) < 8:
            continue
        go, zen = row[6], row[7]
        go_segs = go.get('segments', [])
        zen_segs = zen.get('segments', [])
        go_has_ul = any(u for _, u in go_segs)
        zen_has_ul = any(u for _, u in zen_segs)
        if not go_has_ul and not zen_has_ul:
            continue

        go_full = _normalize_text(_segments_text(go_segs))
        zen_full = _normalize_text(_segments_text(zen_segs))

        strip = False
        # Pass 2: 全文比較（行折り返し差異のみ）
        if go_full == zen_full:
            strip = True
        # Pass 3: UL部分を除いたテキストが反対側と一致
        # ただしULテキストが長い場合（正規化後20文字超）は実質的な変更なので除去しない
        elif go_has_ul and not zen_has_ul:
            if _normalize_text(_segments_text_no_ul(go_segs)) == zen_full:
                ul_len = len(_normalize_text(''.join(t for t, u in go_segs if u)))
                if ul_len <= 20:
                    strip = True
        elif zen_has_ul and not go_has_ul:
            if _normalize_text(_segments_text_no_ul(zen_segs)) == go_full:
                ul_len = len(_normalize_text(''.join(t for t, u in zen_segs if u)))
                if ul_len <= 20:
                    strip = True

        if strip:
            go['segments'] = [(t, False) for t, _ in go_segs]
            zen['segments'] = [(t, False) for t, _ in zen_segs]

    # 処理後にULを失った行 = 偽陽性
    false_positives = set()
    for i in originally_had_ul:
        row = rows[i]
        go, zen = row[6], row[7]
        if (not any(u for _, u in go.get('segments', []))
                and not any(u for _, u in zen.get('segments', []))):
            false_positives.add(i)

    return false_positives


def should_keep_row(row):
    """変更のない行（コンテキスト行）かどうかを判定する。

    以下の場合は変更ありとして保持:
    - ULがある（部分変更）
    - 片方が空（新設/削除）
    - 改正後と改正前のテキストが異なる（全面書き換え）
    """
    if len(row) < 8:
        return False
    go = row[6]
    zen = row[7]

    # ULがある → 変更あり
    if (any(u for _, u in go.get('segments', []))
            or any(u for _, u in zen.get('segments', []))):
        return True

    # 片方が空 → 新設/削除
    go_text = go.get('text', '').strip()
    zen_text = zen.get('text', '').strip()
    if not go_text or not zen_text:
        return True

    # テキスト比較: 異なれば変更あり（全面書き換え）
    return _normalize_text(go_text) != _normalize_text(zen_text)


def fill_missing_sections(kokuji_rows, tsuchi_rows):
    """通知の空白節を告示の情報から補完する。

    通知PDFでは一部の部（例: 第９部 処置）で節見出しが出現しないため、
    告示側の節情報を使って空白を埋める。

    1. 告示の (部, 項目コード) → 節 のマッピングを構築
    2. 通知の空白節を直接マッチで補完
    3. 直接マッチがない場合、同じ部の告示項目順で直近の節を使用
    """
    # 告示から部ごとの項目コード→節の対応を構築（出現順を保持）
    section_by_item = {}
    bu_code_sections = {}  # bu → [(code, section), ...]

    for row in kokuji_rows:
        bu = row[1].get('text', '') if len(row) > 1 else ''
        setu = row[2].get('text', '') if len(row) > 2 else ''
        code_field = row[4].get('text', '') if len(row) > 4 else ''
        code = code_field.split(' ')[0].split('\u3000')[0]
        if bu and setu and code and code != '通則':
            key = (bu, code)
            if key not in section_by_item:
                section_by_item[key] = setu
            if bu not in bu_code_sections:
                bu_code_sections[bu] = []
            if not bu_code_sections[bu] or bu_code_sections[bu][-1] != (code, setu):
                bu_code_sections[bu].append((code, setu))

    filled = 0
    for row in tsuchi_rows:
        setu = row[2].get('text', '') if len(row) > 2 else ''
        if setu:
            continue
        bu = row[1].get('text', '') if len(row) > 1 else ''
        code_field = row[4].get('text', '') if len(row) > 4 else ''
        code = code_field.split(' ')[0].split('\u3000')[0]
        if not bu or not code or code == '通則':
            continue

        # 直接マッチ
        lookup_setu = section_by_item.get((bu, code))

        # フォールバック: 告示の同じ部で、項目コード順で直近の節を使用
        if not lookup_setu and bu in bu_code_sections:
            last_setu = ''
            for kcode, ksetu in bu_code_sections[bu]:
                if kcode > code:
                    break
                last_setu = ksetu
            if last_setu:
                lookup_setu = last_setu

        if lookup_setu:
            row[2] = {'type': 'plain', 'text': lookup_setu,
                      'segments': [(lookup_setu, False)]}
            filled += 1

    return filled


def extract_group_key(row):
    """行から (部, 項目コード全体) のグループキーを抽出する。

    項目コードフィールド全体をキーに使用し、サブ項目単位で
    告示→通知をインターリーブさせる。
    例: "Ａ０００ 初診料" → ("第１部 初・再診料", "Ａ０００ 初診料")
        "Ｂ００１ 特定疾患治療管理料 ７ 難病外来指導管理料"
          → ("第２部 在宅医療", "Ｂ００１ 特定疾患治療管理料 ７ 難病外来指導管理料")
    """
    bu = row[1].get('text', '') if len(row) > 1 else ''
    code_field = row[4].get('text', '') if len(row) > 4 else ''

    return (bu, code_field)


def group_rows_by_key(rows):
    """行をグループキーでグループ化し、出現順を保持して返す。"""
    groups = {}
    order = []
    for row in rows:
        key = extract_group_key(row)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)
    return groups, order


def build_master_order(kokuji_order, tsuchi_order):
    """告示の出現順をベースに、通知のみの項目を適切な位置に挿入したマスター順序を返す。

    通知のみの項目は、通知での直前の共通項目の直後に挿入する。
    """
    kokuji_set = set(kokuji_order)
    master = list(kokuji_order)

    tsuchi_only = [k for k in tsuchi_order if k not in kokuji_set]

    for key in tsuchi_only:
        t_idx = tsuchi_order.index(key)
        # 通知での直前の共通項目（告示にも存在する項目）を探す
        insert_after = None
        for i in range(t_idx - 1, -1, -1):
            if tsuchi_order[i] in kokuji_set:
                insert_after = tsuchi_order[i]
                break

        if insert_after is not None:
            pos = master.index(insert_after) + 1
            # 既に挿入済みの通知のみ項目をスキップして、次の告示項目の手前に挿入
            while pos < len(master) and master[pos] not in kokuji_set:
                pos += 1
            master.insert(pos, key)
        else:
            # 直前に共通項目がない場合、通知での直後の共通項目の手前に挿入
            insert_before = None
            for i in range(t_idx + 1, len(tsuchi_order)):
                if tsuchi_order[i] in kokuji_set:
                    insert_before = tsuchi_order[i]
                    break
            if insert_before is not None:
                pos = master.index(insert_before)
                master.insert(pos, key)
            else:
                # 前後に共通項目がない場合、末尾に追加
                master.append(key)

    return master


def main():
    print("=" * 60, file=sys.stderr)
    print("新旧対照表 統合処理開始", file=sys.stderr)
    print("=" * 60, file=sys.stderr)

    if not os.path.exists(KOKUJI_XLSX):
        print(f"エラー: 告示Excelが見つかりません: {KOKUJI_XLSX}", file=sys.stderr)
        print("先に extract_kokuji_shinkyuu.py を実行してください。", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(TSUCHI_XLSX):
        print(f"エラー: 通知Excelが見つかりません: {TSUCHI_XLSX}", file=sys.stderr)
        sys.exit(1)

    print("\n[Phase 1] Excel読み込み", file=sys.stderr)
    kokuji_headers, kokuji_rows = read_xlsx_with_richtext(KOKUJI_XLSX)
    tsuchi_headers, tsuchi_rows = read_xlsx_with_richtext(TSUCHI_XLSX)

    print("\n[Phase 1.5] 偽陽性UL除去・変更なし行の削除", file=sys.stderr)
    fp_kokuji = strip_false_underlines(kokuji_rows)
    fp_tsuchi = strip_false_underlines(tsuchi_rows)
    print(f"  偽陽性UL除去: 告示{len(fp_kokuji)}件, 通知{len(fp_tsuchi)}件",
          file=sys.stderr)
    kokuji_before = len(kokuji_rows)
    tsuchi_before = len(tsuchi_rows)
    # 偽陽性行とコンテキスト行（テキスト同一でULなし）を削除
    kokuji_rows = [r for i, r in enumerate(kokuji_rows)
                   if i not in fp_kokuji and should_keep_row(r)]
    tsuchi_rows = [r for i, r in enumerate(tsuchi_rows)
                   if i not in fp_tsuchi and should_keep_row(r)]
    print(f"  告示: {kokuji_before}件 → {len(kokuji_rows)}件"
          f"（{kokuji_before - len(kokuji_rows)}件削除）", file=sys.stderr)
    print(f"  通知: {tsuchi_before}件 → {len(tsuchi_rows)}件"
          f"（{tsuchi_before - len(tsuchi_rows)}件削除）", file=sys.stderr)

    # 通知の空白節を告示から補完
    filled = fill_missing_sections(kokuji_rows, tsuchi_rows)
    if filled:
        print(f"  節補完: 通知{filled}件の空白節を告示から補完", file=sys.stderr)

    print("\n[Phase 2] グループ化・並び替え", file=sys.stderr)
    kokuji_groups, kokuji_order = group_rows_by_key(kokuji_rows)
    tsuchi_groups, tsuchi_order = group_rows_by_key(tsuchi_rows)

    master_order = build_master_order(kokuji_order, tsuchi_order)

    common = set(kokuji_order) & set(tsuchi_order)
    kokuji_only = [k for k in kokuji_order if k not in set(tsuchi_order)]
    tsuchi_only = [k for k in tsuchi_order if k not in set(kokuji_order)]
    print(f"  共通項目: {len(common)}グループ", file=sys.stderr)
    print(f"  告示のみ: {len(kokuji_only)}グループ", file=sys.stderr)
    print(f"  通知のみ: {len(tsuchi_only)}グループ", file=sys.stderr)
    print(f"  マスター順序: {len(master_order)}グループ", file=sys.stderr)

    print(f"\n[Phase 3] 統合Excel出力: {OUTPUT_XLSX}", file=sys.stderr)

    # ヘッダー: 種別 + 元のヘッダー
    headers = ['種別'] + kokuji_headers

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
        'underline': True, 'bold': True,
        'text_wrap': True, 'valign': 'top',
    })

    # 列幅設定: 種別(col0) + 元の列を1つ右にシフト
    ws.set_column(0, 0, 8)   # 種別
    ws.set_column(1, 1, 12)  # 章
    ws.set_column(2, 2, 12)  # 部
    ws.set_column(3, 3, 12)  # 節
    ws.set_column(4, 4, 12)  # 款
    ws.set_column(5, 5, 25)  # 通則/項目コード
    ws.set_column(6, 6, 8)   # 注
    ws.set_column(7, 7, 55)  # 改正後（R8）
    ws.set_column(8, 8, 55)  # 改正前（R6）
    ws.set_column(9, 9, 6)   # ページ

    # ヘッダー行
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    current_row = 1

    # マスター順序に従って、各項目ごとに告示→通知の順で出力
    for key in master_order:
        # 告示の行を出力
        if key in kokuji_groups:
            for row_data in kokuji_groups[key]:
                ws.write_string(current_row, 0, '告示', normal_fmt)
                for col, cell_data in enumerate(row_data):
                    write_cell(ws, current_row, col + 1, cell_data,
                               normal_fmt, ul_fmt)
                current_row += 1

        # 通知の行を出力
        if key in tsuchi_groups:
            for row_data in tsuchi_groups[key]:
                ws.write_string(current_row, 0, '通知', normal_fmt)
                for col, cell_data in enumerate(row_data):
                    write_cell(ws, current_row, col + 1, cell_data,
                               normal_fmt, ul_fmt)
                current_row += 1

    wb.close()

    total_rows = current_row - 1
    print(f"\n完了！ 告示{len(kokuji_rows)}件 + 通知{len(tsuchi_rows)}件 = "
          f"計{total_rows}件を統合しました。", file=sys.stderr)


if __name__ == '__main__':
    main()
