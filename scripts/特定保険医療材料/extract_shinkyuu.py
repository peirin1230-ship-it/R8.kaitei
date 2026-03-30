#!/usr/bin/env python3
"""
特定保険医療材料 新旧対照表 抽出スクリプト

R8とR6の特定保険医療材料PDF（告示・定義・留意事項）を比較し、
difflib で文字レベルの差分を検出して XLSX 形式の新旧対照表を生成する。
差分部分にはアンダーライン書式を適用する。

3組のPDFペアを処理:
  1. 告示（材料価格基準）
  2. 定義
  3. 留意事項
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
# パス設定
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(SCRIPT_DIR, "..", "..")
BASE_DIR = os.path.join(REPO_ROOT, "特定保険医療材料")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

CONFIGS = [
    {
        'label': '告示（材料価格基準）',
        'r8_pdf': os.path.join(BASE_DIR, "R8年度",
            "特定保険医療材料及びその材料価格（材料価格基準）の一部を改正する件（令和８年厚生労働省告示第73号）.pdf"),
        'r6_pdf': os.path.join(BASE_DIR, "R6年度",
            "特定保険医療材料及びその材料価格（材料価格基準）の一部を改正する告示　令和６年 厚生労働省告示第61号.pdf"),
        'output': os.path.join(OUTPUT_DIR,
            "特定保険医療材料_告示_新旧対照表.xlsx"),
        'skip_pages': 1,   # 先頭1ページ（表紙）をスキップ
    },
    {
        'label': '定義',
        'r8_pdf': os.path.join(BASE_DIR, "R8年度",
            "特定保険医療材料の定義について（令和８年３月５日保医発0305第４号）.pdf"),
        'r6_pdf': os.path.join(BASE_DIR, "R6年度",
            "特定保険医療材料の定義について（通知）　令和６年３月５日 保医発0305第12号.pdf"),
        'output': os.path.join(OUTPUT_DIR,
            "特定保険医療材料_定義_新旧対照表.xlsx"),
        'skip_pages': 1,
    },
    {
        'label': '留意事項',
        'r8_pdf': os.path.join(BASE_DIR, "R8年度",
            "特定保険医療材料の材料価格算定に関する留意事項について（令和８年３月５日保医発0305第１号）.pdf"),
        'r6_pdf': os.path.join(BASE_DIR, "R6年度",
            "特定保険医療材料の材料価格算定に関する留意事項について（通知） 令和６年３月５日.pdf"),
        'output': os.path.join(OUTPUT_DIR,
            "特定保険医療材料_留意事項_新旧対照表.xlsx"),
        'skip_pages': 1,
    },
]

# ============================================================
# 区分番号・大区分の正規表現
# ============================================================
# 大区分: Ⅰ～Ⅸ（歯科・調剤含む全区分）
RE_SECTION = re.compile(r'^(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)\s+')
# 区分番号: 3桁数字 + 材料名（日本語を含む名称）
# ※ 価格（"116,000 円" 等）を誤認識しないよう、材料名にひらがな・カタカナ・漢字を要求
RE_KUBUN = re.compile(r'^(\d{3})\s+([\u3000-\u9FFF\uF900-\uFAFF].+|削除.*)')
# 複合区分番号: "008，009" のようなカンマ区切り
RE_KUBUN_MULTI = re.compile(r'^(\d{3}[，,]\d{3}(?:[，,]\d{3})*)\s+([\u3000-\u9FFF\uF900-\uFAFF].+|削除.*)')
# 一般的事項の番号: "１ ", "２ " 等
RE_IPPAN = re.compile(r'^[１２３４５６７８９０]+\s+')
# ページ番号行（除外用）
RE_PAGE_NUM = re.compile(r'^\s*-\s*\d+\s*-\s*$')


def normalize_text(text):
    """比較用にテキストを正規化する。"""
    # 全角数字→半角
    t = text
    for zc, hc in zip('０１２３４５６７８９', '0123456789'):
        t = t.replace(zc, hc)
    # 全角スペース→半角
    t = t.replace('\u3000', ' ')
    # 連続空白を1つに
    t = re.sub(r'[ \t]+', ' ', t)
    # 改行前後の空白除去
    t = re.sub(r' *\n *', '\n', t)
    return t.strip()


def extract_blocks_from_pdf(pdf_path, skip_pages=0):
    """
    PDFからテキストを抽出し、区分番号ごとのブロックに分割する。

    Returns:
        list of dict: [{
            'section': 'Ⅰ' or 'Ⅱ' etc.,
            'kubun': '001' or '008，009',
            'name': '材料名',
            'text': 'ブロック全文',
            'page': ページ番号,
        }, ...]
    """
    doc = fitz.open(pdf_path)
    blocks = []
    current_section = ''
    current_kubun = ''
    current_name = ''
    current_lines = []
    current_page = 0

    for pg_idx in range(skip_pages, len(doc)):
        page = doc[pg_idx]
        text = page.get_text()
        lines = text.split('\n')

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            # ページ番号行をスキップ
            if RE_PAGE_NUM.match(stripped):
                continue

            # 大区分の検出
            m_sec = RE_SECTION.match(stripped)
            if m_sec:
                # 前のブロックを保存
                if current_kubun and current_lines:
                    blocks.append(_make_block(
                        current_section, current_kubun, current_name,
                        current_lines, current_page))
                    current_lines = []
                    current_kubun = ''
                    current_name = ''

                current_section = m_sec.group(1)
                # 大区分行のテキストも保持（後のブロックに含めない）
                remaining = stripped[m_sec.end():].strip()
                if remaining:
                    # 大区分と同じ行に区分番号が来ることはないので、
                    # 大区分の説明テキストとして保持
                    pass
                continue

            # 複合区分番号の検出（008，009 のような形式）
            m_multi = RE_KUBUN_MULTI.match(stripped)
            if m_multi:
                if current_kubun and current_lines:
                    blocks.append(_make_block(
                        current_section, current_kubun, current_name,
                        current_lines, current_page))
                current_kubun = m_multi.group(1)
                current_name = m_multi.group(2).strip()
                current_lines = [stripped]
                current_page = pg_idx + 1
                continue

            # 区分番号の検出
            m_kubun = RE_KUBUN.match(stripped)
            if m_kubun:
                if current_kubun and current_lines:
                    blocks.append(_make_block(
                        current_section, current_kubun, current_name,
                        current_lines, current_page))
                current_kubun = m_kubun.group(1)
                current_name = m_kubun.group(2).strip()
                current_lines = [stripped]
                current_page = pg_idx + 1
                continue

            # それ以外: 現在のブロックに追加
            if current_kubun:
                current_lines.append(stripped)
            # 区分番号前のテキスト（一般的事項等）は別途処理
            elif current_section and not current_kubun:
                # 一般的事項等のテキスト
                if not blocks or blocks[-1].get('kubun') != '_general':
                    if current_lines:
                        blocks.append(_make_block(
                            current_section, '_general', '一般的事項',
                            current_lines, current_page))
                    current_lines = [stripped]
                    current_page = pg_idx + 1
                    current_kubun = '_general'
                    current_name = '一般的事項'
                else:
                    current_lines.append(stripped)

    # 最後のブロックを保存
    if current_kubun and current_lines:
        blocks.append(_make_block(
            current_section, current_kubun, current_name,
            current_lines, current_page))

    doc.close()
    return blocks


def _make_block(section, kubun, name, lines, page):
    """ブロック辞書を作成する。"""
    text = '\n'.join(lines)
    return {
        'section': section,
        'kubun': kubun,
        'name': name,
        'text': text,
        'page': page,
    }


def make_block_key(block):
    """マッチング用のキーを生成する。"""
    return (block['section'], block['kubun'])


def text_similarity(t1, t2):
    """2つのテキストの類似度を返す（0.0-1.0）。"""
    n1 = normalize_text(t1)
    n2 = normalize_text(t2)
    if not n1 and not n2:
        return 1.0
    return difflib.SequenceMatcher(None, n1, n2, autojunk=False).ratio()


def compute_diff_segments(r8_text, r6_text):
    """
    R8とR6のテキストを文字レベルで比較し、セグメントリストを返す。

    Returns:
        (r8_segments, r6_segments): [( text, is_changed ), ...]
        差分がなければ (None, None)
    """
    r8_norm = normalize_text(r8_text)
    r6_norm = normalize_text(r6_text)

    if r8_norm == r6_norm:
        return None, None

    # スペース除去後も一致ならスキップ
    if re.sub(r'\s', '', r8_norm) == re.sub(r'\s', '', r6_norm):
        return None, None

    sm = difflib.SequenceMatcher(None, r6_norm, r8_norm, autojunk=False)
    ratio = sm.ratio()

    # 類似度が低すぎる場合は全体をアンダーライン
    if ratio < 0.3:
        return [(r8_text, True)], [(r6_text, True)]

    r8_segments = []
    r6_segments = []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            r6_segments.append((r6_norm[i1:i2], False))
            r8_segments.append((r8_norm[j1:j2], False))
        elif op == 'replace':
            r6_segments.append((r6_norm[i1:i2], True))
            r8_segments.append((r8_norm[j1:j2], True))
        elif op == 'insert':
            r8_segments.append((r8_norm[j1:j2], True))
        elif op == 'delete':
            r6_segments.append((r6_norm[i1:i2], True))

    return r8_segments, r6_segments


def write_rich_cell(ws, row_idx, col_idx, segments, normal_fmt, ul_fmt):
    """セグメントリストをリッチテキストとしてセルに書き込む。"""
    if not segments:
        return

    # 変更がない場合
    if len(segments) == 1 and not segments[0][1]:
        ws.write_string(row_idx, col_idx, segments[0][0], normal_fmt)
        return

    # 全体が変更の場合
    if len(segments) == 1 and segments[0][1]:
        ws.write_string(row_idx, col_idx, segments[0][0], ul_fmt)
        return

    # リッチテキスト構築
    parts = []
    for text, is_changed in segments:
        if not text:
            continue
        fmt = ul_fmt if is_changed else normal_fmt
        parts.extend([fmt, text])

    if len(parts) >= 4:
        try:
            ws.write_rich_string(row_idx, col_idx, *parts, normal_fmt)
        except Exception:
            # フォールバック: プレーンテキスト
            full_text = ''.join(t for t, _ in segments)
            ws.write_string(row_idx, col_idx, full_text, normal_fmt)
    elif parts:
        ws.write_string(row_idx, col_idx, parts[1] if len(parts) >= 2 else '',
                        parts[0] if parts else normal_fmt)


def process_pair(config):
    """1組のPDFペアを処理し、新旧対照表Excelを出力する。"""
    label = config['label']
    r8_pdf = config['r8_pdf']
    r6_pdf = config['r6_pdf']
    output_xlsx = config['output']
    skip_pages = config.get('skip_pages', 0)

    print(f"\n{'='*60}", file=sys.stderr)
    print(f"[{label}] 処理開始", file=sys.stderr)
    print(f"  R8: {os.path.basename(r8_pdf)}", file=sys.stderr)
    print(f"  R6: {os.path.basename(r6_pdf)}", file=sys.stderr)

    # Phase 1: PDFからブロック抽出
    print(f"\n[Phase 1] PDFからブロック抽出中...", file=sys.stderr)
    r8_blocks = extract_blocks_from_pdf(r8_pdf, skip_pages)
    r6_blocks = extract_blocks_from_pdf(r6_pdf, skip_pages)
    print(f"  R8: {len(r8_blocks)} ブロック", file=sys.stderr)
    print(f"  R6: {len(r6_blocks)} ブロック", file=sys.stderr)

    # Phase 2: ブロックマッチングと差分検出
    print(f"\n[Phase 2] ブロックマッチング中...", file=sys.stderr)

    # R6ブロックをキーで索引
    r6_by_key = {}
    for b in r6_blocks:
        key = make_block_key(b)
        r6_by_key.setdefault(key, []).append(b)

    output_rows = []
    matched_r6_keys = set()
    r6_only_count = 0
    changed_count = 0
    new_count = 0

    for r8_block in r8_blocks:
        # 欠番マーカー（"NNN 削除" のみのブロック）はスキップ
        text_stripped = re.sub(r'\s+', '', r8_block['text'])
        kubun_digits = re.sub(r'[，,\s]', '', r8_block['kubun'])
        text_without_kubun = text_stripped
        for d in kubun_digits:
            text_without_kubun = text_without_kubun.replace(d, '', 1)
        if text_without_kubun == '削除' or r8_block['name'] == '削除':
            continue

        key = make_block_key(r8_block)
        r6_candidates = r6_by_key.get(key, [])

        # R6側の欠番マーカーも候補から除外
        r6_candidates = [b for b in r6_candidates
                         if not (re.sub(r'\s+', '', b['text']).endswith('削除')
                                 and len(re.sub(r'\s+', '', b['text'])) < 15)]

        if not r6_candidates:
            # 同じ区分番号がR6の別の大区分にある場合を探す（区分番号のみでマッチ）
            r6_fallback = [b for b in r6_blocks
                           if b['kubun'] == r8_block['kubun']
                           and b['section'] != r8_block['section']
                           and not (re.sub(r'\s+', '', b['text']).endswith('削除')
                                    and len(re.sub(r'\s+', '', b['text'])) < 15)]
            if r6_fallback:
                r6_candidates = r6_fallback

        if not r6_candidates:
            # R8のみ（新設）
            output_rows.append({
                'section': r8_block['section'],
                'kubun': r8_block['kubun'],
                'name': r8_block['name'],
                'r8_segments': [(r8_block['text'], True)],
                'r6_segments': [('（新設）', False)],
                'page': r8_block['page'],
                'status': '新設',
            })
            new_count += 1
            continue

        # 複数候補がある場合は類似度で最良を選択
        best_r6 = None
        best_sim = -1
        for r6_b in r6_candidates:
            sim = text_similarity(r8_block['text'], r6_b['text'])
            if sim > best_sim:
                best_sim = sim
                best_r6 = r6_b

        matched_r6_keys.add((key, id(best_r6)))

        # 差分計算
        r8_segs, r6_segs = compute_diff_segments(r8_block['text'], best_r6['text'])

        if r8_segs is None:
            # 差分なし → スキップ
            continue

        output_rows.append({
            'section': r8_block['section'],
            'kubun': r8_block['kubun'],
            'name': r8_block['name'],
            'r8_segments': r8_segs,
            'r6_segments': r6_segs,
            'page': r8_block['page'],
            'status': '変更',
        })
        changed_count += 1

    # R6のみ（削除）のブロックを検出
    for r6_block in r6_blocks:
        # R6側の欠番マーカーもスキップ
        r6_text_stripped = re.sub(r'\s+', '', r6_block['text'])
        if r6_text_stripped.endswith('削除') and len(r6_text_stripped) < 15:
            continue
        if r6_block['name'] == '削除':
            continue

        key = make_block_key(r6_block)
        # R8に同じキーがあるかチェック
        r8_has_key = any(make_block_key(b) == key for b in r8_blocks)
        # 区分番号のみでもチェック（大区分が変わった場合）
        r8_has_kubun = any(b['kubun'] == r6_block['kubun'] for b in r8_blocks)
        if not r8_has_key and not r8_has_kubun:
            output_rows.append({
                'section': r6_block['section'],
                'kubun': r6_block['kubun'],
                'name': r6_block['name'],
                'r8_segments': [('（削除）', False)],
                'r6_segments': [(r6_block['text'], True)],
                'page': r6_block['page'],
                'status': '削除',
            })
            r6_only_count += 1

    print(f"  変更: {changed_count}件, 新設: {new_count}件, 削除: {r6_only_count}件",
          file=sys.stderr)

    # Phase 3: XLSX出力
    print(f"\n[Phase 3] XLSX出力: {output_xlsx}", file=sys.stderr)
    os.makedirs(os.path.dirname(output_xlsx), exist_ok=True)

    wb = xlsxwriter.Workbook(output_xlsx)

    # 書式定義
    header_fmt = wb.add_format({
        'bold': True,
        'font_name': '游ゴシック',
        'font_size': 11,
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'bg_color': '#D9E1F2',
        'align': 'center',
    })
    normal_fmt = wb.add_format({
        'font_name': '游ゴシック',
        'font_size': 10,
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
    })
    ul_fmt = wb.add_format({
        'font_name': '游ゴシック',
        'font_size': 10,
        'underline': True,
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
    })
    center_fmt = wb.add_format({
        'font_name': '游ゴシック',
        'font_size': 10,
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'align': 'center',
    })
    new_fmt = wb.add_format({
        'font_name': '游ゴシック',
        'font_size': 10,
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'bg_color': '#FFF2CC',
    })
    del_fmt = wb.add_format({
        'font_name': '游ゴシック',
        'font_size': 10,
        'text_wrap': True,
        'valign': 'top',
        'border': 1,
        'bg_color': '#FCE4EC',
    })

    ws = wb.add_worksheet('新旧対照表')

    # ヘッダー
    headers = ['大区分', '区分番号', '材料名', '改正後（R8）', '改正前（R6）',
               'ページ', '状態']
    col_widths = [8, 10, 25, 55, 55, 8, 8]
    for ci, (h, w) in enumerate(zip(headers, col_widths)):
        ws.write(0, ci, h, header_fmt)
        ws.set_column(ci, ci, w)

    # データ行
    row_idx = 1
    for row_data in output_rows:
        status = row_data['status']

        # 状態に応じた背景色
        if status == '新設':
            status_fmt = new_fmt
        elif status == '削除':
            status_fmt = del_fmt
        else:
            status_fmt = center_fmt

        ws.write(row_idx, 0, row_data['section'], center_fmt)
        ws.write(row_idx, 1, row_data['kubun'], center_fmt)
        ws.write(row_idx, 2, row_data['name'], normal_fmt)

        # 改正後（R8）
        write_rich_cell(ws, row_idx, 3, row_data['r8_segments'],
                        normal_fmt, ul_fmt)
        # 改正前（R6）
        write_rich_cell(ws, row_idx, 4, row_data['r6_segments'],
                        normal_fmt, ul_fmt)

        ws.write(row_idx, 5, row_data['page'], center_fmt)
        ws.write(row_idx, 6, status, status_fmt)

        row_idx += 1

    wb.close()
    print(f"  出力完了: {row_idx - 1}行", file=sys.stderr)
    return row_idx - 1


def main():
    print("特定保険医療材料 新旧対照表 抽出スクリプト", file=sys.stderr)
    print(f"出力先: {OUTPUT_DIR}", file=sys.stderr)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    total_rows = 0
    for config in CONFIGS:
        if not os.path.exists(config['r8_pdf']):
            print(f"\n[スキップ] R8 PDF が見つかりません: {config['r8_pdf']}",
                  file=sys.stderr)
            continue
        if not os.path.exists(config['r6_pdf']):
            print(f"\n[スキップ] R6 PDF が見つかりません: {config['r6_pdf']}",
                  file=sys.stderr)
            continue
        total_rows += process_pair(config)

    print(f"\n{'='*60}", file=sys.stderr)
    print(f"全処理完了: 合計 {total_rows} 行の変更を検出", file=sys.stderr)


if __name__ == '__main__':
    main()
