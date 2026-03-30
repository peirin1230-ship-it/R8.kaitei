"""
告示（点数表）と通知（留意事項）を部ごとに統合するスクリプト。

使い方:
    python merge_by_bu.py

各部について、告示69号の該当ページ → 第6号の該当ページ の順に結合し、
1つのPDFとして出力する。
"""

import os
import fitz  # PyMuPDF

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(BASE_DIR, "..", "..")
OUTPUT_DIR = os.path.join(REPO_ROOT, "原文", "分割", "部別_告示通知統合")

# 入力PDFファイル
PDF_KOKUJI = os.path.join(
    REPO_ROOT, "原文",
    "診療報酬の算定方法の一部を改正する件（令和８年厚生労働省告示第69号）医科点数表.pdf",
)
PDF_TSUCHI = os.path.join(
    REPO_ROOT, "原文",
    "診療報酬の算定方法の一部改正に伴う実施上の留意事項について（通知）（令和８年３月５日保医発0305第６号）医科診療報酬点数表.pdf",
)

# ---------------------------------------------------------------------------
# 統合定義: (出力ファイル名, 告示ページ範囲, 通知ページ範囲)
# ページは1始まり。None の場合はその資料に対応セクションなし。
# ---------------------------------------------------------------------------
MERGES = [
    ("第1章_第1部_初再診料.pdf",                (3, 9),     (3, 22)),
    ("第1章_第2部_入院料等.pdf",                (10, 91),   (23, 150)),
    ("第2章_第1部_医学管理等.pdf",              (92, 134),  (151, 242)),
    ("第2章_第2部_在宅医療.pdf",                (135, 173), (243, 324)),
    ("第2章_第3部_検査.pdf",                    (174, 212), (325, 448)),
    ("第2章_第4部_画像診断.pdf",                (213, 220), (449, 469)),
    ("第2章_第5部_投薬.pdf",                    (221, 224), (470, 479)),
    ("第2章_第6部_注射.pdf",                    (225, 228), (480, 488)),
    ("第2章_第7部_リハビリテーション.pdf",       (229, 237), (489, 514)),
    ("第2章_第8部_精神科専門療法.pdf",           (238, 256), (515, 551)),
    ("第2章_第9部_処置.pdf",                    (257, 271), (552, 591)),
    ("第2章_第10部_手術.pdf",                   (272, 358), (592, 690)),
    ("第2章_第11部_麻酔.pdf",                   (359, 364), (691, 701)),
    ("第2章_第12部_放射線治療.pdf",             (365, 368), (702, 710)),
    ("第2章_第13部_病理診断.pdf",               (369, 370), (711, 717)),
    ("第2章_第14部_その他.pdf",                 (371, 400), (718, 719)),
    # 通知のみ（告示に対応セクションなし）
    ("告示69号_目次.pdf",                       (1, 2),     None),
    ("第6号_通則.pdf",                          None,       (1, 2)),
    ("第3章_第1部_併設保険医療機関.pdf",         None,       (720, 721)),
    ("第3章_第2部_併設以外.pdf",                 None,       (722, 724)),
]


def merge_sections():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    src_kokuji = fitz.open(PDF_KOKUJI)
    src_tsuchi = fitz.open(PDF_TSUCHI)

    for filename, kokuji_range, tsuchi_range in MERGES:
        dst = fitz.open()

        # 告示ページを挿入
        if kokuji_range:
            p0, p1 = kokuji_range[0] - 1, kokuji_range[1] - 1
            dst.insert_pdf(src_kokuji, from_page=p0, to_page=p1)
            k_pages = kokuji_range[1] - kokuji_range[0] + 1
        else:
            k_pages = 0

        # 通知ページを挿入
        if tsuchi_range:
            p0, p1 = tsuchi_range[0] - 1, tsuchi_range[1] - 1
            dst.insert_pdf(src_tsuchi, from_page=p0, to_page=p1)
            t_pages = tsuchi_range[1] - tsuchi_range[0] + 1
        else:
            t_pages = 0

        out_path = os.path.join(OUTPUT_DIR, filename)
        dst.save(out_path)
        dst.close()

        # 構成表示
        parts = []
        if k_pages:
            parts.append(f"告示{k_pages}p")
        if t_pages:
            parts.append(f"通知{t_pages}p")
        total = k_pages + t_pages
        print(f"  {filename}  ({'+'.join(parts)} = {total}ページ)")

    src_kokuji.close()
    src_tsuchi.close()


def verify(out_dir):
    ok = True
    for f in sorted(os.listdir(out_dir)):
        if not f.endswith(".pdf"):
            continue
        path = os.path.join(out_dir, f)
        try:
            doc = fitz.open(path)
            if doc.page_count == 0:
                print(f"  [NG] {f}: 0ページ")
                ok = False
            doc.close()
        except Exception as e:
            print(f"  [NG] {f}: {e}")
            ok = False
    return ok


def main():
    print("=" * 60)
    print("部別 告示・通知統合を開始します")
    print("=" * 60)
    print()

    merge_sections()

    print()
    print("検証中...")
    if verify(OUTPUT_DIR):
        print("全ファイル正常です。")
    else:
        print("[警告] 一部ファイルに問題があります。")

    files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf")]
    print(f"\n合計 {len(files)} ファイルを出力しました → {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
