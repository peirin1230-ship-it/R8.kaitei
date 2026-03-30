"""
NotebookLM用に大きい統合PDFを告示/通知に分割するスクリプト。

告示+通知の統合PDFのうち、100ページを超えるものを
告示部分と通知部分に分割して NB1 フォルダに出力する。
"""

import os
import fitz  # PyMuPDF

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(BASE_DIR, "..", "..")
GENMON_DIR = os.path.join(REPO_ROOT, "原文")
SRC_DIR = os.path.join(GENMON_DIR, "分割", "医科診療報酬点数表_告示通知統合")
NB1_DIR = os.path.join(GENMON_DIR, "NotebookLM", "NB1_点数表_告示通知統合")

# 分割対象: (元ファイル名, 告示ページ数)
# merge_by_bu.py の MERGES 定義から算出
SPLITS = [
    ("第1章_第2部_入院料等.pdf", 82),       # 210p → 告示82p + 通知128p
    ("第2章_第1部_医学管理等.pdf", 43),      # 135p → 告示43p + 通知92p
    ("第2章_第2部_在宅医療.pdf", 39),        # 121p → 告示39p + 通知82p
    ("第2章_第3部_検査.pdf", 39),            # 163p → 告示39p + 通知124p
    ("第2章_第10部_手術.pdf", 87),           # 186p → 告示87p + 通知99p
]


def main():
    os.makedirs(NB1_DIR, exist_ok=True)

    for filename, kokuji_pages in SPLITS:
        src_path = os.path.join(SRC_DIR, filename)
        if not os.path.exists(src_path):
            print(f"[スキップ] {filename} が見つかりません")
            continue

        src = fitz.open(src_path)
        total = len(src)
        tsuchi_pages = total - kokuji_pages
        base = filename.replace(".pdf", "")

        # 告示部分
        kokuji_name = f"{base}_告示.pdf"
        dst_k = fitz.open()
        dst_k.insert_pdf(src, from_page=0, to_page=kokuji_pages - 1)
        dst_k.save(os.path.join(NB1_DIR, kokuji_name))
        dst_k.close()

        # 通知部分
        tsuchi_name = f"{base}_通知.pdf"
        dst_t = fitz.open()
        dst_t.insert_pdf(src, from_page=kokuji_pages, to_page=total - 1)
        dst_t.save(os.path.join(NB1_DIR, tsuchi_name))
        dst_t.close()

        # 元の統合ファイルを削除
        merged_in_nb1 = os.path.join(NB1_DIR, filename)
        if os.path.exists(merged_in_nb1):
            os.remove(merged_in_nb1)
            print(f"  [削除] {filename}")

        src.close()
        print(f"  [分割] {filename} ({total}p) → {kokuji_name} ({kokuji_pages}p) + {tsuchi_name} ({tsuchi_pages}p)")

    # 最終確認
    print()
    print("=== NB1 最終ファイル一覧 ===")
    files = sorted(f for f in os.listdir(NB1_DIR) if f.endswith(".pdf"))
    for f in files:
        doc = fitz.open(os.path.join(NB1_DIR, f))
        print(f"  {len(doc):4d}p  {f}")
        doc.close()
    print(f"\n合計 {len(files)} ファイル")


if __name__ == "__main__":
    main()
