"""
NotebookLM 用に全ソースを50ページ以下に分割するスクリプト。

ノートブック2つ構成:
  NB1: 点数表（告示+通知統合）
  NB2: 施設基準（第7号+第8号）

50ページを超えるPDFは自動的に ~50p のチャンクに分割する。
50ページ以下のPDFはそのままコピーする。
"""

import os
import shutil
import fitz  # PyMuPDF

BASE = os.path.dirname(os.path.abspath(__file__))
NB_BASE = os.path.join(BASE, "NotebookLM")
MAX_PAGES = 50

# ───────────────────────────────────────────
# NB1: 点数表（告示+通知統合）
# ───────────────────────────────────────────
NB1_SRC = os.path.join(BASE, "分割", "医科診療報酬点数表_告示通知統合")
NB1_DST = os.path.join(NB_BASE, "NB1_点数表")

# ───────────────────────────────────────────
# NB2: 施設基準（第7号+第8号）
#   除外: 届出書、歯科4、調剤1
# ───────────────────────────────────────────
NB2_SRC_7 = os.path.join(BASE, "分割", "第7号_基本診療料施設基準")
NB2_SRC_8 = os.path.join(BASE, "分割", "第8号_特掲診療料施設基準")
NB2_DST = os.path.join(NB_BASE, "NB2_施設基準")

# NB2 除外キーワード
NB2_EXCLUDE = [
    "届出書",
    "歯科",
    "調剤",
]


def split_pdf(src_path, dst_dir, max_pages=MAX_PAGES):
    """PDFを max_pages 以下のチャンクに分割。小さければそのままコピー。"""
    doc = fitz.open(src_path)
    total = len(doc)
    basename = os.path.splitext(os.path.basename(src_path))[0]

    if total <= max_pages:
        # そのままコピー
        dst_path = os.path.join(dst_dir, os.path.basename(src_path))
        shutil.copy2(src_path, dst_path)
        doc.close()
        return [(os.path.basename(src_path), total)]

    # チャンク数を計算（できるだけ均等に）
    n_chunks = (total + max_pages - 1) // max_pages
    chunk_size = total // n_chunks
    remainder = total % n_chunks

    results = []
    start = 0
    for i in range(n_chunks):
        # 余りを前のチャンクに1ページずつ配分
        size = chunk_size + (1 if i < remainder else 0)
        end = start + size - 1

        chunk_name = f"{basename}_({i+1}of{n_chunks}).pdf"
        chunk_path = os.path.join(dst_dir, chunk_name)

        chunk_doc = fitz.open()
        chunk_doc.insert_pdf(doc, from_page=start, to_page=end)
        chunk_doc.save(chunk_path)
        chunk_doc.close()

        results.append((chunk_name, size))
        start = end + 1

    doc.close()
    return results


def build_nb1():
    """NB1: 点数表を構築"""
    os.makedirs(NB1_DST, exist_ok=True)

    files = sorted(f for f in os.listdir(NB1_SRC) if f.endswith(".pdf"))
    total_sources = 0

    print("=" * 60)
    print("NB1: 点数表（告示+通知統合）")
    print("=" * 60)

    for f in files:
        src = os.path.join(NB1_SRC, f)
        results = split_pdf(src, NB1_DST)
        for name, pages in results:
            split_mark = " [分割]" if len(results) > 1 else ""
            print(f"  {pages:4d}p  {name}{split_mark}")
        total_sources += len(results)

    print(f"\n  → 合計 {total_sources} ソース")
    return total_sources


def build_nb2():
    """NB2: 施設基準を構築"""
    os.makedirs(NB2_DST, exist_ok=True)

    total_sources = 0

    print()
    print("=" * 60)
    print("NB2: 施設基準（第7号+第8号）")
    print("=" * 60)

    for src_dir in [NB2_SRC_7, NB2_SRC_8]:
        for f in sorted(os.listdir(src_dir)):
            if not f.endswith(".pdf"):
                continue

            # 除外チェック
            if any(kw in f for kw in NB2_EXCLUDE):
                print(f"  [除外] {f}")
                continue

            src = os.path.join(src_dir, f)
            results = split_pdf(src, NB2_DST)
            for name, pages in results:
                split_mark = " [分割]" if len(results) > 1 else ""
                print(f"  {pages:4d}p  {name}{split_mark}")
            total_sources += len(results)

    print(f"\n  → 合計 {total_sources} ソース")
    return total_sources


def main():
    # 旧ディレクトリを削除
    for old_dir in [
        "NB1_点数表",
        "NB1_点数表_告示通知統合",
        "NB1a_基本診療料",
        "NB1b_特掲前半_医学管理〜注射",
        "NB1c_特掲後半_リハ〜手術〜病理",
        "NB2_施設基準",
        "NB2_施設基準_第7号_第8号",
    ]:
        path = os.path.join(NB_BASE, old_dir)
        if os.path.exists(path):
            shutil.rmtree(path)
            print(f"[削除] {old_dir}/")

    print()

    n1 = build_nb1()
    n2 = build_nb2()

    print()
    print("=" * 60)
    print(f"完了: NB1={n1}ソース, NB2={n2}ソース")
    print(f"出力先: {NB_BASE}")
    print("=" * 60)


if __name__ == "__main__":
    main()
