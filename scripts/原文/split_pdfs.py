"""
診療報酬関連PDFを章・部・別添単位で分割するスクリプト。

使い方:
    python split_pdfs.py

入力: 原文/ 直下の4つのPDF（告示69号・第6号・第7号・第8号）
出力: 原文/分割/ 配下にサブフォルダごとに分割PDF
"""

import os
import fitz  # PyMuPDF

# ---------------------------------------------------------------------------
# 設定: 入力ファイル名と出力フォルダ
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(BASE_DIR, "..", "..")
GENMON_DIR = os.path.join(REPO_ROOT, "原文")
OUTPUT_ROOT = os.path.join(GENMON_DIR, "分割")

# 入力PDFファイル名
PDF_69 = "診療報酬の算定方法の一部を改正する件（令和８年厚生労働省告示第69号）医科点数表.pdf"
PDF_6 = "診療報酬の算定方法の一部改正に伴う実施上の留意事項について（通知）（令和８年３月５日保医発0305第６号）医科診療報酬点数表.pdf"
PDF_7 = "基本診療料の施設基準等及びその届出に関する手続きの取扱いについて（令和８年３月５日保医発0305第７号）.pdf"
PDF_8 = "特掲診療料の施設基準等及びその届出に関する手続きの取扱いについて（令和８年３月５日保医発0305第８号）.pdf"

# ---------------------------------------------------------------------------
# 分割定義: (出力ファイル名, 開始ページ1始, 終了ページ1始)
# ---------------------------------------------------------------------------

# 告示第69号（医科点数表・400ページ）→ 章・部 で17分割
SPLITS_69 = [
    ("告示69号_目次.pdf", 1, 2),
    ("告示69号_第1章_第1部_初再診料.pdf", 3, 9),
    ("告示69号_第1章_第2部_入院料等.pdf", 10, 91),
    ("告示69号_第2章_第1部_医学管理等.pdf", 92, 134),
    ("告示69号_第2章_第2部_在宅医療.pdf", 135, 173),
    ("告示69号_第2章_第3部_検査.pdf", 174, 212),
    ("告示69号_第2章_第4部_画像診断.pdf", 213, 220),
    ("告示69号_第2章_第5部_投薬.pdf", 221, 224),
    ("告示69号_第2章_第6部_注射.pdf", 225, 228),
    ("告示69号_第2章_第7部_リハビリテーション.pdf", 229, 237),
    ("告示69号_第2章_第8部_精神科専門療法.pdf", 238, 256),
    ("告示69号_第2章_第9部_処置.pdf", 257, 271),
    ("告示69号_第2章_第10部_手術.pdf", 272, 358),
    ("告示69号_第2章_第11部_麻酔.pdf", 359, 364),
    ("告示69号_第2章_第12部_放射線治療.pdf", 365, 368),
    ("告示69号_第2章_第13部_病理診断.pdf", 369, 370),
    ("告示69号_第2章_第14部_その他・処遇改善・ベースアップ.pdf", 371, 400),
]

SPLITS_6 = [
    ("第6号_通則.pdf", 1, 2),
    ("第6号_第1章_第1部_初再診料.pdf", 3, 22),
    ("第6号_第1章_第2部_入院料等.pdf", 23, 150),
    ("第6号_第2章_第1部_医学管理等.pdf", 151, 242),
    ("第6号_第2章_第2部_在宅医療.pdf", 243, 324),
    ("第6号_第2章_第3部_検査.pdf", 325, 448),
    ("第6号_第2章_第4部_画像診断.pdf", 449, 469),
    ("第6号_第2章_第5部_投薬.pdf", 470, 479),
    ("第6号_第2章_第6部_注射.pdf", 480, 488),
    ("第6号_第2章_第7部_リハビリテーション.pdf", 489, 514),
    ("第6号_第2章_第8部_精神科専門療法.pdf", 515, 551),
    ("第6号_第2章_第9部_処置.pdf", 552, 591),
    ("第6号_第2章_第10部_手術.pdf", 592, 690),
    ("第6号_第2章_第11部_麻酔.pdf", 691, 701),
    ("第6号_第2章_第12部_放射線治療.pdf", 702, 710),
    ("第6号_第2章_第13部_病理診断.pdf", 711, 717),
    ("第6号_第2章_第14部_その他.pdf", 718, 719),
    ("第6号_第3章_第1部_併設保険医療機関.pdf", 720, 721),
    ("第6号_第3章_第2部_併設以外.pdf", 722, 724),
]

SPLITS_7 = [
    ("第7号_本文.pdf", 1, 15),
    ("第7号_別添1_初再診料の施設基準等.pdf", 16, 35),
    ("第7号_別添2_入院基本料等の施設基準等.pdf", 36, 77),
    ("第7号_別添3_入院基本料等加算の施設基準等.pdf", 78, 201),
    ("第7号_別添4_特定入院料の施設基準等.pdf", 202, 296),
    ("第7号_別添5_短期滞在手術等基本料の施設基準等.pdf", 297, 298),
    ("第7号_別添6_別紙.pdf", 299, 466),
    ("第7号_別添7_届出書.pdf", 467, 736),
]

SPLITS_8 = [
    ("第8号_本文.pdf", 1, 43),
    ("第8号_別添1_医学管理等.pdf", 44, 103),
    ("第8号_別添1_歯科医学管理.pdf", 104, 110),
    ("第8号_別添1_在宅医療.pdf", 111, 139),
    ("第8号_別添1_歯科訪問.pdf", 139, 140),
    ("第8号_別添1_検査.pdf", 140, 163),
    ("第8号_別添1_画像診断・投薬・注射.pdf", 164, 175),
    ("第8号_別添1_リハビリテーション.pdf", 176, 201),
    ("第8号_別添1_精神科専門療法.pdf", 202, 220),
    ("第8号_別添1_処置.pdf", 221, 233),
    ("第8号_別添1_歯科固有技術.pdf", 234, 237),
    ("第8号_別添1_手術.pdf", 237, 333),
    ("第8号_別添1_麻酔・放射線・病理.pdf", 334, 357),
    ("第8号_別添1_歯科矯正等.pdf", 358, 358),
    ("第8号_別添1_調剤.pdf", 358, 430),
    ("第8号_別添2_届出書.pdf", 431, 1013),
    ("第8号_補則.pdf", 1014, 1021),
]


def split_pdf(src_path: str, splits: list, out_dir: str) -> None:
    """1つのPDFを分割定義に従って複数PDFに分割する。

    Args:
        src_path: 元PDFのパス
        splits: [(出力ファイル名, 開始ページ1始, 終了ページ1始), ...]
        out_dir: 出力ディレクトリ
    """
    os.makedirs(out_dir, exist_ok=True)
    src = fitz.open(src_path)
    total = src.page_count

    for filename, start, end in splits:
        # ページ番号を0始まりに変換
        p0 = start - 1
        p1 = end - 1

        if p0 < 0 or p1 >= total:
            print(f"  [警告] {filename}: ページ範囲 {start}-{end} が元PDF({total}ページ)を超えています。スキップ。")
            continue

        dst = fitz.open()
        dst.insert_pdf(src, from_page=p0, to_page=p1)
        out_path = os.path.join(out_dir, filename)
        dst.save(out_path)
        dst.close()
        print(f"  {filename}  ({end - start + 1}ページ)")

    src.close()


def verify_outputs(out_dir: str) -> bool:
    """分割されたPDFが正常に開けるか検証する。"""
    ok = True
    for root, _dirs, files in os.walk(out_dir):
        for f in sorted(files):
            if not f.endswith(".pdf"):
                continue
            path = os.path.join(root, f)
            try:
                doc = fitz.open(path)
                page_count = doc.page_count
                doc.close()
                if page_count == 0:
                    print(f"  [NG] {f}: 0ページ")
                    ok = False
            except Exception as e:
                print(f"  [NG] {f}: {e}")
                ok = False
    return ok


def main():
    print("=" * 60)
    print("PDF分割を開始します")
    print("=" * 60)

    tasks = [
        (PDF_69, SPLITS_69, "告示69号_医科点数表"),
        (PDF_6, SPLITS_6, "第6号_留意事項通知"),
        (PDF_7, SPLITS_7, "第7号_基本診療料施設基準"),
        (PDF_8, SPLITS_8, "第8号_特掲診療料施設基準"),
    ]

    for pdf_name, splits, folder_name in tasks:
        src_path = os.path.join(GENMON_DIR, pdf_name)
        out_dir = os.path.join(OUTPUT_ROOT, folder_name)

        print(f"\n--- {folder_name} ({len(splits)}分割) ---")

        if not os.path.exists(src_path):
            print(f"  [エラー] 入力ファイルが見つかりません: {pdf_name}")
            continue

        split_pdf(src_path, splits, out_dir)

    # 検証
    print(f"\n{'=' * 60}")
    print("検証中...")
    if verify_outputs(OUTPUT_ROOT):
        print("全ファイル正常です。")
    else:
        print("[警告] 一部ファイルに問題があります。上記を確認してください。")

    # サマリー
    total_files = 0
    for root, _dirs, files in os.walk(OUTPUT_ROOT):
        total_files += sum(1 for f in files if f.endswith(".pdf"))
    print(f"\n合計 {total_files} ファイルを出力しました → {OUTPUT_ROOT}")


if __name__ == "__main__":
    main()
