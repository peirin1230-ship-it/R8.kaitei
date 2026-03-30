"""
NB1（25ソース）を3分割して各ノートブックのソース数を減らすスクリプト。

25ソースでは NotebookLM の RAG 精度が低いため、
以下の3つに分割する：

NB1a: 基本診療料（第1章）+ 通則・目次 — 5ファイル
NB1b: 特掲診療料 前半（第1〜6部）— 10ファイル
NB1c: 特掲診療料 後半（第7〜14部）+ 第3章 — 10ファイル

手術（第10部）は NB1c に入り、ソース数10で検索精度が上がる。
"""

import os
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(BASE, "..", "..")
GENMON_DIR = os.path.join(REPO_ROOT, "原文")
SRC = os.path.join(GENMON_DIR, "NotebookLM", "NB1_点数表_告示通知統合")
NB_BASE = os.path.join(GENMON_DIR, "NotebookLM")

# 分割定義: フォルダ名 → ファイル名に含まれるキーワードのリスト
NOTEBOOKS = {
    "NB1a_基本診療料": [
        "告示69号_目次",
        "第6号_通則",
        "第1章_第1部_初再診料",
        "第1章_第2部_入院料等_告示",
        "第1章_第2部_入院料等_通知",
    ],
    "NB1b_特掲前半_医学管理〜注射": [
        "第2章_第1部_医学管理等_告示",
        "第2章_第1部_医学管理等_通知",
        "第2章_第2部_在宅医療_告示",
        "第2章_第2部_在宅医療_通知",
        "第2章_第3部_検査_告示",
        "第2章_第3部_検査_通知",
        "第2章_第4部_画像診断",
        "第2章_第5部_投薬",
        "第2章_第6部_注射",
    ],
    "NB1c_特掲後半_リハ〜手術〜病理": [
        "第2章_第7部_リハビリテーション",
        "第2章_第8部_精神科専門療法",
        "第2章_第9部_処置",
        "第2章_第10部_手術_告示",
        "第2章_第10部_手術_通知",
        "第2章_第11部_麻酔",
        "第2章_第12部_放射線治療",
        "第2章_第13部_病理診断",
        "第2章_第14部_その他",
        "第3章_第1部_併設保険医療機関",
        "第3章_第2部_併設以外",
    ],
}


def main():
    # ソースディレクトリの全ファイルを取得
    all_files = sorted(f for f in os.listdir(SRC) if f.endswith(".pdf"))
    print(f"ソースファイル数: {len(all_files)}")

    assigned = set()

    for nb_name, keywords in NOTEBOOKS.items():
        dst_dir = os.path.join(NB_BASE, nb_name)
        os.makedirs(dst_dir, exist_ok=True)

        matched = []
        for kw in keywords:
            # キーワードに一致するファイルを検索
            found = [f for f in all_files if f.startswith(kw) and f not in assigned]
            if not found:
                # 完全一致しない場合、部分一致を試行
                found = [f for f in all_files if kw in f and f not in assigned]
            if found:
                matched.append(found[0])
                assigned.add(found[0])
            else:
                print(f"  [警告] キーワード '{kw}' に一致するファイルなし")

        for f in matched:
            src_path = os.path.join(SRC, f)
            dst_path = os.path.join(dst_dir, f)
            shutil.copy2(src_path, dst_path)

        print(f"\n{nb_name}: {len(matched)} ファイル")
        for f in matched:
            print(f"  {f}")

    # 未割当ファイルの確認
    unassigned = [f for f in all_files if f not in assigned]
    if unassigned:
        print(f"\n[警告] 未割当ファイル:")
        for f in unassigned:
            print(f"  {f}")

    print(f"\n完了: {len(assigned)}/{len(all_files)} ファイルを割当済み")


if __name__ == "__main__":
    main()
