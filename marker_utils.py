"""順番ラベル（マーカー）の検出・フィルタリング、および
テキスト正規化のユーティリティ。

対象マーカー:
  - 括弧付き数字: （１）（２）... / (1)(2)...
  - カタカナ1文字: ア イ ウ ... （五十音・いろは両方）
  - 全角小文字アルファベット: ａ ｂ ｃ ...
  - 裸の数字: １ ２ ３ ... / 1 2 3 ...（注の継続番号等）

除外:
  - 「注N」形式（注１ 注２ 等）は構造ラベルのため対象外
"""

import re
import unicodedata

# 順番ラベルのパターン
RE_MARKER_PAREN_FULL = re.compile(r'^([（(][０-９0-9]+[）)])\s*')
RE_MARKER_KANA = re.compile(
    r'^([アイウエオカキクケコサシスセソタチツテト'
    r'ナニヌネノハヒフヘホマミムメモヤユヨ'
    r'ラリルレロワヲン])[\s\u3000]')
RE_MARKER_ALPHA = re.compile(r'^([ａ-ｚ])[\s\u3000]')
# 裸の数字（「注N」は除外。数字+空白+日本語テキストのパターンのみ対象）
RE_MARKER_BARE_NUM = re.compile(r'^([０-９0-9]+)[\s\u3000]')
# 「注N」パターン（除外判定用）
RE_NOTE_PREFIX = re.compile(r'^注[０-９0-9]')

_MARKER_PATTERNS = [RE_MARKER_PAREN_FULL, RE_MARKER_KANA, RE_MARKER_ALPHA,
                    RE_MARKER_BARE_NUM]


def strip_sequence_marker(text):
    """行頭の順番ラベル（マーカー）を分離する。

    「注N」形式（注１ 注２ 等）は構造ラベルのため対象外とする。

    Returns:
        (marker, body): マーカー文字列と残りの本文。
                        マーカーがない場合は ('', text)。
    """
    if not text:
        return '', text

    # 「注N」形式は除外
    if RE_NOTE_PREFIX.match(text):
        return '', text

    for pattern in _MARKER_PATTERNS:
        m = pattern.match(text)
        if m:
            marker = m.group(1)
            body = text[m.end():]
            return marker, body

    return '', text


def normalize_note_refs(text):
    """文章内の注番号参照（「注N」「注１」等）の番号部分をプレースホルダに置換する。

    注の繰り上げ/繰り下げにより文中の「注8」→「注7」のような変更は
    実質的な内容変更ではないため、比較時に吸収する。
    """
    # 「注N」「注１２」等のパターン → 「注_」に統一
    return re.sub(r'注[０-９0-9]+', '注_', text)


def normalize_for_content_compare(text, normalize_fn):
    """内容比較用にテキストを正規化する。

    先頭マーカー除去 + 文中の注番号参照の正規化 + テキスト正規化。
    番号のみ変更（先頭マーカー・文中注参照）を吸収する。
    """
    _, body = strip_sequence_marker(text)
    normalized = normalize_fn(body)
    return normalize_note_refs(normalized)


def is_marker_only_change(r8_text, r6_text, normalize_fn):
    """2つのテキストの差異がマーカー・注番号参照の変更のみかどうかを判定する。

    先頭マーカーの変更と、文中の注番号参照（「注8」→「注7」等）の変更を
    吸収した上で、実質的な内容が同一かどうかを判定する。

    Args:
        r8_text: 改正後テキスト
        r6_text: 改正前テキスト
        normalize_fn: テキスト正規化関数（各スクリプトの normalize_text_for_compare）

    Returns:
        (is_marker_only, r8_marker, r6_marker, r8_body, r6_body)
    """
    r8_marker, r8_body = strip_sequence_marker(r8_text)
    r6_marker, r6_body = strip_sequence_marker(r6_text)

    # マーカーが両方あり異なる、または文中の注番号参照が異なる場合を吸収
    r8_content = normalize_note_refs(normalize_fn(r8_body))
    r6_content = normalize_note_refs(normalize_fn(r6_body))

    if r8_content == r6_content:
        # 先頭マーカーが同じでも文中注番号だけ異なる場合も含む
        if (r8_marker and r6_marker and r8_marker != r6_marker) or \
           r8_content == r6_content:
            return True, r8_marker, r6_marker, r8_body, r6_body

    return False, r8_marker, r6_marker, r8_body, r6_body


# ============================================================
# 全角半角正規化
# ============================================================

def normalize_width(text):
    """全角英数字・記号を半角に、全角スペースを半角スペースに変換する。

    unicodedata.normalize('NFKC') を使用して全角→半角を統一する。
    カタカナ・ひらがな・漢字等の日本語文字はそのまま保持される。
    """
    return unicodedata.normalize('NFKC', text)
