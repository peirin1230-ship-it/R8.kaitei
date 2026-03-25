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


def is_marker_only_change(r8_text, r6_text, normalize_fn):
    """2つのテキストの差異がマーカー部分のみかどうかを判定する。

    Args:
        r8_text: 改正後テキスト
        r6_text: 改正前テキスト
        normalize_fn: テキスト正規化関数（各スクリプトの normalize_text_for_compare）

    Returns:
        (is_marker_only, r8_marker, r6_marker, r8_body, r6_body)
    """
    r8_marker, r8_body = strip_sequence_marker(r8_text)
    r6_marker, r6_body = strip_sequence_marker(r6_text)

    if r8_marker and r6_marker and r8_marker != r6_marker:
        r8_body_norm = normalize_fn(r8_body)
        r6_body_norm = normalize_fn(r6_body)
        if r8_body_norm == r6_body_norm:
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
