"""style.css の検査。

帳票の罫線が消える事故（7 章「罫線の注意」）を防ぐため:
  - 表の要素（table / tr / td / th ...）に当てる規則は、必ず .tw（画面の表）か .sheet（帳票）で囲む
  - .sheet の中で罫線を消す規則（border: 0 / none、border-*-width: 0 など）を書かない
  - :last-child / :first-child など「場所で決まる」規則で罫線をいじらない
また、閉じた網で使うため、外部の読み込み（@import / http）や IE で動かない書き方（var(--) / grid）を止める。
"""
import re

TABLE_EL = re.compile(r"(^|[\s>+~(,])(table|thead|tbody|tfoot|tr|td|th|caption|colgroup|col)(?![\w-])", re.I)


def strip_comments(css):
    return re.sub(r"/\*.*?\*/", "", css, flags=re.S)


def rules(css):
    """(セレクタ, 宣言, @media の外側) を返す。入れ子は 1 段（@media）まで。"""
    css = strip_comments(css)
    out = []
    i = 0
    media = None
    stack = []
    buf = ""
    while i < len(css):
        c = css[i]
        if c == "{":
            sel = buf.strip()
            buf = ""
            if sel.startswith("@media") or sel.startswith("@supports"):
                stack.append(("media", sel))
            else:
                j = css.find("}", i)
                out.append((sel, css[i + 1:j], [m for _, m in stack]))
                i = j + 1
                continue
        elif c == "}":
            if stack:
                stack.pop()
            buf = ""
        else:
            buf += c
        i += 1
    return out


def check(css_text):
    probs = []
    raw = css_text
    if "@import" in raw:
        probs.append("@import は使わない決まりです（ファイルは 1 本）")
    if re.search(r"url\(\s*['\"]?https?:", raw) or re.search(r"https?://", strip_comments(raw)):
        probs.append("外部（http/https）を読み込んでいます。閉じたネットワークでは読めません")
    if "var(--" in raw:
        probs.append("CSS 変数 var(--) は IE モードで効きません")
    if re.search(r"display\s*:\s*grid", raw):
        probs.append("display: grid は IE モードで効きません")
    for sel, decl, media in rules(raw):
        if sel.startswith("@page") or sel.startswith("@font-face"):
            continue
        for one in sel.split(","):
            one = one.strip()
            if not one:
                continue
            if TABLE_EL.search(" " + one) and not re.search(r"\.(tw|sheet)\b", one):
                probs.append(f"表の要素に当てる規則が .tw / .sheet で囲まれていません: {one}（帳票側にも効いて罫線が消える原因になる）")
            if re.search(r"\.sheet\b", one):
                if re.search(r"border(-(top|bottom|left|right))?(-width|-style)?\s*:\s*(none|hidden|0(?![.\d]))", decl):
                    probs.append(f"帳票（.sheet）の中で罫線を消しています: {one} {{{decl.strip()}}}")
                if re.search(r":(last|first|nth)-(child|of-type)", one) and "border" in decl:
                    probs.append(f"帳票（.sheet）で場所により罫線を変えています: {one}")
    return probs
