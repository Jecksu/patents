"""把申请正文中的纯文本符号切分为正文、下标和上标片段。

申请正文按硬约束以纯文本书写符号（`R_11`、`U_A^abn`），不使用反引号与
Markdown 表格。本模块只在生成 DOCX 时把这些记号还原为真正的上下标排版，
不改动 Markdown 源文件，因此此前对源文件做过的全部扫描与复算依然成立。

同时把拼写出来的希腊字母名还原为字形（``tau_cons`` 排为 τ_cons），使其与
正文中已经直接使用的 κ 一致。

判定规则只有两条：

1. 记号必须以数学基名开头。基名是单个拉丁字母、单个希腊字母，或下面
   ``MULTI_CHAR_BASES`` 中列出的多字母基名。``CLOSURE_PASS``、
   ``COMPLETE_FAIL_A``、``closure_snapshot_id``、``candidate_edge_ratio``
   这类状态名与字段名的下划线两侧都是多字母词，不在白名单内，保持平排。
2. 基名前一个字符不得是词字符。这条保证 ``CAD_DIFF_EVENT_V1`` 里的
   ``F_E``、``local_brep_byte_ratio`` 里的 ``l_b`` 不会被误认成符号。
"""

from __future__ import annotations

import re


#: 正文按硬约束把希腊字母拼写为拉丁词，但键对下标 κ 已经是真希腊字母，
#: 于是同一个式子会写成 ``q_κ`` 与 ``tau_cons`` 两种字形。这里统一还原为
#: 希腊字母，正文仍保持纯文本。只收录正文中确已定义为符号的名字。
GREEK_NAMES = {
    "epsilon": "ε",
    "lambda": "λ",
    "alpha": "α",
    "gamma": "γ",
    "delta": "δ",
    "beta": "β",
    "tau": "τ",
    "eps": "ε",
    "rho": "ρ",
    "Phi": "Φ",
    "Pi": "Π",
}

#: 词边界：前一个字符不是标识符字符，后一个字符不是字母或数字。``_`` 与 ``^``
#: 允许紧跟，因为它们开启上下标；这条同时保证 ``eps`` 不会命中 ``epsilon``。
_GREEK_RE = re.compile(
    r"(?<![A-Za-z0-9_])(" + "|".join(sorted(GREEK_NAMES, key=len, reverse=True)) + r")(?![A-Za-z0-9])"
)

#: 单字符希腊基名，替换后可直接作为符号基名参与上下标切分。
_GREEK_CHARS = "".join(sorted(set(GREEK_NAMES.values()))) + "κ"

#: 允许作为基名的多字母记号：端点展平算子 ``end`` 与闭包快照 ``cs``；
#: 其余多字母词一律按普通标识符处理，保持平排。
MULTI_CHAR_BASES = ("end", "cs")

#: 可出现在上下标内部的字符：拉丁字母、数字、κ 和映射方向箭头。
_SCRIPT_RUN = re.compile(r"[A-Za-z0-9κ→]+")
_WORD_CHAR = re.compile(r"[A-Za-z0-9_]")

#: 比较运算符按排版惯例替换，避免 DOCX 里出现 ASCII 拼写。
_OPERATORS = (("<=", "≤"), (">=", "≥"))

Chunk = tuple[str, str | None]


def _base_length(text: str, i: int) -> int:
    """返回位置 i 处基名的长度，不是基名时返回 0。"""
    for base in MULTI_CHAR_BASES:
        if text.startswith(base, i):
            return len(base)
    ch = text[i]
    if ch in _GREEK_CHARS or (ch.isascii() and ch.isalpha()):
        return 1
    return 0


def _read_script(text: str, i: int) -> tuple[str, int]:
    """从 ``_`` 或 ``^`` 之后读取一段上下标内容，返回内容与新位置。"""
    if i < len(text) and text[i] == "+" and text[i - 1] == "^":
        # ``J_κ^+`` 的上标恰为一个加号；其余位置的加号是普通运算符。
        return "+", i + 1
    m = _SCRIPT_RUN.match(text, i)
    if not m:
        return "", i
    content, i = m.group(0), m.end()
    # ``C_cov,X``、``M_edge,valid``、``B_edge,A→B^l`` 的下标含逗号，逗号左边
    # 总是一个多字母词根（cov、edge、local、mapcov）；而 ``a_1,b_1``、
    # ``MANY_TO_ONE(A_h,b)``、``sum(lambda_i*d_i,i∈J)`` 的逗号是符号或实参
    # 之间的分隔符，逗号左边只是单个下标索引。因此只有词根下标才跨逗号续读，
    # 纯数字下标（``R_11,R_1m``）同样不续读。
    while i < len(text) and text[i] == ",":
        if len(content) < 2 or content.isdigit():
            break
        nxt = _SCRIPT_RUN.match(text, i + 1)
        if not nxt or (nxt.end() < len(text) and text[nxt.end()] == "_"):
            break
        content += "," + nxt.group(0)
        i = nxt.end()
    return content, i


def _read_symbol(text: str, i: int) -> tuple[list[Chunk], int] | None:
    """尝试在位置 i 读取一个符号，失败时返回 None。"""
    if i > 0 and _WORD_CHAR.match(text[i - 1]):
        return None
    length = _base_length(text, i)
    if not length:
        return None
    j = i + length
    if j >= len(text) or text[j] not in "_^":
        return None
    chunks: list[Chunk] = [(text[i:j], None)]
    while j < len(text) and text[j] in "_^":
        script = "sub" if text[j] == "_" else "sup"
        content, end = _read_script(text, j + 1)
        if not content:
            break
        chunks.append((content, script))
        j = end
    if len(chunks) == 1:
        return None
    return chunks, j


def split_math(text: str) -> list[Chunk]:
    """把一段文字切分为 ``(片段, script)``，script 取 None、``sub`` 或 ``sup``。"""
    for old, new in _OPERATORS:
        text = text.replace(old, new)
    # 先把拼写出来的希腊字母换成字形，``delta``、``rho`` 这类不带上下标的
    # 裸用法也一并换掉，否则会出现 δ_margin 与 delta 并存。
    text = _GREEK_RE.sub(lambda m: GREEK_NAMES[m.group(1)], text)

    chunks: list[Chunk] = []
    plain: list[str] = []
    i = 0
    while i < len(text):
        found = _read_symbol(text, i)
        if found is None:
            plain.append(text[i])
            i += 1
            continue
        if plain:
            chunks.append(("".join(plain), None))
            plain = []
        symbol, i = found
        chunks.extend(symbol)
    if plain:
        chunks.append(("".join(plain), None))
    return chunks or [(text, None)]


def add_math_runs(paragraph, text: str, style_run) -> None:
    """按 ``split_math`` 的切分结果向段落追加 run，并逐个套用 ``style_run``。"""
    for chunk, script in split_math(text):
        run = paragraph.add_run(chunk)
        style_run(run)
        if script == "sub":
            run.font.subscript = True
        elif script == "sup":
            run.font.superscript = True
