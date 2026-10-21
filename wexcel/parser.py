"""容错的 JSON 解析器。

目标：让人工手写的「类 JSON」数据也能解析成功，而不必严格符合 JSON 规范。

支持（比标准 JSON 更宽松）：
  * 单引号、双引号、中文引号 “ ” ‘ ’ 都可以包字符串；
  * 键名可以不加引号，例如 {l:'0-0', v:'内容'} ；
  * 漏写冒号也能识别，例如 {l'0-1', v:'内容2'} ；
  * 漏写逗号也能识别，例如 {l:'0-0' v:'内容'} ；（在多数情况下）
  * 结尾多余的逗号、分号（; ），也允许；
  * 整段使用中文全角标点：｛｝［］：，； 都可以；
  * 支持 // 行注释 和 /* 块注释 */；
  * true/false/null/none 等常量，未加引号的内容视为字符串；
  * 支持 \\uXXXX 转义。

注意：中文全角逗号、冒号只在「结构位置」被当作分隔符；
字符串内部的中文标点会原样保留（例如 '内容，测试' 不会被拆开）。
"""

from __future__ import annotations

import re
import string
from typing import Any

__all__ = ["ParseError", "loads"]

# 表示「键名后面没有写值」的哨兵对象（区别于用户显式写的 null）
_MISSING = object()

# ---------------------------------------------------------------- 字符常量

_OPEN_OBJ = "{｛"
_CLOSE_OBJ = "}｝"
_OPEN_ARR = "[［"
_CLOSE_ARR = "]］"
_COLON = ":："
_COMMA = ",，;；"

# 引号配对：开 -> 可用的闭引号集合
_QUOTE_FAMILY = {
    '"': {'"', '”'},
    "'": {"'", '’'},
    '\u201c': {'"', '“', '”'},  # “
    '\u2018': {"'", '‘', '’'},  # ‘
}
_QUOTE_CHARS = set(_QUOTE_FAMILY) | {"”", "’"}

# 裸值（未加引号）遇到这些字符就结束
_BARE_STOP = set(_COMMA) | set(_COLON) | set(_CLOSE_OBJ) | set(_CLOSE_ARR)

_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")

_CONSTANTS = {
    "true": True,
    "false": False,
    "null": None,
    "none": None,
    "nil": None,
    "undefined": None,
}

_ESCAPES = {
    "n": "\n",
    "t": "\t",
    "r": "\r",
    "b": "\b",
    "f": "\f",
    "v": "\v",
    "0": "\0",
    '"': '"',
    "'": "'",
    "\\": "\\",
    "/": "/",
    "“": "“",
    "”": "”",
    "‘": "‘",
    "’": "’",
}


class ParseError(ValueError):
    """解析失败。message 是中文提示，附带行列位置。"""

    def __init__(self, message: str, text: str = "", pos: int = 0):
        self.message = message
        self.pos = max(0, min(pos, len(text)))
        self.text = text
        super().__init__(self._format())

    def _format(self) -> str:
        text, pos = self.text, self.pos
        line = text.count("\n", 0, pos) + 1
        line_start = text.rfind("\n", 0, pos) + 1
        col = pos - line_start + 1
        left = text[max(0, pos - 20): pos]
        right = text[pos: pos + 20]
        snippet = (left + right).replace("\n", " ").replace("\t", " ")
        return f"{self.message}（第 {line} 行第 {col} 列，附近：…{snippet}…）"


class _Parser:
    def __init__(self, text: str):
        self.text = text
        self.i = 0
        self.n = len(text)

    # ---------------------------------------------------------- 基础工具

    def fail(self, message: str):
        raise ParseError(message, self.text, self.i)

    def skip(self):
        """跳过空白和注释。"""
        text, n = self.text, self.n
        while self.i < n:
            ch = text[self.i]
            if ch.isspace() or ch == "\ufeff" or ch == "\u200b":
                self.i += 1
            elif text.startswith("//", self.i):
                j = text.find("\n", self.i)
                self.i = n if j < 0 else j + 1
            elif text.startswith("/*", self.i):
                j = text.find("*/", self.i + 2)
                if j < 0:
                    self.fail("块注释 /* 没有对应的 */")
                self.i = j + 2
            else:
                break

    def peek(self) -> str:
        return self.text[self.i] if self.i < self.n else ""

    # ---------------------------------------------------------- 文档入口

    def parse_document(self) -> Any:
        self.skip()
        if self.i >= self.n:
            self.fail("数据为空（txt 文件中没有可解析的内容）")
        value = self.parse_value()
        self.skip()
        # 容忍结尾多余的分隔符
        while self.i < self.n and self.peek() in _COMMA:
            self.i += 1
            self.skip()
        if self.i < self.n:
            self.fail("数据结束后还有多余内容")
        return value

    # ---------------------------------------------------------- 值

    def parse_value(self) -> Any:
        self.skip()
        if self.i >= self.n:
            self.fail("缺少内容")
        ch = self.peek()
        if ch in _OPEN_OBJ:
            return self.parse_object()
        if ch in _OPEN_ARR:
            return self.parse_array()
        if ch in _QUOTE_CHARS:
            return self.parse_string()
        if ch in _CLOSE_OBJ or ch in _CLOSE_ARR:
            self.fail(f"出现了不匹配的结束符号 {ch!r}")
        return self.parse_bare()

    def parse_object(self) -> dict:
        self.i += 1  # 吃掉 {
        obj: dict[str, Any] = {}
        while True:
            self.skip()
            if self.i >= self.n:
                self.fail("对象没有结束，缺少 }")
            ch = self.peek()
            if ch in _CLOSE_OBJ:
                self.i += 1
                return obj
            if ch in _COMMA:
                # 多余逗号，例如 {, l:'0-0'}
                self.i += 1
                continue

            key = self.parse_key()
            self.skip()
            if self.i < self.n and self.peek() in _COLON:
                self.i += 1
                self.skip()
                value = self.parse_value()
            elif self.i < self.n and self.peek() not in _CLOSE_OBJ and self.peek() not in _COMMA:
                # 漏写冒号：{l'0-1', v:'内容2'}
                value = self.parse_value()
            else:
                # 键后面什么都没有
                value = _MISSING

            # 后面的重复键以最后一次为准（方便覆盖）
            obj[key] = value

            self.skip()
            if self.i >= self.n:
                self.fail("对象没有结束，缺少 }")
            ch = self.peek()
            if ch in _COMMA:
                self.i += 1
                continue
            if ch in _CLOSE_OBJ:
                self.i += 1
                return obj
            if ch in _OPEN_OBJ or ch in _QUOTE_CHARS or ch not in _BARE_STOP:
                # 漏写逗号，继续解析下一对
                if ch in _COLON:
                    self.fail("多余的分隔符 :（这里应该写逗号）")
                continue
            self.fail(f"对象中出现意外的字符 {ch!r}")

    def parse_array(self) -> list:
        self.i += 1  # 吃掉 [
        items: list[Any] = []
        while True:
            self.skip()
            if self.i >= self.n:
                self.fail("数组没有结束，缺少 ]")
            ch = self.peek()
            if ch in _CLOSE_ARR:
                self.i += 1
                return items
            if ch in _COMMA:
                # 多余逗号，例如 [,, 1]
                self.i += 1
                continue
            items.append(self.parse_value())
            self.skip()
            if self.i >= self.n:
                self.fail("数组没有结束，缺少 ]")
            ch = self.peek()
            if ch in _COMMA:
                self.i += 1
                continue
            if ch in _CLOSE_ARR:
                self.i += 1
                return items
            if ch in _COLON or ch in _CLOSE_OBJ:
                self.fail(f"数组中出现了意外的字符 {ch!r}")
            # 漏写逗号，继续读下一个元素
            continue

    def parse_key(self) -> str:
        self.skip()
        if self.i >= self.n:
            self.fail("缺少键名")
        ch = self.peek()
        if ch in _QUOTE_CHARS:
            return self.parse_string()
        start = self.i
        while self.i < self.n:
            ch = self.text[self.i]
            if (ch in _COLON or ch in _COMMA or ch in _CLOSE_OBJ
                    or ch in _QUOTE_CHARS or ch.isspace()):
                break
            self.i += 1
        key = self.text[start:self.i]
        if not key:
            self.fail("缺少键名")
        return key

    def parse_string(self) -> str:
        opener = self.peek()
        closers = _QUOTE_FAMILY.get(opener, {opener})
        self.i += 1
        chars: list[str] = []
        while self.i < self.n:
            ch = self.text[self.i]
            if ch == "\\":
                self.i += 1
                if self.i >= self.n:
                    self.fail("字符串末尾有一个孤立的反斜杠 \\")
                esc = self.text[self.i]
                if esc == "u":
                    hex_text = self.text[self.i + 1: self.i + 5]
                    if len(hex_text) == 4 and all(c in string.hexdigits for c in hex_text):
                        chars.append(chr(int(hex_text, 16)))
                        self.i += 5
                        continue
                    chars.append("\\u")  # 不合法就原样保留
                    self.i += 1
                    continue
                if esc in _ESCAPES:
                    chars.append(_ESCAPES[esc])
                elif esc == "\n":
                    pass  # 反斜杠续行
                else:
                    chars.append("\\" + esc)  # 未知转义原样保留
                self.i += 1
                continue
            if ch in closers:
                self.i += 1
                return "".join(chars)
            chars.append(ch)
            self.i += 1
        self.fail(f"字符串没有结束（缺少右引号 {'/'.join(sorted(closers))}）")

    def parse_bare(self) -> Any:
        """未加引号的内容：常量 / 数字 / 普通字符串。"""
        start = self.i
        while self.i < self.n and self.text[self.i] not in _BARE_STOP:
            self.i += 1
        token = self.text[start:self.i].strip()
        if not token:
            self.fail("缺少内容")
        low = token.lower()
        if low in _CONSTANTS:
            return _CONSTANTS[low]
        if _NUMBER_RE.match(token):
            try:
                return int(token)
            except ValueError:
                try:
                    return float(token)
                except ValueError:
                    return token
        return token


def loads(text: str) -> Any:
    """把文本解析成 Python 对象（dict / list / str / int / float / bool / None）。"""
    if not isinstance(text, str):
        raise TypeError("loads() 需要字符串")
    if text.startswith("\ufeff"):
        text = text[1:]
    # 去掉 NUL 等不可见控制字符（常见于从 Excel / 记事本另存的文件）
    text = text.replace("\x00", "")
    return _Parser(text).parse_document()
