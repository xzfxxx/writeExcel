"""把数据写入 Excel 模板（核心逻辑，不涉及命令行交互）。"""

from __future__ import annotations

import json
import os
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .parser import ParseError, _MISSING
from .parser import loads as parse_loads

__all__ = [
    "DataDocument",
    "FillError",
    "FillReport",
    "TEMPLATE_SUFFIXES",
    "normalize_document",
    "normalize_entries",
    "parse_cell_refs",
    "fill",
]

TEMPLATE_SUFFIXES = (".xlsx", ".xlsm")

# 单元格数量上限（防止把 A1:XFD1048576 这种范围写爆）
MAX_RANGE_CELLS = 200_000

_CELL_CHAR_LIMIT = 32_767  # Excel 单元格文本上限

# 位置字段的候选键名（小写比较）
_LOC_KEYS = (
    "l", "loc", "location", "cell", "cells", "pos", "position",
    "位置", "坐标", "单元格", "行-列", "行_列", "行列",
)
# 内容字段的候选键名
_VAL_KEYS = (
    "v", "val", "value", "values", "text", "content", "contents",
    "内容", "值", "数据", "文本",
)
# 工作表字段的候选键名
_SHEET_KEYS = ("s", "sh", "sheet", "worksheet", "工作表", "表", "页签")

_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
# 0-0 这种「行-列」写法，允许全角/各种横线
_RC_RE = re.compile(r"^(\d+)\s*[-–—－‐]\s*(\d+)$")
# A1 写法
_A1_RE = re.compile(r"^([A-Za-z]{1,3})(\d+)$")
# R1C1 写法（1 基）
_R1C1_RE = re.compile(r"^[Rr](\d+)[Cc](\d+)$")
# 多个位置之间的分隔符
_POS_SPLIT_RE = re.compile(r"[;；,，、|｜\s]+")

_MAX_ROW = 1_048_576
_MAX_COL = 16_384


class FillError(Exception):
    """写入失败。problems 里是给人看的中文问题列表。"""

    def __init__(self, message: str, problems: Iterable[str] | None = None):
        self.message = message
        self.problems = list(problems or [])
        super().__init__(message)


@dataclass
class FillReport:
    output: Path | None = None
    written: int = 0
    sheet_used: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    applied: list[tuple[str, int, int, Any]] = field(default_factory=list)  # (表名, 行, 列, 值)
    dry_run: bool = False


# ---------------------------------------------------------------- 数据整理

def _lookup(mapping: dict, keys: tuple[str, ...]) -> tuple[bool, Any]:
    """按候选键名找值（大小写、首尾空格都不敏感）。返回 (是否找到, 值)。"""
    for key, value in mapping.items():
        if isinstance(key, str) and key.strip().lower() in keys:
            return True, value
    return False, None


_DATA_WRAPPER_KEYS = ("data", "datas", "rows", "cells", "list", "items", "数据")


def _unwrap_string(raw: Any) -> Any:
    """容忍整个数据被一对引号包起来："[{l:'0-0',v:'内容'}]"。"""
    for _ in range(2):
        if not isinstance(raw, str):
            break
        text = raw.strip()
        if not text:
            break
        try:
            reparsed = parse_loads(text)
        except ParseError:
            break
        if isinstance(reparsed, str) and reparsed.strip() == text:
            break
        raw = reparsed
    return raw


def _has_data_wrapper(raw: dict) -> bool:
    """判断是不是 {"data": [...]} 这种带外层包装的写法。"""
    for key, value in raw.items():
        if isinstance(key, str) and key.strip().lower() in _DATA_WRAPPER_KEYS:
            if isinstance(value, list):
                return True
    return False


def _looks_like_entry(mapping: Any) -> bool:
    if not isinstance(mapping, dict):
        return False
    found, _ = _lookup(mapping, _LOC_KEYS)
    return found


def _top_level_items(raw: Any) -> list:
    """把「顶层列表 / {data:[...]} / 单条 {l,v}」统一成条目列表。"""
    raw = _unwrap_string(raw)

    if isinstance(raw, str):
        raise FillError(
            "数据结构不对：读到的内容是一段纯文本，不是数组。\n"
            "  请去掉最外层多余的引号，只保留 [{l:'0-0',v:'内容'}] 这样的部分。"
        )
    if isinstance(raw, list):
        return raw
    if isinstance(raw, dict):
        if _has_data_wrapper(raw):
            for key, value in raw.items():
                if isinstance(key, str) and key.strip().lower() in _DATA_WRAPPER_KEYS:
                    if isinstance(value, list):
                        return value
        if _looks_like_entry(raw):
            return [raw]
        raise FillError(
            "数据结构不对：顶层是对象，但既不是 {data:[...]}，也不是单条 {l:'0-0',v:'内容'}，"
            "也不是 {工作表名:[...]} 这种按工作表分组的形式。"
        )
    raise FillError("数据结构不对：顶层应该是 [ ... ] 数组，形如 [{l:'0-0',v:'内容'},{l:'0-1',v:'内容'}]。")


def _entry_from_item(item: Any, tag: str, problems: list[str]) -> dict | None:
    """把一条数据整理成 {loc, value, sheet}；格式不对就把原因记进 problems。"""
    if isinstance(item, (list, tuple)):
        if len(item) != 2:
            problems.append(f"{tag}：应写成 [位置, 内容] 两个元素，实际给了 {len(item)} 个。")
            return None
        item = {"l": item[0], "v": item[1]}
    if not isinstance(item, dict) or not item:
        problems.append(f"{tag}：格式不对，应该是 {{l:'行-列', v:'内容'}} 这样的对象。")
        return None

    found_loc, loc = _lookup(item, _LOC_KEYS)
    if not found_loc or loc is None or str(loc).strip() == "":
        problems.append(f"{tag}：缺少位置字段 l（例如 l:'0-0'）。")
        return None
    found_val, value = _lookup(item, _VAL_KEYS)
    if not found_val:
        problems.append(f"{tag}：缺少内容字段 v（清空单元格请显式写 v:null）。")
        return None
    if value is _MISSING:
        problems.append(f"{tag}：内容字段后面没有值。")
        return None

    found_sheet, sheet = _lookup(item, _SHEET_KEYS)
    return {"loc": loc, "value": value, "sheet": sheet if found_sheet else None}


def normalize_entries(raw: Any) -> list[dict]:
    """把解析出来的 JSON 整理成 [{loc, value, sheet}, ...]，格式不对就报错。"""
    items = _top_level_items(raw)
    if not items:
        raise FillError("数据里没有任何条目（数组是空的）。")

    entries: list[dict] = []
    problems: list[str] = []
    for index, item in enumerate(items, 1):
        entry = _entry_from_item(item, f"第 {index} 条", problems)
        if entry is None:
            continue
        entry["index"] = index
        entries.append(entry)

    if problems:
        raise FillError("数据有 %d 处问题：" % len(problems), problems)
    return entries


@dataclass
class DataDocument:
    """解析后的数据文档。

    * entries:      扁平的写入条目（每条带 sheet 名，可能为 None）
    * sheet_groups: 按工作表分组时的 [(分组名, 该表条目列表), ...]；否则为 None
    * warnings:     解析阶段产生的提示
    """

    entries: list[dict]
    sheet_groups: list[tuple[str, list[dict]]] | None = None
    warnings: list[str] = field(default_factory=list)

    @property
    def is_sheet_map(self) -> bool:
        return self.sheet_groups is not None


def _items_from_value(value: Any) -> list:
    """工作表分组里的值：单条、条目数组、或空（只改表名）。"""
    if value is None:
        raise FillError("后面没有写内容：应写 {l:'0-0',v:'内容'}，或 [{l:...,v:...}, ...]，或空 {} / []")
    if isinstance(value, dict):
        if not value:
            return []
        if _looks_like_entry(value):
            return [value]
        raise FillError("格式不对：应为 {l:'0-0',v:'内容'} 或 [{l:...,v:...}, ...]，实际是缺少 l / v 的对象。")
    if isinstance(value, (list, tuple)):
        return list(value)
    raise FillError(f"格式不对：应为 {{l,v}} 或数组，实际是 {type(value).__name__}")


def _normalize_sheet_map(raw: dict) -> DataDocument:
    """处理 {'工作表名': 条目...} 这种按工作表分组、并按名字重命名的写法。"""
    if not raw:
        raise FillError("数据是空对象 {}，没有任何工作表内容。")

    document = DataDocument(entries=[])
    problems: list[str] = []
    groups: list[tuple[str, list[dict]]] = []

    for key, value in raw.items():
        name = str(key).strip()
        try:
            items = _items_from_value(value)
        except FillError as exc:
            problems.append(f"工作表 {name!r}：{exc.message}")
            continue

        body: list[dict] = []
        for j, item in enumerate(items, 1):
            tag = f"工作表 {name!r} 第 {j} 条"
            entry = _entry_from_item(item, tag, problems)
            if entry is None:
                continue
            if entry["sheet"] is not None and str(entry["sheet"]).strip() != name:
                document.warnings.append(
                    f"{tag}：条目里的工作表字段 s={entry['sheet']!r} 被忽略，按外层分组名 {name!r} 处理。"
                )
            entry["sheet"] = name
            entry["index"] = len(document.entries) + 1
            document.entries.append(entry)
            body.append(entry)

        groups.append((name, body))

    if not groups:
        raise FillError("数据里没有任何工作表分组。", problems)
    if problems:
        raise FillError("数据有 %d 处问题：" % len(problems), problems)

    document.sheet_groups = groups
    return document


def normalize_document(raw: Any) -> DataDocument:
    """把顶层数据统一成 DataDocument（兼容旧的数组写法和新的工作表分组写法）。"""
    raw = _unwrap_string(raw)
    if isinstance(raw, dict) and not _has_data_wrapper(raw) and not _looks_like_entry(raw):
        return _normalize_sheet_map(raw)
    return DataDocument(entries=normalize_entries(raw))


# ---------------------------------------------------------------- 位置解析

def parse_single_ref(text: str) -> tuple[int, int]:
    """解析单个位置，返回 (行, 列)，均为 1 基。"""
    s = text.strip()
    if not s:
        raise ValueError("位置为空")
    m = _R1C1_RE.match(s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = _RC_RE.match(s)
    if m:
        return int(m.group(1)) + 1, int(m.group(2)) + 1
    m = _A1_RE.match(s)
    if m:
        col_letter = m.group(1).upper()
        col = column_index_from_string(col_letter)
        return int(m.group(2)), col
    raise ValueError(f"无法识别的位置写法 {s!r}（应为 '行-列'，如 0-0；或 A1 写法，如 A1）")


def parse_cell_refs(text: Any) -> list[tuple[int, int]]:
    """解析位置表达式，支持多目标与范围。

    * "0-0"            -> 第 1 行第 1 列
    * "A1"             -> 第 1 行第 1 列
    * "0-0,0-1"        -> 两个单元格
    * "0-0:2-3"        -> 区域（行 0~2，列 0~3），内容写到区域里每个单元格
    * "A1:C3"          -> 区域
    """
    raw = str(text).strip()
    if not raw:
        raise ValueError("位置为空")
    refs: list[tuple[int, int]] = []
    for part in _POS_SPLIT_RE.split(raw):
        if not part:
            continue
        if ":" in part or "：" in part:
            left, right = re.split(r"[:：]", part, maxsplit=1)
            r1, c1 = parse_single_ref(left)
            r2, c2 = parse_single_ref(right)
            _check_bounds(r1, c1)
            _check_bounds(r2, c2)
            top, bottom = sorted((r1, r2))
            lcol, rcol = sorted((c1, c2))
            count = (bottom - top + 1) * (rcol - lcol + 1)
            if count > MAX_RANGE_CELLS:
                raise ValueError(f"区域 {part} 太大（{count} 个单元格），请缩小范围")
            for row in range(top, bottom + 1):
                for col in range(lcol, rcol + 1):
                    refs.append((row, col))
        else:
            row, col = parse_single_ref(part)
            _check_bounds(row, col)
            refs.append((row, col))
    if not refs:
        raise ValueError("位置为空")
    return refs


def _check_bounds(row: int, col: int):
    if not (1 <= row <= _MAX_ROW):
        raise ValueError(f"行号 {row} 超出范围（1 ~ {_MAX_ROW}）")
    if not (1 <= col <= _MAX_COL):
        raise ValueError(f"列号 {col} 超出范围（1 ~ {_MAX_COL}）")


# ---------------------------------------------------------------- 工作表

def _resolve_sheet(wb, hint: Any):
    """按名称或序号（1 基）选择工作表；hint 为空时用活动工作表。"""
    if hint is None or str(hint).strip() == "":
        return wb.active
    text = str(hint).strip()
    names = wb.sheetnames
    if text.isdigit():
        idx = int(text)
        if not (1 <= idx <= len(names)):
            raise FillError(f"工作表序号 {idx} 不存在（共 {len(names)} 张表：{'、'.join(names)}）")
        return wb[names[idx - 1]]
    if text.lower() in (name.lower() for name in names):
        for name in names:
            if name.lower() == text.lower():
                return wb[name]
    # 允许用 "sheet1" 这种模糊写法命中 "Sheet 1"
    squashed = re.sub(r"\s+", "", text.lower())
    for name in names:
        if re.sub(r"\s+", "", name.lower()) == squashed:
            return wb[name]
    raise FillError(f"找不到工作表 {text!r}。当前工作簿里有：{'、'.join(names)}")


_INVALID_SHEET_CHARS = set(":\\/?*[]")


def _clean_sheet_name(name: Any) -> str:
    """校验并清理工作表名（Excel 限制：<=31 字符，不能含 : \\ / ? * [ ]）。"""
    text = str(name).strip()
    if not text:
        raise FillError("工作表名不能为空。")
    if len(text) > 31:
        raise FillError(f"工作表名 {text!r} 太长（Excel 工作表名最多 31 个字符）。")
    bad = sorted(_INVALID_SHEET_CHARS.intersection(text))
    if bad:
        raise FillError(f"工作表名 {text!r} 含有 Excel 不允许的字符：{' '.join(bad)}")
    return text


def _apply_sheet_map(wb, groups: list[tuple[str, list[dict]]]) -> tuple[dict[str, Worksheet], list[str]]:
    """按数据里的分组名匹配/重命名工作表。返回 (分组名 -> 工作表, 提示列表)。

    匹配规则：
      1. 分组名与模板中的工作表同名（忽略大小写/首尾空格）→ 用那张表，名字改成分组名写法；
      2. 否则，如果分组数量正好等于模板工作表数量 → 按顺序对应；
      3. 都不满足则报错。
    """
    keys = [_clean_sheet_name(name) for name, _ in groups]
    if len(set(keys)) != len(keys):
        repeated = sorted({k for k in keys if keys.count(k) > 1})
        raise FillError(f"数据里的工作表名重复：{'、'.join(repeated)}。")

    names = list(wb.sheetnames)
    by_lower: dict[str, str] = {}
    for n in names:
        by_lower.setdefault(n.strip().lower(), n)

    if all(k.lower() in by_lower for k in keys):
        pairs = [(by_lower[k.lower()], k) for k in keys]
    elif len(keys) == len(names):
        pairs = list(zip(names, keys))
    else:
        raise FillError(
            f"数据里有 {len(keys)} 个工作表分组（{'、'.join(keys)}），"
            f"模板里有 {len(names)} 张工作表（{'、'.join(names)}），对不上。\n"
            "  规则：分组名要么和模板工作表同名（按名字对应），要么分组数和模板表数相等（按顺序对应）。"
        )

    sources = [src for src, _ in pairs]
    if len(set(sources)) != len(sources):
        raise FillError("多个分组指向了同一张模板工作表，请检查分组名。")

    # 先改成临时名，避免"新名字刚好是另一张表的旧名字"时互相冲突
    warnings: list[str] = []
    pending: list[tuple[Worksheet, str, str]] = []
    for i, (src, dst) in enumerate(pairs):
        if src == dst:
            continue
        ws = wb[src]
        ws.title = f"__writeexcel_rename_{i}__"
        pending.append((ws, dst, src))

    for ws, dst, src in pending:
        ws.title = dst
        warnings.append(f"工作表「{src}」已重命名为「{dst}」")

    mapping = {k: wb[dst] for k, (_src, dst) in zip(keys, pairs)}
    return mapping, warnings


def _merged_anchor(ws: Worksheet, row: int, col: int) -> tuple[int, int]:
    """如果目标单元格在合并区域里，返回区域左上角，否则原样返回。"""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col
    return row, col


def _cell_address(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


# ---------------------------------------------------------------- 值转换

def _convert_value(value: Any, parse_numbers: bool, force_text: bool) -> Any:
    if value is None:
        return None
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, bool):
        if force_text:
            return "TRUE" if value else "FALSE"
        return value
    if force_text:
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    if parse_numbers and isinstance(value, str):
        text = value.strip()
        if _NUMBER_RE.match(text) and not _has_leading_zero(text):
            try:
                return int(text)
            except ValueError:
                return float(text)
    return value


def _has_leading_zero(text: str) -> bool:
    """001、0.5 之外的「01」这类应保留为文本。"""
    t = text.strip().lstrip("+-")
    return len(t) > 1 and t.startswith("0") and t[1] != "."


# ---------------------------------------------------------------- 保存

def _archive_hints(template: Path) -> list[str]:
    """检查模板里有哪些 openpyxl 处理不好的内容，提前提示。"""
    hints: list[str] = []
    try:
        with zipfile.ZipFile(template) as zf:
            names = zf.namelist()
    except Exception:
        return hints
    if any(n.startswith("xl/media/") for n in names):
        hints.append("模板中包含图片：保存后图片有可能丢失，请打开输出文件核对一下。")
    if any(n.startswith("xl/charts/") for n in names):
        hints.append("模板中包含图表：保存后图表有可能丢失，请打开输出文件核对一下。")
    if any(n.startswith("xl/drawings/vmlDrawing") for n in names):
        hints.append("模板中包含批注或表单控件，保存后可能丢失。")
    return hints


def _save_atomic(wb, output: Path):
    tmp = output.with_name(output.name + ".tmp-writing")
    try:
        wb.save(str(tmp))
        os.replace(str(tmp), str(output))
    except PermissionError as exc:
        _safe_unlink(tmp)
        raise FillError(
            f"无法写入 {output}：文件可能正在被 Excel 打开（或没有写权限）。请关闭该文件后重试。"
        ) from exc
    except Exception:
        _safe_unlink(tmp)
        raise


def _safe_unlink(path: Path):
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def _same_value(expected: Any, actual: Any) -> bool:
    if expected is None:
        return actual is None
    if isinstance(expected, bool) or isinstance(actual, bool):
        return bool(expected) == bool(actual)
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(float(expected) - float(actual)) <= 1e-9
    if isinstance(expected, str) and isinstance(actual, str):
        return expected == actual
    return str(expected) == str(actual)


# ---------------------------------------------------------------- 主流程

def fill(
    template: str | os.PathLike,
    raw_data: Any,
    output: str | os.PathLike | None = None,
    *,
    sheet: Any = None,
    overwrite: bool = False,
    parse_numbers: bool = False,
    force_text: bool = False,
    verify: bool = True,
    dry_run: bool = False,
) -> FillReport:
    """把 raw_data（解析好的 JSON 结构）写进模板，另存为 output。

    * template: 模板路径，支持 .xlsx / .xlsm
    * raw_data: 数据，两种形态：
               1. 顶层数组：[{l:'0-0', v:'内容1'}, ...]
               2. 按工作表分组：{'工作表名': {l:'0-0', v:'内容'}, ...}
                  —— 分组名会用来重命名模板里的工作表
    * output:   输出路径；不传则默认「模板名_已填充.xlsx」
    * sheet:    工作表名或序号（1 基），不传用活动工作表
    * overwrite: 输出文件已存在时是否覆盖（同时出现在模板路径时视为允许覆盖模板）
    * parse_numbers: 把「纯数字文本」写成数字
    * force_text:    把所有内容都写成文本
    * verify:   保存后重新打开核对一遍
    * dry_run:  只校验、不写文件（试运行）
    """
    template_path = Path(template).expanduser().resolve()
    _check_template(template_path)

    document = raw_data if isinstance(raw_data, DataDocument) else normalize_document(raw_data)
    entries = document.entries

    output_path = _decide_output(template_path, output, overwrite, dry_run)

    warnings = list(document.warnings)
    warnings.extend(_archive_hints(template_path))

    # ---- 载入工作簿
    keep_vba = template_path.suffix.lower() == ".xlsm"
    try:
        wb = load_workbook(str(template_path), keep_vba=keep_vba)
    except PermissionError as exc:
        raise FillError(
            f"无法读取模板 {template_path}：文件正在被 Excel 占用，请关闭后重试。"
        ) from exc
    except Exception as exc:
        raise FillError(
            f"无法打开模板 {template_path}：{exc}\n"
            "可能的原因：文件损坏、被加密、或不是真正的 Excel 文件（旧版 .xls 请先在 Excel 里另存为 .xlsx）。"
        ) from exc

    # 输出扩展名必须匹配宏状态，否则 openpyxl 会拒绝保存
    if keep_vba and output_path.suffix.lower() != ".xlsm":
        raise FillError("模板是启用宏的 .xlsm，输出文件也必须是 .xlsm 后缀。")

    # ---- 按工作表分组：先把页名改成数据里给的名字
    sheet_map: dict[str, Worksheet] | None = None
    if document.is_sheet_map:
        if sheet not in (None, ""):
            warnings.append("数据已按工作表分组给出，-s / sheet 参数被忽略。")
        sheet_map, rename_warnings = _apply_sheet_map(wb, document.sheet_groups or [])
        warnings.extend(rename_warnings)

    # ---- 逐条校验（先全部校验，任何一条出错都不写文件）
    problems: list[str] = []
    planned: list[tuple[Worksheet, int, int, Any, str]] = []
    used_sheets: list[str] = []
    seen: dict[tuple[str, int, int], int] = {}

    for entry in entries:
        idx = entry["index"]
        tag = f"第 {idx} 条（l={entry['loc']!r}）"
        if sheet_map is not None:
            ws = sheet_map.get(str(entry["sheet"]).strip())
            if ws is None:
                problems.append(f"{tag}：找不到工作表分组 {entry['sheet']!r}。")
                continue
        else:
            try:
                ws = _resolve_sheet(wb, entry["sheet"] if entry["sheet"] is not None else sheet)
            except FillError as exc:
                problems.append(f"{tag}：{exc.message}")
                continue

        try:
            refs = parse_cell_refs(entry["loc"])
        except ValueError as exc:
            problems.append(f"{tag}：{exc}")
            continue

        value = _convert_value(entry["value"], parse_numbers, force_text)
        if isinstance(value, str) and len(value) > _CELL_CHAR_LIMIT:
            problems.append(f"{tag}：内容太长（{len(value)} 字符），超过 Excel 单元格上限 {_CELL_CHAR_LIMIT}。")
            continue
        if isinstance(value, str) and value.startswith("="):
            warnings.append(f"{tag}：内容以 = 开头，会被 Excel 当成公式写入，请核对。")

        for row, col in refs:
            try:
                _check_bounds(row, col)
            except ValueError as exc:
                problems.append(f"{tag}：{exc}")
                continue
            arow, acol = _merged_anchor(ws, row, col)
            if (arow, acol) != (row, col):
                warnings.append(
                    f"{tag}：{_cell_address(row, col)} 在合并单元格 {_cell_address(arow, acol)} 区域内，"
                    f"内容写到了合并区域左上角 {_cell_address(arow, acol)}。"
                )
            key = (ws.title, arow, acol)
            if key in seen:
                warnings.append(
                    f"{tag}：和 {seen[key]} 号条目写同一个单元格 {ws.title}!{_cell_address(arow, acol)}，以后来的为准。"
                )
            seen[key] = idx
            planned.append((ws, arow, acol, value, ws.title))
            if ws.title not in used_sheets:
                used_sheets.append(ws.title)

    if problems:
        raise FillError("数据有 %d 处问题，没有写入任何内容：" % len(problems), problems)

    if sheet_map is not None:
        for name in sheet_map:
            if name not in used_sheets:
                used_sheets.append(name)

    if not planned and sheet_map is None:
        raise FillError("没有可以写入的单元格。")

    # ---- 试运行：只返回计划，不动文件
    if dry_run:
        return FillReport(
            output=output_path,
            written=0,
            sheet_used=used_sheets,
            warnings=_dedup(warnings),
            applied=[(title, row, col, value) for _, row, col, value, title in planned],
            dry_run=True,
        )

    # ---- 写入
    for ws, row, col, value, _ in planned:
        ws.cell(row=row, column=col).value = value

    _save_atomic(wb, output_path)

    # ---- 保存后核对
    if verify:
        warnings.extend(_verify_saved(output_path, planned, keep_vba))

    report = FillReport(
        output=output_path,
        written=len(planned),
        sheet_used=used_sheets,
        warnings=_dedup(warnings),
        applied=[(title, row, col, value) for _, row, col, value, title in planned],
    )
    return report


def _check_template(template_path: Path):
    if not template_path.exists():
        raise FillError(f"找不到模板文件：{template_path}")
    if not template_path.is_file():
        raise FillError(f"模板路径不是文件：{template_path}")
    suffix = template_path.suffix.lower()
    if suffix == ".xls":
        raise FillError(
            "不支持旧版 .xls 格式：请用 Excel 打开该文件，「另存为」→ Excel 工作簿(*.xlsx) 后再试。"
        )
    if suffix == ".xltx":
        raise FillError(
            "不支持 .xltx 模板格式：请用 Excel 打开后「另存为」→ Excel 工作簿(*.xlsx) 后再试。"
        )
    if suffix not in TEMPLATE_SUFFIXES:
        raise FillError(
            f"模板必须是 .xlsx 或 .xlsm 文件，当前是 {suffix or '（无扩展名）'}。"
        )


def _decide_output(template_path: Path, output, overwrite: bool, dry_run: bool = False) -> Path:
    if output is None or str(output).strip() == "":
        stem = template_path.stem
        suffix = template_path.suffix
        candidate = template_path.with_name(f"{stem}_已填充{suffix}")
        if candidate.exists() and not overwrite and not dry_run:
            n = 2
            while True:
                candidate = template_path.with_name(f"{stem}_已填充({n}){suffix}")
                if not candidate.exists():
                    break
                n += 1
        return candidate

    out = Path(output).expanduser().resolve()
    if out.suffix.lower() not in (".xlsx", ".xlsm"):
        # 用户可能只写了文件名，补上后缀
        expected = template_path.suffix.lower()
        out = out.with_suffix(expected)
    if dry_run:
        return out
    if out == template_path and not overwrite:
        raise FillError(
            "输出路径和模板是同一个文件。为了不破坏模板，请换个输出文件名；"
            "确实要直接改模板，请加 --overwrite（-y）参数。"
        )
    if out.exists() and not overwrite:
        raise FillError(f"输出文件已存在：{out}。请换一个名字，或加 --overwrite（-y）参数覆盖。")
    out.parent.mkdir(parents=True, exist_ok=True)
    if out == template_path:
        # 允许覆盖模板：先备份
        backup = template_path.with_name(template_path.stem + ".bak" + template_path.suffix)
        try:
            import shutil
            shutil.copy2(template_path, backup)
        except OSError:
            pass
    return out


def _verify_saved(output_path: Path, planned, keep_vba: bool) -> list[str]:
    warnings: list[str] = []
    try:
        wb2 = load_workbook(str(output_path), keep_vba=keep_vba)
    except Exception as exc:
        return [f"保存成功，但重新打开核对时失败：{exc}"]
    cache: dict[str, Worksheet] = {}
    bad = 0
    for ws, row, col, value, title in planned:
        target = cache.get(title)
        if target is None:
            target = wb2[title]
            cache[title] = target
        actual = target.cell(row=row, column=col).value
        if not _same_value(value, actual):
            bad += 1
            if bad <= 5:
                warnings.append(
                    f"核对不一致：{title}!{_cell_address(row, col)} 期望 {value!r}，实际 {actual!r}"
                )
    if bad > 5:
        warnings.append(f"核对不一致的单元格共 {bad} 个。")
    return warnings


def _dedup(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out
