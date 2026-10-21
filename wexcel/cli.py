"""命令行 / 交互入口。"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

from . import parser as json_parser
from .excel_writer import (
    FillError,
    fill,
    normalize_document,
    _cell_address,
)
from .parser import ParseError

__all__ = ["main"]

PROG = "Excel 模板填充工具"
VERSION = "1.1.0"

EXIT_OK = 0
EXIT_USAGE = 2
EXIT_DATA = 3
EXIT_WRITE = 4


# ---------------------------------------------------------------- 环境准备

def _setup_stdio():
    """保证在 Windows 各种代码页下都不会因为中文/特殊字符崩掉。"""
    if os.name == "nt" and _is_frozen():
        # 打包成 exe 后，把控制台切到 UTF-8，这样英文系统的电脑也能正常显示中文
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)  # type: ignore[attr-defined]
            ctypes.windll.kernel32.SetConsoleCP(65001)  # type: ignore[attr-defined]
        except Exception:
            pass
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
            except Exception:
                pass
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _pause(args=None, text: str = "\n按回车键退出…"):
    """双击运行时让窗口停住，避免一闪而过看不到结果。"""
    if args is not None and getattr(args, "no_pause", False):
        return
    if not sys.stdin or not sys.stdin.isatty():
        return
    try:
        input(text)
    except (EOFError, KeyboardInterrupt):
        pass


# ---------------------------------------------------------------- 文本读取

def read_text_file(path: Path, encoding: str | None = None) -> str:
    """读取 txt，自动识别 UTF-8 / UTF-16 / GBK(GB18030)。"""
    data = path.read_bytes()
    if not data.strip():
        raise ParseError("数据文件是空的。")

    if encoding:
        return data.decode(encoding, errors="replace").lstrip("\ufeff")

    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig")
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16")

    nul_ratio = data.count(b"\x00") / max(1, len(data))
    if nul_ratio > 0.25:
        # 无 BOM 的 UTF-16（Windows 记事本另存为"Unicode"的老习惯）
        for enc in ("utf-16-le", "utf-16-be"):
            try:
                text = data.decode(enc)
            except UnicodeDecodeError:
                continue
            if text.count("\x00") == 0:
                return text

    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        pass

    # 中文 Windows 记事本默认是 GBK
    for enc in ("gb18030", "big5"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


# ---------------------------------------------------------------- 路径辅助

def _clean_path(text: str) -> str:
    """清理拖拽/粘贴进来的路径：去掉成对引号、file:// 前缀、首尾空白。"""
    t = text.strip()
    if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'“”":
        t = t[1:-1].strip()
    if t.lower().startswith("file:///"):
        from urllib.parse import unquote
        t = unquote(t[8:])
        if os.name == "nt":
            t = t.replace("/", "\\")
    # Windows 资源管理器拖拽有时会给带引号的路径，也可能带尾部反斜杠
    return t


def _split_paths(text: str) -> list[str]:
    """把一行里可能包含的多个路径拆开（支持带空格的引号路径）。"""
    parts: list[str] = []
    i, n = 0, len(text)
    while i < n:
        while i < n and text[i].isspace():
            i += 1
        if i >= n:
            break
        if text[i] in "\"'“”":
            quote = text[i]
            i += 1
            start = i
            while i < n and text[i] not in "\"'“”":
                i += 1
            parts.append(text[start:i])
            i += 1
        else:
            start = i
            while i < n and not text[i].isspace():
                i += 1
            parts.append(text[start:i])
    return [p for p in parts if p.strip()]


# ---------------------------------------------------------------- 参数

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="write_excel",
        description=f"{PROG} v{VERSION}：按 txt 里的 JSON 数据填充 Excel 模板。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  write_excel 模板.xlsx 数据.txt\n"
            "  write_excel 模板.xlsx 数据.txt 结果.xlsx -y\n"
            "  write_excel -t 模板.xlsx -d 数据.txt -s Sheet1 --parse-numbers\n"
            "\n"
            "数据文件内容示例：\n"
            "  1) 普通写法（写入活动工作表）：\n"
            "     [{l:'0-0',v:'内容1'},{l:'0-1',v:'内容2'}]\n"
            "  2) 按工作表分组（分组名会用来重命名模板里的工作表）：\n"
            "     {'Sheet1':[{l:'0-0',v:'内容1'}], 'Sheet2':{}}\n"
            "  l 是位置，行-列 从 0 开始；v 是要写入的内容。\n"
            "\n"
            "提示：什么都不带直接运行，会进入交互模式（可把文件直接拖进来）。"
        ),
    )
    p.add_argument("files", nargs="*", help="位置参数：模板 [数据] [输出]，可直接把文件拖到程序上")
    p.add_argument("-t", "--template", help="Excel 模板路径（.xlsx / .xlsm）")
    p.add_argument("-d", "--data", help="数据 txt 文件路径")
    p.add_argument("-o", "--output", help="输出文件路径（默认：模板名_已填充.xlsx）")
    p.add_argument("-s", "--sheet", help="工作表名称或序号（从 1 开始），默认使用活动工作表")
    p.add_argument("-y", "--overwrite", action="store_true", help="允许覆盖已存在的文件（包括直接改模板）")
    p.add_argument("--parse-numbers", action="store_true", help="把「纯数字字符串」写成数字（如 '123' -> 123）")
    p.add_argument("--force-text", action="store_true", help="把所有内容都写成文本（防止被识别成数字/日期）")
    p.add_argument("-n", "--dry-run", action="store_true", help="只校验数据、打印将要写入的位置，不产生文件")
    p.add_argument("--no-verify", action="store_true", help="保存后不再重新打开核对")
    p.add_argument("--encoding", help="强制用指定编码读取 txt（utf-8 / gbk / utf-16 等）")
    p.add_argument("--pause", action="store_true", help="结束后等待回车（默认只有不带参数双击时才等待）")
    p.add_argument("--no-pause", action="store_true", help="强制不等待回车（供脚本调用，保证自动结束）")
    p.add_argument("-V", "--version", action="version", version=f"{PROG} {VERSION}")
    return p


# ---------------------------------------------------------------- 主流程

def run_fill(args, template: str, data: str, output: str | None) -> int:
    started = time.time()

    template_path = Path(_clean_path(template)).expanduser()
    data_path = Path(_clean_path(data)).expanduser()

    if not template_path.exists():
        _error(f"找不到模板文件：{template_path}")
        return EXIT_USAGE
    if not data_path.is_file():
        _error(f"找不到数据文件：{data_path}")
        return EXIT_USAGE

    print(f"[读取] 模板：{template_path}")
    print(f"[读取] 数据：{data_path}")
    try:
        text = read_text_file(data_path, args.encoding)
    except ParseError as exc:
        _error(str(exc))
        return EXIT_DATA
    except OSError as exc:
        _error(f"读取数据文件失败：{exc}")
        return EXIT_DATA

    try:
        raw = json_parser.loads(text)
    except ParseError as exc:
        _error(f"数据格式解析失败：{exc}")
        return EXIT_DATA

    try:
        document = normalize_document(raw)
    except FillError as exc:
        _error(exc.message)
        for line in exc.problems:
            print(f"    - {line}")
        return EXIT_DATA

    if document.is_sheet_map:
        group_names = "、".join(name for name, _ in document.sheet_groups or [])
        print(f"[解析] 共 {len(document.entries)} 条数据，工作表分组：{group_names}")
    else:
        print(f"[解析] 共 {len(document.entries)} 条数据")

    try:
        report = fill(
            template_path,
            document,
            output,
            sheet=args.sheet,
            overwrite=args.overwrite,
            parse_numbers=args.parse_numbers,
            force_text=args.force_text,
            verify=not args.no_verify,
            dry_run=args.dry_run,
        )
    except FillError as exc:
        _error(exc.message)
        for line in exc.problems:
            print(f"    - {line}")
        return EXIT_WRITE
    except PermissionError as exc:
        _error(f"文件被占用或无权限：{exc}")
        return EXIT_WRITE

    for line in report.warnings:
        print(f"[提示] {line}")

    if report.dry_run:
        print(f"[试运行] 工作表：{'、'.join(report.sheet_used)}，计划写入 {len(report.applied)} 个单元格（未生成文件）")
        for title, row, col, value in report.applied[:50]:
            preview = _preview(value)
            print(f"    {title}!{_cell_address(row, col)} ← {preview}")
        if len(report.applied) > 50:
            print(f"    … 其余 {len(report.applied) - 50} 个单元格省略")
        print("[完成] 试运行结束，数据没有发现问题。")
        return EXIT_OK

    print(f"[写入] 工作表：{'、'.join(report.sheet_used)}，共 {report.written} 个单元格")
    print(f"[保存] {report.output}")
    print(f"[完成] 用时 {time.time() - started:.2f} 秒")
    return EXIT_OK


def _preview(value) -> str:
    if value is None:
        return "（清空）"
    text = str(value)
    if len(text) > 40:
        text = text[:40] + "…"
    return repr(text)


def _error(message: str):
    print(f"[错误] {message}", file=sys.stderr)


def run_interactive() -> int:
    print("=" * 58)
    print(f"           {PROG}  v{VERSION}")
    print("=" * 58)
    print("用法：把 Excel 模板和数据 txt 直接拖到本窗口，或按提示输入路径。")
    print("提示：路径可以直接拖进来；带空格也支持。")
    print("-" * 58)

    try:
        raw = input("Excel 模板路径：").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return EXIT_OK
    if not raw:
        print("没有输入模板路径，已退出。")
        return EXIT_OK
    parts = _split_paths(raw)

    template = parts[0] if parts else ""
    data = ""
    output = None
    if len(parts) >= 2:
        data = parts[1]
    if len(parts) >= 3:
        output = parts[2]

    if not data:
        data, output, code = _ask_data_and_output(template, output)
        if code is not None:
            return code

    print("-" * 58)
    args = build_parser().parse_args([])
    code = run_fill(args, template, data, output)
    print("-" * 58)
    return code


def _ask_data_and_output(template: str, output: str | None):
    """追问数据文件路径和输出路径。返回 (data, output, 提前结束时的退出码或 None)。"""
    try:
        data = input("数据 txt 路径：").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return "", None, EXIT_OK
    if not data:
        print("没有输入数据文件路径，已退出。")
        return "", None, EXIT_OK

    if output is None:
        try:
            suffix = Path(_clean_path(template)).suffix or ".xlsx"
            default_name = f"{Path(_clean_path(template)).stem}_已填充{suffix}"
            answer = input(f"输出文件（回车使用默认：{default_name}）：").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return data, None, EXIT_OK
        output = _clean_path(answer) if answer else None
    return data, output, None


def main(argv: list[str] | None = None) -> int:
    _setup_stdio()
    argv = list(sys.argv[1:] if argv is None else argv)
    args = build_parser().parse_args(argv)

    # 位置参数：模板 [数据] [输出]
    files = [_clean_path(f) for f in args.files]
    if files:
        if not args.template and files:
            args.template = files[0]
        if not args.data and len(files) >= 2:
            args.data = files[1]
        if not args.output and len(files) >= 3:
            args.output = files[2]
        if len(files) > 3:
            _error("位置参数最多 3 个：模板 数据 输出")
            return EXIT_USAGE

    if not args.template and not args.data:
        code = run_interactive()
        _pause(args)
        return code

    # 拖进来时顺序可能反了，按扩展名自动纠正
    if args.template and args.data:
        t_suffix = Path(args.template).suffix.lower()
        d_suffix = Path(args.data).suffix.lower()
        if t_suffix == ".txt" and d_suffix in (".xlsx", ".xlsm"):
            args.template, args.data = args.data, args.template

    # 只拖进来一个文件（多半是模板）：在交互终端里追问数据文件
    if args.template and not args.data and len(files) == 1 and sys.stdin and sys.stdin.isatty():
        print(f"[提示] 已识别模板：{args.template}")
        data, extra_output, stop_code = _ask_data_and_output(args.template, args.output)
        if stop_code is not None:
            return stop_code
        args.data = data
        if extra_output:
            args.output = extra_output

    if not args.template or not args.data:
        _error("模板和数据文件都必须提供。例如：write_excel 模板.xlsx 数据.txt")
        _pause(args)
        return EXIT_USAGE

    try:
        code = run_fill(args, args.template, args.data, args.output)
    except KeyboardInterrupt:
        print("\n[中断] 已取消。")
        code = 130
    # 带参数调用（bat / 自动化）默认不暂停，保证能自动结束；需要停在窗口看结果时加 --pause
    if getattr(args, "pause", False):
        _pause(args)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
