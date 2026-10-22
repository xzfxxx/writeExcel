#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自测脚本：不依赖 pytest，直接运行

    python tests/selftest.py

覆盖：容错解析、单元格定位、写入结果、格式保留、合并单元格、试运行、错误提示。
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import traceback
from pathlib import Path

# 输出统一用 UTF-8：避免 Windows 英文系统（cp1252 等）打印中文时崩溃
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402

from wexcel import FillError, loads  # noqa: E402
from wexcel.excel_writer import (  # noqa: E402
    _has_leading_zero,
    fill,
    normalize_document,
    normalize_entries,
    parse_cell_refs,
    parse_single_ref,
)

PASSED = 0
FAILED: list[str] = []


def check(name: str, condition: bool, detail: str = ""):
    global PASSED
    if condition:
        PASSED += 1
    else:
        FAILED.append(f"{name} {detail}".strip())


def expect_error(name: str, func, *args, **kwargs):
    try:
        func(*args, **kwargs)
    except Exception:  # noqa: BLE001
        check(name, True)
        return
    check(name, False, "（应该报错，但没有报错）")


# ---------------------------------------------------------------- 解析测试

def test_parser():
    # 1. 标准 JSON
    check("标准 JSON", loads('[{"l":"0-0","v":"内容1"}]') == [{"l": "0-0", "v": "内容1"}])

    # 2. 单引号 + 无引号键（用户给的写法）
    data = loads("[{l:'0-0',v:'内容1'},{l:'0-1',v:'内容2'}]")
    check("单引号+裸键", data == [{"l": "0-0", "v": "内容1"}, {"l": "0-1", "v": "内容2"}])

    # 3. 漏写冒号：{l'0-1',v:'内容2'}
    data = loads("[{l'0-0',v:'内容1'},{l'0-1',v:'内容2'}]")
    check("漏写冒号", data == [{"l": "0-0", "v": "内容1"}, {"l": "0-1", "v": "内容2"}])

    # 4. 漏写逗号
    data = loads("[{l:'0-0' v:'内容1'} {l:'0-1' v:'内容2'}]")
    check("漏写逗号", len(data) == 2 and data[1]["v"] == "内容2")

    # 5. 结尾多余逗号 / 分号
    check("结尾多余逗号", loads("[{l:'0-0',v:'a'},]") == [{"l": "0-0", "v": "a"}])
    check("结尾多余分号", loads("[{l:'0-0',v:'a'}];") == [{"l": "0-0", "v": "a"}])

    # 6. 全角标点（中文输入法造成的）
    data = loads("［｛l：“0-0”，v：“内容，1”｝］")
    check("全角标点", data == [{"l": "0-0", "v": "内容，1"}], repr(data))

    # 7. 注释
    data = loads("// 说明\n[{l:'0-0',v:'a' /* 行内 */}, {l:'0-1',v:'b'}]")
    check("注释", len(data) == 2)

    # 8. 裸值：数字、布尔、null、带空格的字符串
    data = loads("[{l:'0-0',v:123},{l:'0-1',v:1.5},{l:'1-0',v:true},{l:'1-1',v:null},{l:'1-2',v:Hello World}]")
    check("裸数字", data[0]["v"] == 123 and isinstance(data[0]["v"], int))
    check("裸浮点", data[1]["v"] == 1.5)
    check("布尔", data[2]["v"] is True)
    check("null", data[3]["v"] is None)
    check("带空格字符串", data[4]["v"] == "Hello World", repr(data[4]["v"]))

    # 9. 位置也不加引号（0-0 会被当成字符串）
    data = loads("[{l:0-0,v:内容1},{l:0-1,v:内容2}]")
    check("裸位置", data[0]["l"] == "0-0", repr(data[0]))

    # 10. 转义
    check("转义换行", loads(r"[{l:'0-0',v:'a\nb'}]")[0]["v"] == "a\nb")
    check("unicode 转义", loads(r'[{"l":"0-0","v":"\u4e2d"}]')[0]["v"] == "中")

    # 11. 嵌套 / 单对象
    check("单对象", loads("{l:'0-0',v:'x'}") == {"l": "0-0", "v": "x"})
    check("嵌套容器", loads("{data:[{l:'0-0',v:'x'}]}")["data"][0]["v"] == "x")

    # 12. 错误提示
    expect_error("未闭合字符串", loads, "[{l:'0-0',v:'abc")
    expect_error("未闭合对象", loads, "[{l:'0-0',v:'x'")
    expect_error("空内容", loads, "   ")
    expect_error("多余内容", loads, "[1] 垃圾")


# ---------------------------------------------------------------- 定位测试

def test_positions():
    # 0 基「行-列」 -> 1 基
    check("0-0", parse_single_ref("0-0") == (1, 1))
    check("2-3", parse_single_ref("2-3") == (3, 4))
    check("A1", parse_single_ref("A1") == (1, 1))
    check("C5", parse_single_ref("C5") == (5, 3))
    check("AA10", parse_single_ref("AA10") == (10, 27))
    check("R2C3", parse_single_ref("R2C3") == (2, 3))

    check("多目标", parse_cell_refs("0-0,0-1") == [(1, 1), (1, 2)])
    check("多目标 中文逗号", parse_cell_refs("0-0，0-1") == [(1, 1), (1, 2)])
    check("范围", parse_cell_refs("A1:B2") == [(1, 1), (1, 2), (2, 1), (2, 2)])
    check("范围 0基", parse_cell_refs("0-0:1-1") == [(1, 1), (1, 2), (2, 1), (2, 2)])
    check("范围反向", parse_cell_refs("B2:A1") == [(1, 1), (1, 2), (2, 1), (2, 2)])

    expect_error("非法位置", parse_single_ref, "abc")
    expect_error("非法列", parse_single_ref, "123")
    expect_error("行号越界", parse_cell_refs, "2000000-0")

    # 数字识别
    check("前导零保留", _has_leading_zero("001") is True)
    check("小数不保留", _has_leading_zero("0.5") is False)


# ---------------------------------------------------------------- Excel 测试

def make_template(path: Path) -> None:
    """构造一个带样式的模板：合并标题、表头样式、列宽、冻结窗格、边框。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "客户信息"

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:E1")
    title = ws["A1"]
    title.value = "客户信息登记表"
    title.font = Font(bold=True, size=16, color="1F4E78")
    title.alignment = Alignment(horizontal="center", vertical="center")
    title.fill = PatternFill("solid", fgColor="DDEBF7")

    headers = ["序号", "姓名", "电话", "公司", "备注"]
    for idx, text in enumerate(headers, 1):
        cell = ws.cell(row=2, column=idx, value=text)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for row in range(3, 8):
        for col in range(1, 6):
            ws.cell(row=row, column=col).border = border

    widths = {"A": 8, "B": 14, "C": 18, "D": 26, "E": 30}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False

    ws2 = wb.create_sheet("说明")
    ws2["A1"] = "这个工作表用来测试多表写入"
    wb.save(path)


def test_excel(tmp: Path):
    template = tmp / "模板.xlsx"
    make_template(template)

    raw = loads("[{l:'2-0',v:'张三'},{l:'2-1',v:'13800000000'},{l:'2-2',v:'示例科技有限公司'},"
                "{l:'2-3',v:'重点客户，下周回访'},{l:'3-0',v:'李四'},{l:'3-1',v:'13900000000'}]")
    out = tmp / "out1.xlsx"
    report = fill(template, raw, out)

    check("写入数量", report.written == 6, str(report.written))
    wb = load_workbook(out)
    ws = wb["客户信息"]
    check("A3 张三", ws["A3"].value == "张三", repr(ws["A3"].value))
    check("B3 电话文本", ws["B3"].value == "13800000000")
    check("D3 中文逗号保留", ws["D3"].value == "重点客户，下周回访", repr(ws["D3"].value))

    # 样式是否保留
    check("标题仍是合并单元格", "A1:E1" in {str(r) for r in ws.merged_cells.ranges})
    check("标题字体保留", ws["A1"].font.bold and ws["A1"].font.size == 16)
    check("标题内容未被覆盖", ws["A1"].value == "客户信息登记表")
    check("列宽保留", ws.column_dimensions["D"].width == 26)
    check("冻结窗格保留", ws.freeze_panes == "A3")
    check("边框保留", ws["B3"].border.left.style == "thin")
    check("网格线设置保留", ws.sheet_view.showGridLines is False)
    check("第二张表未受影响", wb["说明"]["A1"].value == "这个工作表用来测试多表写入")

    # ---- 合并单元格重定向
    raw2 = loads("[{l:'0-3',v:'写到合并区域里'}]")  # D1 在 A1:E1 合并区内
    out2 = tmp / "out2.xlsx"
    r2 = fill(template, raw2, out2)
    wb2 = load_workbook(out2)
    check("合并区写入落在左上角", wb2["客户信息"]["A1"].value == "写到合并区域里")
    check("合并重定向有提示", any("合并" in w for w in r2.warnings))

    # ---- 数字/文本控制
    raw3 = loads("[{l:'2-0',v:'001'},{l:'2-1',v:'123'},{l:'2-2',v:'1.50'}]")
    out3 = tmp / "out3.xlsx"
    fill(template, raw3, out3)
    ws3 = load_workbook(out3)["客户信息"]
    check("默认保持文本 001", ws3["A3"].value == "001", repr(ws3["A3"].value))
    check("默认保持文本 123", ws3["B3"].value == "123", type(ws3["B3"].value).__name__)

    out4 = tmp / "out4.xlsx"
    fill(template, raw3, out4, parse_numbers=True)
    ws4 = load_workbook(out4)["客户信息"]
    check("parse_numbers 前导零仍是文本", ws4["A3"].value == "001")
    check("parse_numbers 数字变数值", ws4["B3"].value == 123)
    check("parse_numbers 小数变数值", ws4["C3"].value == 1.5)

    out5 = tmp / "out5.xlsx"
    fill(template, loads("[{l:'2-0',v:12345}]"), out5, force_text=True)
    check("force_text", load_workbook(out5)["客户信息"]["A3"].value == "12345")

    # ---- 清空单元格
    raw6 = loads("[{l:'2-0',v:null}]")
    out6 = tmp / "out6.xlsx"
    fill(template, raw6, out6)
    check("null 清空", load_workbook(out6)["客户信息"]["A3"].value is None)

    # ---- 指定工作表 + 区域写入
    raw7 = loads("[{l:'A1:B2',v:'x',s:'说明'}]")
    out7 = tmp / "out7.xlsx"
    r7 = fill(template, raw7, out7)
    ws7 = load_workbook(out7)["说明"]
    check("区域写入", [ws7[c].value for c in ("A1", "A2", "B1", "B2")] == ["x", "x", "x", "x"])
    check("区域写入数量", r7.written == 4)
    check("sheet_used", r7.sheet_used == ["说明"], str(r7.sheet_used))

    # ---- 试运行不产生文件
    out8 = tmp / "out8.xlsx"
    r8 = fill(template, loads("[{l:'2-0',v:'试运行'}]"), out8, dry_run=True)
    check("试运行不写文件", not out8.exists())
    check("试运行有计划", len(r8.applied) == 1 and r8.dry_run)

    # ---- 默认输出名
    out9 = tmp / "sub" / "模板.xlsx"
    out9.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(template, out9)
    r9 = fill(out9, loads("[{l:'2-0',v:'x'}]"))
    check("默认输出名", r9.output.name == "模板_已填充.xlsx", r9.output.name)

    # ---- 覆盖保护
    expect_error("输出已存在时报错", fill, template, loads("[{l:'2-0',v:'x'}]"), out)

    # ---- 错误数据：不写任何内容
    bad_out = tmp / "bad.xlsx"
    expect_error("位置非法", fill, template, loads("[{l:'??',v:'x'}]"), bad_out)
    check("出错时不生成文件", not bad_out.exists())
    expect_error("缺 v", fill, template, loads("[{l:'0-0'}]"), bad_out)
    expect_error("空数组", fill, template, loads("[]"), bad_out)

    # ---- 区域过大
    expect_error("区域过大", parse_cell_refs, "A1:ZZ1000")

    # ---- 工作表不存在
    expect_error("工作表不存在", fill, template, loads("[{l:'0-0',v:'x'}]"), bad_out, sheet="不存在")

    # ---- 多表写入
    raw10 = loads("[{l:'0-0',v:'主表'},{l:'0-0',v:'副表',s:2}]")
    out10 = tmp / "out10.xlsx"
    r10 = fill(template, raw10, out10)
    wb10 = load_workbook(out10)
    check("多表写入", wb10["客户信息"]["A1"].value == "主表" and wb10["说明"]["A1"].value == "副表")
    check("多表记录", len(r10.sheet_used) == 2)

    # ---- 公式提示
    r11 = fill(template, loads("[{l:'2-0',v:'=1+1'}]"), tmp / "out11.xlsx")
    check("公式给出提示", any("公式" in w for w in r11.warnings))

    # ---- 数据有重复位置时提示
    r12 = fill(template, loads("[{l:'2-0',v:'a'},{l:'2-0',v:'b'}]"), tmp / "out12.xlsx")
    check("重复位置提示", any("同一个单元格" in w for w in r12.warnings))


def test_normalize():
    entries = normalize_entries(
        [{"位置": "0-0", "内容": "a"}, ["1-1", "b"], {"cell": "2-2", "value": 3, "sheet": "说明"}]
    )
    check("中文键名", entries[0]["value"] == "a")
    check("二元数组写法", entries[1]["loc"] == "1-1")
    check("英文键名+工作表", entries[2]["sheet"] == "说明")
    expect_error("缺少位置", normalize_entries, [{"v": "x"}])
    expect_error("缺少内容", normalize_entries, [{"l": "0-0"}])
    expect_error("顶层不是数组", normalize_entries, "abc")

    # 规范化入口：旧格式不算"工作表分组"
    check("数组不算分组", normalize_document(loads("[{l:'0-0',v:'a'}]")).is_sheet_map is False)
    check("data 包装不算分组", normalize_document(loads("{data:[{l:'0-0',v:'a'}]}")).is_sheet_map is False)
    check("单条不算分组", normalize_document(loads("{l:'0-0',v:'a'}")).is_sheet_map is False)
    check("分组被识别", normalize_document(loads("{'S1':{l:'0-0',v:'a'}}")).is_sheet_map is True)


def main() -> int:
    test_parser()
    test_positions()
    test_normalize()

def test_sheet_map(tmp: Path):
    """按工作表分组 + 按参数重命名工作表。"""
    template = tmp / "模板_map.xlsx"
    make_template(template)  # 两张表：客户信息、备注说明

    # 1) 分组名与模板表名一致：写内容，不改名
    raw = loads("{'客户信息':[{l:'2-0',v:'张三'},{l:'2-1',v:'138'}], '说明':{l:'0-0',v:'副表'}}")
    out1 = tmp / "map1.xlsx"
    r1 = fill(template, raw, out1)
    wb1 = load_workbook(out1)
    check("分组同名写入", wb1["客户信息"]["A3"].value == "张三" and wb1["客户信息"]["B3"].value == "138")
    check("分组单条写入", wb1["说明"]["A1"].value == "副表")
    check("分组同名不改名", wb1.sheetnames == ["客户信息", "说明"], str(wb1.sheetnames))
    check("分组上报的表名", r1.sheet_used == ["客户信息", "说明"], str(r1.sheet_used))

    # 2) 分组名不同、数量相同：按顺序重命名 + 写入
    raw2 = loads("{'客户清单':[{l:'2-0',v:'李四'}], '汇总表':{}}")
    out2 = tmp / "map2.xlsx"
    r2 = fill(template, raw2, out2)
    wb2 = load_workbook(out2)
    check("按顺序重命名", wb2.sheetnames == ["客户清单", "汇总表"], str(wb2.sheetnames))
    check("重命名后写入", wb2["客户清单"]["A3"].value == "李四")
    check("空分组也能改名", wb2["汇总表"]["A1"].value == "这个工作表用来测试多表写入")
    check("改名有提示", any("重命名" in w for w in r2.warnings))

    # 3) 分组数与模板表数对不上 -> 报错
    expect_error("分组数对不上", fill, template, loads("{'XX':{}}"), tmp / "map3.xlsx")
    check("分组数对不上不生成文件", not (tmp / "map3.xlsx").exists())

    # 4) 表名不合法
    expect_error("表名过长", fill, template, loads("{'%s':{}}" % ("长" * 32)), tmp / "map4.xlsx")
    expect_error("表名非法字符", fill, template, loads("{'a:b':{}}"), tmp / "map5.xlsx")

    # 5) 空对象
    expect_error("空分组对象", fill, template, loads("{}"), tmp / "map6.xlsx")

    # 6) 分组内条目自带 s 会被忽略并提示
    raw7 = loads("{'客户信息':[{l:'2-0',v:'a',s:'说明'}]}")
    out7 = tmp / "map7.xlsx"
    r7 = fill(template, raw7, out7)
    check("忽略条目 s 并提示", any("被忽略" in w for w in r7.warnings))
    wb7 = load_workbook(out7)
    check("按分组名写入", wb7["客户信息"]["A3"].value == "a")
    check("没写到 s 指定的表", wb7["说明"]["A3"].value is None)

    # 7) 分组格式 + 试运行
    r8 = fill(template, loads("{'表一':{}, '表二':{}}"), None, dry_run=True)
    check("分组试运行", r8.dry_run and r8.sheet_used == ["表一", "表二"], str(r8.sheet_used))

    # 8) 英文表名大小写不同也算同名
    tpl_latin = tmp / "模板_latin.xlsx"
    wb_latin = Workbook()
    wb_latin.active.title = "Sheet1"
    wb_latin.create_sheet("Sheet2")
    wb_latin.save(tpl_latin)
    out9 = tmp / "map9.xlsx"
    fill(tpl_latin, loads("{'SHEET1':{l:'0-0',v:'a'}, 'sheet2':{l:'0-0',v:'b'}}"), out9)
    wb9 = load_workbook(out9)
    check("大小写同名匹配", wb9.sheetnames == ["SHEET1", "sheet2"], str(wb9.sheetnames))
    check("大小写匹配写入", wb9["SHEET1"]["A1"].value == "a" and wb9["sheet2"]["A1"].value == "b")

    # 9) 新旧表名互换（验证内部临时改名机制）
    tpl_swap = tmp / "模板_swap.xlsx"
    wb_swap = Workbook()
    wb_swap.active.title = "A"
    wb_swap.create_sheet("B")
    wb_swap.save(tpl_swap)
    out10 = tmp / "map10.xlsx"
    fill(tpl_swap, loads("{'B':{l:'0-0',v:'来自A'}, 'C':{l:'0-0',v:'来自B'}}"), out10)
    wb10 = load_workbook(out10)
    check("表名互换", wb10.sheetnames == ["B", "C"], str(wb10.sheetnames))
    check("互换后内容正确", wb10["B"]["A1"].value == "来自A" and wb10["C"]["A1"].value == "来自B")

    # 10) 分组模式下 -s 参数被忽略
    r11 = fill(template, loads("{'客户信息':{l:'2-0',v:'x'}}"), tmp / "map11.xlsx", sheet="说明")
    check("分组模式忽略 -s", any("-s" in w for w in r11.warnings))


def main() -> int:
    test_parser()
    test_positions()
    test_normalize()
    tmp = Path(tempfile.mkdtemp(prefix="writeexcel-test-"))
    try:
        test_excel(tmp)
        test_sheet_map(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print(f"\n通过 {PASSED} 项")
    if FAILED:
        print(f"失败 {len(FAILED)} 项：")
        for item in FAILED:
            print(f"  - {item}")
        return 1
    print("全部测试通过（OK）")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        raise SystemExit(1)
