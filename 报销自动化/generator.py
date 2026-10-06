# -*- coding: utf-8 -*-
"""根据模板生成三个文件：支出单(xlsx)、报销说明(docx)、发票订单明细(docx)。"""
import os
import copy
import re
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

from money import rmb_upper, fmt_money

BASE = os.path.dirname(os.path.abspath(__file__))
TPL_XLSX = os.path.join(BASE, "templates", "支出单模板.xlsx")
TPL_SHOU = os.path.join(BASE, "templates", "报销说明模板.docx")

DATA_START = 11        # 支出单数据区起始行（第 11 行）
TPL_DATA_ROWS = 8      # 模板自带的数据行数


# ---------------------------------------------------------------- 工具
def sum_amounts(items) -> float:
    total = Decimal("0")
    for it in items:
        a = it.get("amount")
        if a not in (None, ""):
            total += Decimal(str(a))
    return float(total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _as_int(v, default):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _num(v):
    """数字显示：整数值去掉小数尾，如 59.0->59、39.8->39.8。"""
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError):
        return v


def _summary(item) -> str:
    name = item.get("name") or "物资"
    qty = item.get("quantity") or 1
    unit = item.get("unit") or "个"
    q = _num(qty)
    return f"购买{name} {q}{unit}"


def _note(item) -> str:
    up = item.get("unit_price")
    unit = item.get("unit") or "个"
    if up is None:
        return ""
    return f"单价{fmt_money(up)}元/{unit}"


def _put(cell, value):
    """写入值并把字体强制设为黑色（模板里待填格可能是红色）。"""
    cell.value = value
    f = copy.copy(cell.font)
    f.color = "FF000000"
    cell.font = f
    return cell


# ---------------------------------------------------------------- 支出单 xlsx
def generate_xlsx(meta, items, out_path):
    wb = openpyxl.load_workbook(TPL_XLSX)

    # 找到主表
    main = None
    for name in wb.sheetnames:
        ws = wb[name]
        if ws["A1"].value and "费用支出单" in str(ws["A1"].value):
            main = ws
            break
    if main is None:
        main = wb[wb.sheetnames[0]]
    # 删除其它空表
    for name in [n for n in wb.sheetnames if wb[n] is not main]:
        del wb[name]
    ws = main

    # ---- 表头信息 ----
    y, m, d = meta["date_year"], meta["date_month"], meta["date_day"]
    _put(ws["B2"], f"{y}/{m:02d}/{d:02d}")                 # 填表日期（2025/10/27 格式）
    _put(ws["B3"], meta.get("name", ""))                   # 报销人
    _put(ws["D3"], meta.get("department", "校团委社团管理部"))
    _put(ws["F3"], meta.get("phone", ""))                  # 联系电话
    _put(ws["D4"], _as_int(meta.get("ticket_count"), len(items)))  # 票据张数
    _put(ws["F4"], meta.get("activity", ""))               # 活动事由
    _put(ws["B5"], f"{meta.get('name','')}\n{meta.get('student_id','')}")  # 姓名+学号

    # ---- 调整数据区行数 ----
    # 注意：openpyxl 的 insert_rows/delete_rows 会移动单元格值，但不会移动合并区域，
    # 因此需要先记录、取消合并，行数调整后再按新位置重新合并。
    n = len(items)
    shift = n - TPL_DATA_ROWS
    total_row = DATA_START + n          # 金额合计所在行
    if shift != 0:
        all_merges = list(ws.merged_cells.ranges)
        data_merges = [m for m in all_merges if DATA_START <= m.min_row < DATA_START + TPL_DATA_ROWS]
        fixed_merges = [m for m in all_merges if m.min_row >= DATA_START + TPL_DATA_ROWS]
        for m in data_merges + fixed_merges:
            ws.unmerge_cells(str(m))
        if shift > 0:
            ws.insert_rows(DATA_START + TPL_DATA_ROWS, shift)
            for r in range(DATA_START + TPL_DATA_ROWS, DATA_START + TPL_DATA_ROWS + shift):
                _copy_row_style(ws, DATA_START, r)
                if DATA_START in ws.row_dimensions:
                    ws.row_dimensions[r].height = ws.row_dimensions[DATA_START].height
        else:
            ws.delete_rows(DATA_START + n, -shift)
        # 重新合并固定区（金额合计及以下），整体平移 shift
        for m in fixed_merges:
            ws.merge_cells(start_row=m.min_row + shift, start_column=m.min_col,
                           end_row=m.max_row + shift, end_column=m.max_col)
    # 数据行 B:D 合并（每条明细）
    for i in range(n):
        r = DATA_START + i
        if not _is_merged(ws, r, 2):
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

    # ---- 填充数据行 ----
    for i, item in enumerate(items):
        r = DATA_START + i
        _put(ws.cell(row=r, column=1), item.get("category", "采购费"))
        _put(ws.cell(row=r, column=2), _summary(item))
        amt = item.get("amount")
        _put(ws.cell(row=r, column=5), _num(amt) if amt not in (None, "") else "")
        _put(ws.cell(row=r, column=6), _note(item))

    # ---- 金额合计 ----
    total = sum_amounts(items)
    _put(ws.cell(row=total_row, column=2), f"{rmb_upper(total)} ￥{fmt_money(total)}元")

    wb.save(out_path)
    return total


def _is_merged(ws, row, col) -> bool:
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return True
    return False


def _copy_row_style(ws, src_row, dst_row):
    """把源行的单元格样式（边框/字体/对齐/数字格式）复制到目标行（1..6 列）。"""
    for col in range(1, 7):
        s = ws.cell(row=src_row, column=col)
        d = ws.cell(row=dst_row, column=col)
        if s.has_style:
            d.font = copy.copy(s.font)
            d.border = copy.copy(s.border)
            d.fill = copy.copy(s.fill)
            d.alignment = copy.copy(s.alignment)
            d.number_format = s.number_format
            d.protection = copy.copy(s.protection)


# ---------------------------------------------------------------- 报销说明 docx
def _set_cell_text(cell, text):
    p = cell.paragraphs[0]
    runs = p.runs
    text = str(text)
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._element.getparent().remove(r._element)
    else:
        run = p.add_run(text)
        run.font.name = "宋体"
        run.font.size = Pt(12)
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = rpr.makeelement(qn('w:rFonts'), {})
            rpr.append(rf)
        rf.set(qn('w:eastAsia'), '宋体')


def _set_para_text(para, text):
    runs = para.runs
    if runs:
        runs[0].text = str(text)
        for r in runs[1:]:
            r._element.getparent().remove(r._element)
    else:
        para.add_run(str(text))


def generate_shuoming(meta, items, narrative, out_path):
    doc = Document(TPL_SHOU)
    paras = doc.paragraphs
    # P3 / P4 为叙事段落
    _set_para_text(paras[3], narrative.get("p1", ""))
    _set_para_text(paras[4], narrative.get("p2", ""))

    tbl = doc.tables[0]
    # 模板表结构：第0-1行表头(纵向合并)，第2-12行数据(11行)，第13行合计
    total = sum_amounts(items)
    tpl_data_rows = 11
    n = len(items)
    # 删除多余数据行（从后往前删）
    if n < tpl_data_rows:
        for _ in range(tpl_data_rows - n):
            tr = tbl.rows[2 + n]._tr
            tr.getparent().remove(tr)
    # 插入更多数据行（克隆一个数据行，插入到合计行之前）
    elif n > tpl_data_rows:
        total_tr = tbl.rows[tpl_data_rows + 2]._tr  # 合计行元素（2 行表头 + 11 行数据 = index 13）
        for _ in range(n - tpl_data_rows):
            new_tr = copy.deepcopy(tbl.rows[2]._tr)
            total_tr.addprevious(new_tr)

    # 填充数据行
    for i, item in enumerate(items):
        row = tbl.rows[2 + i]
        cells = row.cells
        _set_cell_text(cells[0], i + 1)
        _set_cell_text(cells[1], f"购买{item.get('name','')}")
        _set_cell_text(cells[2], _num(item.get("quantity") or 1))
        _set_cell_text(cells[3], fmt_money(item.get("unit_price")) if item.get("unit_price") is not None else "")
        amt = item.get("amount")
        _set_cell_text(cells[4], fmt_money(amt) if amt not in (None, "") else "")

    # 合计行（首格 gridSpan=4 的「总计」+ 金额格 + 空列），直接按 <w:tc> 定位金额格
    total_row = tbl.rows[-1]
    from docx.table import _Cell
    tcs = total_row._tr.tc_lst
    _set_cell_text(_Cell(tcs[1], total_row), fmt_money(total))

    # 结尾段落：以上共计 / 日期（按内容定位，避免依赖段落序号）
    y, m, d = meta["date_year"], meta["date_month"], meta["date_day"]
    for p in doc.paragraphs:
        if "以上共计" in p.text:
            _set_para_text(p, f"以上共计：{fmt_money(total)}元。")
            break
    for p in doc.paragraphs:
        if re.fullmatch(r"\s*\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*", p.text):
            _set_para_text(p, f"                                {y}年 {m}月 {d}日")
            break

    doc.save(out_path)
    return total


# ---------------------------------------------------------------- 发票订单明细 docx
def _add_fitted_image(doc, path, max_w_cm=18.0, max_h_cm=26.5, rotate_deg=0):
    """把图片按最大宽高等比缩放进文档，rotate_deg 为顺/逆时针旋转角度（顺时针用负数）。"""
    from PIL import Image
    from io import BytesIO
    im = Image.open(path).convert("RGB")
    if rotate_deg:
        im = im.rotate(rotate_deg, expand=True)
    w, h = im.size
    ar = w / h
    w_cm = max_w_cm
    h_cm = w_cm / ar
    if h_cm > max_h_cm:
        h_cm = max_h_cm
        w_cm = h_cm * ar
    buf = BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    doc.add_picture(buf, width=Cm(w_cm))


def generate_fapiao(meta, items, invoice_imgs, order_imgs, out_path):
    """生成【发票】【订单明细】：一张发票 + 对应订单，以此类推。

    invoice_imgs / order_imgs 为图片路径列表，按索引配对。
    页面为 A4 竖版；发票旋转 90° 后占满竖版页面，订单截图不旋转。
    """
    doc = Document()
    # 页面设置：A4 竖版
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(1.5)
    sec.right_margin = Cm(1.5)
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)

    n = len(invoice_imgs)
    for i in range(n):
        # 每张发票单独起一页，旋转 90° 占满竖版页面
        if i > 0:
            doc.add_page_break()
        if os.path.exists(invoice_imgs[i]):
            _add_fitted_image(doc, invoice_imgs[i], rotate_deg=-90)
        if i < len(order_imgs) and os.path.exists(order_imgs[i]):
            # 订单截图紧跟发票之后，不旋转，空间不足会自动换页
            doc.add_paragraph()
            _add_fitted_image(doc, order_imgs[i])

    doc.save(out_path)
