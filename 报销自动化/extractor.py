# -*- coding: utf-8 -*-
"""从发票文件名解析报销明细：文件名格式「物品+数量+单价+金额」。

例如：表演服+1件+68元+68元.pdf -> 名称=表演服, 数量=1件, 单价=68, 金额=68
     手机支架+1个+20元+20元.jpg -> 名称=手机支架, 数量=1个, 单价=20, 金额=20
"""
import os
import re

from money import parse_money

# PDF 渲染为图片（用于把发票 PDF 嵌入 Word）
try:
    import pypdfium2 as pdfium
except Exception:  # pragma: no cover
    pdfium = None


def render_pdf_to_images(path, scale: float = 3.0):
    """把 PDF 每页渲染成 PIL.Image，scale 越大越清晰（3.0 ≈ 216dpi）。"""
    if pdfium is None:
        raise RuntimeError("缺少 pypdfium2，无法渲染 PDF")
    pdf = pdfium.PdfDocument(path)
    images = []
    for i in range(len(pdf)):
        bitmap = pdf[i].render(scale=scale)
        images.append(bitmap.to_pil())
    return images


def load_image(path):
    from PIL import Image
    return Image.open(path).convert("RGB")


def _to_halfwidth(s: str) -> str:
    """全角数字/加号/小数点/人民币符号转半角。"""
    out = []
    for ch in s:
        code = ord(ch)
        if code == 0xFF0B:            # ＋
            out.append("+")
        elif code == 0xFF0E:          # ．
            out.append(".")
        elif 0xFF10 <= code <= 0xFF19:  # ０-９
            out.append(chr(code - 0xFEE0))
        elif code == 0xFFE5:          # ￥
            out.append("¥")
        else:
            out.append(ch)
    return "".join(out)


def _parse_number(s):
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return None


_TAXI_KEYWORDS = ("打车", "滴滴", "出租车", "的士", "网约车", "出行", "交通费", "车费", "行程", "T3出行", "曹操出行", "高德打车")


def detect_category(name: str) -> str:
    """根据名称判断费用类别：打车费 / 采购费。"""
    if any(k in name for k in _TAXI_KEYWORDS):
        return "打车费"
    return "采购费"


def parse_filename(filename: str) -> dict:
    """解析「物品+数量+单价+金额(.扩展名)」，返回一条明细 dict。"""
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = _to_halfwidth(stem)
    parts = [p.strip() for p in stem.split("+") if p.strip()]

    item = {"category": "采购费", "name": "", "quantity": 1, "unit": "个",
            "amount": None, "unit_price": None}

    if len(parts) >= 4:
        name = "+".join(parts[:-3])   # 名称里可能含 +，合并回来
        qty_str, price_str, amount_str = parts[-3], parts[-2], parts[-1]
    elif len(parts) == 3:             # 物品+数量+金额（缺单价）
        name, qty_str, price_str, amount_str = parts[0], parts[1], None, parts[2]
    elif len(parts) == 2:             # 物品+金额
        name, qty_str, price_str, amount_str = parts[0], None, None, parts[1]
    else:
        name, qty_str, price_str, amount_str = (parts[0] if parts else ""), None, None, None

    item["name"] = name.strip() or "物资"

    # 数量 + 单位，如 "1件" / "6个" / "2双" / "1箱"
    if qty_str:
        m = re.match(r"(\d+(?:\.\d+)?)\s*([^\d]*)", qty_str)
        if m:
            item["quantity"] = _parse_number(m.group(1)) or 1
            unit = m.group(2).strip()
            if unit:
                item["unit"] = unit[0]  # 取第一个字符（个/件/双/箱…）

    if price_str:
        item["unit_price"] = parse_money(price_str)
    if amount_str:
        item["amount"] = parse_money(amount_str)

    # 金额缺失时，用 单价×数量 推算；单价缺失时，用 金额÷数量 推算
    if item["amount"] is None and item["unit_price"] is not None:
        item["amount"] = round(item["unit_price"] * item["quantity"], 2)
    if item["unit_price"] is None and item["amount"] is not None and item["quantity"]:
        item["unit_price"] = round(item["amount"] / item["quantity"], 2)

    item["category"] = detect_category(item["name"])
    return item


if __name__ == "__main__":
    for fn in ["表演服+1件+68元+68元.pdf", "手机支架+1个+20元+20元.jpg",
               "胸针+6个+3.79元+22.76元.pdf", "瓶装水+1箱+11.8元+11.8元.jpg"]:
        print(fn, "->", parse_filename(fn))
