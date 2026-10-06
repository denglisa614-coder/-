# -*- coding: utf-8 -*-
"""社团报销自动化：上传按「物品+数量+单价+金额」命名的发票 → 生成三个文件。"""
import os
import re
import io
import uuid
import base64
import zipfile
import datetime

from flask import Flask, request, jsonify, send_file, send_from_directory

from extractor import parse_filename, render_pdf_to_images, load_image
from generator import generate_xlsx, generate_shuoming, generate_fapiao, sum_amounts
from money import fmt_money

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE, "uploads")
OUTPUT_DIR = os.path.join(BASE, "output")
for d in (UPLOAD_DIR, OUTPUT_DIR):
    os.makedirs(d, exist_ok=True)

app = Flask(__name__, static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024  # 300MB 上限


def _save_upload(file, sid, prefix, idx):
    """保存上传文件，PDF 渲染成 PNG，图片统一转 PNG，返回 PNG 路径（用于嵌入文档）。"""
    ext = os.path.splitext(file.filename)[1].lower()
    data = file.read()
    if ext == ".pdf":
        tmp = os.path.join(UPLOAD_DIR, sid, f"{prefix}{idx}.pdf")
        with open(tmp, "wb") as f:
            f.write(data)
        imgs = render_pdf_to_images(tmp)
        if imgs:
            out = os.path.join(UPLOAD_DIR, sid, f"{prefix}{idx}.png")
            imgs[0].save(out)  # 单页发票；多页只取第一页
            return out
        return None
    img = load_image(io.BytesIO(data))
    out = os.path.join(UPLOAD_DIR, sid, f"{prefix}{idx}.png")
    img.save(out)
    return out


def _thumb_base64(path, max_w=420):
    from PIL import Image
    try:
        with Image.open(path) as im:
            im = im.convert("RGB")
            w, h = im.size
            if w > max_w:
                h = int(h * max_w / w)
                w = max_w
                im = im.resize((w, h))
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=72)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ""


def _sanitize(name):
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def _build_narrative(meta, total):
    activity = meta.get("activity", "") or "社团活动"
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", activity)
    if m:
        date_str = f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    else:
        date_str = f"{meta['date_year']}年{meta['date_month']}月{meta['date_day']}日"
    p1 = f"{date_str}，为开展「{activity}」活动，产生了物资采购等相关费用。"
    p2 = f"为保障活动顺利进行，共计采购 {fmt_money(total)} 元物资，具体如下"
    return {"p1": p1, "p2": p2}


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/extract", methods=["POST"])
def api_extract():
    invoices = [f for f in request.files.getlist("invoices") if f and f.filename]
    orders = [f for f in request.files.getlist("orders") if f and f.filename]

    meta = {
        "name": request.form.get("name", "").strip(),
        "student_id": request.form.get("student_id", "").strip(),
        "phone": request.form.get("phone", "").strip(),
        "department": request.form.get("department", "校团委社团管理部").strip() or "校团委社团管理部",
        "activity": request.form.get("activity", "").strip(),
        "ticket_count": request.form.get("ticket_count", "").strip() or len(invoices),
    }
    today = datetime.date.today()
    meta["date_year"], meta["date_month"], meta["date_day"] = today.year, today.month, today.day

    sid = uuid.uuid4().hex[:12]
    sid_dir = os.path.join(UPLOAD_DIR, sid)
    os.makedirs(sid_dir, exist_ok=True)

    # 保存发票 / 订单，并从发票「文件名」解析明细
    inv_imgs, ord_imgs, items = [], [], []
    for i, f in enumerate(invoices):
        try:
            png = _save_upload(f, sid, "invoice_", i)
        except Exception:
            png = None
        inv_imgs.append(png)
        item = parse_filename(f.filename)
        item["invoice_index"] = i
        items.append(item)
    for i, f in enumerate(orders):
        try:
            png = _save_upload(f, sid, "order_", i)
        except Exception:
            png = None
        ord_imgs.append(png)

    total = sum_amounts(items)
    narrative = _build_narrative(meta, total)

    return jsonify({
        "session_id": sid,
        "meta": meta,
        "items": items,
        "total": total,
        "narrative": narrative,
        "thumbnails": [_thumb_base64(p) if p else "" for p in inv_imgs],
        "invoice_count": len(invoices),
        "order_count": len(orders),
    })


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    sid = data.get("session_id", "")
    meta = data.get("meta", {})
    items = data.get("items", [])
    narrative = data.get("narrative", {"p1": "", "p2": ""})

    # 归一化 items
    norm = []
    for it in items:
        amt = it.get("amount")
        amt = float(amt) if amt not in (None, "") else None
        up = it.get("unit_price")
        up = float(up) if up not in (None, "") else None
        qty = it.get("quantity") or 1
        try:
            qty = float(qty)
            qty = int(qty) if qty == int(qty) else qty
        except (TypeError, ValueError):
            qty = 1
        if up is None and amt is not None and qty:
            up = round(amt / qty, 2)
        norm.append({
            "category": it.get("category", "采购费"),
            "name": (it.get("name") or "").strip(),
            "quantity": qty,
            "unit": (it.get("unit") or "个").strip(),
            "amount": amt,
            "unit_price": up,
        })

    meta["date_year"] = int(meta.get("date_year") or datetime.date.today().year)
    meta["date_month"] = int(meta.get("date_month") or datetime.date.today().month)
    meta["date_day"] = int(meta.get("date_day") or datetime.date.today().day)

    sid_dir = os.path.join(UPLOAD_DIR, sid)
    inv_imgs = [os.path.join(sid_dir, f"invoice_{i}.png") for i in range(len(norm))]
    ord_imgs = [os.path.join(sid_dir, f"order_{i}.png") for i in range(len(norm))]

    # 文件名
    activity = _sanitize(meta.get("activity", "") or "活动")
    name = _sanitize(meta.get("name", "") or "报销人")
    out_dir = os.path.join(OUTPUT_DIR, sid)
    os.makedirs(out_dir, exist_ok=True)
    f_zhi = os.path.join(out_dir, f"【支出单】{name}+{activity}.xlsx")
    f_shou = os.path.join(out_dir, f"【报销说明】{name}+{activity}.docx")
    f_fa = os.path.join(out_dir, f"【发票】【订单明细】{name}+{activity}.docx")

    generate_xlsx(meta, norm, f_zhi)
    generate_shuoming(meta, norm, narrative, f_shou)
    generate_fapiao(meta, norm, inv_imgs, ord_imgs, f_fa)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in (f_zhi, f_shou, f_fa):
            zf.write(fp, os.path.basename(fp))
    buf.seek(0)
    zip_name = f"报销材料_{name}+{activity}.zip"
    return send_file(buf, mimetype="application/zip", as_attachment=True, download_name=zip_name)


if __name__ == "__main__":
    print("报销自动化已启动：http://127.0.0.1:5000")
    # 0.0.0.0 使同一局域网内其它设备也能访问（本机仍可用 127.0.0.1）
    app.run(host="0.0.0.0", port=5000, debug=False)
