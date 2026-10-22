# -*- coding: utf-8 -*-
"""Đọc thư mục PACKING: mỗi model một file PDF "HƯỚNG DẪN ĐÓNG GÓI" (tiếng Việt).

Mỗi PDF gồm:
  - trang sơ đồ đóng gói (ảnh chụp thùng, có đánh số)
  - trang "Vật liệu và quy cách đóng gói": bảng vật liệu, bảng chú thích số, (tuỳ chọn) lưu ý đóng gói
  - các trang ảnh phóng to / ảnh minh hoạ bổ sung
Từ đó lấy: ảnh (ghi ra WebP), bảng vật liệu, chú thích số, lưu ý. Chữ đã là tiếng Việt nên không cần dịch.

Dùng từ build_bom_data.py:  from packing import build_packing
"""
import hashlib
import io
import re
import shutil
import unicodedata
from pathlib import Path

RE_NUM = re.compile(r"^\d{1,2}$")
SOFT = "­"


# ----------------------------------------------------------------- văn bản
def _nfc(s):
    return unicodedata.normalize("NFC", s or "").replace(" ", " ")


def _clean(s):
    return re.sub(r"\s+", " ", _nfc(s).replace(SOFT, "-")).strip()


def _words(page):
    return [(w[0], w[1], w[2], w[3], _nfc(w[4]).replace(SOFT, "-")) for w in page.get_text("words")]


def _line_y(words, text):
    """y0 của từ đầu tiên có nội dung `text` (so khớp từ đầu dòng tiêu đề)."""
    for w in sorted(words, key=lambda w: (w[1], w[0])):
        if w[4] == text:
            return w
    return None


def _join(ws):
    ws = sorted(ws, key=lambda w: (round(w[1] / 4), w[0]))
    return _clean(" ".join(w[4] for w in ws))


def parse_meta(page):
    t = _nfc(page.get_text())
    m = re.search(r"Model công ty:\s*([^|\n]*)", t)
    return _clean(m.group(1)) if m else ""


def parse_guide_page(page):
    """Trang 'Vật liệu và quy cách đóng gói' -> (materials, legend, notes)."""
    words = _words(page)
    footer = _line_y(words, "TÀI")
    fy = footer[1] if footer else page.rect.height
    d_mat = next((w for w in words if w[4] == "DANH"), None)
    d_leg = next((w for w in words if w[4] == "CHÚ"), None)
    d_note = next((w for w in words if w[4] == "LƯU"), None)
    materials, legend, notes = [], [], []

    # ---- bảng vật liệu
    if d_mat:
        end = min([w[1] for w in (d_leg, d_note) if w] + [fy])
        zone = [w for w in words if d_mat[3] <= w[1] < end - 1]
        hdr_y = next((w[1] for w in zone if w[4] == "STT"), None)
        if hdr_y is not None:
            hdr = [w for w in zone if abs(w[1] - hdr_y) < 4]
            x_stt = next(w for w in hdr if w[4] == "STT")
            x_name = next((w for w in hdr if w[4] == "Tên"), None)
            x_spec = next((w for w in hdr if w[4] == "Quy"), None)
            x_qty = next((w for w in hdr if w[4] == "SL"), None)
            body = [w for w in zone if w[1] > hdr_y + 8]
            nums = sorted([w for w in body if abs(w[0] - x_stt[0]) < 14 and RE_NUM.match(w[4])], key=lambda w: w[1])
            for i, n in enumerate(nums):
                y0 = n[1] - 3
                y1 = nums[i + 1][1] - 3 if i + 1 < len(nums) else end
                row = [w for w in body if y0 <= w[1] < y1 and w is not n]
                nm = _join([w for w in row if x_name[0] - 6 <= w[0] < x_spec[0] - 6])
                sp = _join([w for w in row if x_spec[0] - 6 <= w[0] < x_qty[0] - 28])
                q = _join([w for w in row if w[0] >= x_qty[0] - 28])
                if nm:
                    materials.append({"no": int(n[4]), "name": nm, "spec": sp, "qty": q, "cancelled": bool(re.search(r"hủy|huỷ|hủy bỏ", q, re.I))})

    # ---- chú thích số (2 cặp cột STT / Bộ phận)
    if d_leg:
        end = min([w[1] for w in (d_note,) if w] + [fy])
        zone = [w for w in words if d_leg[3] <= w[1] < end - 1]
        hdr_y = next((w[1] for w in zone if w[4] == "STT"), None)
        if hdr_y is not None:
            hdr = sorted([w for w in zone if abs(w[1] - hdr_y) < 4 and w[4] in ("STT", "Bộ")], key=lambda w: w[0])
            stts = [w for w in hdr if w[4] == "STT"]
            body = [w for w in zone if w[1] > hdr_y + 8]
            cols = [(s[0], stts[i + 1][0] if i + 1 < len(stts) else 10 ** 6) for i, s in enumerate(stts)]
            for cx0, cx1 in cols:
                nums = sorted([w for w in body if abs(w[0] - cx0) < 14 and RE_NUM.match(w[4])], key=lambda w: w[1])
                for n in nums:
                    row = [w for w in body if abs(w[1] - n[1]) < 5 and cx0 + 20 <= w[0] < cx1 - 4 and w is not n]
                    nm = _join(row)
                    if nm:
                        legend.append({"no": int(n[4]), "name": nm})
            legend.sort(key=lambda r: r["no"])

    # ---- lưu ý: mỗi dòng là một ý; dòng bắt đầu bằng chữ thường là phần nối tiếp của ý trước
    if d_note:
        zone = [w for w in words if d_note[3] <= w[1] < fy - 1]
        lines = {}
        for w in zone:
            lines.setdefault(round(w[1] / 4), []).append(w)
        for k in sorted(lines):
            t = _join(lines[k])
            if not t or t.startswith("LƯU Ý"):
                continue
            if notes and t[0].islower():
                notes[-1] += " " + t
            else:
                notes.append(t)
    return materials, legend, notes


# ----------------------------------------------------------------- ảnh
def _ahash(im):
    g = im.convert("L").resize((16, 16))
    px = list(g.getdata())
    avg = sum(px) / len(px)
    return sum(1 << i for i, v in enumerate(px) if v >= avg)


def _hamming(a, b):
    return bin(a ^ b).count("1")


def collect_images(doc):
    """Ảnh nhúng trong PDF, bỏ trùng (ảnh trang sơ đồ và ảnh phóng to là cùng một ảnh -> giữ bản lớn)."""
    from PIL import Image
    found = []  # [(order, im, hash)]
    seen_xref = set()
    order = 0
    for p in doc:
        for info in p.get_images(full=True):
            xref = info[0]
            if xref in seen_xref:
                continue
            seen_xref.add(xref)
            try:
                raw = doc.extract_image(xref)
                im = Image.open(io.BytesIO(raw["image"]))
                im.load()
            except Exception:  # noqa
                continue
            if im.width < 120 or im.height < 120:
                continue
            found.append([order, im, _ahash(im)])
            order += 1
    keep = []
    for item in found:
        dup = next((k for k in keep if _hamming(k[2], item[2]) <= 8), None)
        if dup is None:
            keep.append(item)
        elif item[1].width * item[1].height > dup[1].width * dup[1].height:
            dup[1] = item[1]  # giữ bản lớn hơn nhưng giữ vị trí của lần xuất hiện đầu
    keep.sort(key=lambda k: k[0])
    return [k[1] for k in keep]


def _save_webp(im, dest, max_side=1400, quality=78):
    from PIL import Image
    im = im.convert("RGB")
    im.thumbnail((max_side, max_side), Image.LANCZOS)
    dest.parent.mkdir(parents=True, exist_ok=True)
    im.save(dest, "WEBP", quality=quality, method=6)
    return dest.stat().st_size


# ----------------------------------------------------------------- chính
def build_packing(pack_dir, out_dir, models, report):
    """Trả về {model: {"imgs": [...], "materials": [...], "legend": [...], "notes": [...], "company": str, "src": str}}."""
    result = {}
    if not pack_dir.is_dir():
        report.append("Đóng gói: không thấy thư mục " + str(pack_dir))
        return result
    import fitz
    shutil.rmtree(out_dir, ignore_errors=True)
    total, errors = 0, []
    pdfs = sorted(pack_dir.glob("*.pdf"))
    unknown = [p.stem for p in pdfs if p.stem not in models]
    for f in pdfs:
        m = f.stem
        if m not in models:
            continue
        try:
            doc = fitz.open(str(f))
            materials, legend, notes, company = [], [], [], ""
            for p in doc:
                company = company or parse_meta(p)
                if "Vật liệu và quy cách" in _nfc(p.get_text()):
                    mt, lg, nt = parse_guide_page(p)
                    materials += mt
                    legend += lg
                    notes += nt
            images = collect_images(doc)
            doc.close()
        except Exception as e:  # noqa
            errors.append(f"{f.name}: {e!r}")
            continue
        outs = []
        for i, im in enumerate(images, 1):
            dest = out_dir / f"{m}-{i}.webp"
            total += _save_webp(im, dest)
            outs.append(f"PACKING_WEB/{m}-{i}.webp")
        result[m] = {"imgs": outs, "materials": materials, "legend": legend, "notes": notes, "company": company, "src": f.name}
    missing = sorted(models - set(result))
    report.append(f"Đóng gói: {len(result)}/{len(models)} model có hướng dẫn; ảnh WebP {total / 1048576:.1f} MB ({sum(len(v['imgs']) for v in result.values())} ảnh)")
    report.append("Đóng gói - model chưa có file PDF: " + str(missing))
    report.append("Đóng gói - file không khớp model nào (bỏ qua): " + str(unknown))
    report.append("Đóng gói - có bảng vật liệu: %d model; có chú thích số: %d model; có lưu ý: %d model" % (
        sum(1 for v in result.values() if v["materials"]), sum(1 for v in result.values() if v["legend"]), sum(1 for v in result.values() if v["notes"])))
    report.append("Đóng gói - không có ảnh: " + str(sorted(k for k, v in result.items() if not v["imgs"])))
    if errors:
        report.append("Đóng gói - file lỗi: " + str(errors))
    return result
