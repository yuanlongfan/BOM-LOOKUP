# -*- coding: utf-8 -*-
"""
Đọc toàn bộ file BOM (*.xlsx) và nhúng dữ liệu + ảnh (thư mục PICTURE) thẳng vào BOM_Lookup.html để file chạy độc lập ở bất kỳ đâu.

Cách dùng:
    python build_bom_data.py [--bom "D:\\10.Project\\BOM"] [--html "D:\\10.Project\\Claude\\BOM_Lookup.html"] [--pics "D:\\10.Project\\Claude\\PICTURE"]
Mặc định: thư mục BOM = <script>\\..\\..\\BOM, HTML = <script>\\..\\BOM_Lookup.html, ảnh = <script>\\..\\PICTURE
Chỉ ĐỌC thư mục BOM, không ghi gì vào đó.
"""
import argparse
import csv
import datetime as dt
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent

# ----------------------------------------------------------------- hằng số
CATEGORIES = [  # thứ tự: dài trước, ngắn sau
    ("室內吸頂扇", "Quạt ốp trần trong nhà"),
    ("戶外吸頂扇", "Quạt ốp trần ngoài trời"),
    ("室內扇", "Quạt trong nhà"),
    ("戶外扇", "Quạt ngoài trời"),
    ("潮濕扇", "Quạt nơi ẩm ướt"),
]
COLORS = {
    "黑": ("đen", "#1a1a1a"),
    "白": ("trắng", "#f5f5f5"),
    "藍": ("xanh dương", "#1e5bd8"),
    "綠": ("xanh lá", "#1f9d55"),
    "紅": ("đỏ", "#d62828"),
    "黃": ("vàng", "#f2c94c"),
    "棕": ("nâu", "#7b4a2a"),
    "灰": ("xám", "#8a8a8a"),
    "橙": ("cam", "#f08c2b"),
    "紫": ("tím", "#7a3fb0"),
}
COLOR_CHARS = "".join(COLORS)
RE_COLOR_SEG = re.compile(r"^[%s]+$" % COLOR_CHARS)
RE_COLOR_ONE = re.compile(r"^[%s]$" % COLOR_CHARS)
RE_WIRE = re.compile(r"^\d{3,4}#\d+$")
RE_LEN = re.compile(r"^(\d+(?:\.\d+)?)cm$", re.I)
RE_VI = re.compile(
    r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]",
    re.I,
)
RE_CORE = re.compile(r"^粗胚-(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)$")
STATOR_MAP = {"153": "153x130", "172": "172x145", "188": "188x161"}
RE_CJK = re.compile(r"[\u4e00-\u9fff]")


def clean(v):
    if v is None:
        return ""
    return str(v).replace("\n", "").replace("\r", "").strip()


RE_VI_ONLY = re.compile(r"[ăâđêôơưạảẫậằắẳẵặẹẻẽềếểễệỉịọỏồốổỗộờớởỡợụủứừửữựỳỵỷỹ]", re.I)
RE_PINYIN_TONE = re.compile(r"[āēīōūǎěǐǒǔǖǘǚǜ]")


def is_real_vi(t):
    """True nếu chuỗi là tiếng Việt thật (không phải pinyin)."""
    if RE_PINYIN_TONE.search(t):
        return False
    return bool(RE_VI_ONLY.search(t))


# ----------------------------------------------------------------- đọc file
def read_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr_i = None
    for i, r in enumerate(all_rows[:10]):
        if "料件編號" in r:
            hdr_i = i
            break
    if hdr_i is None:
        raise ValueError("Không tìm thấy hàng tiêu đề 料件編號")
    h = list(all_rows[hdr_i])
    ci_code, ci_name, ci_spec = h.index("料件編號"), h.index("品名"), h.index("規格")
    ci_qty = h.index("用量") if "用量" in h else None
    ci_unit = h.index("生產單位") if "生產單位" in h else None
    ci_src = h.index("補給策略") if "補給策略" in h else None
    ci_sup = h.index("供應商名稱") if "供應商名稱" in h else None
    out = []
    for r in all_rows[hdr_i + 1:]:
        def cell(i):
            return clean(r[i]) if i is not None and i < len(r) else ""
        name, spec = cell(ci_name), cell(ci_spec)
        if not name and not spec:
            continue
        vi = cell(ci_spec - 1) if ci_spec - 1 >= 0 else ""
        spec_vi = vi if (vi and vi != spec and RE_VI.search(vi)) else None
        qty = r[ci_qty] if ci_qty is not None and ci_qty < len(r) else None
        # cấp BOM = chỉ số cột đầu tiên có dữ liệu, trước cột 品名
        lv = next((i for i in range(ci_name) if i < len(r) and clean(r[i])), ci_code)
        # tên tiếng Việt thật (nếu có) nằm giữa 品名 và cột quy cách; bỏ pinyin
        vi_name = ""
        for i in range(ci_name + 1, min(ci_spec, len(r))):
            t = cell(i)
            if t and is_real_vi(t):
                vi_name = t
                break
        out.append({"code": cell(lv) or cell(ci_code), "name": name, "qty": qty, "spec": spec, "spec_vi": spec_vi,
                    "lv": lv, "unit": cell(ci_unit), "src": cell(ci_src), "sup": cell(ci_sup), "vi_name": vi_name})
    return out


def split_name(stem):
    m = re.match(r"^(.*)\((.+)\)\s*$", stem)
    if not m:
        return stem, stem
    return m.group(1).strip(), m.group(2).strip()


def category_of(fan_name):
    for cn, vi in CATEGORIES:
        if cn in fan_name:
            code = fan_name.replace(cn, "", 1).strip("-").strip()
            return {"cn": cn, "vi": vi}, code
    return {"cn": "", "vi": "Khác"}, fan_name


def color_obj(ch):
    vi, hx = COLORS[ch]
    return {"cn": ch, "vi": vi, "hex": hx}


# ----------------------------------------------------------------- dây nguồn
def parse_cord_spec(spec):
    segs = spec.split("-")
    wire = next((s for s in segs if RE_WIRE.match(s)), None)
    length = next((s for s in segs if RE_LEN.match(s)), None)
    cseg = next((s for s in segs if RE_COLOR_SEG.match(s)), None)
    colors = []
    if cseg:
        for ch in cseg:
            if ch not in colors:
                colors.append(ch)
    return wire, length, colors


def extract_cord(rows, warnings):
    sets = [r for r in rows if r["name"].startswith("電源線組")]
    if sets:
        first = sets[0]
        wire, length, colors = parse_cord_spec(first["spec"])
        res = {
            "found": True, "source": "cord_set",
            "raw": first["spec"], "rawVi": first["spec_vi"],
            "wireType": wire, "length": length,
            "colors": [color_obj(c) for c in colors], "wireCount": len(colors),
            "extra": [{"raw": s["spec"], "rawVi": s["spec_vi"]} for s in sets[1:]],
        }
        if not colors or not wire or not length:
            warnings.append("Dây nguồn: không tách đủ quy cách (loại/dài/màu) – cần kiểm tra")
        return res
    # --- dự phòng: suy ra từ các dây rời
    groups = defaultdict(set)
    sample = {}
    wtypes = defaultdict(list)
    female = set()            # nhóm có đầu cái (母端) = dây nối giắc bộ đèn, không phải dây nguồn
    ends = defaultdict(list)  # loại đầu nối của từng nhóm (AMP公端, 摩氏公端...)
    for r in rows:
        if r["name"] != "粗胚-電線":
            continue
        segs = r["spec"].split("-")
        wire = next((s for s in segs if RE_WIRE.match(s)), None)
        lm = next((RE_LEN.match(s) for s in segs if RE_LEN.match(s)), None)
        if not wire or not lm or not wire.endswith("#18"):  # dây nguồn cỡ #18; bỏ dây #22 nối mạch
            continue
        col = segs[-1]
        if not RE_COLOR_ONE.match(col) or col == "綠":
            continue
        ln = float(lm.group(1))
        if ln <= 15:
            continue
        key = (ln, lm.group(0))   # gộp theo chiều dài (BOM đôi khi lẫn 1010/1015 trong cùng bộ dây)
        wtypes[key].append(wire)
        groups[key].add(col)
        sample.setdefault(key, r)
        if "母端" in r["spec"]:
            female.add(key)
        for tok in re.findall(r"(?:AMP|摩氏)?公端", r["spec"]):
            if tok not in ends[key]:
                ends[key].append(tok)
    for k in female:           # bỏ nhóm dây đầu cái; nếu loại hết thì giữ nguyên để không mất dữ liệu
        if len(groups) > 1:
            groups.pop(k, None)
    if not groups:
        return {"found": False}
    best = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0][0]))[0]
    (_ln, lstr), cols = best
    wire = Counter(wtypes[(_ln, lstr)]).most_common(1)[0][0]
    ordered = [c for c in "黑白藍紅黃棕灰橙紫" if c in cols]
    warnings.append("Dây nguồn suy ra từ dây rời (đã loại dây nối giắc bộ đèn) – cần kiểm tra")
    conn = "-".join(ends[(_ln, lstr)])
    return {
        "found": True, "source": "inferred",
        "raw": f"{wire}-{lstr}-" + (conn + "-" if conn else "") + "".join(ordered), "rawVi": None,
        "wireType": wire, "length": lstr,
        "colors": [color_obj(c) for c in ordered], "wireCount": len(ordered), "extra": [],
    }


# ----------------------------------------------------------------- tem mô tơ
def extract_label(rows):
    r = next((x for x in rows if "馬達識別標" in x["name"]), None)
    if not r:
        return {"found": False}
    raw = r["spec"]
    m = re.match(r"^(\d+(?:\.\d+)?x\d+(?:\.\d+)?CM)-", raw, re.I)
    size = m.group(1) if m else None
    s = raw.replace('Impedance-Protected"', "ImpedanceProtected-").replace('"', "")
    s = s.replace("Impedance-Protected", "ImpedanceProtected")
    s = re.sub(r"^\d+(?:\.\d+)?x\d+(?:\.\d+)?CM-", "", s, flags=re.I)
    s = s.replace("馬達標-", "")
    s = re.sub(r"-UL(?=-|$)", "", s)
    s = re.sub(r"-銀底黑字$", "", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return {"found": True, "raw": raw, "code": s, "size": size}


# ----------------------------------------------------------------- loại mô tơ
def extract_motor(rows):
    for r in rows:
        m = RE_CORE.match(r["name"])
        if m:
            return {"found": True, "value": "x".join(m.groups()), "source": "core", "raw": r["name"]}
    for r in rows:
        if "馬達組" in r["name"]:
            m = re.match(r"^AC-(\d{3})-(\d+(?:\.\d+)?)", r["spec"])
            if m and m.group(1) in STATOR_MAP:
                return {"found": True, "value": f"{STATOR_MAP[m.group(1)]}x{m.group(2)}",
                        "source": "motor_group", "raw": r["spec"]}
    return {"found": False}


# ----------------------------------------------------------------- tụ điện
def parse_cap(spec):
    uf = None
    m = re.search(r"(\d+(?:\.\d+)?(?:\+\d+(?:\.\d+)?)?)\s*uf", spec, re.I)
    if m:
        uf = m.group(1)
    else:
        m = re.search(r"(\d+(?:\.\d+)?(?:\+\d+(?:\.\d+)?)?)[-/]\s*\d{3}V", spec)
        if m:
            uf = m.group(1)
    mv = re.search(r"(\d{3})V", spec)
    volt = mv.group(1) if mv else None
    me = re.search(r"(\+\d+\+\d+)", spec)
    extra = me.group(1) if me else None
    grade = None
    mg = re.search(r"([A-Z]{1,2}\d?)級", spec)
    if mg:
        grade = mg.group(1)
    elif mv:
        tail = spec[mv.end():].split("-")
        for seg in tail:
            mm = re.match(r"^([A-Z]{1,2})(?![A-Za-z])", seg)
            if mm:
                grade = mm.group(1)
                break
    mw = re.search(r"(\d)線", spec)
    wires = int(mw.group(1)) if mw else None
    parts = []
    if uf:
        parts.append(f"{uf} µF")
    if volt:
        parts.append(f"{volt}V")
    if extra:
        parts.append(extra)
    if grade:
        parts.append(f"cấp {grade}")
    if wires:
        parts.append(f"{wires} dây")
    return {"uF": uf, "V": volt, "extra": extra, "grade": grade, "wires": wires,
            "display": " · ".join(parts) if parts else None}


def extract_caps(rows):
    caps = []
    for r in rows:
        if r["name"].startswith("零件-電容器"):
            c = {"raw": r["spec"], "rawVi": r["spec_vi"], "source": "capacitor"}
            c.update(parse_cap(r["spec"]))
            caps.append(c)
    if caps:
        return caps
    # quạt remote: tụ nằm trong mô-đun đảo chiều (ưu tiên) hoặc trong bộ thu
    for key, src in (("正反轉模組", "reverse_module"), ("接收器", "receiver")):
        for r in rows:
            if key in r["name"] and r["name"].startswith("零件") and re.search(r"\d(?:\.\d)?\s*uf", r["spec"], re.I):
                c = {"raw": r["spec"], "rawVi": r["spec_vi"], "source": src}
                c.update(parse_cap(r["spec"]))
                caps.append(c)
                break
        if caps:
            break
    return caps


# ----------------------------------------------------------------- miếng đệm mô tơ
PAD_INC = re.compile(r"馬達|月眉|半圓|紙墊圈")
PAD_EXC = re.compile(r"吊架|吊盤|葉架|葉片|接頭|燈盤|中圈|分离蓋|螺母|內下殼|防水|上連中|轉盤")
PAD_MAT = [("海綿", "Mút xốp"), ("EVA", "EVA"), ("快巴紙", "Giấy fiber"), ("紙", "Giấy"), ("硬橡膠", "Cao su cứng"), ("橡膠", "Cao su"), ("優能膠", "Keo PU")]
PAD_COLOR = {"黑": "đen", "白": "trắng", "灰": "xám", "紅": "đỏ", "藍": "xanh dương", "綠": "xanh lá", "黃": "vàng", "棕": "nâu"}


def extract_pads(rows):
    out = []
    for r in rows:
        n = r["name"]
        if not (("墊片" in n or "墊圈" in n) and PAD_INC.search(n) and not PAD_EXC.search(n)):
            continue
        s = r["spec"]
        if "月眉" in n:
            shape = "Hình lưỡi liềm (月眉型)"
        elif "半圓" in n:
            shape = "Hình bán nguyệt (半圓型)"
        elif "墊圈" in n:
            shape = "Vòng đệm (墊圈)"
        else:
            shape = "Miếng đệm"
        mat = next((vi for cn, vi in PAD_MAT if cn in s), None) or next((vi for cn, vi in PAD_MAT if cn in n), None)
        col_cn = next((ch for ch in PAD_COLOR if ch in s.split("-")[-1]), None) or next((ch for ch in PAD_COLOR if ch in s), None)
        col = None
        if col_cn:
            col = {"cn": col_cn, "vi": PAD_COLOR[col_cn], "hex": COLORS.get(col_cn, ("", "#888"))[1]}
        mt = re.search(r"\((\d+(?:\.\d+)?)mmT\)|(\d+(?:\.\d+)?)T(?![a-z])", n)
        thick = (mt.group(1) or mt.group(2)) + " mm" if mt else None
        mc = re.search(r"(Y[GU]RA\d+)", s) or re.search(r"(Y[GU]RA\d+)", n)
        q = r["qty"]
        if isinstance(q, float) and q.is_integer():
            q = int(q)
        out.append({
            "name": n.replace("零件-", "").replace("馬達墊片", "").strip("-"), "raw": s, "rawVi": r["spec_vi"],
            "shape": shape, "material": mat, "color": col, "thickness": thick,
            "code": mc.group(1) if mc else None, "adhesive": bool(re.search(r"自粘|自黏", n)), "qty": q,
        })
    return out


# ----------------------------------------------------------------- tem dây đèn
LANG_MAP = [("英西法文", "Anh · Tây Ban Nha · Pháp"), ("英西文", "Anh · Tây Ban Nha"), ("英法文", "Anh · Pháp"), ("英文", "Anh"), ("西文", "Tây Ban Nha")]
LBL_COLOR = {"藍": ("xanh dương", "#1e5bd8"), "白": ("trắng", "#f5f5f5"), "黑": ("đen", "#1a1a1a"), "紅": ("đỏ", "#d62828"),
             "綠": ("xanh lá", "#1f9d55"), "黃": ("vàng", "#f2c94c"), "銀": ("bạc", "#c9ccd1"), "透明": ("trong suốt", "")}


def parse_light_label(spec):
    m = re.match(r"^(\d+(?:\.\d+)?x\d+(?:\.\d+)?)\s*CM", spec, re.I)
    size = m.group(1) + " cm" if m else None
    lang = next((vi for cn, vi in LANG_MAP if cn in spec), None)
    mt = re.search(r"(?:^|-)(FLT|LIGHT)(?=-|$)", spec)
    text = {"LIGHT": "LIGHT", "FLT": "FLT"}.get(mt.group(1)) if mt else None
    bg = fg = None
    mc = re.search(r"(透明|[藍白黑紅綠黃銀])底(透明|[藍白黑紅綠黃銀])字", spec)
    if mc:
        def co(ch):
            vi, hx = LBL_COLOR[ch]
            return {"cn": ch, "vi": vi, "hex": hx}
        bg, fg = co(mc.group(1)), co(mc.group(2))
    return {"size": size, "text": text, "lang": lang, "bg": bg, "fg": fg}


def extract_light_labels(rows):
    out, seen = [], {}
    for r in rows:
        if "燈線標" not in r["name"] or not r["name"].startswith("零件"):
            continue
        s = r["spec"]
        q = r["qty"]
        if isinstance(q, float) and q.is_integer():
            q = int(q)
        # mã linh kiện: lấy từ hàng (cột mã vật tư) nếu có, không thì bỏ
        key = s
        if key in seen:
            if isinstance(q, (int, float)) and isinstance(seen[key]["qty"], (int, float)):
                seen[key]["qty"] = max(seen[key]["qty"], q)
            continue
        d = parse_light_label(s)
        d.update({"raw": s, "rawVi": r["spec_vi"], "qty": q, "partNo": r["code"], "source": "part"})
        seen[key] = d
        out.append(d)
    if out:
        return out
    # dự phòng: nhãn đèn được ghi ngay trong quy cách dây nguồn
    for r in rows:
        if "電源線組" in r["name"]:
            m = re.search(r"燈線標-?\s*(\d+(?:\.\d+)?x\s?\d+(?:\.\d+)?\s?CM.*?)(?=-?(?:其他類標|矽套管|扎帶)|$)", r["spec"].replace(" ", ""))
            if m:
                d = parse_light_label(m.group(1))
                d.update({"raw": m.group(1), "rawVi": None, "qty": 1, "partNo": None, "source": "cord"})
                return [d]
    return []


# ----------------------------------------------------------------- bộ thu / bộ phát
def brand_model(spec):
    out = []
    for i, seg in enumerate(spec.split("-")):
        if i > 0 and (RE_CJK.search(seg) or re.search(r"\d\s*uf", seg, re.I)):
            break
        out.append(seg)
    return "-".join(out)


def extract_rf(rows, key):
    res = []
    for r in rows:
        if key in r["name"] and r["name"].startswith("零件"):
            q = r["qty"]
            if isinstance(q, float) and q.is_integer():
                q = int(q)
            res.append({"raw": r["spec"], "rawVi": r["spec_vi"], "qty": q, "brandModel": brand_model(r["spec"])})
    return res


# ----------------------------------------------------------------- chuẩn hoá tìm kiếm
def norm(s):
    return re.sub(r"\s+", "", s.upper().replace("（", "(").replace("）", ")"))


def build_search(rec):
    parts = [rec["customerModel"], rec["fanCode"], rec["fanName"], rec["file"].rsplit(".", 1)[0]]
    n = [norm(p) for p in parts]
    c = [re.sub(r"[-_+.]", "", x) for x in n]
    return " ".join(dict.fromkeys(n + c))


# ----------------------------------------------------------------- overrides
def apply_overrides(items, path):
    if not path.exists():
        return 0
    by_model = {i["customerModel"]: i for i in items}
    n = 0
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            rec = by_model.get((row.get("customerModel") or "").strip())
            field = (row.get("field") or "").strip()
            if not rec or not field:
                continue
            node = rec
            keys = field.split(".")
            for k in keys[:-1]:
                node = node.setdefault(k, {})
            node[keys[-1]] = row.get("value", "")
            rec.setdefault("overridden", True)
            n += 1
    return n


# ----------------------------------------------------------------- main
def prep_image(path):
    """Cắt viền nền thừa, căn vật thể vào giữa khung 4:3, thu nhỏ -> (mime, bytes). Lỗi/không có Pillow -> ảnh gốc."""
    suffix = path.suffix.lower()
    mimes = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}
    raw = path.read_bytes()
    try:
        from PIL import Image
        import io
        im = Image.open(io.BytesIO(raw))
        im.load()
        if im.mode in ("RGBA", "LA", "P"):
            rgba = im.convert("RGBA")
            base = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            base.alpha_composite(rgba)
            im = base.convert("RGB")
        else:
            im = im.convert("RGB")
        im.thumbnail((1400, 1400))
        w, h = im.size
        px = im.load()
        corners = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
        bg = tuple(sorted(c[k] for c in corners)[1] // 2 + sorted(c[k] for c in corners)[2] // 2 for k in range(3))
        # nhị phân hoá: điểm khác nền
        diff = Image.new("L", (w, h))
        dp = diff.load()
        for y in range(h):
            for x in range(w):
                c = px[x, y]
                dp[x, y] = 255 if abs(c[0] - bg[0]) + abs(c[1] - bg[1]) + abs(c[2] - bg[2]) > 36 else 0
        # bỏ nhiễu: cần >=3 điểm mỗi hàng/cột
        rows = [sum(1 for x in range(w) if dp[x, y]) for y in range(h)]
        cols = [sum(1 for y in range(h) if dp[x, y]) for x in range(w)]
        xs = [i for i, v in enumerate(cols) if v >= 3]
        ys = [i for i, v in enumerate(rows) if v >= 3]
        if xs and ys:
            x0, x1, y0, y1 = xs[0], xs[-1], ys[0], ys[-1]
            im = im.crop((x0, y0, x1 + 1, y1 + 1))
        bw, bh = im.size
        # viền xám mảnh (1-3 px) bao quanh ô trắng -> cắt bỏ
        if bw > 60 and bh > 60:
            _p = im.load()
            if min(_p[0, 0]) < 235 and min(_p[4, 4]) >= 235 and min(_p[bw - 5, bh - 5]) >= 235:
                im = im.crop((4, 4, bw - 4, bh - 4)); bw, bh = im.size
        # nếu sau khi cắt, ảnh có viền trắng (hình nằm trong ô trắng, ngoài là nền xám) -> dùng nền trắng, bỏ khung xám
        cp = im.load()
        cc = [cp[3, 3], cp[bw - 4, 3], cp[3, bh - 4], cp[bw - 4, bh - 4]] if bw > 8 and bh > 8 else [cp[0, 0]]
        if all(min(c) >= 235 for c in cc):
            bg = (255, 255, 255)
        # khung 4:3 với lề đều, vật thể ở giữa
        margin = 0.06
        W = bw / (1 - 2 * margin)
        H = bh / (1 - 2 * margin)
        if W / H > 4 / 3:
            H = W * 3 / 4
        else:
            W = H * 4 / 3
        if min(bg) >= 235:  # nền trắng/gần trắng -> trắng tuyệt đối để hoà với khung
            bg = (255, 255, 255)
            im = im.point(lambda v: 255 if v >= 247 else v)
        canvas = Image.new("RGB", (round(W), round(H)), bg)
        canvas.paste(im, ((canvas.width - bw) // 2, (canvas.height - bh) // 2))
        canvas.thumbnail((1000, 750))
        out = io.BytesIO()
        canvas.save(out, "JPEG", quality=88, optimize=True)
        return "image/jpeg", out.getvalue()
    except Exception as e:  # noqa
        print("  (không xử lý được ảnh", path.name, ":", e, ")")
        return mimes.get(suffix, "image/png"), raw


# ----------------------------------------------------------------- từ điển Trung -> Việt
def load_dict(path):
    import csv
    d = {}
    if not path.exists():
        return d
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            zh = (r.get("zh") or "").strip()
            vi = (r.get("vi") or "").strip()
            if not zh or not vi:
                continue
            alt = [x.strip() for x in (r.get("vi_khac") or "").split("|") if x.strip()]
            d[zh] = [zh, vi, alt, (r.get("nhom") or "").strip()]
    return d


RE_HAN = re.compile(r"[\u3400-\u9fff]")


def seg_han(s, d, maxl=12):
    """Khớp dài nhất từ trái sang. Trả về list (cụm_Hán, đã_dịch). Bỏ qua phần không phải chữ Hán.
    Các chữ Hán chưa dịch nằm liền nhau được gộp thành một cụm."""
    out, i, prev_unk_end = [], 0, -1
    while i < len(s):
        if RE_HAN.match(s[i]):
            hit = next((s[i:i + L] for L in range(min(maxl, len(s) - i), 0, -1) if s[i:i + L] in d), None)
            if hit:
                out.append((hit, True))
                i += len(hit)
            else:
                if prev_unk_end == i and out and not out[-1][1]:
                    out[-1] = (out[-1][0] + s[i], False)
                else:
                    out.append((s[i], False))
                i += 1
                prev_unk_end = i
        else:
            i += 1
    return out


def dict_report(all_rows, d):
    stats = {}
    unk = {}
    for fld in ("name", "spec"):
        tot = hit = 0
        for rows in all_rows.values():
            for r in rows:
                txt = r[fld]
                for t, ok in seg_han(txt, d):
                    tot += 1
                    if ok:
                        hit += 1
                    else:
                        e = unk.setdefault(t, [0, txt, set()])
                        e[0] += 1
                        e[2].add(r["_model"])
        stats[fld] = (hit, tot)
    return stats, unk


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bom", default=str(HERE.parents[1] / "BOM"))
    ap.add_argument("--html", default=str(HERE.parent / "BOM_Lookup.html"))
    ap.add_argument("--pics", default=str(HERE.parent / "PICTURE"))
    a = ap.parse_args()
    bom_dir, html_path, pics_dir = Path(a.bom), Path(a.html), Path(a.pics)
    files = sorted(p for p in bom_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        sys.exit(f"Không có file .xlsx trong {bom_dir}")

    items, errors, rows_by_file = [], [], {}
    for p in files:
        try:
            rows = read_rows(p)
        except Exception as e:  # noqa
            errors.append((p.name, repr(e)))
            continue
        rows_by_file[p.name] = rows
        warnings = []
        fan_name, cm = split_name(p.stem)
        cat, fan_code = category_of(fan_name)
        rec = {
            "id": cm, "file": p.name, "fanName": fan_name, "fanCode": fan_code,
            "category": cat, "customerModel": cm,
            "powerCord": extract_cord(rows, warnings),
            "motorLabel": extract_label(rows),
            "motorType": extract_motor(rows),
            "capacitors": extract_caps(rows),
            "receivers": extract_rf(rows, "接收器"),
            "transmitters": extract_rf(rows, "發射器"),
            "motorPads": extract_pads(rows),
            "lightLabels": extract_light_labels(rows),
            "warnings": warnings,
        }
        for k, v in (("motorType", rec["motorType"]),):
            if v.get("source") == "motor_group":
                warnings.append("Loại mô tơ suy từ cụm mô tơ – cần kiểm tra")
        if any(c["source"] == "receiver" for c in rec["capacitors"]):
            warnings.append("Tụ điện nằm trong bộ thu")
        if any(c["source"] == "reverse_module" for c in rec["capacitors"]):
            warnings.append("Tụ điện nằm trong mô-đun đảo chiều")
        if not rec["capacitors"]:
            warnings.append("Không tìm thấy tụ điện")
        items.append(rec)

    # id duy nhất
    seen = Counter()
    for it in sorted(items, key=lambda x: x["file"]):
        seen[it["id"]] += 1
        if seen[it["id"]] > 1:
            it["id"] = f'{it["id"]}_{seen[it["id"]]}'
    bom_rows = {}
    for it in items:
        rr = rows_by_file[it["file"]]
        for r in rr:
            r["_model"] = it["id"]
        bom_rows[it["id"]] = [
            [r["lv"], r["code"], r["name"], ("" if r["qty"] is None else (int(r["qty"]) if isinstance(r["qty"], float) and r["qty"].is_integer() else (round(r["qty"], 5) if isinstance(r["qty"], float) else r["qty"]))),
             r["spec"], r["unit"], {"自製": "M", "採購": "B"}.get(r["src"], ""), r["sup"].replace("\n", ""), r["vi_name"] or (r["spec_vi"] or "")]
            for r in rr
        ]
    dic = load_dict(HERE / "tu_dien.csv")
    dstats, unk = dict_report({it["id"]: rows_by_file[it["file"]] for it in items}, dic)
    ov = apply_overrides(items, HERE / "overrides.csv")
    for it in items:
        it["search"] = build_search(it)
    items.sort(key=lambda x: x["customerModel"].upper())

    data = {
        "generatedAt": dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).isoformat(timespec="seconds"),
        "sourceDir": str(bom_dir),
        "count": len(items),
        "items": items,
    }
    # ---- ảnh: nhúng dạng data URI (tên file = model khách hàng)
    import base64
    mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}
    models = {i["customerModel"] for i in items}
    pics, pic_bytes, orphan = {}, 0, []
    if pics_dir.is_dir():
        for f in sorted(pics_dir.iterdir()):
            if f.suffix.lower() in mime and f.is_file():
                if f.stem not in models:
                    orphan.append(f.name)
                    continue
                mt, raw = prep_image(f)
                pic_bytes += len(raw)
                pics[f.stem] = f"data:{mt};base64," + base64.b64encode(raw).decode()
    data["pictures"] = len(pics)

    def js(o):
        return json.dumps(o, ensure_ascii=False).replace("</", "<\\/")
    block = ("<!--DATA-START-->\n<script>window.BOM_DATA=" + js(data) + ";window.BOM_PICS=" + js(pics) + ";window.BOM_ROWS=" + js(bom_rows) + ";window.BOM_DICT=" + js(list(dic.values())) + ";</script>\n<!--DATA-END-->")
    html = html_path.read_text(encoding="utf-8")
    new_html, n = re.subn(r"<!--DATA-START-->.*?<!--DATA-END-->", lambda m: block, html, flags=re.S)
    if n != 1:
        sys.exit(f"Không thấy cặp marker DATA-START/DATA-END trong {html_path}")
    html_path.write_text(new_html, encoding="utf-8")
    out = html_path

    # ---- báo cáo
    lines = [f"Số file BOM đọc được: {len(items)}/{len(files)}  (overrides áp dụng: {ov})"]
    miss = {
        "powerCord": [i for i in items if not i["powerCord"]["found"]],
        "motorLabel": [i for i in items if not i["motorLabel"]["found"]],
        "motorType": [i for i in items if not i["motorType"]["found"]],
        "capacitors": [i for i in items if not i["capacitors"]],
    }
    for k, v in miss.items():
        lines.append(f"Thiếu {k}: {len(v)}  {[i['customerModel'] for i in v]}")
    lines.append(f"Có bộ thu: {sum(1 for i in items if i['receivers'])}, có bộ phát: {sum(1 for i in items if i['transmitters'])}")
    lines.append(f"Không có miếng đệm mô tơ: {[i['customerModel'] for i in items if not i['motorPads']]}")
    lines.append(f"Không có tem dây đèn: {[i['customerModel'] for i in items if not i['lightLabels']]}")
    lines.append(f"Tem dây đèn suy từ dây nguồn: {[i['customerModel'] for i in items if any(l['source']=='cord' for l in i['lightLabels'])]}")
    odd = [(i["customerModel"], i["powerCord"]["wireCount"]) for i in items
           if i["powerCord"]["found"] and i["powerCord"]["wireCount"] not in (2, 3)]
    lines.append(f"Dây nguồn số dây lạ (khác 2/3): {odd}")
    nogrd = [i["customerModel"] for i in items for c in i["capacitors"] if not c.get("uF") or not c.get("V")]
    lines.append(f"Tụ không tách được µF/V: {nogrd}")
    lines.append(f"Ảnh nhúng: {len(pics)}/{len(items)} model ({pic_bytes/1e6:.1f} MB)" + (f"; ảnh không khớp model nào: {orphan}" if orphan else ""))
    nopic = [i["customerModel"] for i in items if i["customerModel"] not in pics]
    lines.append(f"Chưa có ảnh ({len(nopic)}): {nopic}")
    cov_n = dstats["name"][0] / max(1, dstats["name"][1])
    cov_s = dstats["spec"][0] / max(1, dstats["spec"][1])
    lines.append(f"Từ điển: {len(dic)} mục. Độ phủ tên (品名): {cov_n:.1%}; quy cách (規格): {cov_s:.1%}")
    top = sorted(unk.items(), key=lambda kv: -kv[1][0])[:30]
    lines.append("30 cụm chưa dịch nhiều nhất: " + ", ".join(f"{k}×{v[0]}" for k, v in top))
    import csv as _csv
    with open(HERE / "tu_chua_dich.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = _csv.writer(f)
        w.writerow(["zh", "so_lan", "vi_du", "model", "vi"])
        for k, v in sorted(unk.items(), key=lambda kv: -kv[1][0]):
            w.writerow([k, v[0], v[1], " ".join(sorted(v[2]))[:80], ""])
    lines.append("--- Cảnh báo theo model ---")
    for i in items:
        if i["warnings"]:
            lines.append(f"{i['customerModel']}: " + "; ".join(i["warnings"]))
    if errors:
        lines.append("--- LỖI ĐỌC FILE ---")
        lines += [f"{n}: {e}" for n, e in errors]
    report = "\n".join(lines)
    (HERE / "build_report.txt").write_text(report, encoding="utf-8")
    print(report)
    print(f"\nĐã ghi dữ liệu + ảnh vào {out}")


if __name__ == "__main__":
    main()
