# -*- coding: utf-8 -*-
"""
Gui bao cao nho (ket qua QC 1 mau thanh pham) len NHOM ZALO.

Zalo KHONG co API cho tai khoan ca nhan gui vao nhom, nen module nay dieu khien
ung dung Zalo PC tren Windows (giong nguoi that thao tac):
    1. dua cua so Zalo len truoc
    2. Ctrl+F -> dan ten nhom -> Enter  (mo dung nhom)
    3. dan ANH phieu bao cao + dan NOI DUNG CHU -> (tuy chon) Enter de gui

Che do "Toi tu mo nhom": phan mem chi dua Zalo len truoc, nguoi dung tu bam vao
nhom can gui roi bam nut "Dan & gui" -> phan mem dan noi dung vao o chat.

Chi dung ctypes + Pillow (khong can cai them thu vien).
"""
import os, io, re, sys, json, time, threading, datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

IS_WIN = sys.platform.startswith("win")


# ============================================================================
# 1) NOI DUNG BAO CAO
# ============================================================================
def build_text(info):
    """Noi dung tin nhan - giong mau QC dang gui tay tren nhom."""
    L = []
    L.append(info.get("date", ""))
    if info.get("order"):
        L.append(info["order"])
    line2 = "   ".join(x for x in (info.get("vmi", ""),
                                  (".DC %s" % info["line_no"]) if info.get("line_no") else "") if x)
    if line2:
        L.append(line2)
    t = (info.get("time") or "").replace(":", "h")
    ca = (info.get("ca") or "").upper()
    if t or ca:
        L.append("   ".join(x for x in (t, ca) if x))
    if info.get("ash"):
        L.append("%%tro = %s" % info["ash"])
    if info.get("mfi"):
        L.append("Mfi :  %s" % info["mfi"])
    if info.get("beads"):
        L.append("Hạt: %s hạt/gam." % info["beads"])
    if info.get("moist"):
        L.append("Độ ẩm: %s%%" % info["moist"])
    if info.get("disp"):
        L.append("Phân tán: %s" % str(info["disp"]).split("/")[0].strip())
    if info.get("bloom"):
        L.append("Hoa màng: %s" % str(info["bloom"]).split("/")[0].strip())
    if info.get("verdict"):
        L.append(info["verdict"])
    if info.get("note"):
        L.append("Ghi chú: %s" % info["note"])
    return "\n".join(L)


# ============================================================================
# 2) ANH PHIEU BAO CAO (PNG)
# ============================================================================
def _font(size, bold=False):
    from PIL import ImageFont
    win = os.environ.get("WINDIR", r"C:\Windows")
    names = (["segoeuib.ttf", "arialbd.ttf", "tahomabd.ttf"] if bold else
             ["segoeui.ttf", "arial.ttf", "tahoma.ttf"])
    cands = [os.path.join(win, "Fonts", n) for n in names]
    cands += ["/usr/share/fonts/truetype/dejavu/DejaVuSans%s.ttf" % ("-Bold" if bold else ""),
              "DejaVuSans%s.ttf" % ("-Bold" if bold else "")]
    for p in cands:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _tw(draw, text, font):
    try:
        b = draw.textbbox((0, 0), text, font=font)
        return b[2] - b[0]
    except Exception:
        return len(text) * font.size // 2


def _wrap(draw, text, font, width):
    words, lines, cur = str(text).split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if _tw(draw, t, font) <= width or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


# ---------------------------------------------------------------------------
# ANH MAU MANG FILM + ANH HAT FILLER (2 khung ben phai phieu) - luu file theo ma mau (sid)
#   <thu muc phan mem>/BaoCaoZalo/AnhMangFilm/<sid>_<thoi gian>.jpg
#   <thu muc phan mem>/BaoCaoZalo/AnhHatFiller/<sid>_<thoi gian>.jpg
# ---------------------------------------------------------------------------
FILM_BASE = [os.path.dirname(os.path.abspath(__file__))]
FILM_MAX = 4
CARD_TITLE = "PHÒNG QC-RD GỬI BÁO CÁO CHẤT LƯỢNG THÀNH PHẨM"
# kind -> (thu muc, tieu de khung tren phieu, ten ngan)
IMG_KINDS = {"film": ("AnhMangFilm", "HÌNH ẢNH MẪU MÀNG FILM", "màng film"),
             "filler": ("AnhHatFiller", "HÌNH ẢNH HẠT FILLER", "hạt Filler")}


def set_base_dir(base_dir):
    if base_dir:
        FILM_BASE[0] = base_dir


def film_dir(kind="film"):
    d = os.path.join(FILM_BASE[0], "BaoCaoZalo", IMG_KINDS.get(kind, IMG_KINDS["film"])[0])
    os.makedirs(d, exist_ok=True)
    return d


def film_files(sid, kind="film"):
    """Danh sach file anh (mang film / hat filler) cua 1 mau (cu -> moi)."""
    if sid is None or sid == "":
        return []
    try:
        pre = "%s_" % sid
        d = film_dir(kind)
        return sorted(os.path.join(d, f) for f in os.listdir(d)
                      if f.startswith(pre) and f.lower().endswith((".jpg", ".jpeg", ".png")))
    except Exception:
        return []


def film_add(sid, img_or_path, kind="film"):
    """Them 1 anh (PIL.Image hoac duong dan file) cho mau sid -> duong dan da luu."""
    from PIL import Image
    im = Image.open(img_or_path) if isinstance(img_or_path, str) else img_or_path
    im.load()
    if im.mode not in ("RGB", "L"):
        bg = Image.new("RGB", im.size, "white")
        try:
            bg.paste(im, mask=im.convert("RGBA").split()[-1])
        except Exception:
            bg.paste(im.convert("RGB"))
        im = bg
    im = im.convert("RGB")
    im.thumbnail((1800, 1800))
    p = os.path.join(film_dir(kind), "%s_%s.jpg" % (sid, dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")))
    im.save(p, "JPEG", quality=90)
    return p


def film_images(info, kind="film"):
    """Anh de ve vao phieu: info['film'] / info['filler'] (list duong dan) hoac tim theo info['sid']."""
    from PIL import Image
    paths = info.get(kind) if info.get(kind) is not None else film_files(info.get("sid"), kind)
    out = []
    for p in (paths or [])[-FILM_MAX:]:
        try:
            im = Image.open(p)
            im.load()
            out.append(im.convert("RGB"))
        except Exception:
            pass
    return out


def delete_sample_images(sids):
    """Xoa anh mang film + anh hat filler cua cac mau (sau khi gui neu co tich 'Tu xoa anh').
    -> so file da xoa"""
    n = 0
    for sid in sids or []:
        for k in IMG_KINDS:
            for p in film_files(sid, k):
                try:
                    os.remove(p)
                    n += 1
                except Exception:
                    pass
    return n


# ---------------------------------------------------------------------------
# ANH RIENG THEO MOC GUI (lich hen gio) + DAY CHUYEN
#   Lich "Mau moi nhat tung day chuyen": luc hen gio mau chua co -> anh gan vao (moc, day chuyen).
#   Khoa luu anh: M<id lich>_DC<so day chuyen>  (cung thu muc AnhMangFilm / AnhHatFiller)
# ---------------------------------------------------------------------------
_SLOT_RE = re.compile(r"^M(\d+)_DC(.+)$")


def slot_key(job_id, line):
    return "M%s_DC%s" % (job_id, line)


def slot_counts(job_id, line):
    """-> (so anh mang film, so anh hat filler) cua 1 moc / 1 day chuyen"""
    k = slot_key(job_id, line)
    return len(film_files(k, "film")), len(film_files(k, "filler"))


def slot_info(info, job):
    """info cua 1 mau + anh RIENG cua moc (lich) theo day chuyen cua mau.
    job['slot_img_only'] = True: moc chua up anh -> de trong (khong lay anh cu cua mau)."""
    out = dict(info)
    key = slot_key(job.get("id"), info.get("line_no", ""))
    only = job.get("slot_img_only", False)
    for kind in IMG_KINDS:
        fs = film_files(key, kind)
        if fs:
            out[kind] = fs
        elif only:
            out[kind] = []
    out["_slot_key"] = key
    return out


def delete_slot_images(job_id):
    """Xoa toan bo anh rieng cua 1 lich (moi day chuyen). -> so file da xoa"""
    n = 0
    pre = "M%s_DC" % job_id
    for kind in IMG_KINDS:
        d = film_dir(kind)
        for f in os.listdir(d):
            if f.startswith(pre):
                try:
                    os.remove(os.path.join(d, f))
                    n += 1
                except Exception:
                    pass
    return n


_HASH_CACHE = {}


def img_hash(im):
    """dHash 16x16 (256 bit): 2 anh giong nhau (chup lai / dan lai / nen JPEG) -> khoang cach nho."""
    g = im.convert("L").resize((17, 16))
    px = list(g.getdata())
    bits = 0
    for y in range(16):
        r = y * 17
        for x in range(16):
            bits = (bits << 1) | (1 if px[r + x] > px[r + x + 1] else 0)
    return bits


def _file_hash(p):
    try:
        st = os.stat(p)
    except OSError:
        return None
    k = (p, st.st_mtime, st.st_size)
    h = _HASH_CACHE.get(k)
    if h is None:
        from PIL import Image
        try:
            with Image.open(p) as im:
                h = img_hash(im)
        except Exception:
            return None
        _HASH_CACHE[k] = h
    return h


def key_of_file(fname):
    """'M123_DC1_20260930_101010_123456.jpg' -> 'M123_DC1' ; '57_2026..._...jpg' -> '57'"""
    return os.path.basename(fname).rsplit("_", 3)[0]


def find_duplicate(im, max_dist=6):
    """Anh im da duoc up o moc / mau khac chua? -> (khoa, kind) hoac None"""
    try:
        h = img_hash(im)
    except Exception:
        return None
    for kind in IMG_KINDS:
        d = film_dir(kind)
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for f in names:
            if not f.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            h2 = _file_hash(os.path.join(d, f))
            if h2 is not None and bin(h ^ h2).count("1") <= max_dist:
                return key_of_file(f), kind
    return None


def describe_key(key, provider=None, sget=None):
    """Mo ta khoa anh cho nguoi dung: 'moc 08:00 · DC1' / 'mau DC3 luc 09:30'"""
    m = _SLOT_RE.match(str(key))
    if m:
        t = "?"
        if sget is not None:
            try:
                j = next((x for x in json.loads(sget("zalo_jobs", "[]") or "[]") if x.get("id") == m.group(1)), None)
                t = j.get("time") if j else "(lịch đã xoá)"
            except Exception:
                pass
        return "mốc gửi %s · DC%s" % (t, m.group(2))
    info = None
    if provider is not None:
        try:
            info = provider.info_of(int(key))
        except Exception:
            info = None
    if info:
        return "mẫu DC%s lúc %s (%s)" % (info.get("line_no", ""), info.get("time", ""), info.get("date", ""))
    return "mẫu #%s" % key


def make_dup_check(provider=None, sget=None):
    """-> ham(im) tra ve cau canh bao neu anh trung anh da up o moc / mau khac, khong thi None"""
    def chk(im):
        hit = find_duplicate(im)
        if not hit:
            return None
        k, kind = hit
        return "Ảnh này GIỐNG ảnh %s đã up cho %s.\nMỗi mốc / mỗi dây chuyền nên có ảnh khác nhau." % (
            IMG_KINDS[kind][2], describe_key(k, provider, sget))
    return chk


def _fit_font(draw, text, width, size, bold=True, min_size=16):
    while size > min_size:
        f = _font(size, bold)
        if _tw(draw, text, f) <= width:
            return f
        size -= 1
    return _font(min_size, bold)


def _is_ok(v):
    v = (v or "").lower()
    return bool(v) and "không" not in v and "chưa" not in v


def _vi(v):
    """'Phân Tán Tốt /Good dispersion' -> 'Phân Tán Tốt' (bo phan tieng Anh cho gon)."""
    v = str(v or "").strip()
    return v.split("/")[0].strip() or v


def _disp_ok(v):
    v = str(v or "").lower()
    return "" if not v else ("Đạt" if "không" not in v and "no " not in v else "Không đạt")


def _bloom_ok(v):
    v = str(v or "").lower().strip()
    return "" if not v else ("Đạt" if v.startswith("không") or v.startswith("no ") else "Không đạt")


def _img_layout(ims, iw, cap):
    """Xep anh vao khung rong iw: 1 anh -> 1 cot, nhieu anh -> 2 cot. -> (cells, tong cao)"""
    if not ims:
        return [], 0
    ncol = 1 if len(ims) == 1 else 2
    cw = iw if ncol == 1 else (iw - 12) // 2
    rows = [ims[i:i + ncol] for i in range(0, len(ims), ncol)]
    rhs = [max(int(im.height * cw / im.width) for im in r) for r in rows]
    need = sum(rhs) + 12 * (len(rows) - 1)
    k = min(1.0, cap / float(need)) if need else 1.0
    cells = [[(im, max(1, int(cw * k))) for im in r] for r in rows]
    return cells, int(need * k)


def _draw_img_frame(img, d, fx0, fx1, fy0, fy1, caption, cells, empty_text, cap_h=44):
    GRID, GRAY, NAVY = (205, 212, 222), (100, 110, 120), (31, 58, 95)
    d.rectangle([fx0, fy0, fx1, fy1], outline=GRID, width=2)
    d.rectangle([fx0, fy0, fx1, fy0 + cap_h], fill=(234, 240, 248))
    f_cap = _fit_font(d, caption, fx1 - fx0 - 16, 20, True, 14)
    d.text((fx0 + (fx1 - fx0 - _tw(d, caption, f_cap)) // 2, fy0 + 10), caption, font=f_cap, fill=NAVY)
    iy = fy0 + cap_h + 12
    if cells:
        tot = sum(max(int(im.height * cw / im.width) for im, cw in r) for r in cells) + 12 * (len(cells) - 1)
        iy += max(0, (fy1 - 12 - iy - tot) // 2)          # canh giua theo chieu doc
        for r in cells:
            rw = sum(cw for _, cw in r) + 12 * (len(r) - 1)
            ix = fx0 + (fx1 - fx0 - rw) // 2
            rh = 0
            for im, cw in r:
                th = max(1, int(im.height * cw / im.width))
                img.paste(im.resize((cw, th)), (ix, iy))
                d.rectangle([ix, iy, ix + cw - 1, iy + th - 1], outline=GRID)
                ix += cw + 12
                rh = max(rh, th)
            iy += rh + 12
    else:
        by0, by1 = iy, fy1 - 12
        bx0, bx1 = fx0 + 12, fx1 - 12
        for xx in range(bx0, bx1, 14):              # khung net dut cho cho chen anh
            d.line([xx, by0, min(xx + 7, bx1), by0], fill=(170, 180, 195), width=2)
            d.line([xx, by1, min(xx + 7, bx1), by1], fill=(170, 180, 195), width=2)
        for yy in range(by0, by1, 14):
            d.line([bx0, yy, bx0, min(yy + 7, by1)], fill=(170, 180, 195), width=2)
            d.line([bx1, yy, bx1, min(yy + 7, by1)], fill=(170, 180, 195), width=2)
        f_l = _fit_font(d, empty_text, bx1 - bx0 - 16, 18, False, 12)
        d.text(((bx0 + bx1 - _tw(d, empty_text, f_l)) // 2, (by0 + by1) // 2 - 12), empty_text, font=f_l, fill=GRAY)


def render_card(info, logo_path=None, film=None, filler=None):
    """Ve phieu ket qua QC -> PIL.Image.
    Ben trai: thong tin + bang ket qua + danh gia.
    Ben phai: 2 khung HINH ANH MAU MANG FILM | HINH ANH HAT FILLER."""
    from PIL import Image, ImageDraw
    WL, P, FWK, GAP = 820, 28, 440, 12    # WL: phan phieu ben trai, FWK: moi khung anh ben phai
    W = WL + FWK * 2 + GAP + P
    NAVY, GRID, GRAY = (31, 58, 95), (205, 212, 222), (100, 110, 120)
    f_s = _font(19)
    f_l, f_v, f_vb = _font(18), _font(20), _font(20, True)
    f_big, f_small = _font(40, True), _font(15)
    HD = 116                            # chieu cao dai tieu de

    tmp = Image.new("RGB", (W, 10))
    d0 = ImageDraw.Draw(tmp)
    film = film_images(info, "film") if film is None else film
    filler = film_images(info, "filler") if filler is None else filler

    # --- du lieu cac bang
    infos = [("Dây chuyền", info.get("line", "")), ("QC", info.get("qc", "")),
             ("Đơn hàng", info.get("order", "")), ("Mã hàng VMI", info.get("vmi", "")),
             ("Mã khách", info.get("cust", "")), ("Lot No", info.get("lot", "")),
             ("Số bao / mẻ", info.get("bags", "")), ("Batch No", info.get("batch", ""))]
    infos = [x for x in infos if x[1]]
    verdict = info.get("verdict", "")
    all_ok = _is_ok(verdict)
    res = []
    if info.get("moist"):
        res.append(["Độ ẩm (%)", info["moist"], "< 0.1", ""])
    if info.get("beads"):
        res.append(["Số hạt / gam", info["beads"], "22 – 25", ""])
    if info.get("mfi"):
        res.append(["MFI" + ((" (%s)" % info["mfi_cond"]) if info.get("mfi_cond") else ""),
                    info["mfi"], info.get("mfi_std", ""), info.get("mfi_ok", "")])
    if info.get("ash"):
        res.append(["% Tro (CaCO3)", info["ash"], info.get("ash_std", ""), info.get("ash_ok", "")])
    if info.get("appear"):
        res.append(["Ngoại quan", _vi(info["appear"]), "", ""])
    # Danh gia PHAN TAN / HOA MANG: luon co 2 dong (chua thu mang -> "—")
    src = info.get("film_src") or ""
    mk = "*" if src else ""
    res.append(["Độ phân tán" + mk, _vi(info.get("disp")) or "—", "Phân tán tốt", _disp_ok(info.get("disp"))])
    res.append(["Độ hoa màng" + mk, _vi(info.get("bloom")) or "—", "Không hoa màng", _bloom_ok(info.get("bloom"))])
    if all_ok:                           # danh gia chung DAT -> cot Xet ghi Dat het (dong da co ket qua)
        for r in res:
            if not r[3] and r[1] != "—":
                r[3] = "Đạt"

    col_w = [200, 210, 190, WL - 2 * P - 600]
    row_h = 40
    half = (WL - 2 * P) // 2
    info_rows = (len(infos) + 1) // 2
    foot = ("* Phân tán / hoa màng lấy theo mẫu thử màng lúc %s" % src) if src else ""
    note_lines = _wrap(d0, "Ghi chú: " + info["note"], f_l, WL - 2 * P) if info.get("note") else []
    HL = (HD + 20 + info_rows * 56 + 20 + (len(res) + 1) * row_h + 30 + (26 if foot else 0) + 90
          + len(note_lines) * 26 + 50)

    # --- kich thuoc 2 khung anh ben phai
    fy0 = HD + 20
    cap_h = 44
    iw = FWK - 24
    min_img_h = HL - 50 - fy0 - cap_h - 24          # chieu cao vung anh theo phan ben trai
    cap = max(min_img_h, 760)                        # anh cao qua -> thu nho lai
    cells_f, h_f = _img_layout(film, iw, cap)
    cells_h, h_h = _img_layout(filler, iw, cap)
    H = max(HL, fy0 + cap_h + max(h_f, h_h, min_img_h) + 24 + 50)

    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    # --- header (trai rong ca phieu)
    d.rectangle([0, 0, W, HD], fill=NAVY)
    x0 = P
    if logo_path is not None and (not isinstance(logo_path, str) or os.path.exists(logo_path)):
        try:
            if hasattr(logo_path, "seek"):
                logo_path.seek(0)
            lg = Image.open(logo_path).convert("RGBA")
            lg.thumbnail((150, 72))
            bg = Image.new("RGBA", (lg.width + 12, lg.height + 12), (255, 255, 255, 255))
            bg.paste(lg, (6, 6), lg)
            img.paste(bg.convert("RGB"), (P, (HD - bg.height) // 2))
            x0 = P + bg.width + 20
        except Exception:
            pass
    f_t = _fit_font(d, CARD_TITLE, W - P - x0, 34)
    d.text((x0, 20), CARD_TITLE, font=f_t, fill="white")
    sub = "  ·  ".join(x for x in (info.get("date", ""), (info.get("ca") or "").upper(),
                                  (info.get("time") or "").replace(":", "h")) if x)
    d.text((x0, 70), sub, font=f_s, fill=(220, 230, 245))

    # --- thong tin chung (2 cot)
    y = HD + 20
    for i, (k, v) in enumerate(infos):
        cx = P + (i % 2) * half
        cy = y + (i // 2) * 56
        d.text((cx, cy), k, font=f_small, fill=GRAY)
        vv = str(v)
        while _tw(d, vv, f_vb) > half - 16 and len(vv) > 4:
            vv = vv[:-2]
        d.text((cx, cy + 20), vv if vv == str(v) else vv + "…", font=f_vb, fill=(20, 20, 20))
    y += info_rows * 56 + 20

    # --- bang ket qua
    if res:
        heads = ["Chỉ tiêu", "Kết quả", "Tiêu chuẩn", "Xét"]
        x = P
        d.rectangle([P, y, WL - P, y + row_h], fill=(234, 240, 248))
        for j, h in enumerate(heads):
            d.text((x + 10, y + 9), h, font=f_vb, fill=NAVY)
            x += col_w[j]
        y += row_h
        for r in res:
            x = P
            for j, v in enumerate(r):
                col = (20, 20, 20)
                bold = j in (1, 3) and bool(v)
                if j == 3 and v:
                    col = (30, 130, 60) if _is_ok(v) else (200, 40, 40)
                if j == 1 and v == "—":
                    col, bold = GRAY, False
                font = _fit_font(d, str(v), col_w[j] - 16, 20, bold, 13)
                d.text((x + 10, y + 9 + max(0, 20 - getattr(font, "size", 20)) // 2), str(v), font=font, fill=col)
                x += col_w[j]
            d.line([P, y + row_h, WL - P, y + row_h], fill=GRID, width=1)
            y += row_h
        d.rectangle([P, y - (len(res) + 1) * row_h, WL - P, y], outline=GRID, width=2)
        y += 8
        if foot:
            d.text((P, y), foot, font=f_small, fill=GRAY)
            y += 26
        y += 22

    # --- ket luan
    color = (30, 140, 70) if all_ok else ((200, 40, 40) if verdict else (150, 150, 150))
    d.rounded_rectangle([P, y, WL - P, y + 72], radius=14, fill=color) if hasattr(d, "rounded_rectangle") \
        else d.rectangle([P, y, WL - P, y + 72], fill=color)
    txt = ("ĐÁNH GIÁ: " + verdict.upper()) if verdict else "CHƯA ĐÁNH GIÁ"
    d.text(((WL - _tw(d, txt, f_big)) // 2, y + 12), txt, font=f_big, fill="white")
    y += 90
    for ln in note_lines:
        d.text((P, y), ln, font=f_l, fill=(60, 60, 60))
        y += 26

    # --- 2 khung anh ben phai: MANG FILM | HAT FILLER
    fy1 = H - 50
    fx0 = WL
    _draw_img_frame(img, d, fx0, fx0 + FWK, fy0, fy1, IMG_KINDS["film"][1], cells_f,
                    "(Chưa chèn ảnh mẫu màng film)", cap_h)
    fx0 = WL + FWK + GAP
    _draw_img_frame(img, d, fx0, fx0 + FWK, fy0, fy1, IMG_KINDS["filler"][1], cells_h,
                    "(Chưa chèn ảnh hạt Filler)", cap_h)

    d.text((P, H - 34), "HITECHPLAST · tạo lúc %s" % dt.datetime.now().strftime("%H:%M %d/%m/%Y"),
           font=f_small, fill=GRAY)
    return img


# ============================================================================
# 3) DIEU KHIEN ZALO PC (Windows, ctypes)
# ============================================================================
class ZaloError(Exception):
    pass


if IS_WIN:
    import ctypes
    import ctypes.wintypes as wt

    _u32 = ctypes.WinDLL("user32", use_last_error=True)
    _k32 = ctypes.WinDLL("kernel32", use_last_error=True)

    _u32.OpenClipboard.argtypes = [wt.HWND]
    _u32.OpenClipboard.restype = wt.BOOL
    _u32.EmptyClipboard.restype = wt.BOOL
    _u32.CloseClipboard.restype = wt.BOOL
    _u32.SetClipboardData.argtypes = [wt.UINT, wt.HANDLE]
    _u32.SetClipboardData.restype = wt.HANDLE
    _u32.RegisterClipboardFormatW.argtypes = [wt.LPCWSTR]
    _u32.RegisterClipboardFormatW.restype = wt.UINT
    _k32.GlobalAlloc.argtypes = [wt.UINT, ctypes.c_size_t]
    _k32.GlobalAlloc.restype = ctypes.c_void_p
    _k32.GlobalLock.argtypes = [ctypes.c_void_p]
    _k32.GlobalLock.restype = ctypes.c_void_p
    _k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    _k32.GlobalFree.argtypes = [ctypes.c_void_p]

    _EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
    _u32.EnumWindows.argtypes = [_EnumProc, wt.LPARAM]
    _u32.GetWindowTextLengthW.argtypes = [wt.HWND]
    _u32.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
    _u32.IsWindowVisible.argtypes = [wt.HWND]
    _u32.IsIconic.argtypes = [wt.HWND]
    _u32.ShowWindow.argtypes = [wt.HWND, ctypes.c_int]
    _u32.SetForegroundWindow.argtypes = [wt.HWND]
    _u32.BringWindowToTop.argtypes = [wt.HWND]
    _u32.GetForegroundWindow.restype = wt.HWND
    _u32.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]
    _u32.GetWindowRect.argtypes = [wt.HWND, ctypes.POINTER(wt.RECT)]
    _k32.OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
    _k32.OpenProcess.restype = wt.HANDLE
    _k32.CloseHandle.argtypes = [wt.HANDLE]
    _k32.QueryFullProcessImageNameW.argtypes = [wt.HANDLE, wt.DWORD, wt.LPWSTR, ctypes.POINTER(wt.DWORD)]

VK_CONTROL, VK_MENU, VK_RETURN, VK_ESCAPE, VK_DOWN = 0x11, 0x12, 0x0D, 0x1B, 0x28
KEYUP = 0x0002
CF_UNICODETEXT, CF_DIB = 13, 8


def _key(vk, up=False):
    _u32.keybd_event(vk, 0, KEYUP if up else 0, 0)


def _tap(vk, delay=0.05):
    _key(vk)
    time.sleep(delay)
    _key(vk, True)
    time.sleep(delay)


def _combo(mod, vk):
    _key(mod)
    time.sleep(0.05)
    _tap(vk)
    _key(mod, True)
    time.sleep(0.08)


def _set_clipboard(items):
    """items = [(format, bytes), ...] dat vao clipboard cung luc."""
    for _ in range(20):
        if _u32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        raise ZaloError("Không mở được clipboard (đang bị chương trình khác giữ).")
    try:
        _u32.EmptyClipboard()
        for fmt, data in items:
            h = _k32.GlobalAlloc(0x0002, len(data))          # GMEM_MOVEABLE
            p = _k32.GlobalLock(h)
            ctypes.memmove(p, data, len(data))
            _k32.GlobalUnlock(h)
            if not _u32.SetClipboardData(fmt, h):
                _k32.GlobalFree(h)
    finally:
        _u32.CloseClipboard()


def clip_text(text):
    _set_clipboard([(CF_UNICODETEXT, (text + "\0").encode("utf-16-le"))])


def clip_image(pil_img):
    img = pil_img.convert("RGB")
    b = io.BytesIO()
    img.save(b, "BMP")
    dib = b.getvalue()[14:]                              # bo BITMAPFILEHEADER
    p = io.BytesIO()
    img.save(p, "PNG")
    png_fmt = _u32.RegisterClipboardFormatW("PNG")
    _set_clipboard([(CF_DIB, dib), (png_fmt, p.getvalue())])


def _proc_name(hwnd):
    pid = wt.DWORD()
    _u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    h = _k32.OpenProcess(0x1000, False, pid.value)       # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(520)
        n = wt.DWORD(520)
        if _k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)):
            return os.path.basename(buf.value).lower()
    finally:
        _k32.CloseHandle(h)
    return ""


def find_zalo():
    """Cua so chinh cua Zalo PC (lon nhat, process zalo.exe)."""
    found = []

    def cb(hwnd, _):
        if not _u32.IsWindowVisible(hwnd) and not _u32.IsIconic(hwnd):
            return True
        n = _u32.GetWindowTextLengthW(hwnd)
        if n <= 0:
            return True
        buf = ctypes.create_unicode_buffer(n + 1)
        _u32.GetWindowTextW(hwnd, buf, n + 1)
        title = buf.value
        if _proc_name(hwnd) == "zalo.exe" or title.strip() == "Zalo":
            r = wt.RECT()
            _u32.GetWindowRect(hwnd, ctypes.byref(r))
            found.append(((r.right - r.left) * (r.bottom - r.top), hwnd, title))
        return True

    _u32.EnumWindows(_EnumProc(cb), 0)
    if not found:
        return None
    found.sort(key=lambda x: -x[0])
    return found[0][1]


def launch_zalo():
    for p in (os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Zalo", "Zalo.exe"),
              os.path.join(os.environ.get("PROGRAMFILES", ""), "Zalo", "Zalo.exe"),
              os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Zalo", "Zalo.exe")):
        if p and os.path.exists(p):
            os.startfile(p)
            return True
    return False


def activate_zalo(wait_launch=15):
    h = find_zalo()
    if not h:
        if not launch_zalo():
            raise ZaloError("Không thấy ứng dụng Zalo PC.\nHãy mở Zalo trên máy tính và đăng nhập trước.")
        for _ in range(wait_launch * 2):
            time.sleep(0.5)
            h = find_zalo()
            if h:
                time.sleep(2)
                break
        if not h:
            raise ZaloError("Đã mở Zalo nhưng chưa thấy cửa sổ – hãy đăng nhập Zalo rồi thử lại.")
    if _u32.IsIconic(h):
        _u32.ShowWindow(h, 9)                            # SW_RESTORE
    _tap(VK_MENU)                                        # meo cho phep SetForegroundWindow
    _u32.BringWindowToTop(h)
    _u32.SetForegroundWindow(h)
    time.sleep(0.4)
    if _u32.GetForegroundWindow() != h:
        _u32.ShowWindow(h, 5)
        _u32.SetForegroundWindow(h)
        time.sleep(0.4)
    if _u32.GetForegroundWindow() != h:
        raise ZaloError("Không đưa được cửa sổ Zalo lên trước. Hãy bấm vào Zalo một lần rồi thử lại.")
    return h


def _still_zalo(h):
    if _u32.GetForegroundWindow() != h:
        raise ZaloError("Cửa sổ Zalo bị mất tiêu điểm (có ai bấm chuột/bàn phím?) – đã dừng gửi.")


def open_group(h, name, open_key="enter"):
    _still_zalo(h)
    _tap(VK_ESCAPE)
    _combo(VK_CONTROL, 0x46)                             # Ctrl+F: o tim kiem cua Zalo
    time.sleep(0.5)
    _combo(VK_CONTROL, 0x41)                             # Ctrl+A: xoa chu cu
    clip_text(name)
    _combo(VK_CONTROL, 0x56)                             # Ctrl+V: dan ten nhom
    time.sleep(1.6)                                      # cho Zalo hien ket qua
    _still_zalo(h)
    if open_key == "down_enter":
        _tap(VK_DOWN)
    _tap(VK_RETURN)
    time.sleep(1.3)


def test_group(name, open_key="enter"):
    """Mo Zalo + tim & mo nhom theo ten (KHONG gui gi) de nguoi dung kiem tra ten nhom dung chua."""
    if not IS_WIN:
        raise ZaloError("Chỉ dùng được trên Windows có cài Zalo PC.")
    if not _SEND_LOCK.acquire(blocking=False):
        raise ZaloError("Đang có một lượt gửi Zalo – đợi xong rồi thử lại.")
    try:
        h = activate_zalo()
        open_group(h, name, open_key)
    finally:
        _SEND_LOCK.release()


def open_zalo_search():
    """Dua Zalo len truoc + mo o tim kiem (Ctrl+F) de nguoi dung xem / chon nhom."""
    if not IS_WIN:
        raise ZaloError("Chỉ dùng được trên Windows có cài Zalo PC.")
    h = activate_zalo()
    _tap(VK_ESCAPE)
    _combo(VK_CONTROL, 0x46)
    return h


# ----------------------------------------------------------------------------
# 3a') TU LAY TEN NHOM DANG MO TREN ZALO (chup thanh tieu de cuoc tro chuyen + OCR)
# ----------------------------------------------------------------------------
# Chu co dinh tren giao dien Zalo (khong phai ten nhom) - so sanh da bo dau
_UI_WORDS = ("zalo", "tim kiem", "search", "tat ca", "chua doc", "phan loai", "uu tien", "khac",
             "thanh vien", "truy cap", "vua truy cap", "dang hoat dong", "online", "cong dong",
             "them ban", "tao nhom", "cloud cua toi", "my documents", "ban be", "danh ba",
             "nhan tin", "goi dien", "gan the", "phut truoc", "gio truoc", "ngay truoc", "khong co")


def _fold(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or "").lower()).replace("đ", "d")
    return " ".join("".join(c for c in s if unicodedata.category(c) != "Mn").split())


def _is_ui_word(t):
    f = _fold(t).strip(" .:·-|")
    if len(f) < 2:
        return True
    return any(f == w or f.startswith(w + " ") or f.endswith(" " + w) for w in _UI_WORDS)


def _win_rect_px(h):
    """Khung cua so theo PIXEL THAT (dung ca khi man hinh phong to 125%/150%)."""
    r = wt.RECT()
    try:
        dwm = ctypes.WinDLL("dwmapi")
        if dwm.DwmGetWindowAttribute(wt.HWND(h), 9, ctypes.byref(r), ctypes.sizeof(r)) == 0:
            return r.left, r.top, r.right, r.bottom            # DWMWA_EXTENDED_FRAME_BOUNDS
    except Exception:
        pass
    _u32.GetWindowRect(h, ctypes.byref(r))
    return r.left, r.top, r.right, r.bottom


def _win_scale(h):
    try:
        d = _u32.GetDpiForWindow(wt.HWND(h))
        if d:
            return max(1.0, d / 96.0)
    except Exception:
        pass
    return 1.0


def grab_zalo_header(h):
    """Chup dai phia tren cua so Zalo (co thanh tieu de cuoc tro chuyen). -> (PIL.Image, scale)"""
    from PIL import ImageGrab
    x0, y0, x1, y1 = _win_rect_px(h)
    s = _win_scale(h)
    hh = min(y1 - y0, int(115 * s))
    if x1 - x0 < 200 or hh < 30:
        raise ZaloError("Cửa sổ Zalo đang thu nhỏ – hãy phóng to Zalo rồi thử lại.")
    return ImageGrab.grab(bbox=(x0, y0, x1, y0 + hh), all_screens=True), s


def read_chat_title(img, s=1.0):
    """Tim TEN cuoc tro chuyen dang mo trong anh dai tieu de Zalo. Tra ve '' neu khong thay."""
    import image_recipe as ir
    boxes = [b for b in ir.ocr_image(img) if b.text.strip()]
    W = img.width
    # khung o tim kiem cua danh sach ben trai -> ten nhom nam ben phai no
    left = 230 * s
    for b in boxes:
        if _fold(b.text).startswith(("tim kiem", "search")):
            left = max(left, b.x1 + 40 * s)
    cand = [b for b in boxes if b.x0 >= left and b.x1 <= W - 20 * s and b.has_letters
            and not _is_ui_word(b.text) and b.score >= 0.5]
    if not cand:
        return ""
    hmax = max(b.h for b in cand)
    big = [b for b in cand if b.h >= 0.8 * hmax]          # ten nhom: chu dam, to nhat
    first = min(big, key=lambda b: (round(b.cy / (8 * s)), b.x0))
    # ghep cac manh cung dong (OCR doi khi tach ten dai thanh 2-3 khung)
    row = sorted([b for b in big if abs(b.cy - first.cy) < 0.5 * first.h and b.x0 >= first.x0], key=lambda b: b.x0)
    parts, last = [], None
    for b in row:
        if last is not None and b.x0 - last.x1 > 3 * first.h:
            break
        parts.append(b.text.strip())
        last = b
    name = " ".join(" ".join(parts).split()).strip(" .:·|")
    return name[:120]


def current_chat_name(activate=True):
    """Dua Zalo len truoc (tuy chon) + doc ten cuoc tro chuyen dang mo. Tra ve ten ('' neu khong doc duoc)."""
    if not IS_WIN:
        raise ZaloError("Chỉ dùng được trên Windows có cài Zalo PC.")
    h = activate_zalo() if activate else find_zalo()
    if not h:
        raise ZaloError("Chưa mở Zalo PC.")
    if activate:
        time.sleep(0.3)
    img, s = grab_zalo_header(h)
    return read_chat_title(img, s)


def zalo_foreground():
    """hwnd cua Zalo neu Zalo dang la cua so tren cung, nguoc lai None."""
    if not IS_WIN:
        return None
    try:
        f = _u32.GetForegroundWindow()
        if f and _proc_name(f) == "zalo.exe":
            return find_zalo() or f          # co the la cua so phu (xem anh...) -> lay cua so chinh
        return None
    except Exception:
        return None


def paste_payload(h, text, images, auto_send):
    """Dan anh + chu vao o chat dang mo. images = list PIL.Image."""
    for im in images:
        _still_zalo(h)
        clip_image(im)
        _combo(VK_CONTROL, 0x56)
        time.sleep(1.2)
        if auto_send:
            _tap(VK_RETURN)
            time.sleep(1.2)
    if text:
        _still_zalo(h)
        clip_text(text)
        _combo(VK_CONTROL, 0x56)
        time.sleep(0.5)
        if auto_send:
            _tap(VK_RETURN)
            time.sleep(0.6)


# ============================================================================
# 3b) GUI NHIEU TIN VAO NHIEU NHOM (chay trong thread, 1 luot 1 lan)
# ============================================================================
_SEND_LOCK = threading.Lock()
# Muc dac biet trong danh sach nhom: gui vao cuoc tro chuyen DANG MO san tren Zalo (khong tim theo ten)
CUR_CHAT = "📌 Nhóm đang mở trên Zalo"


def is_cur_chat(g):
    return (g or "").strip() == CUR_CHAT


def sender_busy():
    return _SEND_LOCK.locked()


def send_messages(groups, messages, auto_send=True, open_key="enter", progress=None):
    """groups: ten nhom ([None] = nhom dang mo san trong Zalo).
    messages: [(text, [PIL.Image, ...]), ...]. Goi trong THREAD. Tra ve list nhom da gui."""
    if not IS_WIN:
        raise ZaloError("Chỉ tự gửi được trên Windows có cài Zalo PC.")
    if not _SEND_LOCK.acquire(blocking=False):
        raise ZaloError("Đang có một lượt gửi Zalo khác – đợi xong rồi thử lại.")
    done = []
    try:
        for g in groups:
            if progress:
                progress("Đang gửi vào: %s …" % ("nhóm đang mở" if (not g or is_cur_chat(g)) else g))
            h = activate_zalo()
            if g and not is_cur_chat(g):
                open_group(h, g, open_key)
            else:
                time.sleep(0.3)
            for text, images in messages:
                paste_payload(h, text, images, auto_send)
                time.sleep(0.5)
            done.append("nhóm đang mở" if (not g or is_cur_chat(g)) else g)
            time.sleep(0.8)
    finally:
        _SEND_LOCK.release()
    return done


# ============================================================================
# 3c) BAO CAO TONG HOP (nhieu mau)
# ============================================================================
def _vn_verdict(v):
    return v or "—"


def build_summary_text(title, infos):
    L = [title]
    ok = bad = 0
    for i in infos:
        parts = ["DC%s %s" % (i.get("line_no", ""), (i.get("time") or "").replace(":", "h"))]
        if i.get("order"):
            parts.append(i["order"])
        vals = []
        if i.get("moist"):
            vals.append("Ẩm %s" % i["moist"])
        if i.get("beads"):
            vals.append("Hạt %s" % i["beads"])
        if i.get("mfi"):
            vals.append("MFI %s" % i["mfi"])
        if i.get("ash"):
            vals.append("Tro %s" % i["ash"])
        v = i.get("verdict", "")
        if v:
            vals.append(v)
            if "không" in v.lower():
                bad += 1
            else:
                ok += 1
        L.append(" ".join(parts) + (" | " + " | ".join(vals) if vals else ""))
    L.append("Tổng: %d mẫu – Đạt %d, Không đạt %d" % (len(infos), ok, bad))
    return "\n".join(L)


def render_summary_card(title, subtitle, infos, logo_path=None):
    from PIL import Image, ImageDraw
    NAVY, GRID, GRAY = (31, 58, 95), (205, 212, 222), (100, 110, 120)
    cols = [("DC", 55), ("Giờ", 75), ("Đơn hàng", 190), ("Ẩm", 70), ("Hạt/g", 70),
            ("MFI", 75), ("% Tro", 105), ("Kết quả", 130)]
    P = 24
    W = P * 2 + sum(w for _, w in cols)
    rh = 36
    H = 110 + 20 + (len(infos) + 1) * rh + 70 + 40
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    f_t, f_s, f_h, f_v, f_vb, f_small = _font(28, True), _font(18), _font(17, True), _font(17), _font(17, True), _font(15)
    d.rectangle([0, 0, W, 110], fill=NAVY)
    x0 = P
    if logo_path is not None and (not isinstance(logo_path, str) or os.path.exists(logo_path)):
        try:
            if hasattr(logo_path, "seek"):
                logo_path.seek(0)
            lg = Image.open(logo_path).convert("RGBA")
            lg.thumbnail((150, 70))
            bg = Image.new("RGBA", (lg.width + 12, lg.height + 12), (255, 255, 255, 255))
            bg.paste(lg, (6, 6), lg)
            img.paste(bg.convert("RGB"), (P, (110 - bg.height) // 2))
            x0 = P + bg.width + 18
        except Exception:
            pass
    d.text((x0, 20), title, font=f_t, fill="white")
    d.text((x0, 64), subtitle, font=f_s, fill=(220, 230, 245))
    y = 130
    d.rectangle([P, y, W - P, y + rh], fill=(234, 240, 248))
    x = P
    for h, w in cols:
        d.text((x + 8, y + 8), h, font=f_h, fill=NAVY)
        x += w
    y += rh
    ok = bad = 0
    for n, i in enumerate(infos):
        v = i.get("verdict", "")
        isbad = "không" in v.lower() if v else False
        if isbad:
            d.rectangle([P, y, W - P, y + rh], fill=(255, 228, 228))
            bad += 1
        elif v:
            ok += 1
        vals = [i.get("line_no", ""), (i.get("time") or "").replace(":", "h"), i.get("order", ""),
                i.get("moist", ""), i.get("beads", ""), i.get("mfi", ""), i.get("ash", ""), _vn_verdict(v)]
        x = P
        for j, (val, (_, w)) in enumerate(zip(vals, cols)):
            col, font = (20, 20, 20), f_v
            if j == 7:
                font = f_vb
                col = (200, 40, 40) if isbad else ((30, 130, 60) if v else GRAY)
            s = str(val)
            while _tw(d, s, font) > w - 12 and len(s) > 3:
                s = s[:-2]
            d.text((x + 8, y + 8), s, font=font, fill=col)
            x += w
        d.line([P, y + rh, W - P, y + rh], fill=GRID)
        y += rh
    d.rectangle([P, 130, W - P, y], outline=GRID, width=2)
    y += 18
    txt = "Tổng %d mẫu  ·  Đạt %d  ·  Không đạt %d" % (len(infos), ok, bad)
    d.text((P, y), txt, font=_font(22, True), fill=(200, 40, 40) if bad else (30, 130, 60))
    d.text((P, H - 32), "HITECHPLAST · tạo lúc %s" % dt.datetime.now().strftime("%H:%M %d/%m/%Y"),
           font=f_small, fill=GRAY)
    return img


# ============================================================================
# 4) HOP THOAI XEM TRUOC & GUI
# ============================================================================
SEP = "\n────────────\n"


def _groups_of(sget):
    try:
        g = json.loads(sget("zalo_groups", "[]") or "[]")
        return [x for x in g if isinstance(x, str) and x.strip()]
    except Exception:
        return []


def _flag(sget, key, default):
    v = sget(key, "")
    return default if v == "" else v == "1"


class ZaloSendDialog(tk.Toplevel):
    """Xem truoc + sua noi dung + chon nhom -> gui.
    info: 1 mau (dict) HOAC messages: [(text, card PIL.Image), ...]."""

    def __init__(self, master, info=None, store_get=None, store_set=None, base_dir=".", logo_path=None,
                 messages=None, title=None, on_sent=None, infos=None):
        super().__init__(master)
        set_base_dir(base_dir)
        self.title(title or "📤 Gửi báo cáo lên nhóm Zalo")
        self.sget, self.sset = store_get, store_set
        self.base_dir, self.logo, self.on_sent = base_dir, logo_path, on_sent
        self.info = info or {}
        if messages is None:
            messages = [(build_text(self.info), render_card(self.info, logo_path))]
            infos = [self.info]
        self.messages = list(messages)
        self.infos = list(infos or [None] * len(self.messages))
        self.extra = []
        self.card_paths = []
        self.transient(master)
        self.geometry("1000x660")
        self.minsize(820, 560)
        self._build()
        self.grab_set()
        self.focus_set()

    def _build(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=3)
        root.columnconfigure(1, weight=2)
        root.rowconfigure(0, weight=1)
        L = ttk.Frame(root)
        L.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        n = len(self.messages)
        ttk.Label(L, text="Nội dung tin nhắn (sửa được)%s:" % ((" – %d tin, ngăn cách bằng ────" % n) if n > 1 else ""),
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.txt = tk.Text(L, height=10, wrap="word", font=("Segoe UI", 11))
        self.txt.pack(fill="x", pady=(2, 6))
        self.txt.insert("1.0", SEP.join(t for t, _ in self.messages))
        opt = ttk.Frame(L)
        opt.pack(fill="x")
        self.v_text = tk.BooleanVar(value=_flag(self.sget, "zalo_send_text", True))
        self.v_card = tk.BooleanVar(value=_flag(self.sget, "zalo_send_card", True))
        ttk.Checkbutton(opt, text="Gửi nội dung chữ", variable=self.v_text).pack(side="left")
        ttk.Checkbutton(opt, text="Gửi ảnh phiếu", variable=self.v_card).pack(side="left", padx=12)
        ttk.Button(opt, text="🖼 Thêm ảnh khác…", command=self._add_img).pack(side="left", padx=4)
        ttk.Button(opt, text="✖", width=3, command=self._clear_img).pack(side="left")
        ttk.Button(opt, text="🎞 Ảnh màng film…", command=self._film).pack(side="left", padx=(8, 0))
        ttk.Button(opt, text="⚪ Ảnh hạt Filler…", command=lambda: self._film("filler")).pack(side="left", padx=(4, 0))
        opt2 = ttk.Frame(L)
        opt2.pack(fill="x", pady=(2, 0))
        self.v_del = tk.BooleanVar(value=_flag(self.sget, "zalo_img_autodel", False))
        ttk.Checkbutton(opt2, text="🗑 Tự xoá ảnh màng film + ảnh hạt Filler sau khi gửi (không tích = giữ lại)",
                        variable=self.v_del,
                        command=lambda: self.sset("zalo_img_autodel", "1" if self.v_del.get() else "0")
                        ).pack(side="left")
        self.lbl_extra = ttk.Label(L, text="", foreground="#666")
        self.lbl_extra.pack(anchor="w", pady=(2, 4))
        pv = ttk.LabelFrame(L, text=" Xem trước ảnh%s " % ((" (1/%d)" % n) if n > 1 else ""))
        pv.pack(fill="both", expand=True)
        self.cv = tk.Label(pv, bg="#eef1f5")
        self.cv.pack(fill="both", expand=True)
        self._pi = 0
        if n > 1:
            nav = ttk.Frame(pv)
            nav.pack(fill="x")
            ttk.Button(nav, text="◀", width=3, command=lambda: self._nav(-1)).pack(side="left")
            ttk.Button(nav, text="▶", width=3, command=lambda: self._nav(1)).pack(side="left")
            self._pv = pv
        self._show_preview()

        R = ttk.Frame(root)
        R.grid(row=0, column=1, sticky="nsew")
        self.groups = GroupBox(R, self.sget, self.sset)
        self.groups.pack(fill="both", expand=True)
        m = ttk.LabelFrame(R, text=" Cách gửi ", padding=6)
        m.pack(fill="x", pady=(8, 0))
        self.v_mode = tk.StringVar(value=self.sget("zalo_mode", "auto") or "auto")
        ttk.Radiobutton(m, text="Tự tìm & mở nhóm theo tên", value="auto", variable=self.v_mode).pack(anchor="w")
        ttk.Radiobutton(m, text="Tôi tự mở nhóm trong Zalo (phần mềm chỉ dán)", value="manual",
                        variable=self.v_mode).pack(anchor="w")
        self.v_send = tk.BooleanVar(value=_flag(self.sget, "zalo_auto_send", True))
        ttk.Checkbutton(m, text="Tự bấm Gửi (Enter) sau khi dán", variable=self.v_send).pack(anchor="w", pady=(4, 0))
        self.v_down = tk.BooleanVar(value=self.sget("zalo_open_key", "enter") == "down_enter")
        ttk.Checkbutton(m, text="Mở nhóm bằng ↓ + Enter", variable=self.v_down).pack(anchor="w")
        self.btn_send = tk.Button(R, text="📤  GỬI LÊN ZALO", font=("Segoe UI", 12, "bold"),
                                  bg="#0068ff", fg="white", activebackground="#0052cc",
                                  relief="flat", pady=6, command=self.send)
        self.btn_send.pack(fill="x", pady=(10, 0))
        b2 = ttk.Frame(R)
        b2.pack(fill="x", pady=(6, 0))
        ttk.Button(b2, text="📋 Chép chữ", command=self._copy_text).pack(side="left", fill="x", expand=True)
        ttk.Button(b2, text="🖼 Chép ảnh", command=self._copy_img).pack(side="left", fill="x", expand=True, padx=3)
        ttk.Button(b2, text="💾 Lưu ảnh…", command=self._save_as).pack(side="left", fill="x", expand=True)
        ttk.Button(R, text="Đóng", command=self.destroy).pack(fill="x", pady=(6, 0))
        self.status = ttk.Label(R, text="", foreground="#1f3a5f", wraplength=340, justify="left")
        self.status.pack(anchor="w", pady=(8, 0))
        if not IS_WIN:
            self._st("Máy này không phải Windows: chỉ chép được nội dung, không tự gửi Zalo.")

    def _nav(self, k):
        self._pi = (self._pi + k) % len(self.messages)
        self._pv.configure(text=" Xem trước ảnh (%d/%d) " % (self._pi + 1, len(self.messages)))
        self._show_preview()

    def _show_preview(self):
        try:
            from PIL import ImageTk
            im = self.messages[self._pi][1].copy()
            im.thumbnail((520, 360))
            self._tkimg = ImageTk.PhotoImage(im)
            self.cv.configure(image=self._tkimg)
        except Exception as ex:
            self.cv.configure(text="(không xem trước được ảnh: %s)" % ex)

    def _st(self, s, err=False):
        try:
            self.status.configure(text=s, foreground="#b00020" if err else "#1f3a5f")
        except tk.TclError:
            pass

    def _add_img(self):
        ps = filedialog.askopenfilenames(parent=self, title="Chọn ảnh gửi kèm",
                                         filetypes=[("Ảnh", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Tất cả", "*.*")])
        if ps:
            self.extra += list(ps)
            self.lbl_extra.configure(text="Ảnh gửi kèm: " + ", ".join(os.path.basename(p) for p in self.extra))

    def _film(self, kind="film"):
        info = self.infos[self._pi] if self._pi < len(self.infos) else None
        if not info or info.get("sid") in (None, ""):
            messagebox.showinfo("Ảnh " + IMG_KINDS[kind][2], "Phiếu đang xem là phiếu tổng hợp – chỉ chèn ảnh "
                                "màng film / hạt Filler cho phiếu từng mẫu.", parent=self)
            return
        i = self._pi

        def redraw():
            try:
                self.messages[i] = (self.messages[i][0], render_card(info, self.logo))
                self.card_paths = []
                if self._pi == i:
                    self._show_preview()
            except Exception as ex:
                self._st("Lỗi vẽ lại phiếu: %s" % ex, True)

        FilmDialog(self, info["sid"], "DC%s %s" % (info.get("line_no", ""), info.get("time", "")), redraw, kind)

    def _clear_img(self):
        self.extra = []
        self.lbl_extra.configure(text="")

    def _payload(self):
        from PIL import Image
        texts = self.txt.get("1.0", "end").strip().split(SEP.strip())
        texts = [t.strip() for t in texts]
        msgs = []
        for i, (_, card) in enumerate(self.messages):
            t = texts[i] if i < len(texts) and self.v_text.get() else ""
            msgs.append((t, [card] if self.v_card.get() else []))
        extra = []
        for p in self.extra:
            try:
                im = Image.open(p)
                im.load()
                extra.append(im)
            except Exception:
                pass
        if extra:
            msgs.append(("", extra))
        return [(t, ims) for t, ims in msgs if t or ims]

    def save_cards(self):
        if self.card_paths:
            return self.card_paths
        self.card_paths = [save_card_file(c, self.base_dir, self.info) for _, c in self.messages]
        return self.card_paths

    def _save_as(self):
        c = self.messages[self._pi][1]
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".png",
                                         initialfile=os.path.basename(self.save_cards()[self._pi]),
                                         filetypes=[("PNG", "*.png")])
        if p:
            c.save(p)
            self._st("Đã lưu ảnh: " + p)

    def _copy_text(self):
        self.clipboard_clear()
        self.clipboard_append(self.txt.get("1.0", "end").strip())
        self._st("Đã chép nội dung – vào Zalo bấm Ctrl+V để dán.")

    def _copy_img(self):
        if not IS_WIN:
            self._st("Chỉ chép ảnh được trên Windows. Dùng 'Lưu ảnh…' rồi gửi file.", True)
            return
        try:
            clip_image(self.messages[self._pi][1])
            self._st("Đã chép ảnh – vào Zalo bấm Ctrl+V để dán.")
        except Exception as ex:
            self._st("Lỗi chép ảnh: %s" % ex, True)

    def _remember(self):
        self.sset("zalo_send_text", "1" if self.v_text.get() else "0")
        self.sset("zalo_send_card", "1" if self.v_card.get() else "0")
        self.sset("zalo_auto_send", "1" if self.v_send.get() else "0")
        self.sset("zalo_mode", self.v_mode.get())
        self.sset("zalo_open_key", "down_enter" if self.v_down.get() else "enter")

    def _missing_images(self):
        """-> [(ten phieu, 'màng film' / 'hạt Filler' / 'màng film + hạt Filler')] cac phieu TUNG MAU con thieu anh"""
        out = []
        for info in self.infos:
            if not info:
                continue
            if info.get("sid") in (None, "") and info.get("film") is None and info.get("filler") is None:
                continue                      # phieu tong hop - khong co o anh
            miss = []
            for k in ("film", "filler"):
                paths = info.get(k) if info.get(k) is not None else film_files(info.get("sid"), k)
                if not [p for p in (paths or []) if p and os.path.exists(p)]:
                    miss.append(IMG_KINDS[k][2])
            if miss:
                name = ("DC%s %s" % (info.get("line_no", ""), info.get("time", ""))).strip() or \
                    "mẫu #%s" % info.get("sid")
                out.append((name, " + ".join(miss)))
        return out

    def send(self):
        miss = self._missing_images()
        if miss:
            messagebox.showwarning(
                "Chưa đủ ảnh – không gửi",
                "Phải chèn ĐỦ 2 ảnh (ảnh mẫu màng film + ảnh hạt Filler) mới được gửi.\n\nCòn thiếu:\n" +
                "\n".join("  • %s: thiếu ảnh %s" % m for m in miss) +
                "\n\nBấm \"🎞 Ảnh màng film…\" / \"⚪ Ảnh hạt Filler…\" để chèn ảnh rồi gửi lại.",
                parent=self)
            self._st("✖ Chưa đủ ảnh: " + "; ".join("%s thiếu %s" % m for m in miss), True)
            return
        msgs = self._payload()
        if not msgs:
            messagebox.showinfo("Zalo", "Chưa chọn gửi gì (chữ hoặc ảnh).", parent=self)
            return
        try:
            self.save_cards()
        except Exception:
            pass
        if not IS_WIN:
            self._copy_text()
            return
        self._remember()
        auto_send = self.v_send.get()
        open_key = "down_enter" if self.v_down.get() else "enter"
        if self.v_mode.get() == "manual":
            try:
                activate_zalo()
            except Exception as ex:
                messagebox.showwarning("Zalo", str(ex), parent=self)
                return
            ManualPasteToast(self, lambda: self._run([None], msgs, auto_send, open_key))
            return
        groups = self.groups.selected()
        if not groups:
            messagebox.showinfo("Zalo", "Chọn (hoặc thêm) ít nhất 1 nhóm Zalo.", parent=self)
            return
        self.groups.remember()
        if len(groups) > 1 and not messagebox.askyesno(
                "Gửi Zalo", "Gửi lần lượt vào %d nhóm:\n%s ?" % (len(groups), "\n".join(groups)), parent=self):
            return
        self._run(groups, msgs, auto_send, open_key)

    def _run(self, groups, msgs, auto_send, open_key):
        self.btn_send.configure(state="disabled")
        self._st("Đang gửi… KHÔNG chạm chuột / bàn phím cho tới khi xong.")

        def prog(s):
            self.after(0, lambda: self._st(s))

        def work():
            done, err = [], None
            try:
                done = send_messages(groups, msgs, auto_send, open_key, prog)
            except Exception as ex:
                err = ex
            self.after(0, lambda: self._finish(done, err, auto_send))

        threading.Thread(target=work, daemon=True).start()

    def _card_sids(self):
        """Ma mau cua cac phieu TUNG MAU da gui kem anh phieu (de tu xoa anh neu co tich)."""
        if not self.v_card.get():
            return []
        return [i.get("sid") for i in self.infos if i and i.get("sid") not in (None, "")]

    def _finish(self, done, err, auto_send):
        try:
            self.btn_send.configure(state="normal")
            self.lift()
        except tk.TclError:
            return
        self.log_del = 0
        if done and not err and self.v_del.get():
            self.log_del = delete_sample_images(self._card_sids())
        if self.on_sent:
            try:
                self.on_sent(done, err)
            except Exception:
                pass
        if err:
            self._st("⚠ %s" % err, True)
            messagebox.showwarning("Gửi Zalo", str(err), parent=self)
        elif auto_send:
            self._st("✔ Đã gửi vào: %s%s" % (", ".join(done), ("  ·  đã xoá %d ảnh màng film/hạt Filler"
                                                              % getattr(self, "log_del", 0))
                                             if getattr(self, "log_del", 0) else ""))
        else:
            self._st("✔ Đã dán vào: %s – vào Zalo bấm Enter để gửi." % ", ".join(done))


def save_card_file(img, base_dir, info=None):
    info = info or {}
    d = os.path.join(base_dir, "BaoCaoZalo", dt.date.today().strftime("%Y-%m"))
    os.makedirs(d, exist_ok=True)
    name = "QC_%s_%s.png" % (("DC" + info["line_no"]) if info.get("line_no") else "TongHop",
                             dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3])
    p = os.path.join(d, name)
    img.save(p)
    return p


class ManualPasteToast(tk.Toplevel):
    def __init__(self, master, go):
        super().__init__(master)
        self.title("Gửi Zalo")
        self.attributes("-topmost", True)
        self.resizable(False, False)
        self.geometry("+%d+%d" % (self.winfo_screenwidth() - 420, 60))
        ttk.Label(self, text="1. Trong Zalo, bấm mở NHÓM cần gửi\n2. Bấm vào ô nhập tin nhắn\n3. Bấm nút dưới đây",
                  padding=12, font=("Segoe UI", 11)).pack()

        def ok():
            self.destroy()
            go()

        tk.Button(self, text="▶  Dán & gửi", font=("Segoe UI", 12, "bold"), bg="#0068ff", fg="white",
                  relief="flat", padx=20, pady=6, command=ok).pack(pady=(0, 8))
        ttk.Button(self, text="Huỷ", command=self.destroy).pack(pady=(0, 10))


class FilmDialog(tk.Toplevel):
    """Chen / xoa anh MAU MANG FILM hoac anh HAT FILLER cho 1 mau (khung ben phai phieu bao cao).
    kind = 'film' | 'filler'"""

    def __init__(self, master, sid, title="", on_change=None, kind="film", dup_check=None):
        super().__init__(master)
        self.sid, self.on_change = sid, on_change
        self.dup_check = dup_check      # ham(PIL.Image) -> canh bao anh trung (moc / mau khac) hoac None
        self.kind = kind if kind in IMG_KINDS else "film"
        _, cap, short = IMG_KINDS[self.kind]
        self.name = "Ảnh " + short
        self.title(("🎞 " if self.kind == "film" else "⚪ ") + "Ảnh %s – " % short + (title or "mẫu #%s" % sid))
        self.transient(master)
        self.geometry("760x520")
        self.minsize(560, 420)
        top = ttk.Frame(self, padding=(10, 8, 10, 0))
        top.pack(fill="x")
        ttk.Label(top, text="Ảnh sẽ hiện trong khung '%s' bên phải phiếu (tối đa %d ảnh)." % (cap, FILM_MAX),
                  foreground="#1f3a5f").pack(anchor="w")
        bt = ttk.Frame(self, padding=(10, 6))
        bt.pack(fill="x")
        ttk.Button(bt, text="📂 Chọn file ảnh…", command=self.add_file).pack(side="left")
        ttk.Button(bt, text="📋 Dán ảnh (Ctrl+V)", command=self.paste).pack(side="left", padx=4)
        ttk.Button(bt, text="🗑 Xoá ảnh đang chọn", command=self.delete).pack(side="left")
        ttk.Button(bt, text="Xong", command=self.destroy).pack(side="right")
        self.cv = tk.Canvas(self, bg="#eef1f5", highlightthickness=0)
        self.cv.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.status = ttk.Label(self, text="", foreground="#1f3a5f")
        self.status.pack(anchor="w", padx=10, pady=(0, 8))
        self.bind("<Control-v>", lambda e: self.paste())
        self.bind("<Control-V>", lambda e: self.paste())
        self.cv.bind("<Configure>", lambda e: self.show())
        self._sel = None
        self._tk = []
        self.grab_set()
        self.focus_set()

    def _changed(self, msg):
        self.status.configure(text=msg)
        self.show()
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass

    def _room(self):
        if len(film_files(self.sid, self.kind)) >= FILM_MAX:
            messagebox.showinfo(self.name, "Đã đủ %d ảnh – xoá bớt trước khi thêm." % FILM_MAX, parent=self)
            return False
        return True

    def add_file(self):
        ps = filedialog.askopenfilenames(parent=self, title="Chọn " + self.name.lower(),
                                         filetypes=[("Ảnh", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Tất cả", "*.*")])
        n = 0
        for p in ps:
            if not self._room():
                break
            try:
                from PIL import Image
                im = Image.open(p)
                im.load()
                if not self._dup_ok(im, os.path.basename(p)):
                    continue
                film_add(self.sid, im, self.kind)
                n += 1
            except Exception as ex:
                messagebox.showwarning(self.name, "Không mở được ảnh:\n%s\n%s" % (p, ex), parent=self)
        if n:
            self._changed("✔ Đã thêm %d ảnh." % n)

    def _dup_ok(self, im, name=""):
        """Anh trung anh da up o moc / mau khac -> hoi co van them khong."""
        if not self.dup_check:
            return True
        try:
            msg = self.dup_check(im)
        except Exception:
            msg = None
        if not msg:
            return True
        return messagebox.askyesno(self.name, ("%s\n\n" % name if name else "") + msg + "\n\nVẫn thêm ảnh này?",
                                   icon="warning", parent=self)

    def paste(self):
        if not self._room():
            return
        try:
            from PIL import ImageGrab
            got = ImageGrab.grabclipboard()
        except Exception as ex:
            messagebox.showwarning(self.name, "Không đọc được clipboard: %s" % ex, parent=self)
            return
        items = got if isinstance(got, list) else ([got] if got is not None else [])
        n = 0
        for it in items:
            if len(film_files(self.sid, self.kind)) >= FILM_MAX:
                break
            try:
                from PIL import Image
                im = Image.open(it) if isinstance(it, str) else it
                im.load()
                if not self._dup_ok(im):
                    continue
                film_add(self.sid, im, self.kind)
                n += 1
            except Exception:
                pass
        if n:
            self._changed("✔ Đã dán %d ảnh." % n)
        elif not items:
            messagebox.showinfo(self.name, "Clipboard chưa có ảnh.\n(Chụp màn hình / Copy ảnh rồi bấm lại.)",
                                parent=self)

    def delete(self):
        fs = film_files(self.sid, self.kind)
        if self._sel is None or self._sel >= len(fs):
            messagebox.showinfo(self.name, "Bấm vào 1 ảnh để chọn trước khi xoá.", parent=self)
            return
        if not messagebox.askyesno(self.name, "Xoá ảnh đang chọn?", parent=self):
            return
        try:
            os.remove(fs[self._sel])
        except Exception as ex:
            messagebox.showwarning(self.name, str(ex), parent=self)
        self._sel = None
        self._changed("Đã xoá ảnh.")

    def show(self):
        from PIL import Image, ImageTk
        self.cv.delete("all")
        self._tk = []
        fs = film_files(self.sid, self.kind)
        w, h = max(200, self.cv.winfo_width()), max(150, self.cv.winfo_height())
        if not fs:
            self.cv.create_text(w // 2, h // 2, text="Chưa có ảnh – bấm 'Chọn file ảnh…' hoặc Ctrl+V để dán",
                                fill="#667", font=("Segoe UI", 11))
            return
        n = len(fs)
        cw = (w - 12 * (n + 1)) // n
        for i, p in enumerate(fs):
            try:
                im = Image.open(p)
                im.thumbnail((cw, h - 30))
                t = ImageTk.PhotoImage(im)
            except Exception:
                continue
            self._tk.append(t)
            x = 12 + i * (cw + 12)
            self.cv.create_image(x + cw // 2, (h - 20) // 2, image=t, tags=("im%d" % i,))
            if i == self._sel:
                self.cv.create_rectangle(x - 4, 4, x + cw + 4, h - 22, outline="#0068ff", width=3)
            self.cv.create_text(x + cw // 2, h - 10, text="Ảnh %d" % (i + 1), fill="#334")
            self.cv.tag_bind("im%d" % i, "<Button-1>", lambda e, k=i: self._pick(k))

    def _pick(self, k):
        self._sel = k
        self.show()


def _tmin(t):
    """'HH:MM' -> phut tinh tu 06:00 ngay SX (0h-5h59 xep cuoi ngay)."""
    try:
        h, m = [int(x) for x in str(t).strip().replace("h", ":").split(":")[:2]]
        v = h * 60 + m
        return v + 1440 if v < 360 else v
    except Exception:
        return -1


def latest_per_line(items, lines=None):
    """[(sid, info)] -> mau moi nhat cua tung day chuyen (loc theo lines = ['1','3',...] neu co)."""
    by = {}
    for sid, i in items:
        ln = i.get("line_no", "")
        if lines and ln not in lines:
            continue
        cur = by.get(ln)
        if cur is None or (_tmin(i.get("time")), sid) > (_tmin(cur[1].get("time")), cur[0]):
            by[ln] = (sid, i)
    return [by[k] for k in sorted(by, key=lambda x: (len(x), x))]


class SamplePicker(tk.Toplevel):
    """Chon tung DAY CHUYEN / tung MOC GIO (mau) de gui Zalo.
    provider can: samples_on(date, ca) -> [(sid, info)], context() -> (date, ca), shifts() -> [ca]"""

    ON, OFF, HALF = "☑", "☐", "▣"

    def __init__(self, master, provider, sids=None, summary=False, on_ok=None, date=None, ca=None, sget=None):
        super().__init__(master)
        self.title("🎯 Chọn dây chuyền / mốc giờ để gửi Zalo")
        self.transient(master)
        self.geometry("1080x600")
        self.sget = sget
        self.minsize(640, 420)
        self.p, self.on_ok = provider, on_ok
        self.checked = set(sids or [])
        self.items = []
        d0, c0 = provider.context() if hasattr(provider, "context") else (dt.date.today(), "")
        d0 = date or d0
        c0 = c0 if ca is None else ca
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Ngày SX:").pack(side="left")
        self.v_date = tk.StringVar(value=d0.strftime("%d/%m/%Y") if hasattr(d0, "strftime") else str(d0))
        e = ttk.Entry(top, textvariable=self.v_date, width=12)
        e.pack(side="left", padx=(4, 10))
        e.bind("<Return>", lambda ev: self.load())
        ttk.Label(top, text="Ca:").pack(side="left")
        shifts = ["(cả ngày)"] + list(provider.shifts() if hasattr(provider, "shifts") else [])
        self.v_ca = tk.StringVar(value=c0 or "(cả ngày)")
        cb = ttk.Combobox(top, textvariable=self.v_ca, values=shifts, state="readonly", width=14)
        cb.pack(side="left", padx=(4, 10))
        cb.bind("<<ComboboxSelected>>", lambda ev: self.load())
        ttk.Button(top, text="🔄 Tải", command=self.load).pack(side="left")
        qb = ttk.Frame(self, padding=(10, 0))
        qb.pack(fill="x")
        ttk.Button(qb, text="☑ Chọn tất cả", command=lambda: self._all(True)).pack(side="left")
        ttk.Button(qb, text="☐ Bỏ chọn", command=lambda: self._all(False)).pack(side="left", padx=4)
        ttk.Button(qb, text="⏱ Mới nhất mỗi DC", command=self._latest).pack(side="left")
        ttk.Label(qb, text="  Bấm dây chuyền = chọn cả DC · bấm mốc giờ = chọn từng mẫu · "
                           "bấm cột 🎞 / ⚪ = up ảnh RIÊNG cho mốc đó",
                  foreground="#666").pack(side="left", padx=8)
        tf = ttk.Frame(self, padding=(10, 6))
        tf.pack(fill="both", expand=True)
        cols = (("time", "Mốc giờ", 80), ("order", "Đơn hàng", 140), ("vmi", "Mã hàng", 140),
                ("res", "Kết quả", 190), ("vd", "Đánh giá", 80), ("fi", "🎞 Màng film", 95),
                ("fl", "⚪ Hạt Filler", 95))
        self.tv = ttk.Treeview(tf, columns=[c for c, _, _ in cols], show="tree headings", selectmode="none")
        self.tv.heading("#0", text="Dây chuyền")
        self.tv.column("#0", width=230, stretch=False)
        for c, h, w in cols:
            self.tv.heading(c, text=h)
            self.tv.column(c, width=w, anchor="center" if c in ("time", "vd", "fi", "fl") else "w")
        vs = ttk.Scrollbar(tf, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=vs.set)
        self.tv.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.tv.tag_configure("line", font=("Segoe UI", 10, "bold"), background="#eaf0f8")
        self.tv.tag_configure("bad", foreground="#c62828")
        self.tv.tag_configure("noimg", background="#fff4e5")
        self.tv.bind("<Button-1>", self._click)
        bt = ttk.Frame(self, padding=(10, 4, 10, 10))
        bt.pack(fill="x")
        self.v_sum = tk.BooleanVar(value=summary)
        ttk.Checkbutton(bt, text="Gộp các mẫu đã chọn thành 1 phiếu TỔNG HỢP", variable=self.v_sum).pack(side="left")
        self.lbl = ttk.Label(bt, text="", foreground="#1f3a5f")
        self.lbl.pack(side="left", padx=12)
        ttk.Button(bt, text="Huỷ", command=self.destroy).pack(side="right")
        tk.Button(bt, text="✔ Dùng các mẫu đã chọn", bg="#0068ff", fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=10, command=self.ok).pack(side="right", padx=6)
        self.load()
        self.grab_set()

    def _ctx(self):
        try:
            d = dt.datetime.strptime(self.v_date.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            messagebox.showwarning("Chọn mẫu", "Ngày không hợp lệ (dd/mm/yyyy).", parent=self)
            return None, None
        ca = self.v_ca.get()
        return d, ("" if ca.startswith("(") else ca)

    def load(self):
        d, ca = self._ctx()
        if d is None:
            return
        try:
            self.items = list(self.p.samples_on(d, ca))
        except Exception as ex:
            self.items = []
            messagebox.showwarning("Chọn mẫu", "Không tải được danh sách mẫu:\n%s" % ex, parent=self)
        self._fill()

    def _fill(self):
        self.tv.delete(*self.tv.get_children())
        by = {}
        for sid, i in self.items:
            by.setdefault(i.get("line_no", "") or "?", []).append((sid, i))
        for ln in sorted(by, key=lambda x: (len(x), x)):
            rows = sorted(by[ln], key=lambda x: (_tmin(x[1].get("time")), x[0]))
            pid = "L" + ln
            self.tv.insert("", "end", iid=pid, text="", open=True, tags=("line",))
            for sid, i in rows:
                res = "  ".join(x for x in (("Ẩm %s" % i["moist"]) if i.get("moist") else "",
                                            ("Hạt %s" % i["beads"]) if i.get("beads") else "",
                                            ("MFI %s" % i["mfi"]) if i.get("mfi") else "",
                                            ("Tro %s" % i["ash"]) if i.get("ash") else "") if x)
                v = i.get("verdict", "")
                nf, nh = len(film_files(sid, "film")), len(film_files(sid, "filler"))
                tags = ("bad",) if v and not _is_ok(v) else ()
                if not (nf and nh):
                    tags += ("noimg",)
                self.tv.insert(pid, "end", iid="S%s" % sid, text="",
                               values=((i.get("time") or "").replace(":", "h"), i.get("order", ""),
                                       i.get("vmi", ""), res, v or "—", self._cnt(nf), self._cnt(nh)),
                               tags=tags)
        self._paint()

    @staticmethod
    def _cnt(n):
        return ("✔ %d ảnh" % n) if n else "＋ thêm"

    def _img(self, iid, kind):
        """Up anh rieng cho 1 moc (mau) - moi moc / moi day chuyen 1 bo anh khac nhau."""
        sid = self._sid(iid)
        info = next((i for s, i in self.items if s == sid), {}) or {}
        dlg = FilmDialog(self, sid, "DC%s · mốc %s · %s" % (info.get("line_no", ""), info.get("time", ""),
                                                            info.get("order", "")),
                         kind=kind, dup_check=make_dup_check(self.p, self.sget))
        self.wait_window(dlg)
        try:
            nf, nh = len(film_files(sid, "film")), len(film_files(sid, "filler"))
            self.tv.set(iid, "fi", self._cnt(nf))
            self.tv.set(iid, "fl", self._cnt(nh))
            tags = tuple(t for t in self.tv.item(iid, "tags") if t != "noimg")
            self.tv.item(iid, tags=tags if (nf and nh) else tags + ("noimg",))
            self._paint()
            self.grab_set()
        except tk.TclError:
            pass

    def _paint(self):
        valid = {sid for sid, _ in self.items}
        n = 0
        for pid in self.tv.get_children():
            kids = self.tv.get_children(pid)
            on = 0
            for k in kids:
                sid = self._sid(k)
                c = sid in self.checked
                on += c
                self.tv.item(k, text="    " + (self.ON if c else self.OFF))
            n += on
            mark = self.ON if kids and on == len(kids) else (self.HALF if on else self.OFF)
            full = sum(1 for k in kids if "noimg" not in self.tv.item(k, "tags"))
            self.tv.item(pid, text="%s  Dây chuyền %s  (%d/%d)" % (mark, pid[1:], on, len(kids)),
                         values=("", "", "", "", "", "đủ ảnh %d/%d" % (full, len(kids)), ""))
        other = len([s for s in self.checked if s not in valid])
        self.lbl.configure(text="Đã chọn %d mẫu%s" % (n, (" (+%d mẫu ngày khác)" % other) if other else ""))

    def _sid(self, iid):
        v = iid[1:]
        try:
            return int(v)
        except ValueError:
            return v

    def _click(self, ev):
        iid = self.tv.identify_row(ev.y)
        if not iid or self.tv.identify_element(ev.x, ev.y) == "Treeitem.indicator":
            return
        col = self.tv.identify_column(ev.x)
        if iid.startswith("S") and col in ("#6", "#7"):          # cot anh mang film / hat filler
            self.after(10, lambda: self._img(iid, "film" if col == "#6" else "filler"))
            return "break"
        if iid.startswith("L"):
            kids = [self._sid(k) for k in self.tv.get_children(iid)]
            if all(k in self.checked for k in kids):
                self.checked.difference_update(kids)
            else:
                self.checked.update(kids)
        else:
            self.checked ^= {self._sid(iid)}
        self._paint()
        return "break"

    def _all(self, on):
        ids = {sid for sid, _ in self.items}
        if on:
            self.checked |= ids
        else:
            self.checked -= ids
        self._paint()

    def _latest(self):
        self.checked -= {sid for sid, _ in self.items}
        self.checked |= {sid for sid, _ in latest_per_line(self.items)}
        self._paint()

    def ok(self):
        if not self.checked:
            messagebox.showinfo("Chọn mẫu", "Chưa chọn mẫu nào.", parent=self)
            return
        order = {sid: n for n, (sid, _) in enumerate(sorted(
            self.items, key=lambda x: (len(x[1].get("line_no", "")), x[1].get("line_no", ""),
                                       _tmin(x[1].get("time")), x[0])))}
        sids = sorted(self.checked, key=lambda s: (order.get(s, 10 ** 9), str(s)))
        summ = self.v_sum.get()
        self.destroy()
        if self.on_ok:
            self.on_ok(sids, summ)


def pick_desc(infos):
    """Mo ta ngan: 'DC1: 08h, 10h · DC3: 09h'"""
    by = {}
    for i in infos:
        by.setdefault(i.get("line_no", "") or "?", []).append((i.get("time") or "").replace(":", "h"))
    return " · ".join("DC%s: %s" % (k, ", ".join(v)) for k, v in sorted(by.items(), key=lambda x: (len(x[0]), x[0])))



class AddGroupDialog(tk.Toplevel):
    """Them nhom Zalo khi o ten dang trong:
    - mo Zalo -> bam vao nhom nao thi phan mem TU LAY TEN nhom do (doc thanh tieu de Zalo) -> Them
    - hoac go / dan ten -> Kiem tra (mo thu nhom, khong gui) -> Them
    - hoac dung muc dac biet 'Nhom dang mo tren Zalo' (tu bam vao nhom trong Zalo, phan mem chi dan)."""

    WATCH_MS = 1200          # bao lau kiem tra Zalo 1 lan khi dang "theo doi"
    WATCH_MAX_SEC = 300      # tu dung theo doi sau 5 phut

    def __init__(self, master, on_add, sget=None):
        super().__init__(master)
        self.on_add, self.sget = on_add, sget
        self._watch_until = 0
        self._busy = False
        self._last_sig = None
        self._last_name = ""
        self._typed = False           # nguoi dung tu go tay -> khong ghi de nua
        self.title("➕ Thêm nhóm Zalo")
        self.transient(master)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Cách 1 – Thêm theo TÊN nhóm (phần mềm tự tìm & mở nhóm khi gửi):",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(f, text="① Bấm 'Mở Zalo & tự lấy tên' → trong Zalo BẤM VÀO NHÓM cần gửi\n"
                          "     → phần mềm tự đọc tên nhóm đó và điền vào ô dưới (có thể sửa lại)\n"
                          "② Bấm 'Kiểm tra trên Zalo' để thử mở nhóm (không gửi gì) → đúng thì bấm 'Thêm'\n"
                          "     (Vẫn gõ tay hoặc copy rồi bấm 'Dán' được như cũ)",
                  foreground="#555", justify="left").pack(anchor="w", pady=(2, 6))
        r = ttk.Frame(f)
        r.pack(fill="x")
        self.v = tk.StringVar()
        self.e = ttk.Entry(r, textvariable=self.v, width=40, font=("Segoe UI", 11))
        self.e.pack(side="left", fill="x", expand=True)
        self.e.bind("<Return>", lambda ev: self.ok())
        self.e.bind("<Key>", self._on_type)
        ttk.Button(r, text="📋 Dán", command=self._paste).pack(side="left", padx=(4, 0))
        b = ttk.Frame(f)
        b.pack(fill="x", pady=(6, 0))
        self.btn_watch = tk.Button(b, text="💬 Mở Zalo & tự lấy tên", bg="#ffb300", fg="#222", relief="flat",
                                   activebackground="#ffca28", font=("Segoe UI", 9, "bold"), padx=6,
                                   command=self._open)
        self.btn_watch.pack(side="left")
        ttk.Button(b, text="📥 Lấy tên nhóm đang mở", command=self._grab_now).pack(side="left", padx=4)
        ttk.Button(b, text="🔎 Kiểm tra trên Zalo", command=self._test).pack(side="left")
        tk.Button(b, text="➕ Thêm", bg="#0068ff", fg="white", relief="flat", font=("Segoe UI", 10, "bold"),
                  padx=14, command=self.ok).pack(side="right")
        ttk.Separator(f).pack(fill="x", pady=10)
        ttk.Label(f, text="Cách 2 – Không cần gõ tên:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(f, text="Thêm mục '%s'. Khi gửi, bạn chỉ cần mở sẵn nhóm trong Zalo,\n"
                          "phần mềm dán & gửi vào đúng cuộc trò chuyện đang mở." % CUR_CHAT,
                  foreground="#555", justify="left").pack(anchor="w", pady=(2, 4))
        ttk.Button(f, text="📌 Thêm '%s'" % CUR_CHAT, command=self._cur).pack(anchor="w")
        self.st = ttk.Label(f, text="", foreground="#1f3a5f", wraplength=520, justify="left")
        self.st.pack(anchor="w", pady=(8, 0))
        ttk.Button(f, text="Đóng", command=self.destroy).pack(anchor="e", pady=(6, 0))
        self.e.focus_set()
        self.grab_set()

    def _msg(self, t, err=False):
        try:
            self.st.configure(text=t, foreground="#b00020" if err else "#1f3a5f")
        except tk.TclError:
            pass

    def _alive(self):
        try:
            return bool(self.winfo_exists())
        except tk.TclError:
            return False

    def _on_type(self, ev):
        if ev.keysym not in ("Return", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            if ev.keysym not in ("Left", "Right", "Home", "End", "Up", "Down") and (not (ev.state & 0x4) or ev.keysym.lower() in ("v", "x")):
                self._typed = True

    def _paste(self):
        try:
            t = self.clipboard_get()
        except tk.TclError:
            t = ""
        t = " ".join(str(t).split())[:120]
        if t:
            self.v.set(t)
            self._typed = True
        else:
            self._msg("Clipboard chưa có chữ.", True)

    def _set_name(self, name):
        if not name or not self._alive():
            return
        self._last_name = name
        self._typed = False
        self.v.set(name)
        self.e.icursor("end")
        self._msg("✔ Đã lấy tên nhóm: '%s'\nKiểm tra lại chữ/dấu cho đúng (sửa được) → bấm 'Kiểm tra trên Zalo' "
                  "hoặc 'Thêm'. Bấm nhóm khác trong Zalo thì tên sẽ tự đổi theo." % name)

    # ---- lay ten 1 lan (Zalo dang mo san 1 nhom)
    def _grab_now(self):
        if self._busy:
            return
        self._busy = True
        self._msg("Đang đọc tên nhóm đang mở trên Zalo… (lần đầu hơi lâu vài giây)")

        def work():
            name, err = "", None
            try:
                name = current_chat_name(activate=True)
            except Exception as ex:
                err = ex
            self._busy = False

            def done():
                if not self._alive():
                    return
                try:
                    self.lift()
                    self.focus_force()
                except tk.TclError:
                    pass
                if err:
                    self._msg("⚠ %s" % err, True)
                elif name:
                    self._set_name(name)
                else:
                    self._msg("⚠ Chưa đọc được tên nhóm. Hãy mở (bấm vào) 1 nhóm trong Zalo, phóng to cửa sổ Zalo "
                              "rồi bấm lại – hoặc gõ tay tên nhóm.", True)
            self.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    # ---- mo Zalo + theo doi: nguoi dung bam nhom nao -> tu lay ten nhom do
    def _open(self):
        try:
            open_zalo_search()
        except Exception as ex:
            self._msg(str(ex), True)
            return
        self._watch_until = time.time() + self.WATCH_MAX_SEC
        self._last_sig = None
        try:
            self.btn_watch.configure(text="👀 Đang theo dõi Zalo…", bg="#43a047", fg="white")
        except tk.TclError:
            pass
        self._msg("Đã mở Zalo – tìm và BẤM VÀO NHÓM cần gửi. Phần mềm sẽ tự đọc tên nhóm và điền vào ô "
                  "(xong quay lại cửa sổ này bấm 'Thêm').")
        threading.Thread(target=self._warm, daemon=True).start()
        self.after(self.WATCH_MS, self._watch)

    @staticmethod
    def _warm():
        try:                                   # nap san bo doc chu (OCR) cho lan doc dau nhanh hon
            import image_recipe as ir
            ir._rapid_engine()
        except Exception:
            pass

    def _stop_watch(self):
        self._watch_until = 0
        try:
            self.btn_watch.configure(text="💬 Mở Zalo & tự lấy tên", bg="#ffb300", fg="#222")
        except tk.TclError:
            pass

    def _watch(self):
        if not self._alive() or not self._watch_until:
            return
        if time.time() > self._watch_until:
            self._stop_watch()
            return
        h = zalo_foreground()
        if h and not self._busy and not self._typed:
            self._busy = True
            threading.Thread(target=self._watch_work, args=(h,), daemon=True).start()
        self.after(self.WATCH_MS, self._watch)

    def _watch_work(self, h):
        name = ""
        try:
            img, s = grab_zalo_header(h)
            # chi doc chu khi thanh tieu de thay doi (bam sang nhom khac)
            sig = img.convert("L").resize((64, 8)).tobytes()
            if sig != self._last_sig:
                self._last_sig = sig
                name = read_chat_title(img, s)
        except Exception:
            name = ""
        finally:
            self._busy = False
        if name and name != self._last_name:
            try:
                self.after(0, lambda: (not self._typed) and self._set_name(name))
            except (tk.TclError, RuntimeError):
                pass

    def _test(self):
        n = self.v.get().strip()
        if not n:
            self._msg("Gõ tên nhóm trước khi kiểm tra.", True)
            return
        self._stop_watch()
        key = (self.sget("zalo_open_key", "enter") if self.sget else "enter") or "enter"
        self._msg("Đang mở thử nhóm trên Zalo… (không chạm chuột/bàn phím)")

        def work():
            err = None
            try:
                test_group(n, key)
            except Exception as ex:
                err = ex
            self.after(0, lambda: self._msg(
                ("⚠ %s" % err) if err else
                "✔ Đã mở thử. Xem Zalo: nếu đúng nhóm '%s' đang mở → bấm 'Thêm'. "
                "Nếu mở nhầm, sửa tên cho đúng/đầy đủ hơn rồi kiểm tra lại." % n, bool(err)))

        threading.Thread(target=work, daemon=True).start()

    def destroy(self):
        self._watch_until = 0
        super().destroy()

    def _cur(self):
        self.destroy()
        self.on_add(CUR_CHAT)

    def ok(self):
        n = self.v.get().strip()
        if not n:
            self._msg("Chưa có tên nhóm.", True)
            return
        self.destroy()
        self.on_add(n)


class GroupBox(ttk.LabelFrame):
    """Danh sach nhom Zalo (luu trong cai dat 'zalo_groups'), chon nhieu."""

    def __init__(self, master, sget, sset, height=7, key_last="zalo_last_groups"):
        super().__init__(master, text=" Nhóm Zalo (chọn 1 hoặc nhiều) ", padding=6)
        self.sget, self.sset, self.key_last = sget, sset, key_last
        lf = ttk.Frame(self)
        lf.pack(fill="both", expand=True)
        self.lb = tk.Listbox(lf, selectmode="extended", height=height, font=("Segoe UI", 10),
                             exportselection=False)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        ad = ttk.Frame(self)
        ad.pack(fill="x", pady=(4, 0))
        self.v_new = tk.StringVar()
        e = ttk.Entry(ad, textvariable=self.v_new)
        e.pack(side="left", fill="x", expand=True)
        e.bind("<Return>", lambda ev: self.add())
        ttk.Button(ad, text="➕", width=3, command=self.add).pack(side="left", padx=2)
        ttk.Button(ad, text="🗑", width=3, command=self.delete).pack(side="left")
        self.reload()

    def all(self):
        return _groups_of(self.sget)

    def reload(self, select=None):
        g = self.all()
        sel = select if select is not None else (self.sget(self.key_last, "") or "").split("\n")
        self.lb.delete(0, "end")
        for i, x in enumerate(g):
            self.lb.insert("end", x)
            if x in sel:
                self.lb.selection_set(i)

    def selected(self):
        return [self.lb.get(i) for i in self.lb.curselection()]

    def remember(self):
        self.sset(self.key_last, "\n".join(self.selected()))

    def add(self):
        n = self.v_new.get().strip()
        if not n:                      # o trong -> mo hop thoai: mo Zalo tim nhom / dung nhom dang mo
            AddGroupDialog(self.winfo_toplevel(), self._add_name, self.sget)
            return
        self._add_name(n)
        self.v_new.set("")

    def _add_name(self, n):
        n = (n or "").strip()
        if not n:
            return
        g = self.all()
        if n not in g:
            g.append(n)
            self.sset("zalo_groups", json.dumps(g, ensure_ascii=False))
        self.reload(self.selected() + [n])

    def delete(self):
        names = self.selected()
        if not names:
            return
        if not messagebox.askyesno("Xoá nhóm", "Xoá khỏi danh sách:\n%s ?" % "\n".join(names), parent=self):
            return
        g = [x for x in self.all() if x not in names]
        self.sset("zalo_groups", json.dumps(g, ensure_ascii=False))
        self.reload([])


# ============================================================================
# 5) PHAN "GUI ZALO & HEN GIO" (nhung vao cuoi tab Thanh pham)
# ============================================================================
CONTENTS = [("selected", "Mẫu đang chọn"),
            ("latest", "Mẫu mới nhất từng dây chuyền"),
            ("summary", "Tổng hợp ca (tất cả mẫu)"),
            ("pick", "Chọn dây chuyền / mốc giờ")]
REPEATS = [("once", "Một lần"), ("daily", "Hằng ngày"), ("interval", "Lặp mỗi … phút")]
GRACE_MIN = 15          # tre qua 15 phut (vd may tat) -> bo qua lan do


def _lbl(pairs, key):
    return dict(pairs).get(key, key)


def _pkey(pairs, label):
    return next((k for k, v in pairs if v == label), pairs[0][0])


def _hm(s):
    try:
        h, m = str(s).strip().replace("h", ":").split(":")[:2]
        h, m = int(h), int(m)
        if 0 <= h < 24 and 0 <= m < 60:
            return h, m
    except Exception:
        pass
    return None


def job_next(job, now=None):
    """Thoi diem gui ke tiep (datetime) hoac None."""
    now = now or dt.datetime.now()
    hm = _hm(job.get("time")) or (0, 0)
    last = dt.datetime.fromisoformat(job["last"]) if job.get("last") else None
    rep = job.get("repeat")
    if rep == "once":
        try:
            d = dt.date.fromisoformat(job.get("date"))
        except Exception:
            return None
        t = dt.datetime.combine(d, dt.time(*hm))
        return None if last else t
    if rep == "daily":
        t = dt.datetime.combine(now.date(), dt.time(*hm))
        if last and last >= t:
            t += dt.timedelta(days=1)
        elif now - t > dt.timedelta(minutes=GRACE_MIN):
            t += dt.timedelta(days=1)
        return t
    if rep == "interval":
        every = max(1, int(job.get("every") or 30))
        anchor = dt.datetime.combine(now.date(), dt.time(*hm))
        if last is None:
            return anchor if now < anchor else now
        n = last + dt.timedelta(minutes=every)
        return max(n, anchor) if last < anchor else n
    return None


def job_due(job, now=None):
    now = now or dt.datetime.now()
    if not job.get("on", True):
        return False
    t = job_next(job, now)
    if t is None or now < t:
        return False
    if job.get("repeat") in ("once", "daily") and now - t > dt.timedelta(minutes=GRACE_MIN):
        return False
    return True


class JobDialog(tk.Toplevel):
    def __init__(self, master, sget, sset, job=None, on_ok=None, sel_sid=None, provider=None, open_imgs=None):
        super().__init__(master)
        self.p = provider
        self.open_imgs = open_imgs      # ham(job_id) mo bang anh theo moc / day chuyen
        self.title("⏰ Hẹn giờ gửi Zalo")
        self.transient(master)
        self.resizable(False, False)
        self.sget, self.on_ok, self.sel_sid = sget, on_ok, sel_sid
        self.job = dict(job or {})
        now = dt.datetime.now()
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        r = 0

        def row(label, w):
            nonlocal r
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="w", pady=3, padx=(0, 8))
            w.grid(row=r, column=1, sticky="we", pady=3)
            r += 1

        self.v_rep = tk.StringVar(value=_lbl(REPEATS, self.job.get("repeat", "daily")))
        cb = ttk.Combobox(f, textvariable=self.v_rep, values=[v for _, v in REPEATS], state="readonly", width=24)
        row("Kiểu lặp:", cb)
        self.v_time = tk.StringVar(value=self.job.get("time") or (now + dt.timedelta(minutes=5)).strftime("%H:%M"))
        tf_ = ttk.Frame(f)
        ttk.Entry(tf_, textvariable=self.v_time, width=30).pack(side="left")
        row("Giờ gửi (HH:MM):", tf_)
        if not self.job.get("id"):
            ttk.Label(f, text="Nhiều mốc: gõ cách nhau dấu phẩy, vd  08:00, 10:00, 14:00, 16:30\n"
                              "→ mỗi mốc thành 1 lịch riêng, up ảnh màng film / hạt RIÊNG cho từng mốc + dây chuyền.",
                      foreground="#0068ff").grid(row=r, column=1, sticky="w")
            r += 1
        self.v_date = tk.StringVar(value=self._vn(self.job.get("date")) or now.strftime("%d/%m/%Y"))
        self.e_date = ttk.Entry(f, textvariable=self.v_date, width=12)
        row("Ngày (lặp một lần):", self.e_date)
        self.v_every = tk.StringVar(value=str(self.job.get("every") or 60))
        self.e_every = ttk.Spinbox(f, from_=5, to=720, increment=5, textvariable=self.v_every, width=8)
        row("Lặp mỗi (phút):", self.e_every)
        self.v_content = tk.StringVar(value=_lbl(CONTENTS, self.job.get("content", "latest")))
        cc = ttk.Combobox(f, textvariable=self.v_content, values=[v for _, v in CONTENTS], state="readonly", width=32)
        row("Nội dung:", cc)
        cc.bind("<<ComboboxSelected>>", lambda e: self._sync())
        # loc day chuyen (cho 'moi nhat' / 'tong hop'): khong tich = tat ca
        lf = ttk.Frame(f)
        self.line_vars = {}
        jl = set(self.job.get("lines") or [])
        for ln in (provider.line_nos() if provider is not None and hasattr(provider, "line_nos")
                   else _default_lines()):
            v = tk.BooleanVar(value=ln in jl)
            self.line_vars[ln] = v
            ttk.Checkbutton(lf, text="DC" + ln, variable=v).pack(side="left", padx=(0, 4))
        row("Chỉ dây chuyền:", lf)
        ttk.Label(f, text="(không tích = tất cả dây chuyền)", foreground="#888").grid(
            row=r, column=1, sticky="w")
        r += 1
        # chon tung moc (mau cu the)
        pf = ttk.Frame(f)
        self.pick_sids = list(self.job.get("sids") or [])
        self.pick_sum = bool(self.job.get("pick_summary"))
        self.btn_pick = ttk.Button(pf, text="🎯 Chọn DC / mốc…", command=self._pick)
        self.btn_pick.pack(side="left")
        self.lbl_pick = ttk.Label(pf, text="", foreground="#1f3a5f", wraplength=260)
        self.lbl_pick.pack(side="left", padx=6)
        row("Mẫu đã chọn:", pf)
        self._pick_lbl()
        self.v_new = tk.BooleanVar(value=self.job.get("only_new", True))
        ttk.Checkbutton(f, text="Chỉ gửi khi có mẫu MỚI (chỉ áp dụng cho 'Lặp mỗi … phút')", variable=self.v_new
                        ).grid(row=r, column=0, columnspan=2, sticky="w")
        r += 1
        # anh rieng theo moc (lich 'Mau moi nhat tung day chuyen'): lich moi mac dinh BAT
        self.v_slot = tk.BooleanVar(value=self.job.get("slot_img_only", not self.job.get("id")))
        sf = ttk.Frame(f)
        sf.grid(row=r, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(sf, text="Mỗi mốc dùng ẢNH RIÊNG theo dây chuyền (không lấy lại ảnh cũ của mẫu)",
                        variable=self.v_slot).pack(side="left")
        if self.job.get("id") and self.open_imgs:
            ttk.Button(sf, text="🖼 Ảnh của mốc này…",
                       command=lambda: self.open_imgs(self.job["id"])).pack(side="left", padx=6)
        r += 1
        self.v_send = tk.BooleanVar(value=self.job.get("auto_send", True))
        ttk.Checkbutton(f, text="Tự bấm Gửi (Enter)", variable=self.v_send).grid(row=r, column=0, columnspan=2, sticky="w")
        r += 1
        self.gb = GroupBox(f, sget, sset, height=5, key_last="__none__")
        self.gb.grid(row=r, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        self.gb.reload(self.job.get("groups") or (sget("zalo_last_groups", "") or "").split("\n"))
        r += 1
        ttk.Label(f, text="Lịch chỉ chạy khi phần mềm đang mở và Zalo PC đã đăng nhập.\n"
                          "Trước khi gửi sẽ đếm ngược 10 giây để bạn huỷ nếu đang gõ phím.",
                  foreground="#666").grid(row=r, column=0, columnspan=2, sticky="w", pady=(6, 0))
        r += 1
        bt = ttk.Frame(f)
        bt.grid(row=r, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(bt, text="💾 Lưu lịch", command=self.ok).pack(side="left", padx=3)
        ttk.Button(bt, text="Huỷ", command=self.destroy).pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._sync())
        self._sync()
        self.grab_set()

    @staticmethod
    def _vn(iso):
        try:
            return dt.date.fromisoformat(iso).strftime("%d/%m/%Y")
        except Exception:
            return ""

    def _pick_lbl(self):
        if not self.pick_sids:
            self.lbl_pick.configure(text="(chưa chọn)")
            return
        infos = [i for i in (self.p.info_of(s) if self.p is not None else None for s in self.pick_sids) if i]
        self.lbl_pick.configure(text="%d mẫu%s – %s" % (len(self.pick_sids), ", gộp tổng hợp" if self.pick_sum else "",
                                                        pick_desc(infos)))

    def _pick(self):
        if self.p is None or not hasattr(self.p, "samples_on"):
            messagebox.showinfo("Hẹn giờ", "Chưa hỗ trợ chọn mẫu ở đây.", parent=self)
            return

        def done(sids, summ):
            self.pick_sids, self.pick_sum = sids, summ
            self._pick_lbl()
            self.v_content.set(_lbl(CONTENTS, "pick"))
            self._sync()

        SamplePicker(self, self.p, self.pick_sids, self.pick_sum, done, sget=self.sget)

    def _sync(self):
        rep = _pkey(REPEATS, self.v_rep.get())
        self.e_date.configure(state="normal" if rep == "once" else "disabled")
        self.e_every.configure(state="normal" if rep == "interval" else "disabled")

    def ok(self):
        raw = [t for t in re.split(r"[,;\s]+", self.v_time.get().strip()) if t]
        hms = [_hm(t) for t in raw]
        if not raw or not all(hms):
            bad = ", ".join(t for t, h in zip(raw, hms) if not h)
            messagebox.showwarning("Hẹn giờ", "Giờ không hợp lệ%s (vd 14:30 hoặc 08:00, 10:00, 14:00)." %
                                   ((": " + bad) if bad else ""), parent=self)
            return
        hms = sorted(set(hms))
        rep = _pkey(REPEATS, self.v_rep.get())
        if len(hms) > 1 and self.job.get("id"):
            messagebox.showwarning("Hẹn giờ", "Khi SỬA lịch chỉ nhập 1 giờ.\nMuốn thêm mốc: bấm '➕ Thêm lịch'.",
                                   parent=self)
            return
        if len(hms) > 1 and rep == "interval":
            messagebox.showwarning("Hẹn giờ", "Kiểu 'Lặp mỗi … phút' chỉ nhập 1 giờ bắt đầu.", parent=self)
            return
        hm = hms[0]
        job = dict(self.job)
        job.update({"repeat": rep, "time": "%02d:%02d" % hm, "content": _pkey(CONTENTS, self.v_content.get()),
                    "only_new": self.v_new.get(), "auto_send": self.v_send.get(),
                    "slot_img_only": self.v_slot.get(),
                    "groups": self.gb.selected(), "on": job.get("on", True)})
        if not job["groups"]:
            messagebox.showwarning("Hẹn giờ", "Chọn ít nhất 1 nhóm Zalo.", parent=self)
            return
        if rep == "once":
            try:
                job["date"] = dt.datetime.strptime(self.v_date.get().strip(), "%d/%m/%Y").date().isoformat()
            except ValueError:
                messagebox.showwarning("Hẹn giờ", "Ngày không hợp lệ (dd/mm/yyyy).", parent=self)
                return
            job["last"] = None
        if rep == "interval":
            try:
                job["every"] = max(5, int(self.v_every.get()))
            except ValueError:
                job["every"] = 60
        job["lines"] = [k for k, v in self.line_vars.items() if v.get()]
        if job["content"] == "pick":
            if not self.pick_sids:
                messagebox.showwarning("Hẹn giờ", "Bấm 'Chọn DC / mốc…' để chọn mẫu cần gửi.", parent=self)
                return
            job["sids"], job["pick_summary"] = list(self.pick_sids), self.pick_sum
        if job["content"] == "selected":
            if self.sel_sid is None and not job.get("sid"):
                messagebox.showwarning("Hẹn giờ", "Chưa chọn mẫu nào ở bảng mẫu phía trên.", parent=self)
                return
            if self.sel_sid is not None:
                job["sid"] = self.sel_sid
        if job.get("id"):
            jobs = [job]
        else:                                     # lich moi: moi moc gio = 1 lich (anh rieng tung moc)
            base = int(time.time() * 1000)
            jobs = []
            for n, h in enumerate(hms):
                j = dict(job)
                j["time"] = "%02d:%02d" % h
                j["id"] = "%d" % (base + n)
                if len(hms) > 1:
                    j["grp"] = "%d" % base
                jobs.append(j)
        self.destroy()
        if self.on_ok:
            self.on_ok(jobs if len(jobs) > 1 else jobs[0])


class CountdownToast(tk.Toplevel):
    """Bao truoc 'sap gui Zalo' - cho huy neu dang go phim."""

    def __init__(self, master, text, seconds, go, cancel):
        super().__init__(master)
        self.title("Sắp gửi Zalo")
        self.attributes("-topmost", True)
        self.resizable(False, False)
        self.geometry("+%d+%d" % (self.winfo_screenwidth() - 440, 60))
        self.go, self.cancel_cb, self.left = go, cancel, seconds
        tk.Label(self, text="⏰ " + text, font=("Segoe UI", 11, "bold"), wraplength=380, justify="left",
                 padx=12, pady=8).pack(anchor="w")
        self.lbl = tk.Label(self, text="", font=("Segoe UI", 10), fg="#b00020", padx=12)
        self.lbl.pack(anchor="w")
        bt = tk.Frame(self, padx=12, pady=10)
        bt.pack(fill="x")
        tk.Button(bt, text="▶ Gửi ngay", bg="#0068ff", fg="white", relief="flat", padx=12,
                  command=self._go).pack(side="left")
        tk.Button(bt, text="✖ Huỷ lần này", relief="groove", padx=10, command=self._cancel).pack(side="left", padx=8)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._tick()

    def _tick(self):
        if self.left <= 0:
            self._go()
            return
        self.lbl.configure(text="Tự gửi sau %d giây – đừng chạm chuột/bàn phím khi đang gửi." % self.left)
        self.left -= 1
        self._aid = self.after(1000, self._tick)

    def _go(self):
        try:
            self.after_cancel(self._aid)
        except Exception:
            pass
        self.destroy()
        self.go()

    def _cancel(self):
        try:
            self.after_cancel(self._aid)
        except Exception:
            pass
        self.destroy()
        self.cancel_cb()


def _default_lines():
    """So cac day chuyen dang dung (line_config.py) - mac dinh 1..6."""
    try:
        import line_config
        return [str(i) for i in line_config.nos()]
    except Exception:
        return [str(i) for i in range(1, 7)]


def job_lines(job, provider=None):
    """Day chuyen cua 1 lich: da tich trong lich, khong tich = moi day chuyen."""
    ls = [str(x) for x in (job.get("lines") or [])]
    if ls:
        return ls
    try:
        return [str(x) for x in provider.line_nos()] if provider is not None else _default_lines()
    except Exception:
        return _default_lines()


def job_uses_slot_imgs(job):
    """Lich nao dung anh theo moc: 'Mau moi nhat tung day chuyen' (luc hen gio mau chua co).
    Lich 'Chon DC / moc' / 'Mau dang chon' dung anh rieng cua TUNG MAU (moc) da chon."""
    return job.get("content") == "latest"


def slot_missing(job, lines):
    """-> [(day chuyen, 'màng film' / 'hạt Filler' / 'màng film + hạt Filler')] con thieu anh"""
    out = []
    for ln in lines:
        nf, nh = slot_counts(job.get("id"), ln)
        miss = [IMG_KINDS[k][2] for k, n in (("film", nf), ("filler", nh)) if not n]
        if miss:
            out.append((ln, " + ".join(miss)))
    return out


class SlotImagesDialog(tk.Toplevel):
    """Bang UP ANH theo MOC GUI x DAY CHUYEN.
    Moi o = 1 moc (lich) cua 1 day chuyen: 🎞 anh mang film | ⚪ anh hat Filler (moi o 1 bo anh rieng)."""

    def __init__(self, master, panel, only_job=None):
        super().__init__(master)
        self.panel, self.only_job = panel, only_job
        self.title("🖼 Ảnh màng film / hạt Filler theo MỐC GỬI & DÂY CHUYỀN")
        self.transient(master)
        self.geometry("1100x560")
        self.minsize(720, 380)
        top = ttk.Frame(self, padding=(10, 8, 10, 2))
        top.pack(fill="x")
        ttk.Label(top, text="Mỗi ô = ảnh RIÊNG của 1 mốc gửi cho 1 dây chuyền.  Bấm 🎞 (màng film) / ⚪ (hạt Filler) "
                            "để chọn file, dán Ctrl+V hoặc xoá ảnh.   Xanh = đã có ảnh · Đỏ = còn thiếu.",
                  foreground="#1f3a5f", wraplength=1050).pack(anchor="w")
        ttk.Label(top, text="Áp dụng cho lịch 'Mẫu mới nhất từng dây chuyền'. Lịch 'Chọn DC / mốc' dùng ảnh của "
                            "từng mẫu: bấm '🎯 Chọn…' rồi bấm cột 🎞 / ⚪ của mốc đó.  Ảnh trùng mốc khác sẽ được cảnh báo.",
                  foreground="#666", wraplength=1050).pack(anchor="w")
        body = ttk.Frame(self, padding=(10, 4))
        body.pack(fill="both", expand=True)
        self.cv = tk.Canvas(body, highlightthickness=0, bg="#f7f9fc")
        vs = ttk.Scrollbar(body, orient="vertical", command=self.cv.yview)
        hs = ttk.Scrollbar(body, orient="horizontal", command=self.cv.xview)
        self.cv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")
        self.cv.pack(side="left", fill="both", expand=True)
        self.gf = tk.Frame(self.cv, bg="#f7f9fc")
        self.cv.create_window((0, 0), window=self.gf, anchor="nw")
        self.gf.bind("<Configure>", lambda e: self.cv.configure(scrollregion=self.cv.bbox("all")))
        bt = ttk.Frame(self, padding=(10, 4, 10, 10))
        bt.pack(fill="x")
        self.lbl = ttk.Label(bt, text="", foreground="#1f3a5f")
        self.lbl.pack(side="left")
        ttk.Button(bt, text="Đóng", command=self.destroy).pack(side="right")
        ttk.Button(bt, text="🔄 Làm mới", command=self.fill).pack(side="right", padx=6)
        self.fill()
        self.grab_set()

    def _jobs(self):
        js = [j for j in self.panel.jobs() if job_uses_slot_imgs(j)]
        if self.only_job:
            js = [j for j in js if j["id"] == self.only_job] or js
        now = dt.datetime.now()
        return sorted(js, key=lambda j: (not j.get("on", True), _hm(j.get("time")) or (99, 0),
                                         job_next(j, now) or dt.datetime.max))

    def fill(self):
        for w in self.gf.winfo_children():
            w.destroy()
        js = self._jobs()
        if not js:
            # Co the dang co lich 'Chon DC/moc' / 'Mau dang chon' -> huong dan dat anh dung cho
            others = [j for j in self.panel.jobs()
                      if (j.get("content") in ("pick", "selected")) and not self.only_job]
            if others:
                msg = ("Cửa sổ này CHỈ dùng cho lịch 'Mẫu mới nhất từng dây chuyền'.\n\n"
                       "Lịch bạn đang có là 'Chọn DC / mốc' (hoặc 'Mẫu đang chọn') — ảnh gắn theo "
                       "TỪNG MẪU đã chọn, KHÔNG đặt ở đây. Cách làm:\n"
                       "   1) Ở bảng lịch bên dưới, chọn đúng lịch đó → bấm '✏ Sửa'.\n"
                       "   2) Bấm '🎯 Chọn DC / mốc…' để mở bảng chọn mẫu.\n"
                       "   3) Trong bảng đó, bấm cột 🎞 (màng film) / ⚪ (hạt Filler) của từng mốc "
                       "để chọn ảnh cho dây chuyền.\n\n"
                       "Nếu muốn gửi theo 'mẫu mới nhất từng dây chuyền' (lúc hẹn giờ mẫu chưa có), "
                       "hãy bấm '➕ Thêm lịch' → Nội dung: Mẫu mới nhất từng dây chuyền.")
            else:
                msg = ("Chưa có lịch 'Mẫu mới nhất từng dây chuyền'.\n"
                       "Bấm '➕ Thêm lịch' → Nội dung: Mẫu mới nhất từng dây chuyền → giờ gửi "
                       "(nhiều mốc: 08:00, 10:00, 14:00).")
            tk.Label(self.gf, text=msg, bg="#f7f9fc", fg="#555", font=("Segoe UI", 10),
                     justify="left", wraplength=980).grid(padx=20, pady=20, sticky="w")
            self.lbl.configure(text="")
            return
        all_lines = []
        for ln in job_lines({}, self.panel.p):
            if ln not in all_lines:
                all_lines.append(ln)
        for j in js:
            for ln in job_lines(j, self.panel.p):
                if ln not in all_lines:
                    all_lines.append(ln)
        hdr = dict(bg="#1f3a5f", fg="white", font=("Segoe UI", 10, "bold"), padx=8, pady=5)
        tk.Label(self.gf, text="Mốc gửi", **hdr).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        for c, ln in enumerate(all_lines, 1):
            tk.Label(self.gf, text="Dây chuyền %s" % ln, **hdr).grid(row=0, column=c, sticky="nsew", padx=1, pady=1)
        tk.Label(self.gf, text="", **hdr).grid(row=0, column=len(all_lines) + 1, sticky="nsew", padx=1, pady=1)
        now = dt.datetime.now()
        full = tot = 0
        for r, j in enumerate(js, 1):
            n = job_next(j, now) if j.get("on", True) else None
            on = j.get("on", True)
            txt = "%s  %s\n%s%s" % ("⏰" if on else "⏸", self.panel._when(j),
                                    ("lần tới %s" % n.strftime("%d/%m %H:%M")) if n else "(đang tắt)",
                                    "" if j.get("slot_img_only") else "  · thiếu → dùng ảnh mẫu")
            tk.Label(self.gf, text=txt, bg="#eaf0f8" if on else "#eeeeee", fg="#1f3a5f" if on else "#888",
                     font=("Segoe UI", 9, "bold"), justify="left", anchor="w", padx=8, pady=4
                     ).grid(row=r, column=0, sticky="nsew", padx=1, pady=1)
            mine = job_lines(j, self.panel.p)
            for c, ln in enumerate(all_lines, 1):
                cell = tk.Frame(self.gf, bg="white", padx=4, pady=4)
                cell.grid(row=r, column=c, sticky="nsew", padx=1, pady=1)
                if ln not in mine:
                    tk.Label(cell, text="—", bg="white", fg="#bbb").pack()
                    continue
                nf, nh = slot_counts(j["id"], ln)
                tot += 1
                full += bool(nf and nh)
                for kind, cnt, ic in (("film", nf, "🎞"), ("filler", nh, "⚪")):
                    tk.Button(cell, text="%s %s" % (ic, cnt if cnt else "＋"), relief="flat", width=5,
                              bg="#2e7d32" if cnt else "#ffdede", fg="white" if cnt else "#b00020",
                              activebackground="#43a047" if cnt else "#ffc4c4", font=("Segoe UI", 9, "bold"),
                              command=lambda jj=j, l=ln, k=kind: self.edit(jj, l, k)).pack(side="left", padx=2)
            tk.Button(self.gf, text="🗑 Xoá ảnh mốc", relief="groove", font=("Segoe UI", 8),
                      command=lambda jj=j: self.clear(jj)).grid(row=r, column=len(all_lines) + 1, padx=4)
        self.lbl.configure(text="Đủ ảnh (màng film + hạt): %d / %d ô" % (full, tot))

    def edit(self, job, line, kind):
        dlg = FilmDialog(self, slot_key(job["id"], line), "Mốc %s · DC%s" % (job.get("time"), line), kind=kind,
                         dup_check=make_dup_check(self.panel.p, self.panel.sget))
        self.wait_window(dlg)
        try:
            self.fill()
            self.grab_set()
        except tk.TclError:
            return
        self.panel.refresh_jobs()

    def clear(self, job):
        if not messagebox.askyesno("Xoá ảnh mốc", "Xoá TẤT CẢ ảnh màng film + hạt Filler của mốc %s?"
                                   % self.panel._when(job), parent=self):
            return
        n = delete_slot_images(job["id"])
        self.fill()
        self.panel.refresh_jobs()
        self.lbl.configure(text="Đã xoá %d ảnh của mốc %s." % (n, job.get("time")))


class ZaloPanel(ttk.Frame):
    """Phan GUI ZALO & HEN GIO o cuoi tab Thanh pham.
    provider can co: selected() -> (sid, info)|None, info_of(sid) -> info|None,
    latest(now) -> [(sid, info)], summary(now) -> (title, subtitle, [(sid, info)]), logo()"""

    TICK_MS = 15000
    RETRY_SEC = 60          # gui loi (Zalo mat tieu diem / dang ban) -> thu lai sau 60 giay
    MAX_TRY = 3

    def __init__(self, master, provider, store_get, store_set, base_dir, on_toggle=None, is_admin=None):
        super().__init__(master)
        self.p, self.sget, self.sset, self.base_dir = provider, store_get, store_set, base_dir
        self.is_admin = is_admin      # True/False hoac ham tra ve True/False (tai khoan Toan quyen)
        set_base_dir(base_dir)
        self.on_toggle = on_toggle
        try:
            self.pick_sids = json.loads(self.sget("zalo_pick_sids", "[]") or "[]")
        except Exception:
            self.pick_sids = []
        self.pick_sum = _flag(self.sget, "zalo_pick_summary", False)
        self._toast = None
        self._running = False
        self._job_busy = None       # id lich dang dem nguoc / dang gui
        self._fails = {}            # id lich -> (so lan loi, thoi diem duoc thu lai)
        hd = tk.Frame(self, bg="#0068ff")
        hd.pack(fill="x")
        tk.Label(hd, text="📤  GỬI ZALO & HẸN GIỜ GỬI", bg="#0068ff", fg="white",
                 font=("Segoe UI", 11, "bold"), padx=10, pady=5).pack(side="left")
        self.btn_tg = tk.Button(hd, text="", bg="#0068ff", fg="white", relief="flat", activebackground="#0052cc",
                                font=("Segoe UI", 9, "bold"), command=self.toggle)
        self.btn_tg.pack(side="right", padx=6)
        self.lbl_next = tk.Label(hd, text="", bg="#0068ff", fg="#dfeaff", font=("Segoe UI", 9))
        self.lbl_next.pack(side="right", padx=10)
        self.body = ttk.Frame(self, padding=(6, 6, 6, 4))
        self._build_body()
        self.collapsed = self.sget("zalo_panel_collapsed", "0") == "1"
        self._apply_collapse()
        self.refresh_jobs()
        self.after(3000, self._tick)

    # ---------------- giao dien
    def _build_body(self):
        b = self.body
        b.columnconfigure(0, weight=2)
        b.columnconfigure(1, weight=2, minsize=260)
        b.columnconfigure(2, weight=5)
        self.gb = GroupBox(b, self.sget, self.sset, height=6)
        self.gb.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        now = ttk.LabelFrame(b, text=" Gửi ngay ", padding=6)
        now.grid(row=0, column=1, sticky="nsew", padx=(0, 6))
        self.v_content = tk.StringVar(value=self.sget("zalo_now_content", "selected") or "selected")
        for k, v in CONTENTS:
            if k == "pick":
                pr = ttk.Frame(now)
                pr.pack(fill="x")
                ttk.Radiobutton(pr, text="Chọn DC / mốc:", value=k, variable=self.v_content).pack(side="left")
                tk.Button(pr, text="🎯 Chọn…", bg="#ffb300", fg="#222", relief="flat", activebackground="#ffca28",
                          font=("Segoe UI", 9, "bold"), padx=6, command=self.pick).pack(side="left", padx=(4, 0))
                self.lbl_pick = ttk.Label(now, text="", foreground="#1f3a5f", wraplength=230, justify="left")
                self.lbl_pick.pack(anchor="w", padx=(20, 0))
            else:
                ttk.Radiobutton(now, text=v, value=k, variable=self.v_content).pack(anchor="w")
        self.v_text = tk.BooleanVar(value=_flag(self.sget, "zalo_send_text", True))
        self.v_card = tk.BooleanVar(value=_flag(self.sget, "zalo_send_card", True))
        self.v_send = tk.BooleanVar(value=_flag(self.sget, "zalo_auto_send", True))
        self.v_down = tk.BooleanVar(value=self.sget("zalo_open_key", "enter") == "down_enter")
        o = ttk.Frame(now)
        o.pack(fill="x", pady=(4, 0))
        ttk.Checkbutton(o, text="Chữ", variable=self.v_text).pack(side="left")
        ttk.Checkbutton(o, text="Ảnh phiếu", variable=self.v_card).pack(side="left", padx=6)
        ttk.Checkbutton(now, text="Tự bấm Gửi (Enter)", variable=self.v_send).pack(anchor="w")
        ttk.Checkbutton(now, text="Mở nhóm bằng ↓ + Enter", variable=self.v_down).pack(anchor="w")
        bt = ttk.Frame(now)
        bt.pack(fill="x", pady=(6, 0))
        self.btn_now = tk.Button(bt, text="📤 Gửi ngay", bg="#0068ff", fg="white", relief="flat",
                                 font=("Segoe UI", 10, "bold"), activebackground="#0052cc", command=self.send_now)
        self.btn_now.pack(side="left", fill="x", expand=True)
        ttk.Button(bt, text="👁 Xem trước…", command=self.preview).pack(side="left", padx=(4, 0))
        ib = ttk.Frame(now)
        ib.pack(fill="x", pady=(4, 0))
        ttk.Button(ib, text="🎞 Ảnh màng film…", command=self.film).pack(side="left", fill="x", expand=True)
        ttk.Button(ib, text="⚪ Ảnh hạt Filler…", command=lambda: self.film("filler")).pack(
            side="left", fill="x", expand=True, padx=(4, 0))
        self.v_del = tk.BooleanVar(value=_flag(self.sget, "zalo_img_autodel", False))
        ttk.Checkbutton(now, text="Tự xoá ảnh màng film + hạt Filler sau khi gửi", variable=self.v_del,
                        command=lambda: self.sset("zalo_img_autodel", "1" if self.v_del.get() else "0")
                        ).pack(anchor="w", pady=(2, 0))
        self.v_ask = tk.BooleanVar(value=_flag(self.sget, "zalo_auto_ask", False))
        ttk.Checkbutton(now, text="Mở cửa sổ gửi sau khi lưu mẫu mới", variable=self.v_ask,
                        command=lambda: self.sset("zalo_auto_ask", "1" if self.v_ask.get() else "0")
                        ).pack(anchor="w", pady=(4, 0))

        nb = ttk.Notebook(b)
        nb.grid(row=0, column=2, sticky="nsew")
        jf = ttk.Frame(nb, padding=4)
        nb.add(jf, text=" ⏰ Hẹn giờ gửi ")
        cols = (("on", "Bật", 40), ("when", "Thời gian", 150), ("content", "Nội dung", 190),
                ("imgs", "Ảnh mốc", 120), ("groups", "Nhóm", 150), ("next", "Lần tới", 110),
                ("last", "Lần gửi cuối", 110))
        tf = ttk.Frame(jf)
        tf.pack(fill="both", expand=True)
        self.jt = ttk.Treeview(tf, columns=[c for c, _, _ in cols], show="headings", height=6, selectmode="browse")
        for c, h, w in cols:
            self.jt.heading(c, text=h)
            self.jt.column(c, width=w, anchor="center" if c in ("on", "next", "last", "imgs") else "w",
                           stretch=c == "groups")
        vs = ttk.Scrollbar(tf, orient="vertical", command=self.jt.yview)
        self.jt.configure(yscrollcommand=vs.set)
        self.jt.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.jt.tag_configure("off", foreground="#999")
        self.jt.tag_configure("miss", background="#fff1e0")
        self.jt.bind("<Double-1>", lambda e: self.edit_job())
        jb = ttk.Frame(jf)
        jb.pack(fill="x", pady=(4, 0))
        ttk.Button(jb, text="➕ Thêm lịch", command=self.add_job).pack(side="left")
        ttk.Button(jb, text="✎ Sửa", command=self.edit_job).pack(side="left", padx=3)
        ttk.Button(jb, text="⏯ Bật/Tắt", command=self.toggle_job).pack(side="left")
        ttk.Button(jb, text="🗑 Xoá", command=self.del_job).pack(side="left", padx=3)
        ttk.Button(jb, text="▶ Chạy thử ngay", command=self.run_selected_job).pack(side="left")
        tk.Button(jb, text="🖼 Ảnh theo mốc / DC…", bg="#2e7d32", fg="white", relief="flat",
                  activebackground="#43a047", font=("Segoe UI", 9, "bold"), padx=8,
                  command=self.slot_images).pack(side="left", padx=(6, 0))
        lf = ttk.Frame(nb, padding=4)
        nb.add(lf, text=" 📜 Nhật ký gửi ")
        if self._admin():             # chi tai khoan admin: xoa nhat ky gui
            lbar = ttk.Frame(lf)
            lbar.pack(side="bottom", fill="x", pady=(4, 0))
            ttk.Button(lbar, text="🗑 Xoá dòng đã chọn", command=self.del_log_sel).pack(side="left")
            ttk.Button(lbar, text="🧹 Xoá toàn bộ nhật ký", command=self.clear_log).pack(side="left", padx=4)
            ttk.Label(lbar, text="(chỉ admin · Ctrl/Shift + bấm để chọn nhiều dòng)",
                      foreground="#888").pack(side="left", padx=6)
        lw = ttk.Frame(lf)
        lw.pack(fill="both", expand=True)
        self.log_lb = tk.Listbox(lw, height=7, font=("Consolas", 9),
                                 selectmode="extended" if self._admin() else "browse")
        ls = ttk.Scrollbar(lw, orient="vertical", command=self.log_lb.yview)
        self.log_lb.configure(yscrollcommand=ls.set)
        self.log_lb.pack(side="left", fill="both", expand=True)
        ls.pack(side="right", fill="y")
        if self._admin():
            self.log_lb.bind("<Delete>", lambda e: self.del_log_sel())
        for x in self._log_rows():
            self.log_lb.insert("end", x)
        # --- NHAC NHAP BAO CAO (rieng tung tai khoan) - nam ngay duoi phan gui Zalo
        self.remind = None
        try:
            self.remind = RemindBox(b, self)
            self.remind.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        except Exception as ex:
            print("RemindBox:", ex)
        self.status = ttk.Label(b, text="", foreground="#1f3a5f")
        self.status.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self._pick_lbl()
        if not IS_WIN:
            self._st("Máy này không phải Windows – không tự gửi Zalo được (chỉ xem trước / chép nội dung).")

    def toggle(self):
        self.collapsed = not self.collapsed
        self.sset("zalo_panel_collapsed", "1" if self.collapsed else "0")
        self._apply_collapse()

    def _apply_collapse(self):
        if self.collapsed:
            self.body.pack_forget()
            self.btn_tg.configure(text="▲ Mở rộng")
        else:
            self.body.pack(fill="both", expand=True)
            self.btn_tg.configure(text="▼ Thu gọn")
        if self.on_toggle:
            self.after_idle(self.on_toggle)

    def _st(self, s, err=False):
        try:
            self.status.configure(text=s, foreground="#b00020" if err else "#1f3a5f")
        except tk.TclError:
            pass

    # ---------------- nhat ky
    def _admin(self):
        a = self.is_admin
        try:
            return bool(a() if callable(a) else a)
        except Exception:
            return False

    def _reload_log(self):
        try:
            self.log_lb.delete(0, "end")
            for x in self._log_rows():
                self.log_lb.insert("end", x)
        except tk.TclError:
            pass

    def del_log_sel(self):
        """Admin: xoa cac dong nhat ky dang chon."""
        if not self._admin():
            return
        idx = list(self.log_lb.curselection())
        if not idx:
            self._st("Chọn dòng nhật ký cần xoá trước (Ctrl/Shift + bấm để chọn nhiều dòng).", True)
            return
        lines = [self.log_lb.get(i) for i in idx]
        if not messagebox.askyesno("Xoá nhật ký gửi", "Xoá %d dòng nhật ký đã chọn?" % len(lines), parent=self):
            return
        rows = self._log_rows()
        for ln in lines:                      # xoa theo noi dung (an toan khi may khac vua ghi them)
            if ln in rows:
                rows.remove(ln)
        self.sset("zalo_log", json.dumps(rows, ensure_ascii=False))
        self._reload_log()
        self._st("🗑 Đã xoá %d dòng nhật ký gửi." % len(lines))

    def clear_log(self):
        """Admin: xoa toan bo nhat ky gui."""
        if not self._admin():
            return
        n = len(self._log_rows())
        if not n:
            self._st("Nhật ký gửi đang trống.")
            return
        if not messagebox.askyesno("Xoá toàn bộ nhật ký gửi",
                                   "Xoá TOÀN BỘ %d dòng nhật ký gửi Zalo?\nKhông khôi phục lại được." % n,
                                   icon="warning", parent=self):
            return
        self.sset("zalo_log", "[]")
        self._reload_log()
        self._st("🧹 Đã xoá toàn bộ nhật ký gửi.")

    def _log_rows(self):
        try:
            return json.loads(self.sget("zalo_log", "[]") or "[]")
        except Exception:
            return []

    def log(self, msg):
        line = "%s  %s" % (dt.datetime.now().strftime("%d/%m %H:%M:%S"), msg)
        rows = ([line] + self._log_rows())[:150]
        self.sset("zalo_log", json.dumps(rows, ensure_ascii=False))
        try:
            self.log_lb.insert(0, line)
            if self.log_lb.size() > 150:
                self.log_lb.delete(150, "end")
        except tk.TclError:
            pass

    # ---------------- chon DC / moc + anh mang film
    def _pick_lbl(self):
        infos = [i for i in (self.p.info_of(x) for x in self.pick_sids) if i]
        t = ("%d mẫu%s: %s" % (len(infos), " (gộp)" if self.pick_sum else "", pick_desc(infos))) if infos \
            else "(chưa chọn mẫu)"
        try:
            self.lbl_pick.configure(text=t)
        except (tk.TclError, AttributeError):
            pass

    def pick(self):
        if not hasattr(self.p, "samples_on"):
            messagebox.showinfo("Zalo", "Phiên bản này chưa hỗ trợ chọn mẫu.", parent=self)
            return

        def done(sids, summ):
            self.pick_sids, self.pick_sum = sids, summ
            self.sset("zalo_pick_sids", json.dumps(sids))
            self.sset("zalo_pick_summary", "1" if summ else "0")
            self.v_content.set("pick")
            self._pick_lbl()

        SamplePicker(self.winfo_toplevel(), self.p, self.pick_sids, self.pick_sum, done, sget=self.sget)

    def film(self, kind="film"):
        x = self.p.selected()
        if not x:
            messagebox.showinfo("Ảnh " + IMG_KINDS[kind][2], "Chọn 1 mẫu ở bảng mẫu phía trên trước.", parent=self)
            return
        sid, info = x
        FilmDialog(self.winfo_toplevel(), sid, "DC%s %s %s" % (info.get("line_no", ""), info.get("time", ""),
                                                                info.get("order", "")), kind=kind,
                   dup_check=make_dup_check(self.p, self.sget))

    @staticmethod
    def _img_sids(items, card=True):
        """Ma mau cua cac phieu TUNG MAU (co ve anh mang film / hat filler) trong lan gui."""
        if not card:
            return []
        out = []
        for it in items:
            if it[3] is None or isinstance(it[0], tuple):
                continue
            out.append(it[0])
            if it[3].get("_slot_key"):              # anh rieng cua moc (lich hen gio)
                out.append(it[3]["_slot_key"])
        return out

    # ---------------- noi dung
    def _cards(self, pairs, job=None):
        """job = lich 'Mau moi nhat tung day chuyen' -> ve anh RIENG cua moc do theo tung day chuyen."""
        logo = self.p.logo()
        out = []
        for s, i in pairs:
            if job is not None and job.get("id") and job_uses_slot_imgs(job):
                i = slot_info(i, job)
            out.append((s, build_text(i), render_card(i, logo), i))
        return out

    def _pool(self, now):
        """Tat ca mau cua ngay/ca hien tai (moi day chuyen) -> [(sid, info)]"""
        if hasattr(self.p, "samples_on") and hasattr(self.p, "context"):
            d, ca = self.p.context(now)
            return list(self.p.samples_on(d, ca))
        return list(self.p.summary(now)[2])

    def _items(self, content, now=False, sid=None, job=None):
        """-> (tieu de mo ta, [(sid|tuple sid, text, card, info|None)])"""
        job = job or {}
        lines = [str(x) for x in (job.get("lines") or [])]
        logo = self.p.logo()
        if content == "selected":
            if sid is not None:
                info = self.p.info_of(sid)
                if not info:
                    return "mẫu #%s (đã xoá?)" % sid, []
                items = [(sid, info)]
            else:
                x = self.p.selected()
                if not x:
                    return "mẫu đang chọn", []
                items = [x]
            return "mẫu DC%s %s" % (items[0][1].get("line_no", ""), items[0][1].get("time", "")), self._cards(items)
        if content == "latest":
            items = latest_per_line(self._pool(now), lines) if lines else self.p.latest(now)
            return "mẫu mới nhất %d dây chuyền" % len(items), self._cards(items, job if job else None)
        if content == "pick":
            sids = job.get("sids") if job else self.pick_sids
            summ = job.get("pick_summary") if job else self.pick_sum
            items = [(x, i) for x, i in ((x, self.p.info_of(x)) for x in (sids or [])) if i]
            if not items:
                return "mẫu đã chọn", []
            desc = "%d mẫu chọn (%s)" % (len(items), pick_desc([i for _, i in items]))
            if not summ:
                return desc, self._cards(items)
            infos = [i for _, i in items]
            dates = sorted({i.get("date", "") for i in infos if i.get("date")})
            sub = "%s  ·  %d mẫu  ·  %s" % (", ".join(dates), len(infos), pick_desc(infos))
            title = "TỔNG HỢP QC THÀNH PHẨM"
            return desc, [(tuple(x for x, _ in items), build_summary_text(title + " – " + sub, infos),
                           render_summary_card(title, sub, infos, logo), None)]
        title, sub, items = self.p.summary(now)
        if lines:
            items = [(x, i) for x, i in (self._pool(now) if hasattr(self.p, "samples_on") else items)
                     if i.get("line_no", "") in lines]
            sub = sub + "  ·  DC " + ", ".join(lines)
        if not items:
            return title, []
        infos = [i for _, i in items]
        return title, [(tuple(s for s, _ in items), build_summary_text(title + " – " + sub, infos),
                        render_summary_card(title, sub, infos, logo), None)]

    def _msgs(self, items, text=True, card=True):
        return [(it[1] if text else "", [it[2]] if card else []) for it in items]

    def _save_opts(self):
        self.sset("zalo_send_text", "1" if self.v_text.get() else "0")
        self.sset("zalo_send_card", "1" if self.v_card.get() else "0")
        self.sset("zalo_auto_send", "1" if self.v_send.get() else "0")
        self.sset("zalo_open_key", "down_enter" if self.v_down.get() else "enter")
        self.sset("zalo_now_content", self.v_content.get())

    def preview(self):
        self._save_opts()
        desc, items = self._items(self.v_content.get())
        if not items:
            messagebox.showinfo("Zalo", "Không có mẫu nào để gửi (%s)." % desc, parent=self)
            return
        info = items[0][3] if len(items) == 1 else None
        ZaloSendDialog(self.winfo_toplevel(), info, self.sget, self.sset, self.base_dir, self.p.logo(),
                       messages=[(it[1], it[2]) for it in items], title="📤 Gửi Zalo – " + desc,
                       infos=[it[3] for it in items],
                       on_sent=lambda done, err: self.log(("✔ Gửi %s → %s" % (desc, ", ".join(done)))
                                                          if not err else "⚠ Lỗi gửi %s: %s" % (desc, err)))

    def send_now(self):
        self._save_opts()
        groups = self.gb.selected()
        if not groups:
            messagebox.showinfo("Zalo", "Chọn (hoặc thêm) ít nhất 1 nhóm Zalo ở cột bên trái.", parent=self)
            return
        self.gb.remember()
        desc, items = self._items(self.v_content.get())
        if not items:
            messagebox.showinfo("Zalo", "Không có mẫu nào để gửi (%s)." % desc, parent=self)
            return
        msgs = [m for m in self._msgs(items, self.v_text.get(), self.v_card.get()) if m[0] or m[1]]
        if not msgs:
            messagebox.showinfo("Zalo", "Chưa chọn gửi Chữ hoặc Ảnh phiếu.", parent=self)
            return
        self._send(groups, msgs, self.v_send.get(), desc, img_sids=self._img_sids(items, self.v_card.get()))

    def _send(self, groups, msgs, auto_send, desc, after=None, img_sids=None):
        """Gui trong thread. after(ok) goi khi xong. -> True neu da bat dau gui."""
        if not IS_WIN:
            self._st("Không tự gửi được trên máy này.", True)
            if after:
                after(False)
            return False
        if sender_busy() or self._running:
            self._st("Đang có lượt gửi khác – đợi xong.", True)
            if after:
                after(False)
            return False
        self._running = True
        self.btn_now.configure(state="disabled")
        open_key = "down_enter" if self.v_down.get() else "enter"
        for _, ims in msgs:
            for im in ims:
                try:
                    save_card_file(im, self.base_dir)
                except Exception:
                    pass

        def prog(s):
            self.after(0, lambda: self._st(s + "  (không chạm chuột/bàn phím)"))

        def work():
            done, err = [], None
            try:
                done = send_messages(groups, msgs, auto_send, open_key, prog)
            except Exception as ex:
                err = ex
            self.after(0, lambda: self._done(done, err, desc, after, img_sids))

        threading.Thread(target=work, daemon=True).start()
        return True

    def _done(self, done, err, desc, after, img_sids=None):
        self._running = False
        try:
            self.btn_now.configure(state="normal")
        except tk.TclError:
            return
        if err:
            self._st("⚠ %s" % err, True)
            self.log("⚠ Lỗi gửi %s: %s%s" % (desc, err, (" (đã gửi: %s)" % ", ".join(done)) if done else ""))
        else:
            extra = ""
            if img_sids and _flag(self.sget, "zalo_img_autodel", False):
                n = delete_sample_images(img_sids)
                if n:
                    extra = "  ·  đã xoá %d ảnh màng film/hạt Filler" % n
            self._st("✔ Đã gửi %s → %s%s" % (desc, ", ".join(done), extra))
            self.log("✔ Gửi %s → %s%s" % (desc, ", ".join(done), extra))
        if after:
            after(not err)

    # ---------------- lich hen
    def jobs(self):
        try:
            return json.loads(self.sget("zalo_jobs", "[]") or "[]")
        except Exception:
            return []

    def save_jobs(self, jobs):
        self.sset("zalo_jobs", json.dumps(jobs, ensure_ascii=False))
        self.refresh_jobs()

    def _cdesc(self, j):
        c = _lbl(CONTENTS, j.get("content"))
        if j.get("content") == "pick":
            infos = [i for i in (self.p.info_of(x) for x in (j.get("sids") or [])) if i]
            c = "Chọn %d mẫu: %s" % (len(infos), pick_desc(infos))
        elif j.get("lines"):
            c += " – DC " + ",".join(j["lines"])
        return c

    def _when(self, j):
        rep = j.get("repeat")
        if rep == "once":
            return "Một lần %s %s" % (JobDialog._vn(j.get("date")), j.get("time"))
        if rep == "daily":
            return "Hằng ngày %s" % j.get("time")
        return "Mỗi %s phút (từ %s)" % (j.get("every"), j.get("time"))

    def refresh_jobs(self):
        sel = self.jt.selection()
        self.jt.delete(*self.jt.get_children())
        now = dt.datetime.now()
        nxt = []
        for j in self.jobs():
            n = job_next(j, now) if j.get("on", True) else None
            if n:
                nxt.append(n)
            last = dt.datetime.fromisoformat(j["last"]).strftime("%d/%m %H:%M") if j.get("last") else ""
            imgs, miss = self._img_state(j)
            tags = () if j.get("on", True) else ("off",)
            if miss and j.get("on", True):
                tags += ("miss",)
            self.jt.insert("", "end", iid=j["id"], tags=tags, values=(
                "✔" if j.get("on", True) else "—", self._when(j), self._cdesc(j), imgs,
                ", ".join(j.get("groups") or []), n.strftime("%d/%m %H:%M") if n else "", last))
        for s in sel:
            if self.jt.exists(s):
                self.jt.selection_set(s)
        self.lbl_next.configure(text=("Lịch gửi tới: " + min(nxt).strftime("%H:%M %d/%m")) if nxt else "")

    def _img_state(self, j):
        """Cot 'Anh moc' tren bang lich -> (chu hien thi, con thieu anh?)"""
        c = j.get("content")
        if c == "summary" or (c == "pick" and j.get("pick_summary")):
            return "— (phiếu tổng hợp)", False
        if c in ("pick", "selected"):
            sids = j.get("sids") if c == "pick" else ([j["sid"]] if j.get("sid") is not None else [])
            sids = sids or []
            ok = sum(1 for x in sids if film_files(x, "film") and film_files(x, "filler"))
            return ("ảnh theo mẫu %d/%d" % (ok, len(sids))) if sids else "ảnh theo mẫu", ok < len(sids)
        lines = job_lines(j, self.p)
        miss = slot_missing(j, lines)
        return "🖼 đủ %d/%d DC" % (len(lines) - len(miss), len(lines)), bool(miss)

    def slot_images(self, job_id=None):
        """Mo bang up anh theo MOC GUI x DAY CHUYEN."""
        if job_id is None:
            j = self._sel_job()
            job_id = j["id"] if j and job_uses_slot_imgs(j) else None
        SlotImagesDialog(self.winfo_toplevel(), self, job_id)

    def _sel_job(self):
        s = self.jt.selection()
        return next((j for j in self.jobs() if j["id"] == s[0]), None) if s else None

    def _upsert(self, job):
        new = job if isinstance(job, list) else [job]
        ids = {j["id"] for j in new}
        old_ids = {j["id"] for j in self.jobs()}
        js = [j for j in self.jobs() if j["id"] not in ids]
        js += new
        js.sort(key=lambda j: (_hm(j.get("time")) or (99, 0)))
        self.save_jobs(js)
        if len(new) > 1:
            self.log("⏰ Lưu %d lịch (%s): %s – %s" % (len(new), _lbl(REPEATS, new[0].get("repeat")),
                                                   ", ".join(j["time"] for j in new),
                                                   _lbl(CONTENTS, new[0]["content"])))
        else:
            self.log("⏰ Lưu lịch: %s – %s" % (self._when(new[0]), _lbl(CONTENTS, new[0]["content"])))
        # lich MOI 'Mau moi nhat tung day chuyen' + anh rieng -> mo luon bang up anh theo moc / day chuyen
        fresh = [j for j in new if j["id"] not in old_ids and job_uses_slot_imgs(j) and j.get("slot_img_only")]
        if fresh:
            self.after(150, lambda: SlotImagesDialog(self.winfo_toplevel(), self,
                                                     fresh[0]["id"] if len(fresh) == 1 else None))

    def add_job(self):
        x = self.p.selected()
        JobDialog(self.winfo_toplevel(), self.sget, self.sset, None, self._upsert, x[0] if x else None, self.p,
                  open_imgs=self.slot_images)

    def edit_job(self):
        j = self._sel_job()
        if j:
            x = self.p.selected() if j.get("content") == "selected" and not j.get("sid") else None
            JobDialog(self.winfo_toplevel(), self.sget, self.sset, j, self._upsert, x[0] if x else None, self.p,
                      open_imgs=self.slot_images)

    def toggle_job(self):
        j = self._sel_job()
        if j:
            j["on"] = not j.get("on", True)
            self._upsert(j)

    def del_job(self):
        j = self._sel_job()
        if j and messagebox.askyesno("Xoá lịch", "Xoá lịch: %s ?\n(ảnh riêng của mốc này cũng bị xoá)"
                                     % self._when(j), parent=self):
            self.save_jobs([x for x in self.jobs() if x["id"] != j["id"]])
            n = delete_slot_images(j["id"])
            if n:
                self._st("Đã xoá lịch + %d ảnh của mốc." % n)

    def run_selected_job(self):
        j = self._sel_job()
        if j:
            self._run_job(j, forced=True)

    def _tick(self):
        """Kiem tra lich moi TICK_MS. Moi lich den gio duoc chay LAN LUOT (lich nay xong moi toi lich sau);
        lich chua gui duoc (dang co luot gui khac / loi Zalo) se cho va thu lai, khong bi bo qua."""
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            now = dt.datetime.now()
            if not (self._toast or self._running or self._job_busy or sender_busy()):
                js = self.jobs()
                changed = False
                for j in js:        # lich 'mot lan' da qua han lau (may tat) -> tat di
                    if j.get("repeat") == "once" and j.get("on", True) and not j.get("last"):
                        t = job_next(j, now)
                        if t and now - t > dt.timedelta(minutes=GRACE_MIN):
                            j["on"] = False
                            changed = True
                            self.log("⏭ Bỏ lỡ lịch %s (phần mềm tắt lúc đó)" % self._when(j))
                if changed:
                    self.save_jobs(js)
                due = []
                for j in js:
                    if not job_due(j, now):
                        self._fails.pop(j["id"], None)     # het gio cho -> xoa dem loi cu
                        continue
                    f = self._fails.get(j["id"])
                    if f and now < f[1]:
                        continue            # dang cho thu lai
                    due.append((job_next(j, now) or now, j))
                if due:
                    due.sort(key=lambda x: x[0])     # lich den gio som nhat chay truoc
                    self._run_job(due[0][1])
                self._remind_images(js, now)
            if self.remind is not None:      # nhac nhap bao cao: chi gui khi khong co luot gui / lich nao khac
                free = not (self._toast or self._running or self._job_busy or sender_busy())
                self.remind.tick(now, can_send=free)
            self.refresh_jobs()
        except Exception as ex:
            self._st("Lỗi lịch hẹn: %s" % ex, True)
        self.after(self.TICK_MS, self._tick)

    REMIND_MIN = 10         # truoc gio gui 10 phut: nhac moc con thieu anh

    def _remind_images(self, js, now):
        """Moc sap gui (<= 10 phut) ma day chuyen dang chay con thieu anh rieng -> nhac 1 lan / moc."""
        warned = getattr(self, "_warned", None)
        if warned is None:
            warned = self._warned = set()
        pool_lines = None
        for j in js:
            if not (j.get("on", True) and job_uses_slot_imgs(j) and j.get("slot_img_only")):
                continue
            t = job_next(j, now)
            if t is None or not (dt.timedelta(0) <= t - now <= dt.timedelta(minutes=self.REMIND_MIN)):
                continue
            k = (j["id"], t.isoformat())
            if k in warned:
                continue
            warned.add(k)
            if j.get("lines"):
                lines = [str(x) for x in j["lines"]]
            else:
                if pool_lines is None:
                    try:
                        pool_lines = sorted({i.get("line_no", "") for _, i in self._pool(True)} - {""},
                                            key=lambda x: (len(x), x))
                    except Exception:
                        pool_lines = []
                lines = pool_lines
            miss = slot_missing(j, lines)
            if miss:
                txt = "; ".join("DC%s thiếu %s" % m for m in miss)
                self.log("🖼 Mốc %s (%s) sắp gửi – %s" % (j.get("time"), t.strftime("%H:%M"), txt))
                self._st("🖼 Mốc %s sắp gửi còn thiếu ảnh: %s – bấm '🖼 Ảnh theo mốc / DC…'" % (j.get("time"), txt),
                         True)

    def _mark(self, job, sids=None):
        """Ghi 'Lan gui cuoi' (lich nay xong cho lan toi) + danh sach mau da gui."""
        js = self.jobs()
        for j in js:
            if j["id"] == job["id"]:
                j["last"] = dt.datetime.now().isoformat(timespec="seconds")
                if sids:
                    j["sent"] = (list(sids) + [x for x in (j.get("sent") or []) if x not in sids])[:300]
        self._fails.pop(job["id"], None)
        self.save_jobs(js)

    def _failed(self, job, why):
        n = self._fails.get(job["id"], (0, None))[0] + 1
        if n >= self.MAX_TRY:
            self._mark(job)
            self.log("⚠ Lịch %s: gửi lỗi %d lần – bỏ qua lần này (%s)" % (self._when(job), n, why))
        else:
            self._fails[job["id"]] = (n, dt.datetime.now() + dt.timedelta(seconds=self.RETRY_SEC))
            self.log("⚠ Lịch %s: chưa gửi được (%s) – thử lại sau %d giây (lần %d/%d)"
                     % (self._when(job), why, self.RETRY_SEC, n + 1, self.MAX_TRY))

    def _run_job(self, job, forced=False):
        """forced = 'Chay thu ngay': gui ngay, KHONG ghi 'lan gui cuoi' / 'mau da gui' -> lich that van chay."""
        try:
            desc, items = self._items(job.get("content"), now=True, sid=job.get("sid"), job=job)
        except Exception as ex:
            if forced:
                self._st("⚠ Lỗi tạo nội dung: %s" % ex, True)
            else:
                self._failed(job, "lỗi tạo nội dung: %s" % ex)
            return
        sent = set(job.get("sent") or [])
        # 'Chi gui mau moi' CHI ap dung cho lich 'Lap moi … phut'. Lich 'Hang ngay' / 'Mot lan' LUON gui
        # dung gio (truoc day lich thu 2 hay bi bo qua vi mau da duoc lich khac / 'Chay thu' gui roi).
        if job.get("only_new", True) and job.get("repeat") == "interval" and not forced:
            def new(k):
                ks = k if isinstance(k, tuple) else (k,)
                return any(x not in sent for x in ks)
            items = [it for it in items if new(it[0])]
        if not items:
            if not forced:
                self._mark(job)
            self.log("⏰ %s: không có mẫu %sđể gửi – bỏ qua" % (self._when(job),
                                                              "mới " if job.get("repeat") == "interval" else ""))
            return
        sids = []
        for it in items:
            k = it[0]
            sids += list(k) if isinstance(k, tuple) else [k]
        card = _flag(self.sget, "zalo_send_card", True)
        msgs = self._msgs(items, _flag(self.sget, "zalo_send_text", True), card)
        msgs = [m for m in msgs if m[0] or m[1]]
        groups = job.get("groups") or []
        if not groups or not msgs:
            if not forced:
                self._mark(job)
            self.log("⏰ %s: chưa chọn nhóm / nội dung – bỏ qua" % self._when(job))
            return
        img_sids = self._img_sids(items, card)
        if card and job_uses_slot_imgs(job):
            miss = []
            for it in items:
                inf = it[3] or {}
                if inf.get("_slot_key"):
                    m = [IMG_KINDS[k][2] for k in IMG_KINDS if not film_files(inf["_slot_key"], k)]
                    if m:
                        miss.append("DC%s thiếu %s" % (inf.get("line_no", ""), " + ".join(m)))
            if miss:
                self.log("🖼 Mốc %s: %s%s" % (job.get("time"), "; ".join(miss),
                                               "" if job.get("slot_img_only") else " (dùng ảnh của mẫu)"))
        self._job_busy = job["id"]

        def after(ok):
            self._job_busy = None
            if forced:
                return
            if ok:
                self._mark(job, sids)
            else:
                self._failed(job, "Zalo lỗi / đang bận")

        def go():
            self._toast = None
            self._send(groups, msgs, job.get("auto_send", True), ("[chạy thử] " if forced else "[lịch] ") + desc,
                       after=after, img_sids=img_sids)

        def cancel():
            self._toast = None
            self._job_busy = None
            if not forced:
                self._mark(job)
            self.log("✖ Huỷ lượt gửi theo lịch %s" % self._when(job))

        self._toast = CountdownToast(self.winfo_toplevel(), "Sắp gửi Zalo: %s → %s" % (desc, ", ".join(groups)),
                                     10, go, cancel)


# ============================================================================
# 6) NHAC NHAP BAO CAO (tu gui tin nhac khi qua lau chua nhap mau thanh pham moi)
#    - Lan nhap cuoi = gio lay mau (cot 'Gio') moi nhat cua tung day chuyen.
#    - Sau [first] phut ma chua co mau moi -> gui nhac lan 1; sau do cu [repeat] phut gui lai,
#      toi da [max] lan (0 = khong gioi han). Co mau moi -> tu dat lai tu dau.
#    - Cai dat + trang thai LUU RIENG THEO TAI KHOAN DANG NHAP (khoa 'zalo_remind::<tai khoan>').
#    - Gui toi: My Documents (Cloud cua toi) / 1 nguoi / 1 nhom - tuy chon.
# ============================================================================
REMIND_TARGETS = [("mydoc", "My Documents"), ("person", "Người"), ("group", "Nhóm")]
MYDOC_NAMES = ("My Documents", "Cloud của tôi")
REMIND_HEAD = ("🔔 NHẮC NHẬP BÁO CÁO QC THÀNH PHẨM ({bay_gio})\n"
               "{ds}\n"
               "→ Vui lòng vào phần mềm nhập kết quả mẫu mới.")
REMIND_ITEM = "• DC{dc}: nhập cuối lúc {gio} ngày {ngay} – đã {tre} chưa nhập mẫu mới (nhắc lần {lan})"
REMIND_DEFAULT = {"on": False, "auto": True, "first": 60, "repeat": 30, "max": 4, "stale_h": 8, "lines": [],
                  "target": "mydoc", "mydoc": "My Documents", "person": "", "group": "",
                  "head": REMIND_HEAD, "item": REMIND_ITEM}


class _SafeFmt(dict):
    def __missing__(self, k):
        return "{" + k + "}"


def fmt_minutes(m):
    m = max(0, int(m))
    h, mm = divmod(m, 60)
    if h and mm:
        return "%d giờ %d phút" % (h, mm)
    return ("%d giờ" % h) if h else ("%d phút" % mm)


def _iso(t):
    return t.isoformat(timespec="seconds") if t else None


def _from_iso(s):
    try:
        return dt.datetime.fromisoformat(s) if s else None
    except Exception:
        return None


def remind_plan(cfg, state, reports, now):
    """Tinh trang thai nhac cho tung day chuyen.
    reports: {line_no: (datetime nhap cuoi, info dict)}; state: {line_no: {last, n, at}} (bi sua tai cho
    khi co mau moi -> dat lai). -> (due: [line_no], rows: [(line_no, chu trang thai)], next_dt|None)"""
    first = max(1, int(cfg.get("first") or 60))
    rep = max(1, int(cfg.get("repeat") or 30))
    mx = max(0, int(cfg.get("max") or 0))
    stale = max(0.0, float(cfg.get("stale_h") or 0))
    want = [str(x) for x in (cfg.get("lines") or [])]
    due, rows, nxt = [], [], []
    for ln in sorted(reports, key=lambda x: (len(x), x)):
        if want and ln not in want:
            continue
        last, _info = reports[ln]
        st = state.setdefault(ln, {})
        if st.get("last") != _iso(last):          # co mau moi -> dat lai dem nhac
            st.clear()
            st["last"] = _iso(last)
        n = int(st.get("n") or 0)
        late = (now - last).total_seconds() / 60.0
        head = "DC%s · cuối %s" % (ln, last.strftime("%H:%M"))
        if stale and late > stale * 60:
            rows.append((ln, head + " · ngưng > %g giờ – không nhắc" % stale))
            continue
        if mx and n >= mx:
            rows.append((ln, head + " · đã nhắc đủ %d lần" % mx))
            continue
        at = _from_iso(st.get("at"))
        t = last + dt.timedelta(minutes=first) if (n == 0 or at is None) else at + dt.timedelta(minutes=rep)
        if now >= t:
            due.append(ln)
            rows.append((ln, head + " · trễ %s · ĐẾN GIỜ NHẮC" % fmt_minutes(late)))
        else:
            nxt.append(t)
            rows.append((ln, head + " · %snhắc lúc %s" % (("đã nhắc %d · " % n) if n else "",
                                                         t.strftime("%H:%M"))))
    return due, rows, (min(nxt) if nxt else None)


def remind_text(cfg, reports, state, lines, now, test=False):
    """Noi dung tin nhac cho cac day chuyen 'lines'."""
    items = []
    for ln in lines:
        last, info = reports[ln]
        info = info or {}
        n = int((state.get(ln) or {}).get("n") or 0) + (0 if test else 1)
        items.append((cfg.get("item") or REMIND_ITEM).format_map(_SafeFmt(
            dc=ln, gio=last.strftime("%H:%M"), ngay=last.strftime("%d/%m"),
            tre=fmt_minutes((now - last).total_seconds() / 60.0), lan=max(1, n),
            don=info.get("order", ""), qc=info.get("qc", ""), ma=info.get("vmi", ""), ca=info.get("ca", ""))))
    txt = (cfg.get("head") or REMIND_HEAD).format_map(_SafeFmt(
        ds="\n".join(items), bay_gio=now.strftime("%H:%M %d/%m"), so_dc=len(lines)))
    return ("[GỬI THỬ] " if test else "") + txt


class RemindAskToast(tk.Toplevel):
    """Hop hoi 'Gui nhac ngay?' (che do KHONG tu dong gui)."""

    def __init__(self, master, text, go, cancel):
        super().__init__(master)
        self.title("Nhắc nhập báo cáo")
        self.attributes("-topmost", True)
        self.resizable(False, False)
        self.geometry("+%d+%d" % (self.winfo_screenwidth() - 440, 60))
        self.go_cb, self.cancel_cb = go, cancel
        tk.Label(self, text="🔔 " + text, font=("Segoe UI", 11, "bold"), wraplength=380, justify="left",
                 padx=12, pady=8).pack(anchor="w")
        tk.Label(self, text="Bấm Gửi để gửi tin nhắc qua Zalo (đừng chạm chuột/bàn phím khi đang gửi).",
                 font=("Segoe UI", 9), fg="#555", wraplength=380, justify="left", padx=12).pack(anchor="w")
        bt = tk.Frame(self, padx=12, pady=10)
        bt.pack(fill="x")
        tk.Button(bt, text="▶ Gửi", bg="#0068ff", fg="white", relief="flat", padx=14,
                  command=self._go).pack(side="left")
        tk.Button(bt, text="✖ Bỏ qua lần này", relief="groove", padx=10, command=self._cancel).pack(side="left", padx=8)
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _go(self):
        self.destroy()
        self.go_cb()

    def _cancel(self):
        self.destroy()
        self.cancel_cb()


class RemindTemplateDialog(tk.Toplevel):
    """Sua mau tin nhac."""

    def __init__(self, master, cfg, on_ok):
        super().__init__(master)
        self.title("✎ Mẫu tin nhắc nhập báo cáo")
        self.transient(master)
        self.on_ok = on_ok
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Khung tin  ({ds} = danh sách dây chuyền, {bay_gio}, {so_dc}):").pack(anchor="w")
        self.t_head = tk.Text(f, width=80, height=5, font=("Segoe UI", 10), wrap="word")
        self.t_head.pack(fill="both", expand=True, pady=(2, 8))
        self.t_head.insert("1.0", cfg.get("head") or REMIND_HEAD)
        ttk.Label(f, text="Mỗi dây chuyền  ({dc} {gio} {ngay} {tre} {lan} {don} {qc} {ma} {ca}):").pack(anchor="w")
        self.t_item = tk.Text(f, width=80, height=3, font=("Segoe UI", 10), wrap="word")
        self.t_item.pack(fill="both", expand=True, pady=(2, 8))
        self.t_item.insert("1.0", cfg.get("item") or REMIND_ITEM)
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="↺ Mẫu mặc định", command=self._reset).pack(side="left")
        ttk.Button(b, text="Huỷ", command=self.destroy).pack(side="right")
        ttk.Button(b, text="✔ Lưu", command=self._ok).pack(side="right", padx=4)
        self.grab_set()

    def _reset(self):
        self.t_head.delete("1.0", "end")
        self.t_head.insert("1.0", REMIND_HEAD)
        self.t_item.delete("1.0", "end")
        self.t_item.insert("1.0", REMIND_ITEM)

    def _ok(self):
        self.on_ok(self.t_head.get("1.0", "end-1c").strip() or REMIND_HEAD,
                   self.t_item.get("1.0", "end-1c").strip() or REMIND_ITEM)
        self.destroy()


class RemindBox(ttk.LabelFrame):
    """Khung '🔔 Nhắc nhập báo cáo' ben duoi phan Gui Zalo. Moi tai khoan dang nhap 1 cai dat rieng.
    provider can co: last_reports() -> {line_no: (datetime, info)}; (tuy chon) user(), line_nos()."""

    CHECK_SEC = 30            # bao lau doc lai du lieu / xet nhac 1 lan
    RETRY_SEC = 90            # gui loi -> thu lai sau
    MAX_TRY = 3

    def __init__(self, master, panel):
        super().__init__(master, padding=(6, 3, 6, 4))
        self.panel = panel
        self.p, self.sget, self.sset = panel.p, panel.sget, panel.sset
        self._last_check = None
        self._reports = {}
        self._busy = False
        self._fail = (0, None)     # (so lan loi, thu lai luc)
        self.cfg = self._load_cfg()
        self.configure(text=" 🔔 Nhắc nhập báo cáo – tài khoản: %s " % self.user_label())
        self._build()
        self.after(1500, lambda: self.tick(dt.datetime.now(), force=True))

    # ---------------- tai khoan + luu tru rieng
    def user(self):
        u = ""
        f = getattr(self.p, "user", None)
        try:
            u = (f() if callable(f) else (f or "")) or ""
        except Exception:
            u = ""
        u = str(u).strip()
        if not u:
            import socket
            u = "@" + socket.gethostname()          # khong dang nhap -> rieng theo may
        return u

    def user_label(self):
        u = self.user()
        return ("máy %s (chưa đăng nhập)" % u[1:]) if u.startswith("@") else u

    def _k(self, what):
        return "zalo_remind%s::%s" % (what, self.user())

    def _load_cfg(self):
        c = dict(REMIND_DEFAULT)
        try:
            c.update(json.loads(self.sget(self._k(""), "") or "{}"))
        except Exception:
            pass
        return c

    def _save_cfg(self):
        self.sset(self._k(""), json.dumps(self.cfg, ensure_ascii=False))

    def _state(self):
        try:
            s = json.loads(self.sget(self._k("_state"), "") or "{}")
            return s if isinstance(s, dict) else {}
        except Exception:
            return {}

    def _save_state(self, s):
        self.sset(self._k("_state"), json.dumps(s, ensure_ascii=False))

    # ---------------- giao dien
    def _build(self):
        c = self.cfg
        r1 = ttk.Frame(self)
        r1.pack(fill="x")
        self.v_on = tk.BooleanVar(value=bool(c.get("on")))
        tk.Checkbutton(r1, text="Bật nhắc", variable=self.v_on, font=("Segoe UI", 9, "bold"),
                       fg="#c62828", command=self.save).pack(side="left")
        self.v_auto = tk.BooleanVar(value=bool(c.get("auto", True)))
        self.cb_auto = tk.Checkbutton(r1, text="⚡ Tự động gửi", variable=self.v_auto,
                                      font=("Segoe UI", 9, "bold"), fg="#2e7d32", command=self.save)
        self.cb_auto.pack(side="left", padx=(4, 0))

        def spin(parent, label, key, lo, hi, w=5, after=""):
            ttk.Label(parent, text=label).pack(side="left", padx=(10, 2))
            v = tk.StringVar(value=str(c.get(key)))
            sp = ttk.Spinbox(parent, from_=lo, to=hi, width=w, textvariable=v, command=self.save)
            sp.pack(side="left")
            sp.bind("<FocusOut>", lambda e: self.save())
            sp.bind("<Return>", lambda e: self.save())
            if after:
                ttk.Label(parent, text=after).pack(side="left", padx=(2, 0))
            return v

        self.v_first = spin(r1, "Nhắc lần đầu sau", "first", 1, 1440, after="phút từ lần nhập cuối")
        self.v_rep = spin(r1, "· Nhắc lại mỗi", "repeat", 1, 1440, after="phút nếu vẫn chưa nhập")
        self.v_max = spin(r1, "· Tối đa", "max", 0, 99, w=4, after="lần (0 = mãi)")
        self.v_stale = spin(r1, "· Bỏ qua DC dừng >", "stale_h", 0, 72, w=4, after="giờ")

        r2 = ttk.Frame(self)
        r2.pack(fill="x", pady=(3, 0))
        bb = ttk.Frame(r2)                 # nut ben phai: pack truoc -> luon hien du khi cua so hep
        bb.pack(side="right")
        ttk.Button(bb, text="✎ Mẫu tin…", command=self.edit_template).pack(side="left", padx=(6, 0))
        ttk.Button(bb, text="👁 Xem tin", command=self.preview).pack(side="left", padx=(4, 0))
        tk.Button(bb, text="▶ Gửi thử", bg="#0068ff", fg="white", relief="flat", activebackground="#0052cc",
                  font=("Segoe UI", 9, "bold"), padx=8, command=self.send_test).pack(side="left", padx=(4, 0))
        self.lines_mb = ttk.Menubutton(r2, text="")
        self.lines_mb.pack(side="left", padx=(0, 12))
        self.lines_menu = tk.Menu(self.lines_mb, tearoff=False, postcommand=self._fill_lines)
        self.lines_mb["menu"] = self.lines_menu
        self._lines_lbl()
        ttk.Label(r2, text="Gửi tới:").pack(side="left")
        self.v_target = tk.StringVar(value=c.get("target") or "mydoc")
        ttk.Radiobutton(r2, text="My Documents", value="mydoc", variable=self.v_target,
                        command=self.save).pack(side="left", padx=(4, 0))
        self.v_mydoc = tk.StringVar(value=c.get("mydoc") or MYDOC_NAMES[0])
        cb = ttk.Combobox(r2, textvariable=self.v_mydoc, values=MYDOC_NAMES, width=13)
        cb.pack(side="left", padx=(2, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: self._pick("mydoc"))
        cb.bind("<FocusOut>", lambda e: self.save())
        ttk.Radiobutton(r2, text="Người:", value="person", variable=self.v_target,
                        command=self.save).pack(side="left", padx=(10, 0))
        self.v_person = tk.StringVar(value=c.get("person") or "")
        e = ttk.Entry(r2, textvariable=self.v_person, width=15)
        e.pack(side="left", padx=(2, 0))
        e.bind("<FocusOut>", lambda ev: self.save())
        e.bind("<Key>", lambda ev: self.v_target.set("person"))
        ttk.Radiobutton(r2, text="Nhóm:", value="group", variable=self.v_target,
                        command=self.save).pack(side="left", padx=(10, 0))
        self.v_group = tk.StringVar(value=c.get("group") or "")
        self.cb_group = ttk.Combobox(r2, textvariable=self.v_group, width=17, postcommand=self._fill_groups)
        self.cb_group.pack(side="left", padx=(2, 0))
        self.cb_group.bind("<<ComboboxSelected>>", lambda e: self._pick("group"))
        self.cb_group.bind("<FocusOut>", lambda e: self.save())

        self.lbl_state = tk.Label(self, text="", anchor="w", justify="left", fg="#1f3a5f",
                                  font=("Segoe UI", 9))
        self.lbl_state.pack(fill="x", pady=(3, 0))
        self.bind("<Configure>", lambda e: self.lbl_state.configure(wraplength=max(300, e.width - 20)))

    def _pick(self, kind):
        self.v_target.set(kind)
        self.save()

    def _fill_groups(self):
        self.cb_group.configure(values=_groups_of(self.sget))

    def _all_lines(self):
        f = getattr(self.p, "line_nos", None)
        try:
            ls = list(f()) if callable(f) else []
        except Exception:
            ls = []
        ls = [str(x) for x in ls if str(x)] or list(self._reports)
        return sorted(set(ls), key=lambda x: (len(x), x))

    def _fill_lines(self):
        m = self.lines_menu
        m.delete(0, "end")
        want = set(self.cfg.get("lines") or [])
        m.add_command(label="✔ Tất cả dây chuyền", command=lambda: self._set_lines([]))
        m.add_separator()
        for ln in self._all_lines():
            v = tk.BooleanVar(value=ln in want)
            m.add_checkbutton(label="DC " + ln, variable=v, command=lambda ln=ln: self._toggle_line(ln))

    def _toggle_line(self, ln):
        cur = list(self.cfg.get("lines") or [])
        cur = [x for x in cur if x != ln] if ln in cur else cur + [ln]
        self._set_lines(sorted(cur, key=lambda x: (len(x), x)))

    def _set_lines(self, ls):
        self.cfg["lines"] = ls
        self._save_cfg()
        self._lines_lbl()
        self.tick(dt.datetime.now(), force=True)

    def _lines_lbl(self):
        ls = self.cfg.get("lines") or []
        self.lines_mb.configure(text="DC theo dõi: " + (", ".join(ls) if ls else "Tất cả") + " ▾")

    @staticmethod
    def _int(v, d, lo=0):
        try:
            return max(lo, int(float(str(v).strip())))
        except Exception:
            return d

    def save(self):
        c = self.cfg
        old_on = bool(c.get("on"))
        c["on"] = bool(self.v_on.get())
        old_auto = bool(c.get("auto", True))
        c["auto"] = bool(self.v_auto.get())
        c["first"] = self._int(self.v_first.get(), 60, 1)
        c["repeat"] = self._int(self.v_rep.get(), 30, 1)
        c["max"] = self._int(self.v_max.get(), 4, 0)
        c["stale_h"] = self._int(self.v_stale.get(), 8, 0)
        c["target"] = self.v_target.get()
        c["mydoc"] = self.v_mydoc.get().strip() or MYDOC_NAMES[0]
        c["person"] = self.v_person.get().strip()
        c["group"] = self.v_group.get().strip()
        self._save_cfg()
        try:
            self.cb_auto.configure(state="normal" if c["on"] else "disabled")
        except tk.TclError:
            pass
        if c["on"] and c["auto"] != old_auto:
            self.panel.log("🔔 Nhắc nhập (%s): %s" % (self.user_label(), "TỰ ĐỘNG gửi" if c["auto"]
                                                     else "hỏi trước khi gửi"))
        if c["on"] != old_on:
            self.panel.log("🔔 %s nhắc nhập báo cáo (%s) → %s" % ("Bật" if c["on"] else "Tắt",
                                                                   self.user_label(), self.target_desc()))
        self.tick(dt.datetime.now(), force=True)

    def target(self):
        c = self.cfg
        k = c.get("target")
        return (c.get("mydoc") or MYDOC_NAMES[0]) if k == "mydoc" else (c.get(k) or "").strip()

    def target_desc(self):
        t = self.target()
        return t or "(chưa nhập tên %s)" % ("người" if self.cfg.get("target") == "person" else "nhóm")

    def edit_template(self):
        def ok(head, item):
            self.cfg["head"], self.cfg["item"] = head, item
            self._save_cfg()
        RemindTemplateDialog(self.winfo_toplevel(), self.cfg, ok)

    # ---------------- du lieu
    def _load_reports(self):
        f = getattr(self.p, "last_reports", None)
        if not callable(f):
            return {}
        try:
            return f() or {}
        except Exception as ex:
            self.panel._st("Lỗi đọc lần nhập cuối: %s" % ex, True)
            return {}

    def _show(self, rows, nxt):
        if not self.cfg.get("on"):
            txt = "⏸ Đang tắt. "
        else:
            txt = "%s → %s. " % ("⚡ Tự động gửi" if self.cfg.get("auto", True) else "❓ Hỏi trước khi gửi",
                                self.target_desc())
        txt += ("   |   ".join(r for _, r in rows)) if rows else "Chưa có mẫu nào gần đây."
        if nxt and self.cfg.get("on"):
            txt += "   →  Lần nhắc tới: %s" % nxt.strftime("%H:%M %d/%m")
        try:
            self.lbl_state.configure(text=txt, fg="#1f3a5f" if self.cfg.get("on") else "#888")
        except tk.TclError:
            pass

    # ---------------- vong kiem tra (goi tu ZaloPanel._tick)
    def tick(self, now, force=False, can_send=False):
        if not force and self._last_check and (now - self._last_check).total_seconds() < self.CHECK_SEC:
            return
        self._last_check = now
        self._reports = self._load_reports()
        st = self._state()
        before = json.dumps(st, sort_keys=True)
        due, rows, nxt = remind_plan(self.cfg, st, self._reports, now)
        if json.dumps(st, sort_keys=True) != before:
            self._save_state(st)                     # co mau moi -> luu dat lai dem nhac
        self._show(rows, nxt)
        if not (can_send and due and self.cfg.get("on")) or self._busy:
            return
        if self._fail[1] and now < self._fail[1]:
            return
        tg = self.target()
        if not tg:
            self.panel._st("🔔 Nhắc nhập báo cáo: chưa nhập tên người / nhóm nhận tin.", True)
            return
        self._fire(due, st, now, tg)

    def _fire(self, due, st, now, tg, test=False):
        text = remind_text(self.cfg, self._reports, st, due, now, test=test)
        desc = "%s[nhắc nhập] DC %s" % ("[gửi thử] " if test else "", ", ".join(due))
        self._busy = True
        pnl = self.panel
        lasts = {ln: self._reports[ln][0] for ln in due}   # giu lai gio nhap cuoi luc bat dau gui

        def after(ok):
            self._busy = False
            if test:
                return
            if ok:
                self._fail = (0, None)
                s2 = self._state()                     # doc lai (may khac co the vua ghi)
                t = dt.datetime.now()
                for ln in due:
                    x = s2.setdefault(ln, {})
                    x["last"] = _iso(lasts[ln])
                    x["n"] = int(x.get("n") or 0) + 1
                    x["at"] = _iso(t)
                self._save_state(s2)
            else:
                n = self._fail[0] + 1
                if n >= self.MAX_TRY:                  # loi nhieu lan -> tinh nhu da nhac, cho lan sau
                    self._fail = (0, None)
                    s2 = self._state()
                    for ln in due:
                        x = s2.setdefault(ln, {})
                        x["n"] = int(x.get("n") or 0) + 1
                        x["at"] = _iso(dt.datetime.now())
                    self._save_state(s2)
                    pnl.log("⚠ Nhắc nhập DC %s: lỗi %d lần – bỏ qua lượt này" % (", ".join(due), n))
                else:
                    self._fail = (n, dt.datetime.now() + dt.timedelta(seconds=self.RETRY_SEC))
            self.tick(dt.datetime.now(), force=True)

        def go():
            if not pnl._send([tg], [(text, [])], True, desc, after=after):
                pass                                   # _send da goi after(False)

        def cancel():
            self._busy = False
            if not test:                               # huy -> tinh la 1 lan nhac, cho lan nhac lai sau
                s2 = self._state()
                for ln in due:
                    x = s2.setdefault(ln, {})
                    x["n"] = int(x.get("n") or 0) + 1
                    x["at"] = _iso(dt.datetime.now())
                self._save_state(s2)
                pnl.log("✖ Huỷ nhắc nhập DC %s" % ", ".join(due))

        if test:                                       # Gui thu: nguoi dung vua bam -> gui luon
            go()
            return
        if not self.cfg.get("auto", True):
            # khong tu dong: hien hop hoi, NGUOI DUNG bam Gui moi gui (khong chan lich hen gio khac)
            self._ask = RemindAskToast(self.winfo_toplevel(), "Đến giờ NHẮC NHẬP BÁO CÁO (DC %s) → %s"
                                       % (", ".join(due), tg), go, cancel)
            return
        # TU DONG GUI: gui luon, khong hoi / khong dem nguoc
        pnl.log("⚡ Tự động gửi nhắc nhập DC %s → %s" % (", ".join(due), tg))
        go()

    def _lines_for_test(self):
        want = self.cfg.get("lines") or []
        ls = [x for x in self._reports if not want or x in want]
        return sorted(ls, key=lambda x: (len(x), x))

    def preview(self):
        self.save()
        ls = self._lines_for_test()
        if not ls:
            messagebox.showinfo("Nhắc nhập báo cáo", "Chưa có mẫu thành phẩm nào gần đây để làm ví dụ.", parent=self)
            return
        txt = remind_text(self.cfg, self._reports, self._state(), ls, dt.datetime.now())
        messagebox.showinfo("Xem tin nhắc → %s" % self.target_desc(), txt, parent=self)

    def send_test(self):
        self.save()
        tg = self.target()
        if not tg:
            messagebox.showinfo("Nhắc nhập báo cáo", "Nhập tên người / chọn nhóm nhận tin trước.", parent=self)
            return
        ls = self._lines_for_test()
        if not ls:
            messagebox.showinfo("Nhắc nhập báo cáo", "Chưa có mẫu thành phẩm nào gần đây để gửi thử.", parent=self)
            return
        if self._busy or self.panel._toast or self.panel._running or sender_busy():
            messagebox.showinfo("Nhắc nhập báo cáo", "Đang có lượt gửi Zalo khác – đợi xong.", parent=self)
            return
        self._fire(ls, self._state(), dt.datetime.now(), tg, test=True)
