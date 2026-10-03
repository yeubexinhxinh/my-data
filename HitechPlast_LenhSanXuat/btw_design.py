# -*- coding: utf-8 -*-
"""
btw_design.py - TAO MAU TEM MOI (khung bang tuy y, kich thuoc tuy y) -> file BarTender .btw

  design_from_image(anh)   : nhan dien khung bang (dong ke ngang / doc) + chu (OCR) -> thiet ke
  render_design(tk, ...)   : ve xem truoc DUNG ti le mm (giong het file .btw se in)
  export_design(tk, ...)   : ghi file .btw (dung tem_mau_AA06.btw lam 'phoi': doi kho giay, doi vi tri
                             duong ke / khung / o chu, co chu, noi dung)

Thiet ke (dict):
  name, w_mm, h_mm, margin (mm), border (ke khung), split (ti le cot trai 0..1), bold, stretch (% rong chu),
  rows: [{'cells': ['Nhan', 'Gia tri'] hoac ['1 o gop'], 'h': ti le cao, 'pt': [co chu | 0 = tu dong],
          'line': True (ke duoi dong)}]
  Moi o co the nhieu dong chu (xuong dong '\\n').

Gioi han cua phoi AA06: toi da 17 dong chu, 7 duong ke + 3 khung chu nhat.
"""
import os
import re
import copy
import btw_core as B

MIL = B.MIL_PER_MM
PRESETS = [('60 x 90 mm', 60, 90), ('60 x 50 mm', 60, 50), ('100 x 120 mm', 100, 120), ('20 x 30 mm', 20, 30),
           ('57 x 50 mm (AA06)', 60, 50)]
MAX_TEXTS = 17
# thu tu dung cac o chu cua phoi (Text 17 neo giua, Text 13 can trai, Text 12 lay tu Excel -> dung sau cung)
TEXT_ORDER = ['Text 17', 'Text 11', 'Text 9', 'Text 1', 'Text 5', 'Text 6', 'Text 7', 'Text 8', 'Text 10',
              'Text 2', 'Text 3', 'Text 4', 'Text 14', 'Text 15', 'Text 16', 'Text 13', 'Text 12']
LINE_H = 1.15          # chieu cao dong = 1,15 em
PT_MIL = 13.889        # 1 pt = 13,889 mil


def new_design(w=60, h=50):
    return {'name': '', 'w_mm': w, 'h_mm': h, 'margin': 1.5, 'border': True, 'split': 0.45, 'bold': True,
            'stretch': 100, 'uniform': True, 'autowrap': True, 'logo': '', 'logo_h': 0.0,
            'rows': [{'cells': ['TÊN HÀNG'], 'h': 1.0, 'pt': [0], 'line': True},
                     {'cells': ['NET WEIGHT', '25 KGS'], 'h': 1.0, 'pt': [0, 0], 'line': True},
                     {'cells': ['BATCH NO', 'VI. 1.1'], 'h': 1.0, 'pt': [0, 0], 'line': True}]}


def normalize(dz):
    d = copy.deepcopy(dz)
    for k, v in new_design().items():
        d.setdefault(k, v)
    rows = []
    for r in d.get('rows') or []:
        cells = [str(c or '') for c in (r.get('cells') or [''])][:2] or ['']
        pt = list(r.get('pt') or []) + [0, 0]
        rows.append({'cells': cells, 'h': max(0.05, float(r.get('h') or 1.0)),
                     'pt': [float(x or 0) for x in pt[:len(cells)]], 'line': r.get('line', True)})
    d['rows'] = rows or new_design()['rows']
    d['w_mm'] = max(5.0, min(400.0, float(d['w_mm'])))
    d['h_mm'] = max(5.0, min(400.0, float(d['h_mm'])))
    d['split'] = max(0.1, min(0.9, float(d['split'])))
    return d


def count_texts(d):
    return sum(len([x for x in c.split('\n') if x.strip()]) or 0 for r in d['rows'] for c in r['cells'])


# =====================================================================
#  DO CHU (Times New Roman dam - cung phong BarTender dung)
# =====================================================================
_MEASURE = {}


def text_w_per_pt(s, bold=True):
    """Chieu rong (mil) cua chu s o co 1 pt, do rong 100%."""
    key = (s, bold)
    if key not in _MEASURE:
        from PIL import Image, ImageDraw
        f = B._font(200, bold)
        dr = ImageDraw.Draw(Image.new('L', (4, 4)))
        bb = dr.textbbox((0, 0), s, font=f)
        _MEASURE[key] = (bb[2] - bb[0]) / 200.0 * PT_MIL
    return _MEASURE[key]


# =====================================================================
#  TINH BO CUC (don vi mil, toa do trong vung mau - goc tren trai)
# =====================================================================
def compute_layout(dz):
    """-> dict: tpl_w, tpl_h, rect, hlines, vlines, texts[{s, cx, cy, w, pt, cell}]"""
    d = normalize(dz)
    W, H = d['w_mm'] * MIL, d['h_mm'] * MIL
    tpl_w = W - 100                       # le trai/phai 50 mil nhu phoi
    tpl_h = H
    m = max(20.0, d['margin'] * MIL)          # le tinh tu MEP GIAY; vung mau bat dau cach mep trai 50 mil
    x0, x1 = max(5.0, m - 50), min(tpl_w - 5, tpl_w + 50 - m)
    y0, y1 = m, tpl_h - m
    # Chua cho LOGO o dau tem (neu co) -> bang xuong duoi, giu nguyen kho tem
    logo_rect = None
    if d.get('logo'):
        band = (float(d.get('logo_h') or 0)) * MIL
        if band <= 0:
            band = min((y1 - y0) * 0.22, 26.0 * MIL)
        band = min(band, (y1 - y0) * 0.5)
        logo_rect = (x0, y0, x1, y0 + band)
        y0 = y0 + band
    tot = sum(r['h'] for r in d['rows'])
    ys = [y0]
    for r in d['rows']:
        ys.append(ys[-1] + (y1 - y0) * r['h'] / tot)
    xs = x0 + d['split'] * (x1 - x0)
    sx = d['stretch'] / 100.0
    texts = []
    cells = []
    for i, r in enumerate(d['rows']):
        ra, rb = ys[i], ys[i + 1]
        spans = [(x0, x1)] if len(r['cells']) == 1 else [(x0, xs), (xs, x1)]
        for c, (ca, cb) in enumerate(spans):
            lines = [x for x in r['cells'][c].split('\n') if x.strip()]
            cells.append((i, c, ca, cb, ra, rb, lines))
    # co chu tu dong: vua chieu cao + chieu rong o
    def fit(lines, ca, cb, ra, rb):
        n = len(lines)
        pad = min(60.0, (cb - ca) * 0.05)
        ph = (rb - ra) * 0.80 / n / (LINE_H * PT_MIL)
        pw = min(((cb - ca) - 2 * pad) / max(1e-6, text_w_per_pt(s, d['bold']) * sx) for s in lines)
        return max(2.0, min(ph, pw, 72.0))

    def split2(s):
        sp = [m.start() for m in re.finditer(' ', s)]
        if not sp:
            return None
        k = min(sp, key=lambda p: abs(text_w_per_pt(s[:p]) - text_w_per_pt(s[p + 1:])))
        return [s[:k], s[k + 1:]]

    auto = {}
    for (i, c, ca, cb, ra, rb, lines) in cells:
        if lines:
            auto[(i, c)] = fit(lines, ca, cb, ra, rb)
    if d.get('autowrap', True):
        # o 'nghen' (chu dai, o hep) -> tu xuong 2 dong neu chu to hon ro, trong gioi han 17 o chu
        budget = MAX_TEXTS - sum(len(x[6]) for x in cells)
        tried = set()
        while budget > 0:
            cand = [(auto[(x[0], x[1])], n) for n, x in enumerate(cells)
                    if len(x[6]) == 1 and (x[0], x[1]) in auto and n not in tried
                    and not (d['rows'][x[0]]['pt'][x[1]] if x[1] < len(d['rows'][x[0]]['pt']) else 0)]
            if not cand:
                break
            _, n = min(cand)
            tried.add(n)
            i, c, ca, cb, ra, rb, lines = cells[n]
            two = split2(lines[0])
            if not two:
                continue
            p2 = fit(two, ca, cb, ra, rb)
            if p2 > auto[(i, c)] * 1.15:
                cells[n] = (i, c, ca, cb, ra, rb, two)
                auto[(i, c)] = p2
                budget -= 1
    if d.get('uniform'):
        # cac o cua dong 2 cot dung chung 1 co chu (nho nhat) -> deu dep;
        # dong gop 1 o (tieu de / dong cuoi) toi da 1,35 lan co chu do
        ks = [k for k in auto if len(d['rows'][k[0]]['cells']) == 2]
        if ks:
            v = min(auto[k] for k in ks)
            for k in ks:
                auto[k] = v
            for k in auto:
                if k not in ks:
                    auto[k] = min(auto[k], v * 1.35)
    for (i, c, ca, cb, ra, rb, lines) in cells:
        if not lines:
            continue
        up = d['rows'][i]['pt'][c] if c < len(d['rows'][i]['pt']) else 0
        pt = round(up if up else auto[(i, c)] * 2) / 2.0          # buoc 0,5 pt
        pt = max(2.0, pt)
        lh = pt * PT_MIL * LINE_H
        top = (ra + rb) / 2.0 - lh * len(lines) / 2.0
        for k, s in enumerate(lines):
            cy = top + lh * (k + 0.5)
            texts.append({'s': s.strip(), 'cx': (ca + cb) / 2.0, 'cy': cy, 'w': cb - ca, 'pt': pt, 'cell': (i, c)})
    # duong ke
    hl, vl = [], []
    for i in range(len(d['rows']) - 1):
        if d['rows'][i].get('line', True):
            hl.append((x0, ys[i + 1], x1, ys[i + 1]))
    i = 0
    while i < len(d['rows']):
        if len(d['rows'][i]['cells']) == 2:
            j = i
            while j + 1 < len(d['rows']) and len(d['rows'][j + 1]['cells']) == 2:
                j += 1
            vl.append((xs, ys[i], xs, ys[j + 1]))
            i = j + 1
        else:
            i += 1
    return {'d': d, 'W': W, 'H': H, 'tpl_w': tpl_w, 'tpl_h': tpl_h, 'rect': (x0, y0, x1, y1),
            'hlines': hl if d['border'] else [], 'vlines': vl if d['border'] else [], 'texts': texts, 'ys': ys,
            'logo_rect': logo_rect}


def _text_num(n):
    try:
        return int(str(n).split()[-1])
    except Exception:
        return 99999


def usable_texts(g):
    """Danh sach TEN o chu dung duoc cua phoi: uu tien thu tu TEXT_ORDER (cho cac o dac biet
    Text 17 neo giua / Text 13 can trai / Text 12 lay tu Excel), roi them MOI o 'Text N' khac
    co trong phoi -> phoi co cang nhieu o chu thi lam duoc cang nhieu dong."""
    present = [n for n in TEXT_ORDER if n in g.texts]
    extra = sorted((n for n in g.texts if n not in TEXT_ORDER and str(n).startswith('Text ')),
                   key=_text_num)
    return present + extra


def donor_capacity(donor_path):
    """(so o chu, so duong ke HIEU DUNG) ma phoi .btw nay cung cap duoc. Loi -> (MAX_TEXTS, 11).
    So duong ke hieu dung = so duong ke + 2*(so khung chu nhat - 1): moi khung phu thay cho 2 duong
    ngang (xem export_design). Phoi mac dinh AA06: 7 duong + 3 khung -> 7 + 2*2 = 11 (giu nguyen nhu cu)."""
    try:
        btw = B.BtwFile(donor_path)
        g = B.Geometry(btw)
        eff_lines = len(g.lines) + 2 * max(0, len(g.boxes) - 1)
        return max(1, len(usable_texts(g))), max(1, eff_lines)
    except Exception:
        return MAX_TEXTS, 11


def donor_info(donor_path):
    """Mo ta phoi .btw: {'texts','lines','boxes','page','cap_lines'}. Loi -> None."""
    try:
        btw = B.BtwFile(donor_path)
        g = B.Geometry(btw)
        return {'texts': len(usable_texts(g)), 'lines': len(g.lines), 'boxes': len(g.boxes),
                'page': bool(g.page_pat), 'cap_lines': len(g.lines) + 2 * max(0, len(g.boxes) - 1)}
    except Exception:
        return None


def check_limits(dz, cap_texts=MAX_TEXTS, cap_lines=11):
    """-> danh sach loi (rong = xuat duoc). cap_texts/cap_lines = kha nang cua phoi .btw dang chon."""
    L = compute_layout(dz)
    err = []
    if len(L['texts']) > cap_texts:
        err.append('Tem có %d dòng chữ – phôi .btw đang dùng chỉ có %d ô chữ. '
                   'Gộp bớt / xoá bớt dòng, hoặc chọn phôi .btw có nhiều ô chữ hơn.'
                   % (len(L['texts']), cap_texts))
    if len(L['hlines']) + len(L['vlines']) > cap_lines:
        err.append('Quá nhiều đường kẻ (%d, phôi chỉ có %d). Bỏ kẻ dưới vài dòng '
                   'hoặc chọn phôi .btw có nhiều đường kẻ hơn.'
                   % (len(L['hlines']) + len(L['vlines']), cap_lines))
    return err


# =====================================================================
#  VE XEM TRUOC (dung ti le mm)
# =====================================================================
def render_design(dz, px_per_mm=8, frame=True, ruler=False):
    from PIL import Image, ImageDraw
    L = compute_layout(dz)
    d = L['d']
    k = px_per_mm / MIL                      # px / mil
    lab_w, lab_h = int(round(d['w_mm'] * px_per_mm)), int(round(d['h_mm'] * px_per_mm))
    off = 50 * k                              # vung mau lech trai 50 mil so voi mep giay
    img = Image.new('RGB', (lab_w, lab_h), 'white')
    dr = ImageDraw.Draw(img)
    lw = max(1, int(round(14 * k)))
    if d['border']:
        x0, y0, x1, y1 = L['rect']
        dr.rectangle([off + x0 * k, y0 * k, off + x1 * k, y1 * k], outline='black', width=lw)
    for (a, b, c, e) in L['hlines'] + L['vlines']:
        dr.line([off + a * k, b * k, off + c * k, e * k], fill='black', width=lw)
    sx = d['stretch'] / 100.0
    for t in L['texts']:
        f = B._font(max(4, int(round(t['pt'] * PT_MIL * k))), d['bold'])
        B.draw_text_scaled(img, t['s'], f, off + t['cx'] * k, t['cy'] * k, sx)
    # --- LOGO o dau tem ---
    lr = L.get('logo_rect')
    lp = d.get('logo')
    if lr and lp and os.path.exists(lp):
        try:
            from PIL import Image
            lx0, ly0, lx1, ly1 = lr
            bw = int(round((lx1 - lx0) * k))
            bh = int(round((ly1 - ly0) * k))
            lg = Image.open(lp).convert('RGBA')
            scale = min(bw / lg.width, bh / lg.height)
            lw2 = max(1, int(lg.width * scale))
            lh2 = max(1, int(lg.height * scale))
            lg = lg.resize((lw2, lh2), Image.LANCZOS)
            px = int(round(off + lx0 * k)) + (bw - lw2) // 2
            py = int(round(ly0 * k)) + (bh - lh2) // 2
            img.paste(lg, (px, py), lg)
        except Exception:
            pass
    if frame:
        pad = 14
        out = Image.new('RGB', (lab_w + 2 * pad, lab_h + 2 * pad), (150, 150, 150))
        try:
            ImageDraw.Draw(out).rounded_rectangle([pad - 2, pad - 2, pad + lab_w + 1, pad + lab_h + 1],
                                                  radius=8, fill='white')
        except AttributeError:
            pass
        out.paste(img, (pad, pad))
        img = out
    return img


# =====================================================================
#  NHAN DIEN KHUNG TU ANH
# =====================================================================
def _runs(vals):
    """list bool -> [(bat dau, ket thuc)] cac doan True lien tiep"""
    out, st = [], None
    for i, v in enumerate(vals):
        if v and st is None:
            st = i
        elif not v and st is not None:
            out.append((st, i - 1))
            st = None
    if st is not None:
        out.append((st, len(vals) - 1))
    return out


def _seg_cov(s):
    """Ti le ky tu cua s nam trong tu dien (dung de chon cach sua chu OCR tot nhat)."""
    u = s.upper()
    n = len(u)
    if not n:
        return 0.0
    words = [w for w in B.VOCAB if len(w) >= 2]
    best = [-1] * (n + 1)
    best[0] = 0
    for i in range(n):
        if best[i] < 0:
            continue
        if best[i + 1] < best[i]:
            best[i + 1] = best[i]
        for w in words:
            if u.startswith(w, i) and best[i + len(w)] < best[i] + len(w):
                best[i + len(w)] = best[i] + len(w)
    return best[n] / float(n)


def _fix_word(t):
    """1 tu IN HOA dinh chu: 'GRADENAMIE' -> 'GRADE NAME', 'FILLERMIASTERBATCH' -> 'FILLER MASTERBATCH'.
    Font co chan (Times) hay bi OCR doc 'M' thanh 'MI' -> thu bo 'I' xem co khop tu dien hon khong."""
    if not (len(t) >= 5 and t.isalpha() and t.isupper()):
        return t
    cands = [t]
    if 'MI' in t:
        cands.append(t.replace('MI', 'M'))
        for m in re.finditer('MI', t):
            cands.append(t[:m.start()] + 'M' + t[m.end():])
    if 'RN' in t:
        cands.append(t.replace('RN', 'M'))
    best = max(cands, key=lambda c: (round(_seg_cov(c), 3), c == t))
    if _seg_cov(best) < 0.6:
        return t
    out = B._segment(best)
    # 'MASTER BATCH' -> giu dang 1 tu pho bien
    return out.replace('MASTER BATCH', 'MASTERBATCH')


def _fix_text(s):
    """Sua chu OCR: tach le chu cai ('TR A D IN G'), dinh chu ('MadeinVietnam'), '25KGS' -> '25 KGS'."""
    s = re.sub(r'\s+', ' ', s or '').strip()
    toks = s.split(' ')
    keep = ('in', 'of', 'no', 'IN', 'OF', 'NO', 'No', 'In', 'Of', 'PE', 'PP', 'PS', 'A', 'I', 'to', 'TO', 'by', 'BY')
    frag = [t for t in toks if len(t) <= 2 and t.isalpha() and t not in keep]
    if len(frag) >= 2 or any(len(t) == 1 and t.islower() for t in frag):
        s = ''.join(toks)                                # manh le -> ghep lai roi tach theo tu dien
    elif frag:
        s = re.sub(r'(?<=[A-Za-z]) (?=%s\b)' % re.escape(frag[0]), '', s)
    s = re.sub(r'(\d)\s*(KGS|KG|LBS)\b', r'\1 \2', s, flags=re.I)
    s = re.sub(r'\b[A-Z]{5,}\b', lambda m: _fix_word(m.group(0)), s)
    s = re.sub(r'[A-Za-z]{9,}', lambda m: B._segment(m.group(0)), s)
    return re.sub(r'\s+', ' ', s).strip()


# ---------------------------------------------------------------------
#  Anh -> mat na "diem den" (thich ung: ca anh scan/chup co duong ke XAM, mo, hoi nghieng)
# ---------------------------------------------------------------------
def _dark_mask(g):
    import numpy as np
    a = np.asarray(g, dtype=np.int16)
    bg = float(np.percentile(a, 90))                    # mau nen (giay)
    ink = float(np.percentile(a, 1))                    # mau muc dam nhat
    thr = ink + (bg - ink) * 0.62                       # nguong giua nen va muc
    thr = max(90.0, min(thr, bg - 25.0))
    m = a < thr
    # bo cac MANG toi lon (nen xam quanh anh chup man hinh, o to den, logo...) -> chi giu net manh
    try:
        from PIL import Image, ImageFilter
        H, W = m.shape
        k = max(7, int(0.02 * min(W, H)) | 1)
        mi = Image.fromarray((m * 255).astype(np.uint8))
        blob = mi.filter(ImageFilter.MinFilter(k)).filter(ImageFilter.MaxFilter(k))
        m &= ~(np.asarray(blob) > 0)
    except Exception:
        pass
    return m


def _longest_run(v):
    """v: mang bool 1 chieu -> (do dai, bat dau, ket thuc) cua doan True dai nhat"""
    import numpy as np
    if not v.any():
        return 0, 0, 0
    d = np.diff(np.concatenate(([0], v.astype(np.int8), [0])))
    st = np.flatnonzero(d == 1)
    en = np.flatnonzero(d == -1) - 1
    k = int(np.argmax(en - st))
    return int(en[k] - st[k] + 1), int(st[k]), int(en[k])


def _close1d(v, gap):
    """Noi cac khe False ngan (<= gap) nam giua 2 doan True"""
    import numpy as np
    v = v.copy()
    idx = np.flatnonzero(v)
    if len(idx) < 2:
        return v
    d = np.diff(idx)
    for k in np.flatnonzero((d > 1) & (d <= gap + 1)):
        v[idx[k]:idx[k + 1]] = True
    return v


def _find_hlines(dark, min_frac=0.35):
    """Duong ke ngang (cho phep nghieng nhe) -> [(y, x_dau, x_cuoi)].
    Nhom hang lien tiep; nhom qua day (vd. vung chu dam) bi loai."""
    H, W = dark.shape
    dh = dark.copy()
    dh[1:] |= dark[:-1]                                  # gian doc 1 px -> duong nghieng / mo van lien
    hl = []
    for y in range(H):
        n, a, b = _longest_run(dh[y])
        if n >= min_frac * W:
            hl.append((y, a, b))
    groups = []
    for y, a, b in hl:
        if groups and y - groups[-1][-1][0] <= 2:
            groups[-1].append((y, a, b))
        else:
            groups.append([(y, a, b)])
    maxth = max(6, int(0.04 * H))
    return [(sum(v[0] for v in gp) / len(gp), min(v[1] for v in gp), max(v[2] for v in gp))
            for gp in groups if gp[-1][0] - gp[0][0] <= maxth]


def _vsplit(dark, xa, xb, ya, yb):
    """Tim duong ke doc chia cot trong 1 dai hang -> toa do x hoac None"""
    import numpy as np
    h = yb - ya
    y_a, y_b = int(ya + h * 0.18), int(yb - h * 0.18)
    if y_b - y_a < 2:
        return None
    x_a, x_b = int(xa) + 6, int(xb) - 6
    if x_b <= x_a:
        return None
    band = dark[y_a:y_b + 1, x_a:x_b]
    bd = band.copy()
    bd[:, 1:] |= band[:, :-1]                           # gian ngang 1 px (duong nghieng nhe)
    frac = bd.mean(axis=0)
    ok = frac >= 0.85
    if not ok.any():
        return None
    d = np.diff(np.concatenate(([0], ok.astype(np.int8), [0])))
    st = np.flatnonzero(d == 1)
    en = np.flatnonzero(d == -1) - 1
    mids = [x_a + (a + b) / 2.0 for a, b in zip(st, en) if b - a <= max(6, 0.03 * (xb - xa))]
    if not mids:
        return None
    return min(mids, key=lambda m: abs(m - (xa + xb) / 2.0))


def design_from_image(src, w_mm=None, h_mm=None):
    """Anh tem (chup / scan / chup man hinh) -> thiet ke. Kich thuoc mac dinh theo ti le anh.

    Nhan dien:
      - khung ngoai + duong ke ngang (ca duong XAM / mo cua anh scan)
      - duong ke doc chia 2 cot (Nhan | Gia tri) theo TUNG dong
      - 1 dai hang KHONG co duong ke ngan nhung chua nhieu dong chu -> tach thanh nhieu dong (khong ke)
    """
    import numpy as np
    import image_recipe
    img = image_recipe._to_pil(src)
    g = img.convert('L')
    W, H = g.size
    dark = _dark_mask(g)
    hlines = _find_hlines(dark)
    boxes = []
    try:
        boxes = image_recipe.ocr_image(img, use_cls=False)
    except TypeError:
        boxes = image_recipe.ocr_image(img)
    except Exception:
        boxes = []
    d = new_design()
    d['rows'] = []
    table = None
    if len(hlines) >= 2:
        wid = max(hlines, key=lambda h: h[2] - h[1])
        xa, xb = wid[1], wid[2]
        # vien trai: cot doc gan xa
        best = None
        gap = max(4, int(0.06 * H))                 # vien scan hay bi dut quang -> noi cac khe ngan
        for x in range(max(0, int(xa) - 3), min(W, int(xa) + 8)):
            col = dark[:, x].copy()
            if x + 1 < W:
                col |= dark[:, x + 1]
            n, a, b = _longest_run(_close1d(col, gap))
            if n > (best[1] - best[0] if best else 0):
                best = (a, b)
        if best and best[1] - best[0] > 0.2 * H:
            top, bot = best
        else:
            top, bot = hlines[0][0], hlines[-1][0]
        ls = [h for h in hlines if top - 4 <= h[0] <= bot + 4 and (h[2] - h[1]) >= 0.6 * (xb - xa)]
        if ls and bot - ls[-1][0] > 12:          # anh bi cat mat canh duoi -> dong khung o cuoi vien trai
            ls.append((float(bot), xa, xb))
        if ls and ls[0][0] - top > 12:
            ls.insert(0, (float(top), xa, xb))
        if len(ls) >= 2:
            table = (xa, xb, ls)
    if table:
        xa, xb, ls = table
        ys = [h[0] for h in ls]
        bands = [(ys[i], ys[i + 1]) for i in range(len(ys) - 1) if ys[i + 1] - ys[i] > 6]
        splits = []
        rows = []
        for ya, yb in bands:
            sp = _vsplit(dark, xa, xb, ya, yb)
            if sp is not None:
                splits.append((sp - xa) / float(xb - xa))
            rows.append((ya, yb, sp))
        split = round(sum(splits) / len(splits), 3) if splits else 0.45
        d['split'] = split
        hmed = sorted(yb - ya for ya, yb, _ in rows)[len(rows) // 2] if rows else 1
        for ya, yb, sp in rows:
            bx = [b for b in boxes if ya < b.cy < yb and xa - 4 < b.cx < xb + 4]
            if sp is not None:
                cells = [_cell_text([b for b in bx if b.cx < sp]), _cell_text([b for b in bx if b.cx >= sp])]
                d['rows'].append({'cells': cells, 'h': round((yb - ya) / float(H), 4), 'pt': [0, 0],
                                  'line': True})
                continue
            lines = _group_lines(bx)
            if len(lines) >= 2 and (yb - ya) >= 1.5 * hmed:
                # dai hang cao chua nhieu dong chu, khong co duong ke ngan -> moi dong chu = 1 dong tem
                cys = [sum(b.cy for b in ln) / len(ln) for ln in lines]
                cuts = [ya] + [(cys[k] + cys[k + 1]) / 2.0 for k in range(len(cys) - 1)] + [yb]
                for k, ln in enumerate(lines):
                    two = _two_cols(ln, xa + split * (xb - xa))
                    cells = two or [_cell_text(ln)]
                    d['rows'].append({'cells': cells, 'h': round((cuts[k + 1] - cuts[k]) / float(H), 4),
                                      'pt': [0] * len(cells), 'line': k == len(lines) - 1})
                continue
            cells = [_cell_text(bx)]
            d['rows'].append({'cells': cells, 'h': round((yb - ya) / float(H), 4), 'pt': [0], 'line': True})
        d['border'] = True
        aspect = (xb - xa) / float(max(1, ls[-1][0] - ls[0][0]))
    else:
        # khong co khung: moi dong chu = 1 dong
        d['border'] = False
        lines = _group_lines(boxes)
        if not lines:
            raise B.BtwError('Không đọc được chữ nào trong ảnh.')
        hs = [max(b.h for b in ln) for ln in lines]
        for ln, h in zip(lines, hs):
            d['rows'].append({'cells': [_cell_text(ln)], 'h': round(h * 1.6 / H, 4), 'pt': [0], 'line': False})
        xa = min(b.x0 for ln in lines for b in ln)
        xb = max(b.x1 for ln in lines for b in ln)
        aspect = (xb - xa) * 1.15 / float(sum(h * 1.6 for h in hs))
    if w_mm and h_mm:
        d['w_mm'], d['h_mm'] = float(w_mm), float(h_mm)
    else:
        d['w_mm'] = 60.0
        d['h_mm'] = round(max(10.0, min(200.0, 60.0 / max(0.2, aspect))), 1)
    d['name'] = next((c for r in d['rows'] for c in r['cells'] if c.strip()), 'Mẫu tem')[:40].replace('\n', ' ')
    return normalize(d)


def _two_cols(ln, sp):
    """1 dong chu co ca 2 phia duong chia (khong ke doc) va khoang trong lon o giua -> 2 o, nguoc lai None"""
    left = [b for b in ln if b.x1 <= sp + 2]
    right = [b for b in ln if b.x0 >= sp - 2]
    if left and right and len(left) + len(right) == len(ln):
        gap = min(b.x0 for b in right) - max(b.x1 for b in left)
        if gap > 2.5 * max(b.h for b in ln):
            return [_cell_text(left), _cell_text(right)]
    return None


def _group_lines(boxes):
    if not boxes:
        return []
    hmed = sorted(b.h for b in boxes)[len(boxes) // 2]
    out = []
    for b in sorted(boxes, key=lambda b: b.cy):
        if out and abs(b.cy - sum(x.cy for x in out[-1]) / len(out[-1])) <= 0.5 * hmed:
            out[-1].append(b)
        else:
            out.append([b])
    return [sorted(ln, key=lambda b: b.x0) for ln in out]


def _cell_text(bx):
    return '\n'.join(_fix_text(' '.join(b.text for b in ln)) for ln in _group_lines(bx))


# =====================================================================
#  XUAT FILE .btw
# =====================================================================
def export_design(dz, donor_path, out_path, db_values=None):
    """Ghi thiet ke thanh file .btw. -> (duong dan .btw, duong dan Excel du lieu)
    db_values: noi dung cho o chu lay tu Excel (Text 12) - nhieu gia tri = in nhieu tem."""
    L = compute_layout(dz)
    d = L['d']
    btw = B.BtwFile(donor_path)
    g = B.Geometry(btw)
    need = usable_texts(g)                         # tat ca o chu dung duoc cua phoi (khong gioi han 17)
    errs = check_limits(dz, cap_texts=len(need), cap_lines=len(g.lines))
    if errs:
        raise B.BtwError('\n'.join(errs))
    miss = []
    if len(need) < len(L['texts']):
        miss.append('ô chữ Text (phôi có %d, tem cần %d)' % (len(need), len(L['texts'])))
    if len(g.lines) < 1:
        miss.append('đường kẻ Line (cần ≥1)')
    if not g.boxes:
        miss.append('khung chữ nhật Box/Rectangle (cần ≥1 – để vẽ viền bảng)')
    if not g.page_pat:
        miss.append('khổ tem/khổ giấy (page size) – phôi chưa đặt kích thước tem')
    if miss:
        raise B.BtwError('File phôi .btw thiếu: ' + '; '.join(miss) + '.\n\n'
                         'Phôi .btw phải là 1 tem BarTender có sẵn: nhiều ô chữ (Text), '
                         'ít nhất 1 đường kẻ (Line) và 1 khung chữ nhật (Box/Rectangle bao ngoài), '
                         'và đã đặt khổ tem. Phần mềm chỉ MƯỢN LẠI các đối tượng này (không tạo mới được).')
    g.set_page(d['w_mm'] * MIL, d['h_mm'] * MIL)
    x0, y0, x1, y1 = L['rect']
    # --- khung + duong ke
    segs = list(L['hlines']) + list(L['vlines'])
    box_names = sorted(g.boxes, key=lambda n: -g.boxes[n]['h'])
    line_names = sorted(g.lines, key=lambda n: int(n.split()[1]))
    far = L['tpl_w'] + 3000                    # khong ke khung -> dua ra ngoai vung in
    if d['border']:
        g.set_box(box_names[0], x0, y0, x1 - x0, y1 - y0)
    else:
        g.set_box(box_names[0], far, 0, 10, 10)
    extra = box_names[1:]
    hs = [s for s in segs if s[1] == s[3]]
    vs = [s for s in segs if s[0] == s[2]]
    # thua duong ke -> dung khung chu nhat phu (2 canh ngang = 2 duong ke)
    while len(hs) + len(vs) > len(line_names) and extra and len(hs) >= 2:
        a, b = hs.pop(0), hs.pop(0)
        g.set_box(extra.pop(0), x0, a[1], x1 - x0, b[1] - a[1])
    for n in extra:                             # khung phu khong dung -> trung vien ngoai / ra ngoai
        if d['border']:
            g.set_box(n, x0, y0, x1 - x0, y1 - y0)
        else:
            g.set_box(n, far, 0, 10, 10)
    segs = hs + vs
    for k, n in enumerate(line_names):
        if k < len(segs):
            g.set_line(n, *segs[k])
        elif d['border']:
            g.set_line(n, x0, y0, x1, y0)      # trung canh tren -> khong thay
        else:
            g.set_line(n, far, 0, far + 10, 0)
    # --- chu
    sx = d['stretch'] / 100.0
    weight = 700 if d['bold'] else 400
    db_name = 'Text 12' if 'Text 12' in g.texts and g.texts['Text 12'] and \
        any(o.name == 'Text 12' and o.db for o in btw.objects) else None
    names = [n for n in need if n != db_name]
    items = list(L['texts'])
    db_text = None
    if db_name and len(items) > len(names):
        db_text = items.pop()                  # o cuoi cung lay qua Excel
    for n, t in zip(names, items):
        _place(g, btw, n, t, sx, weight)
    for n in names[len(items):]:               # o chu thua: rong, dat vao goc
        btw.set_text(n, ' ')
        g.set_text_frame(n, x0 + 10, y0 + 10, 40, pt=2, weight=weight, stretch=sx)
    data_xlsx = None
    if db_name:
        if db_text is not None:
            _place(g, btw, db_name, db_text, sx, weight, set_value=False)
            vals = db_values or [db_text['s']]
        else:
            g.set_text_frame(db_name, x0 + 10, y0 + 10, 40, pt=2, weight=weight, stretch=sx)
            vals = db_values or [' ']
        btw.set_text(db_name, vals[0] if vals else '')
        data_xlsx = os.path.splitext(out_path)[0] + '_DATA.xlsx'
        btw.set_excel_path(os.path.abspath(data_xlsx))
        B.write_data_xlsx(btw, data_xlsx, vals)
    thumb = render_design(dz, px_per_mm=max(2, 640.0 / max(d['w_mm'], d['h_mm'])), frame=False)
    from PIL import Image
    sq = Image.new('RGB', (664, 664), (128, 128, 128))
    thumb.thumbnail((640, 640))
    sq.paste(thumb, ((664 - thumb.width) // 2, (664 - thumb.height) // 2))
    btw.save(out_path, preview=sq)
    return out_path, data_xlsx


def _place(g, btw, name, t, sx, weight, set_value=True):
    tw = text_w_per_pt(t['s']) * t['pt'] * sx
    lh = t['pt'] * PT_MIL * LINE_H
    top = t['cy'] - lh / 2.0 - B.TEXT_DROP * t['pt'] * PT_MIL
    if name in ('Text 13', 'Text 12'):         # o can trai -> khung vua khit chu, dat giua o
        w = tw * 1.08 + 20
    else:
        w = max(tw * 1.08 + 20, t['w'] * 0.96)
    g.set_text_frame(name, t['cx'] - w / 2.0, top, w, pt=t['pt'], weight=weight, stretch=sx)
    if set_value:
        btw.set_text(name, t['s'])
