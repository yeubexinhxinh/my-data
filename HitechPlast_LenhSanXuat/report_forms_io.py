# -*- coding: utf-8 -*-
"""
report_forms_io.py - DU LIEU cho tab "Báo cáo chất lượng NVL" (report_forms.py)

    FormStore    : danh sach MAU form (rmqr_forms.json) - them / sua / xoa / nhan ban / nhap tu Excel
    ReportStore  : cac BAO CAO da luu (moi bao cao = 1 file JSON + anh trong thu muc images/)
    export_pdf   : xuat 1 bao cao ra PDF (reportlab) - vua 1 trang A4 giong ExportAsFixedFormat
    export_excel : xuat 1 bao cao ra .xlsx (openpyxl) - giu o gop, vien, mau, co chu, anh

Cau truc 1 form (giong 1 khoi 8 cot trong sheet 'Initial report'):
    {id, name, group:[...], ncol, nrow, cols:[do rong Excel], rowh:[chieu cao pt],
     merges:[[r0,c0,r1,c1]], cells:[{r,c,v|ref,b,i,sz,color,al,va,wrap,fill,bd:[l,t,r,b],nf}],
     images:[{r0,c0,r1,c1,file}], lists:{"r,c":[...]}, name_cell:[r,c], date_cell:[r,c], font}
Cau truc 1 bao cao:
    {id, form_id, form_name, group, title, created, updated, user, form:{...ban sao form...},
     values:{"r,c": "chu"}, images:[{r0,c0,r1,c1,file}], hide_tpl:[chi so anh mau bi an]}
"""
import copy
import datetime as dt
import io
import json
import os
import re
import shutil
import sys
import uuid

IMG_EXT = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp')


# ---------------------------------------------------------------- duong dan
def resource_dir():
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))


def writable_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


SETTINGS_FILE = os.path.join(writable_dir(), 'rmqr_settings.json')


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(s):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(s, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def storage_dir():
    """Thu muc luu bao cao (co the la thu muc chia se tren mang LAN: \\\\may-chu\\BaoCaoNVL)."""
    d = load_settings().get('storage_dir') or os.path.join(writable_dir(), 'BaoCaoNVL')
    for sub in ('', 'reports', 'images', 'form_media', 'pdf'):
        try:
            os.makedirs(os.path.join(d, sub), exist_ok=True)
        except Exception:
            pass
    return d


def key(r, c):
    return '%d,%d' % (r, c)


def unkey(k):
    a, b = str(k).split(',')
    return int(a), int(b)


def now_str(fmt='%d/%m/%Y %H:%M'):
    return dt.datetime.now().strftime(fmt)


# ---------------------------------------------------------------- anh
def resolve_image(name):
    """Tim file anh theo ten: anh bao cao -> anh mau (nguoi dung them) -> anh mau goc (rmqr_media)."""
    if not name:
        return None
    if os.path.isabs(name) and os.path.exists(name):
        return name
    sd = storage_dir()
    for p in (os.path.join(sd, 'images', name), os.path.join(sd, 'form_media', name),
              os.path.join(resource_dir(), 'rmqr_media', name),
              os.path.join(writable_dir(), 'rmqr_media', name)):
        if os.path.exists(p):
            return p
    return None


def import_image(src, sub='images', max_side=1600):
    """Chep anh (duong dan hoac PIL.Image) vao thu muc luu -> tra ve TEN file."""
    from PIL import Image
    if isinstance(src, Image.Image):
        im = src
    else:
        im = Image.open(src)
        im.load()
    try:
        from PIL import ImageOps
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass
    im = im.copy()
    im.thumbnail((max_side, max_side))
    alpha = im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info)
    name = dt.datetime.now().strftime('%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:6]
    folder = os.path.join(storage_dir(), sub)
    os.makedirs(folder, exist_ok=True)
    if alpha:
        name += '.png'
        im.save(os.path.join(folder, name), 'PNG', optimize=True)
    else:
        name += '.jpg'
        im.convert('RGB').save(os.path.join(folder, name), 'JPEG', quality=88)
    return name


# ---------------------------------------------------------------- tinh toan form
class FormModel:
    """Truy cap nhanh 1 form: o theo (r,c), vung gop, chu hien thi."""

    def __init__(self, form):
        self.f = form
        self.cells = {(c['r'], c['c']): c for c in form.get('cells', [])}
        self.span = {}
        self.covered = {}
        for (r0, c0, r1, c1) in form.get('merges', []):
            self.span[(r0, c0)] = (r1, c1)
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    if (rr, cc) != (r0, c0):
                        self.covered[(rr, cc)] = (r0, c0)

    @property
    def nrow(self):
        return int(self.f.get('nrow') or len(self.f.get('rowh', [])))

    @property
    def ncol(self):
        return int(self.f.get('ncol') or len(self.f.get('cols', [])))

    def anchor(self, r, c):
        return self.covered.get((r, c), (r, c))

    def area(self, r, c):
        r, c = self.anchor(r, c)
        r1, c1 = self.span.get((r, c), (r, c))
        return r, c, r1, c1

    def style(self, r, c):
        return self.cells.get((r, c), {})

    def border_lines(self):
        """Cac doan vien can ve: [(r, c, side)] side: 0=trai 1=tren 2=phai 3=duoi.
        Bo vien NAM TRONG vung gop (Excel khong ve vien ben trong o da gop)."""
        out = []
        nb = {0: (0, -1), 1: (-1, 0), 2: (0, 1), 3: (1, 0)}
        for (r, c), cell in self.cells.items():
            bd = cell.get('bd') or (0, 0, 0, 0)
            if not any(bd):
                continue
            a = self.area(r, c)
            for side in range(4):
                if not bd[side]:
                    continue
                dr, dc = nb[side]
                if self.area(r + dr, c + dc) == a and (r + dr, c + dc) != (r, c):
                    continue
                out.append((r, c, side))
        return out

    @staticmethod
    def fmt_value(v, nf=''):
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, (int, float)):
            if nf and '%' in nf:
                dec = 0
                m = re.search(r'0\.(0+)%', nf)
                if m:
                    dec = len(m.group(1))
                return f'{v * 100:.{dec}f}%'
            if float(v).is_integer():
                return str(int(v))
            return f'{v:.10g}'
        return '' if v is None else str(v)

    def text(self, r, c, values=None, _depth=0):
        values = values or {}
        k = key(r, c)
        if k in values:
            return values[k]
        cell = self.cells.get((r, c))
        if not cell:
            return ''
        if cell.get('ref') and _depth < 8:
            rr, cc = cell['ref']
            return self.text(rr, cc, values, _depth + 1)
        return self.fmt_value(cell.get('v'), cell.get('nf', ''))


def blank_form(name='Form mới', group=None, ncol=8, nrow=30):
    cells = []
    for r in range(1, nrow + 1):
        for c in range(1, ncol + 1):
            cells.append({'r': r, 'c': c, 'b': 0, 'i': 0, 'sz': 12.0, 'color': '#000000', 'al': 'left',
                          'va': 'center', 'wrap': 1, 'fill': '', 'bd': [1, 1, 1, 1] if r >= 3 else [0, 0, 0, 0]})
    cells[0].update({'v': name.upper(), 'b': 1, 'sz': 16.0, 'al': 'center'})
    return {'id': '', 'name': name, 'group': group or ['Form tự tạo'], 'ncol': ncol, 'nrow': nrow,
            'cols': [14.0] * ncol, 'rowh': [30.0] * nrow, 'merges': [[1, 1, 2, ncol]], 'cells': cells,
            'images': [], 'lists': {}, 'font': 'Times New Roman', 'name_cell': None, 'date_cell': None, 'concl_cell': None,
            'note': ''}


def shift_form(form, axis, at, n):
    """Chen (n>0) / xoa (n<0) hang hoac cot tai vi tri `at` (1-based). axis = 'r' | 'c'."""
    f = form
    i = 0 if axis == 'r' else 1

    def mv(x):
        return x + n if x >= at else x
    if n < 0:
        lo, hi = at, at - n - 1          # vung bi xoa
        f['cells'] = [c for c in f['cells'] if not (lo <= (c['r'] if axis == 'r' else c['c']) <= hi)]
        newm = []
        for m in f['merges']:
            a, b = (m[0], m[2]) if axis == 'r' else (m[1], m[3])
            if a >= lo and b <= hi:
                continue
            na = a if a < lo else (lo if a <= hi else a + n)
            nb = b if b < lo else (max(lo - 1, na) if b <= hi else b + n)
            if axis == 'r':
                m = [na, m[1], nb, m[3]]
            else:
                m = [m[0], na, m[2], nb]
            if m[0] != m[2] or m[1] != m[3]:
                newm.append(m)
        f['merges'] = newm
        imgs = []
        for im in f.get('images', []):
            a, b = (im['r0'], im['r1']) if axis == 'r' else (im['c0'], im['c1'])
            if a >= lo and b <= hi:
                continue
            na = a if a < lo else (lo if a <= hi else a + n)
            nb = b if b < lo else (max(lo - 1, na) if b <= hi else b + n)
            im = dict(im)
            if axis == 'r':
                im['r0'], im['r1'] = na, max(na, nb)
            else:
                im['c0'], im['c1'] = na, max(na, nb)
            imgs.append(im)
        f['images'] = imgs
        for c in f['cells']:
            if axis == 'r' and c['r'] > hi:
                c['r'] += n
            if axis == 'c' and c['c'] > hi:
                c['c'] += n
            if c.get('ref'):
                if c['ref'][i] > hi:
                    c['ref'][i] += n
        f['lists'] = {k: v for k, v in f.get('lists', {}).items()
                      if not (lo <= unkey(k)[i] <= hi)}
        f['lists'] = {(key(unkey(k)[0] + (n if axis == 'r' and unkey(k)[0] > hi else 0),
                           unkey(k)[1] + (n if axis == 'c' and unkey(k)[1] > hi else 0))): v
                      for k, v in f['lists'].items()}
        if axis == 'r':
            del f['rowh'][lo - 1:hi]
            f['nrow'] = len(f['rowh'])
        else:
            del f['cols'][lo - 1:hi]
            f['ncol'] = len(f['cols'])
    else:
        for c in f['cells']:
            if axis == 'r':
                c['r'] = mv(c['r'])
            else:
                c['c'] = mv(c['c'])
            if c.get('ref'):
                c['ref'][i] = mv(c['ref'][i])
        newm = []
        for m in f['merges']:
            if axis == 'r':
                m = [mv(m[0]), m[1], m[2] + n if m[2] >= at else m[2], m[3]]
            else:
                m = [m[0], mv(m[1]), m[2], m[3] + n if m[3] >= at else m[3]]
            newm.append(m)
        f['merges'] = newm
        for im in f.get('images', []):
            if axis == 'r':
                im['r0'], im['r1'] = mv(im['r0']), (im['r1'] + n if im['r1'] >= at else im['r1'])
            else:
                im['c0'], im['c1'] = mv(im['c0']), (im['c1'] + n if im['c1'] >= at else im['c1'])
        f['lists'] = {key(mv(unkey(k)[0]) if axis == 'r' else unkey(k)[0],
                          mv(unkey(k)[1]) if axis == 'c' else unkey(k)[1]): v
                      for k, v in f.get('lists', {}).items()}
        if axis == 'r':
            ref_h = f['rowh'][at - 2] if at >= 2 and at - 2 < len(f['rowh']) else 30.0
            for _ in range(n):
                f['rowh'].insert(at - 1, ref_h)
            f['nrow'] = len(f['rowh'])
            # chep vien cua hang tren cho hang moi
            above = [c for c in f['cells'] if c['r'] == at - 1]
            for k2 in range(n):
                for c in above:
                    nc = {x: copy.deepcopy(y) for x, y in c.items() if x not in ('v', 'ref', 'covered')}
                    nc['r'] = at + k2
                    f['cells'].append(nc)
        else:
            ref_w = f['cols'][at - 2] if at >= 2 and at - 2 < len(f['cols']) else 12.0
            for _ in range(n):
                f['cols'].insert(at - 1, ref_w)
            f['ncol'] = len(f['cols'])
    for key_ in ('name_cell', 'date_cell', 'concl_cell'):
        v = f.get(key_)
        if v:
            if n < 0 and lo <= v[i] <= hi:
                f[key_] = None
            elif n < 0 and v[i] > hi:
                v[i] += n
            elif n > 0 and v[i] >= at:
                v[i] += n
    return f


def shift_values(values, axis, at, n):
    """Doi toa do cac o da nhap khi chen / xoa hang, cot trong BAO CAO."""
    i = 0 if axis == 'r' else 1
    out = {}
    for k, v in values.items():
        rc = list(unkey(k))
        if n < 0 and at <= rc[i] <= at - n - 1:
            continue
        if n < 0 and rc[i] > at - n - 1:
            rc[i] += n
        elif n > 0 and rc[i] >= at:
            rc[i] += n
        out[key(*rc)] = v
    return out


# ---------------------------------------------------------------- MAU FORM
class FormStore:
    FILE = 'rmqr_forms.json'

    def __init__(self):
        self.path_user = os.path.join(storage_dir(), self.FILE)
        self.path_base = os.path.join(resource_dir(), self.FILE)
        self.forms = []
        self.load()

    def load(self):
        for p in (self.path_user, self.path_base, os.path.join(writable_dir(), self.FILE)):
            if os.path.exists(p):
                with open(p, encoding='utf-8') as fh:
                    doc = json.load(fh)
                self.forms = doc.get('forms', doc if isinstance(doc, list) else [])
                return
        self.forms = []

    def save(self):
        tmp = self.path_user + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as fh:
            json.dump({'version': 1, 'forms': self.forms}, fh, ensure_ascii=False, separators=(',', ':'))
        os.replace(tmp, self.path_user)

    def get(self, fid):
        return next((f for f in self.forms if f.get('id') == fid), None)

    def new_id(self):
        ids = {f.get('id') for f in self.forms}
        n = 1
        while 'U%03d' % n in ids:
            n += 1
        return 'U%03d' % n

    def add(self, form, after_id=None):
        form = copy.deepcopy(form)
        form['id'] = self.new_id()
        idx = next((i for i, f in enumerate(self.forms) if f.get('id') == after_id), None)
        if idx is None:
            self.forms.append(form)
        else:
            self.forms.insert(idx + 1, form)
        self.save()
        return form

    def update(self, form):
        for i, f in enumerate(self.forms):
            if f.get('id') == form.get('id'):
                self.forms[i] = copy.deepcopy(form)
                self.save()
                return
        self.forms.append(copy.deepcopy(form))
        self.save()

    def delete(self, fid):
        self.forms = [f for f in self.forms if f.get('id') != fid]
        self.save()

    def move(self, fid, step):
        i = next((i for i, f in enumerate(self.forms) if f.get('id') == fid), None)
        if i is None:
            return
        j = max(0, min(len(self.forms) - 1, i + step))
        self.forms.insert(j, self.forms.pop(i))
        self.save()

    def groups(self):
        seen = []
        for f in self.forms:
            g = ' > '.join(f.get('group') or [])
            if g not in seen:
                seen.append(g)
        return seen


# ---------------------------------------------------------------- BAO CAO DA LUU
class ReportStore:
    def __init__(self):
        self.dir = os.path.join(storage_dir(), 'reports')
        os.makedirs(self.dir, exist_ok=True)

    def list(self):
        out = []
        for fn in os.listdir(self.dir):
            if not fn.endswith('.json'):
                continue
            try:
                with open(os.path.join(self.dir, fn), encoding='utf-8') as fh:
                    r = json.load(fh)
                out.append({k: r.get(k, '') for k in ('id', 'title', 'form_id', 'form_name', 'group',
                                                       'created', 'updated', 'user', 'material')})
            except Exception:
                continue
        out.sort(key=lambda r: r.get('updated') or r.get('created') or '', reverse=True)
        return out

    def load(self, rid):
        with open(os.path.join(self.dir, rid + '.json'), encoding='utf-8') as fh:
            return json.load(fh)

    def save(self, rec):
        if not rec.get('id'):
            rec['id'] = dt.datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:4]
            rec['created'] = rec.get('created') or dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        rec['updated'] = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        p = os.path.join(self.dir, rec['id'] + '.json')
        tmp = p + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as fh:
            json.dump(rec, fh, ensure_ascii=False, separators=(',', ':'))
        os.replace(tmp, p)
        return rec

    def delete(self, rid):
        p = os.path.join(self.dir, rid + '.json')
        if os.path.exists(p):
            os.remove(p)


# ---------------------------------------------------------------- hinh hoc chung
def geometry(form, px_col=7.0, px_pad=5.0, row_k=1.0, zoom=1.0):
    """Toa do bien cot / hang. Tra ve (xs, ys) - xs[c-1] = mep trai cot c."""
    xs = [0.0]
    for w in form.get('cols', []):
        xs.append(xs[-1] + (w * px_col + px_pad) * zoom if w else xs[-1])
    ys = [0.0]
    for h in form.get('rowh', []):
        ys.append(ys[-1] + (h or 15.0) * row_k * zoom)
    return xs, ys


def fit_box(iw, ih, bw, bh):
    """Anh iw x ih vua khung bw x bh (giu ti le) -> (w, h, dx, dy)."""
    if iw <= 0 or ih <= 0:
        return bw, bh, 0, 0
    k = min(bw / iw, bh / ih)
    w, h = iw * k, ih * k
    return w, h, (bw - w) / 2, (bh - h) / 2


def all_images(rec):
    """Anh mau (khong bi an) + anh cua bao cao."""
    form = rec['form']
    hide = set(rec.get('hide_tpl') or [])
    out = [dict(im, tpl=i) for i, im in enumerate(form.get('images', [])) if i not in hide]
    out += [dict(im, tpl=None, idx=j) for j, im in enumerate(rec.get('images', []))]
    return out


# ---------------------------------------------------------------- KET LUAN
# So sanh "ket qua" voi "tieu chuan" de GOI Y ket luan (nguoi lap tu sua lai).
def concl_num(s):
    """Doc 1 so don tu chuoi: '3,5' '3.5%' ' 3.5 ' -> 3.5 ; None neu khong phai."""
    if s is None:
        return None
    t = str(s).strip().replace('%', '').replace(',', '.')
    m = re.fullmatch(r'[-+]?\d+(?:\.\d+)?', t)
    return float(m.group()) if m else None


def concl_range(s):
    """Doc dai tieu chuan -> (lo, hi) (lo/hi co the None). None neu khong phai dai.
    Ho tro: 'a-b', 'a~b', 'a den b', '<=b' '≤b' 'max b', '>=a' '≥a' 'min a', 'a±d'."""
    if s is None:
        return None
    t = str(s).strip().replace(',', '.')
    low = t.lower()
    if not low:
        return None
    m = re.fullmatch(r'([-+]?\d+(?:\.\d+)?)\s*(?:±|\+/-|\+-)\s*(\d+(?:\.\d+)?)', t)
    if m:
        a, d = float(m.group(1)), float(m.group(2))
        return (a - d, a + d)
    m = re.fullmatch(r'(?:≤|<=|<|max\.?|tối đa|toi da)\s*([-+]?\d+(?:\.\d+)?)', low)
    if m:
        return (None, float(m.group(1)))
    m = re.fullmatch(r'(?:≥|>=|>|min\.?|tối thiểu|toi thieu)\s*([-+]?\d+(?:\.\d+)?)', low)
    if m:
        return (float(m.group(1)), None)
    m = re.fullmatch(r'([-+]?\d+(?:\.\d+)?)\s*(?:-|–|—|~|đến|den|to)\s*([-+]?\d+(?:\.\d+)?)', low)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        return (min(a, b), max(a, b))
    return None


def _concl_fmt(v):
    return str(int(v)) if float(v).is_integer() else f'{v:g}'


def _concl_fmt_range(lo, hi):
    if lo is not None and hi is not None:
        return f'{_concl_fmt(lo)}–{_concl_fmt(hi)}'
    if hi is not None:
        return f'≤ {_concl_fmt(hi)}'
    if lo is not None:
        return f'≥ {_concl_fmt(lo)}'
    return ''


def _tpl_cell_v(fm, r, c):
    """Gia tri MAU (template) cua o (theo o neo neu bi gop). None neu khong co."""
    a = fm.anchor(r, c)
    cell = fm.cells.get(a)
    return cell.get('v') if cell else None


def _find_std_range(fm, r, c):
    """Tim DAI tieu chuan cho o ket qua (r,c):
    1) cung HANG: o nao co gia tri mau la dai (vd 'Chi tieu | 2-3 | <ket qua>')
    2) cung COT, quet LEN tren: dai tieu chuan gan nhat (form 2 phan: TC o phan I)."""
    for cc in range(1, fm.ncol + 1):
        if cc == c:
            continue
        pr = concl_range(_tpl_cell_v(fm, r, cc))
        if pr:
            return pr
    for rr in range(r - 1, 0, -1):
        pr = concl_range(_tpl_cell_v(fm, rr, c))
        if pr:
            return pr
    return None


def _result_label(fm, values, r, c):
    """Ten chi tieu cho o ket qua: tieu de cung cot (quet len), hoac chu dau hang."""
    for rr in range(r - 1, 0, -1):
        s = (fm.text(rr, c, values) or '').strip()
        if s and concl_num(s) is None and concl_range(s) is None:
            return s
    for cc in range(1, fm.ncol + 1):
        s = (fm.text(r, cc, values) or '').strip()
        if s and concl_num(s) is None and concl_range(s) is None:
            return s
    return f'Chỉ tiêu (hàng {r}, cột {c})'


def suggest_conclusion(fm, values):
    """CHI danh gia cac o NGUOI DUNG DA NHAP (trong values) la so, roi tim tieu
    chuan tuong ung (cung hang / cung cot phia tren). Tra ve (checked, issues)."""
    issues = []
    checked = 0
    for k, v in (values or {}).items():
        val = concl_num(v)
        if val is None:
            continue
        try:
            r, c = unkey(k)
        except Exception:
            continue
        if r > fm.nrow or c > fm.ncol:
            continue
        rng = _find_std_range(fm, r, c)
        if not rng:
            continue
        checked += 1
        lo, hi = rng
        how = None
        if lo is not None and val < lo:
            how = 'thấp hơn'
        elif hi is not None and val > hi:
            how = 'vượt'
        if how:
            issues.append((_result_label(fm, values, r, c), val, rng, how))
    return checked, issues


def conclusion_text(fm, values):
    """Chuoi ket luan GOI Y dua tren so sanh (nguoi lap co the sua lai)."""
    checked, issues = suggest_conclusion(fm, values)
    if checked == 0:
        return ('KẾT LUẬN: ....................................................\n'
                '(Người lập tự đánh giá: ĐẠT / KHÔNG ĐẠT so với tiêu chuẩn)')
    if not issues:
        return 'KẾT LUẬN: Các chỉ tiêu nằm trong tiêu chuẩn → Nguyên vật liệu ĐẠT yêu cầu.'
    lines = ['KẾT LUẬN: Nguyên vật liệu KHÔNG ĐẠT — các chỉ tiêu ngoài tiêu chuẩn:']
    for (label, val, (lo, hi), how) in issues:
        lines.append(f'• {label}: {_concl_fmt(val)} ({how} tiêu chuẩn {_concl_fmt_range(lo, hi)})')
    return '\n'.join(lines)


# ---------------------------------------------------- SO THEO DOI (tong hop)
def _norm2(s):
    """Bo dau tieng Viet + chu thuong (dung cho do nhan field)."""
    import unicodedata
    t = unicodedata.normalize('NFD', str(s or ''))
    t = ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')
    return t.replace('đ', 'd').replace('Đ', 'D').lower()


_FIELD_LABELS = {
    'material': ('ten nguyen vat lieu', 'ten nvl', 'ten san pham', 'ten vat tu'),
    'supplier': ('nha cung cap', 'supplier'),
    'address':  ('dia chi', 'address'),
    'date_recv': ('ngay nhan mau', 'sample received', 'ngay nhan'),
    'date_result': ('ngay tra kq', 'ngay tra', 'test result'),
}


def _value_right_of(fm, r, c, values):
    """Gia tri nam ben phai o nhan (cung hang), bo qua o trong."""
    parts = []
    for cc in range(c + 1, fm.ncol + 1):
        if fm.anchor(r, cc) == (r, c):
            continue
        t = (fm.text(r, cc, values) or '').strip()
        if t:
            parts.append(t)
    return ' '.join(parts).strip()


def extract_fields(fm, values):
    """Trich cac field tieu de tu 1 bao cao: material/supplier/address/date_recv/date_result."""
    out = {'material': '', 'supplier': '', 'address': '', 'date_recv': '', 'date_result': ''}
    nc = fm.f.get('name_cell')
    dc = fm.f.get('date_cell')
    if nc:
        out['material'] = (fm.text(nc[0], nc[1], values) or '').strip()
    if dc:
        out['date_result'] = (fm.text(dc[0], dc[1], values) or '').strip()
    for r in range(1, fm.nrow + 1):
        for c in range(1, fm.ncol + 1):
            if (r, c) in fm.covered:
                continue
            lab = _norm2(fm.text(r, c, values))
            if not lab:
                continue
            for name, keys in _FIELD_LABELS.items():
                if out[name]:
                    continue
                if any(k in lab for k in keys):
                    val = _value_right_of(fm, r, c, values)
                    if not val:
                        raw = fm.text(r, c, values) or ''
                        if ':' in raw:
                            val = raw.split(':', 1)[1]
                        val = re.sub(r'\([^)]*\)', '', val).strip()
                    out[name] = val
    return out


def compare_all(fm, values):
    """Moi o NGUOI DUNG NHAP la so + co tieu chuan -> (nhan, tc_goc, kq_mau, dat?)."""
    rows = []
    for k, v in (values or {}).items():
        val = concl_num(v)
        if val is None:
            continue
        try:
            r, c = unkey(k)
        except Exception:
            continue
        if r > fm.nrow or c > fm.ncol:
            continue
        rng = _find_std_range(fm, r, c)
        if not rng:
            continue
        lo, hi = rng
        ok = not ((lo is not None and val < lo) or (hi is not None and val > hi))
        rows.append((_result_label(fm, values, r, c), _concl_fmt_range(lo, hi), _concl_fmt(val), ok))
    return rows


def report_summary(rec):
    """Tom tat 1 bao cao thanh 1 dong cho so theo doi."""
    form = rec.get('form') or {}
    fm = FormModel(form)
    values = rec.get('values') or {}
    fields = extract_fields(fm, values)
    rows = compare_all(fm, values)
    checked, issues = suggest_conclusion(fm, values)
    verdict = '' if checked == 0 else ('KHÔNG ĐẠT' if issues else 'ĐẠT')
    std_str = '; '.join(f'{lb}: {std}' for (lb, std, val, ok) in rows)
    sample_str = '; '.join(f'{lb}: {val}' for (lb, std, val, ok) in rows)
    return {
        'title': rec.get('title') or '',
        'form_name': rec.get('form_name') or '',
        'material': fields['material'] or rec.get('material') or '',
        'supplier': fields['supplier'],
        'address': fields['address'],
        'date_recv': fields['date_recv'],
        'date_result': fields['date_result'],
        'verdict': verdict,
        'std': std_str,
        'sample': sample_str,
        'user': rec.get('user') or '',
    }


def export_log_excel(summaries, path):
    """Xuat SO THEO DOI (tong hop nhieu bao cao) ra 1 sheet Excel."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'So theo doi NVL'
    headers = ['STT', 'Tên nguyên vật liệu', 'Nhà cung cấp', 'Địa chỉ',
               'Ngày nhận mẫu', 'Ngày trả KQ', 'Đánh giá kết quả',
               'Tiêu chuẩn gốc', 'Kết quả mẫu', 'Người lập']
    widths = [5, 26, 22, 28, 15, 15, 14, 34, 30, 16]
    thin = Side(style='thin', color='BBBBBB')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill('solid', fgColor='17426B')
    head_font = Font(bold=True, color='FFFFFF', name='Calibri', size=11)
    ws.append([])  # chua tieu de lon
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(1, 1, 'SỔ THEO DÕI CHẤT LƯỢNG NGUYÊN VẬT LIỆU')
    c.font = Font(bold=True, size=14, color='17426B')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24
    ws.append(headers)
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
        hc = ws.cell(2, j)
        hc.fill = head_fill
        hc.font = head_font
        hc.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        hc.border = border
    for i, s in enumerate(summaries, 1):
        row = [i, s['material'], s['supplier'], s['address'], s['date_recv'],
               s['date_result'], s['verdict'], s['std'], s['sample'], s['user']]
        ws.append(row)
        rr = ws.max_row
        for j in range(1, len(headers) + 1):
            cell = ws.cell(rr, j)
            cell.border = border
            cell.alignment = Alignment(vertical='center',
                                       wrap_text=(j in (2, 3, 4, 8, 9)),
                                       horizontal=('center' if j in (1, 5, 6, 7) else 'left'))
            if j == 7 and s['verdict']:   # to mau cot danh gia
                if s['verdict'] == 'ĐẠT':
                    cell.fill = PatternFill('solid', fgColor='D7F0DD')
                    cell.font = Font(bold=True, color='1E7145')
                else:
                    cell.fill = PatternFill('solid', fgColor='FBD9D7')
                    cell.font = Font(bold=True, color='B0413E')
    ws.freeze_panes = 'A3'
    wb.save(path)
    return path


# ---------------------------------------------------------------- PDF
_PDF_FONTS = {}


def _pdf_fonts():
    if _PDF_FONTS:
        return _PDF_FONTS
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    W = r'C:\Windows\Fonts'
    cand = [
        (W + r'\times.ttf', W + r'\timesbd.ttf', W + r'\timesi.ttf', W + r'\timesbi.ttf'),
        (W + r'\arial.ttf', W + r'\arialbd.ttf', W + r'\ariali.ttf', W + r'\arialbi.ttf'),
        ('/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf',
         '/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf',
         '/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf',
         '/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf'),
        ('/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSerif-BoldItalic.ttf'),
        ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf'),
        ('/Library/Fonts/Times New Roman.ttf', '/Library/Fonts/Times New Roman Bold.ttf',
         '/Library/Fonts/Times New Roman Italic.ttf', '/Library/Fonts/Times New Roman Bold Italic.ttf'),
    ]
    for n, b, i, bi in cand:
        if not os.path.exists(n):
            continue
        try:
            pdfmetrics.registerFont(TTFont('RF', n))
            pdfmetrics.registerFont(TTFont('RF-B', b if os.path.exists(b) else n))
            pdfmetrics.registerFont(TTFont('RF-I', i if os.path.exists(i) else n))
            pdfmetrics.registerFont(TTFont('RF-BI', bi if os.path.exists(bi) else (b if os.path.exists(b) else n)))
            _PDF_FONTS.update({(0, 0): 'RF', (1, 0): 'RF-B', (0, 1): 'RF-I', (1, 1): 'RF-BI'})
            return _PDF_FONTS
        except Exception:
            continue
    _PDF_FONTS.update({(0, 0): 'Times-Roman', (1, 0): 'Times-Bold', (0, 1): 'Times-Italic',
                       (1, 1): 'Times-BoldItalic'})
    return _PDF_FONTS


def _rgb(h, default=(0, 0, 0)):
    try:
        h = str(h).lstrip('#')
        return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception:
        return default


def _wrap(text, font, size, width, sw):
    lines = []
    for para in str(text).split('\n'):
        words = para.split(' ')
        cur = ''
        for w in words:
            t = (cur + ' ' + w) if cur else w
            if sw(t, font, size) <= width or not cur:
                cur = t
                while sw(cur, font, size) > width and len(cur) > 1:   # tu qua dai -> cat
                    k = len(cur)
                    while k > 1 and sw(cur[:k], font, size) > width:
                        k -= 1
                    lines.append(cur[:k])
                    cur = cur[k:]
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
    return lines


def pdf_geometry(form):
    return geometry(form, px_col=7 * 0.75, px_pad=5 * 0.75, row_k=1.0)


def safe_breaks(form, images=()):
    """Tap cac hang r ma co the NGAT TRANG TRUOC hang r (khong cat ngang o gop / anh)."""
    fm = FormModel(form)
    bad = set()
    spans = [(r0, r1) for (r0, c0, r1, c1) in form.get('merges', [])]
    spans += [(im['r0'], im['r1']) for im in images]
    for (r0, r1) in spans:
        for r in range(r0 + 1, r1 + 1):
            bad.add(r)
    return [r for r in range(2, fm.nrow + 1) if r not in bad]


def paginate(form, images, k, page_h, first_row=1):
    """Chia hang thanh cac trang: [(r_from, r_to), ...] - moi trang cao toi da page_h (pt)."""
    fm = FormModel(form)
    xs, ys = pdf_geometry(form)
    ok = set(safe_breaks(form, images))
    pages = []
    start = first_row
    n = fm.nrow
    img_rows = set()
    for im in images:
        img_rows.update(range(im['r0'], im['r1'] + 1))

    def blank(r):
        """Hang trong han: khong chu, khong vien, khong mau nen, khong anh, khong o gop."""
        if r in img_rows:
            return False
        for cc in range(1, fm.ncol + 1):
            if (r, cc) in fm.covered or (r, cc) in fm.span:
                return False
            st = fm.cells.get((r, cc)) or {}
            if st.get('fill') or any(st.get('bd') or ()) or fm.text(r, cc):
                return False
        # vien duoi cua hang tren cung tinh la vien tren cua hang nay
        for cc in range(1, fm.ncol + 1):
            st = fm.cells.get((r - 1, cc)) or {}
            if (st.get('bd') or (0, 0, 0, 0))[3]:
                return False
        return True
    while start <= n:
        if pages:                                   # trang tiep: bo hang trong o dau trang
            while start < n and blank(start):
                start += 1
        end = start
        last_ok = None
        r = start
        while r <= n and (ys[r] - ys[start - 1]) * k <= page_h + 0.01:
            end = r
            if (r + 1) in ok:
                last_ok = r
            r += 1
        if end >= n:
            pages.append((start, n))
            break
        if last_ok is not None and last_ok >= start:
            end = last_ok
        pages.append((start, end))
        start = end + 1
    return pages


def choose_scale(form, images, k_full, page_h, first_row=1, max_shrink=0.12):
    """Ti le vua chieu ngang; neu thu nho <= 12% ma bot duoc 1 trang (VD trang cuoi chi con
    phan ky ten) thi dung ti le lon nhat bot duoc trang do."""
    n0 = len(paginate(form, images, k_full, page_h, first_row))
    if n0 <= 1:
        return k_full, n0
    k = k_full
    while k >= k_full * (1 - max_shrink):
        n = len(paginate(form, images, k, page_h, first_row))
        if n < n0:
            return k, n
        k -= k_full * 0.01
    return k_full, n0


def draw_report_pdf(c, rec, x0, y0, avail_w, avail_h, k=None, r_from=1, r_to=None):
    """Ve 1 bao cao len canvas reportlab trong khung (x0,y0 = goc TREN-TRAI, toa do PDF).
    k = ti le (None = vua khung, khong phong to); r_from..r_to = chi ve cac hang nay (ngat trang)."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.lib.utils import ImageReader
    form = rec['form']
    fm = FormModel(form)
    fonts = _pdf_fonts()
    xs, ys = pdf_geometry(form)
    W, H = xs[-1], ys[-1]
    r_to = fm.nrow if r_to is None else r_to
    if k is None:
        k = min(avail_w / W, avail_h / H, 1.0)
    ox = x0 + (avail_w - W * k) / 2
    y_top = ys[r_from - 1]

    def X(cc):
        return ox + xs[min(max(cc, 1), len(xs)) - 1] * k

    def Y(rr):
        return y0 - (ys[min(max(rr, 1), len(ys)) - 1] - y_top) * k

    def inside(r):
        return r_from <= r <= r_to
    values = rec.get('values') or {}
    # nen
    for (r, cc), cell in fm.cells.items():
        if (r, cc) in fm.covered or not cell.get('fill') or not inside(r):
            continue
        r0, c0, r1, c1 = fm.area(r, cc)
        c.setFillColorRGB(*_rgb(cell['fill'], (1, 1, 1)))
        c.rect(X(c0), Y(r1 + 1), X(c1 + 1) - X(c0), Y(r0) - Y(r1 + 1), stroke=0, fill=1)
    # anh
    for im in all_images(rec):
        if not inside(im['r0']):
            continue
        p = resolve_image(im.get('file'))
        if not p:
            continue
        try:
            ir = ImageReader(p)
            iw, ih = ir.getSize()
            bx, by = X(im['c0']), Y(im['r1'] + 1)
            bw, bh = X(im['c1'] + 1) - bx, Y(im['r0']) - by
            pad = 2
            w, h, dx, dy = fit_box(iw, ih, bw - 2 * pad, bh - 2 * pad)
            c.drawImage(ir, bx + pad + dx, by + pad + dy, w, h, mask='auto')
        except Exception:
            continue
    # chu
    for r in range(r_from, r_to + 1):
        for cc in range(1, fm.ncol + 1):
            if (r, cc) in fm.covered:
                continue
            t = fm.text(r, cc, values)
            if not t:
                continue
            st = fm.style(r, cc)
            r0, c0, r1, c1 = fm.area(r, cc)
            fn = fonts[(1 if st.get('b') else 0, 1 if st.get('i') else 0)]
            size = float(st.get('sz') or 11) * k
            bx0, bx1 = X(c0) + 2 * k + 1, X(c1 + 1) - 2 * k - 1
            top, bot = Y(r0), Y(r1 + 1)
            single = (r0 == r1 and c0 == c1)
            if not st.get('wrap', 1) and single:
                lines = t.split('\n')                # khong xuong dong -> chu tran sang o ben (giong Excel)
            else:
                lines = _wrap(t, fn, size, max(4, bx1 - bx0), stringWidth)
            lh = size * 1.18
            total = lh * len(lines)
            while total > (top - bot) and size > 4:           # chu khong vua o -> thu nho
                size *= 0.92
                lh = size * 1.18
                if st.get('wrap', 1) or not single:
                    lines = _wrap(t, fn, size, max(4, bx1 - bx0), stringWidth)
                total = lh * len(lines)
            va = st.get('va') or 'bottom'
            if va == 'top':
                yy = top - size
            elif va == 'center':
                yy = (top + bot) / 2 + total / 2 - size
            else:
                yy = bot + total - size + 1
            al = st.get('al') or ''
            if not al:
                al = 'right' if re.fullmatch(r'-?[\d.,]+%?', t.strip()) else 'left'
            c.setFillColorRGB(*_rgb(st.get('color') or '#000000'))
            c.setFont(fn, size)
            for ln in lines:
                if al in ('center', 'centerContinuous', 'justify', 'distributed'):
                    c.drawCentredString((bx0 + bx1) / 2, yy, ln)
                elif al == 'right':
                    c.drawRightString(bx1, yy, ln)
                else:
                    c.drawString(bx0, yy, ln)
                yy -= lh
    # vien
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.6)
    for (r, cc, side) in fm.border_lines():
        if not inside(r):
            continue
        xa, xb, ya, yb = X(cc), X(cc + 1), Y(r), Y(r + 1)
        if side == 0:
            c.line(xa, ya, xa, yb)
        elif side == 1:
            c.line(xa, ya, xb, ya)
        elif side == 2:
            c.line(xb, ya, xb, yb)
        else:
            c.line(xa, yb, xb, yb)


# ---------------------------------------------------------------- GIAY TIEU DE (lay tu TDS)
def letterhead_on():
    """Bat/tat giay tieu de HITECH PLAST (header + footer cua TDS) khi xuat BCCLNVL.
    Tat: them "letterhead": false vao rmqr_settings.json."""
    return bool(load_settings().get('letterhead', True))


def letterhead_spec():
    """(bg, box) cua giay tieu de - dung chung anh doc_bg_tds.png voi TDS (doc_params 'rmqr')."""
    if not letterhead_on():
        return None
    try:
        import doc_params
        spec = doc_params.page_spec('rmqr')
    except Exception:
        spec = None
    if not spec:
        return None
    bg = []
    for (f, bx, by, bw, bh) in spec.get('bg') or []:
        for p in (os.path.join(resource_dir(), f), os.path.join(writable_dir(), f)):
            if os.path.exists(p):
                bg.append((p, bx, by, bw, bh))
                break
    return (bg, spec['box']) if bg else None


def _drop_blank_row1(rec):
    """Hang 1 cua form (cao 81pt) la cho trong de header cu -> co giay tieu de thi an di."""
    form = rec.get('form') or {}
    if not form.get('rowh'):
        return rec
    values = rec.get('values') or {}
    fm = FormModel(form)
    if any(fm.text(1, cc, values) for cc in range(1, fm.ncol + 1)):
        return rec
    if any(im.get('r0') == 1 for im in all_images(rec)):
        return rec
    rec = dict(rec)
    rec['form'] = dict(form)
    rec['form']['rowh'] = [0.001] + list(form['rowh'][1:])
    return rec


def draw_report_pages(c, rec, page_w, page_h, landscape=False):
    """Ve BCCLNVL: phong to vua CHIEU NGANG trang, phan con thua sang trang sau.
    Moi trang deu co giay tieu de TDS (neu bat)."""
    lh = None if landscape else letterhead_spec()
    if lh:
        bg, (bl, bt, br, bb) = lh
        rec = _drop_blank_row1(rec)
    else:
        bg, m = [], 16
        bl, bt, br, bb = m, m, page_w - m, page_h - m
    form = rec['form']
    xs, ys = pdf_geometry(form)
    aw, ah = br - bl, bb - bt
    first = 2 if (form.get('rowh') and form['rowh'][0] < 0.01) else 1
    k, _n = choose_scale(form, all_images(rec), aw / xs[-1], ah, first)
    pages = paginate(form, all_images(rec), k, ah, first)
    for i, (r0, r1) in enumerate(pages):
        if i:
            c.showPage()
        for (p, bx, by, bw, bh) in bg:
            try:
                c.drawImage(p, bx, page_h - by - bh, bw, bh, mask='auto')
            except Exception:
                pass
        draw_report_pdf(c, rec, bl, page_h - bt, aw, ah, k=k, r_from=r0, r_to=r1)
    return len(pages)


def export_pdf(rec, path, landscape=False):
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4, landscape as _ls
    size = _ls(A4) if landscape else A4
    c = canvas.Canvas(path, pagesize=size)
    c.setTitle(rec.get('title') or rec.get('form_name') or 'Report')
    draw_report_pages(c, rec, size[0], size[1], landscape)
    c.showPage()
    c.save()
    return path


def export_pdf_many(recs, path):
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    c = canvas.Canvas(path, pagesize=A4)
    for rec in recs:
        draw_report_pages(c, rec, A4[0], A4[1])
        c.showPage()
    c.save()
    return path


def _excel_fit_width(ws, rec, fm):
    """Excel: phong to vua CHIEU NGANG trang, phan thua sang trang 2 (ngat trang dung cho,
    khong cat ngang o gop / anh) - giong PDF."""
    import math
    from openpyxl.worksheet.pagebreak import Break, RowBreak
    import doc_params as dp
    pm = ws.page_margins
    aw = dp.PAGE_W - (pm.left + pm.right) * 72.0
    ah = dp.PAGE_H - (pm.top + pm.bottom) * 72.0
    xs, _ys = pdf_geometry(rec['form'])
    first = int(str(ws.print_area).split('!')[-1].split(':')[0].replace('$', '').lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ') or 1)
    imgs = all_images(rec)
    k, _n = choose_scale(rec['form'], imgs, aw / xs[-1], ah * 0.985, first)
    scale = max(10, min(400, int(math.floor(k * 100)) - 1))
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.page_setup.scale = scale
    pages = paginate(rec['form'], imgs, scale / 100.0, ah * 0.985, first)
    rb = RowBreak()
    for (_a, b) in pages[:-1]:
        rb.append(Break(id=b))
    ws.row_breaks = rb


# ---------------------------------------------------------------- EXCEL
def export_excel(rec, path, sheet_title=None):
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter as L
    from openpyxl.worksheet.datavalidation import DataValidation
    form = rec['form']
    fm = FormModel(form)
    wb = openpyxl.Workbook()
    ws = wb.active
    t = re.sub(r'[\\/*?:\[\]]', ' ', sheet_title or rec.get('title') or form.get('name') or 'Report')[:31]
    ws.title = t.strip() or 'Report'
    fname = form.get('font') or 'Times New Roman'
    for i, w in enumerate(form.get('cols', []), 1):
        ws.column_dimensions[L(i)].width = w
    for i, h in enumerate(form.get('rowh', []), 1):
        ws.row_dimensions[i].height = h
    thin = Side(style='thin', color='000000')
    values = rec.get('values') or {}
    for (r, cc), st in fm.cells.items():
        cell = ws.cell(r, cc)
        bd = st.get('bd') or [0, 0, 0, 0]
        cell.border = Border(left=thin if bd[0] else Side(), top=thin if bd[1] else Side(),
                             right=thin if bd[2] else Side(), bottom=thin if bd[3] else Side())
        if (r, cc) in fm.covered:
            continue
        cell.font = Font(name=fname, size=st.get('sz') or 11, bold=bool(st.get('b')), italic=bool(st.get('i')),
                         color=(st.get('color') or '#000000').lstrip('#'))
        cell.alignment = Alignment(horizontal=st.get('al') or None, vertical=st.get('va') or 'bottom',
                                   wrap_text=bool(st.get('wrap', 1)))
        if st.get('fill'):
            cell.fill = PatternFill('solid', fgColor=st['fill'].lstrip('#'))
    for r in range(1, fm.nrow + 1):
        for cc in range(1, fm.ncol + 1):
            if (r, cc) in fm.covered:
                continue
            txt = fm.text(r, cc, values)
            if txt == '':
                continue
            v = txt
            if re.fullmatch(r'-?\d+(\.\d+)?', txt.strip()):
                v = float(txt) if '.' in txt else int(txt)
            ws.cell(r, cc).value = v
    for (r0, c0, r1, c1) in form.get('merges', []):
        if r1 <= fm.nrow and c1 <= fm.ncol:
            ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
    for k_, items in (form.get('lists') or {}).items():
        r, cc = unkey(k_)
        if r > fm.nrow:
            continue
        f1 = '"' + ','.join(items) + '"'
        if len(f1) <= 255:
            dv = DataValidation(type='list', formula1=f1, allow_blank=True, showErrorMessage=False)
            ws.add_data_validation(dv)
            dv.add(f'{L(cc)}{r}')
    # anh
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU
    from PIL import Image
    xs, ys = geometry(form, px_col=7, px_pad=5, row_k=96 / 72)
    for im in all_images(rec):
        p = resolve_image(im.get('file'))
        if not p:
            continue
        try:
            pil = Image.open(p)
            iw, ih = pil.size
            bw = xs[im['c1']] - xs[im['c0'] - 1] - 6
            bh = ys[im['r1']] - ys[im['r0'] - 1] - 6
            w, h, dx, dy = fit_box(iw, ih, bw, bh)
            xi = XLImage(p)
            xi.width, xi.height = int(w), int(h)
            # vi tri: dat tai o goc + lech (dx,dy) de can giua khung
            col0, x_off = im['c0'] - 1, dx + 3
            while col0 + 1 < len(xs) - 1 and xs[col0 + 1] - xs[im['c0'] - 1] <= x_off:
                col0 += 1
            x_off -= xs[col0] - xs[im['c0'] - 1]
            row0, y_off = im['r0'] - 1, dy + 3
            while row0 + 1 < len(ys) - 1 and ys[row0 + 1] - ys[im['r0'] - 1] <= y_off:
                row0 += 1
            y_off -= ys[row0] - ys[im['r0'] - 1]
            marker = AnchorMarker(col=col0, colOff=pixels_to_EMU(int(x_off)), row=row0,
                                  rowOff=pixels_to_EMU(int(y_off)))
            xi.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(int(w)),
                                                                          pixels_to_EMU(int(h))))
            ws.add_image(xi)
        except Exception:
            continue
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = ws.page_margins.right = 0.2
    ws.page_margins.top = ws.page_margins.bottom = 0.3
    ws.print_area = f'A1:{L(fm.ncol)}{fm.nrow}'
    ws.sheet_view.showGridLines = False
    # giay tieu de TDS -> Header giua (&G) cua Excel, le tren/duoi chua header/footer
    hf = None
    lh = letterhead_spec()
    if lh:
        try:
            import doc_tools
            _bg, (bl, bt, br, bb) = lh
            import doc_params as dp
            tmpdir = os.path.join(storage_dir(), 'pdf')
            png = os.path.join(tmpdir, '_letterhead_%s.png' % uuid.uuid4().hex[:6])
            got = doc_tools.compose_page_bg('rmqr', png)
            if got:
                _p, bx, by, bw, bh = got
                ws.oddHeader.center.text = '&G'
                ws.page_margins.header = round(by / 72.0, 3)
                ws.page_margins.top = round(bt / 72.0, 3)
                ws.page_margins.bottom = round((dp.PAGE_H - bb) / 72.0, 3)
                ws.page_margins.footer = 0.2
                ws.page_margins.left = round(bl / 72.0, 3)
                ws.page_margins.right = round((dp.PAGE_W - br) / 72.0, 3)
                r0 = 2 if _drop_blank_row1(rec) is not rec else 1
                ws.print_area = f'A{r0}:{L(fm.ncol)}{fm.nrow}'
                # anh header KHONG co gian theo trang (neu khong se bi thu nho cung noi dung)
                ws.HeaderFooter.scaleWithDoc = False
                ws.HeaderFooter.alignWithMargins = False
                hf = (png, bw, bh)
        except Exception:
            hf = None
    _excel_fit_width(ws, rec, fm)
    wb.save(path)
    if hf:
        try:
            import doc_tools
            doc_tools._inject_header_image(path, hf[0], 'C', hf[1], hf[2])
        finally:
            try:
                os.remove(hf[0])
            except Exception:
                pass
    return path


# ---------------------------------------------------------------- NHAP FORM TU EXCEL
def form_from_excel(path, sheet, rng, name, group):
    """Doc 1 vung (VD 'A1:H39') cua sheet Excel (.xlsx/.xlsm) -> form moi (kem anh trong vung)."""
    import openpyxl
    from openpyxl.utils import column_index_from_string as CI
    m = re.fullmatch(r'\s*([A-Za-z]{1,3})(\d+)\s*:\s*([A-Za-z]{1,3})(\d+)\s*', rng or '')
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    if m:
        c0, r0, c1, r1 = CI(m.group(1).upper()), int(m.group(2)), CI(m.group(3).upper()), int(m.group(4))
    else:
        c0, r0, c1, r1 = 1, 1, ws.max_column, ws.max_row
    ncol, nrow = c1 - c0 + 1, r1 - r0 + 1
    from openpyxl.utils import get_column_letter as L

    def col_hex(cobj):
        try:
            if cobj is not None and cobj.type == 'rgb' and isinstance(cobj.rgb, str):
                return '#' + cobj.rgb[-6:].upper()
        except Exception:
            pass
        return ''
    cells = []
    font_name = None
    for r in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            cell = ws.cell(r, cc)
            v = cell.value
            f, a, b = cell.font, cell.alignment, cell.border
            fill = ''
            try:
                if cell.fill and cell.fill.fill_type == 'solid':
                    fill = col_hex(cell.fill.fgColor)
                    if fill == '#FFFFFF':
                        fill = ''
            except Exception:
                pass
            bd = [1 if (s is not None and s.style) else 0 for s in (b.left, b.top, b.right, b.bottom)]
            if v is None and not fill and not any(bd):
                continue
            d = {'r': r - r0 + 1, 'c': cc - c0 + 1, 'b': 1 if f.b else 0, 'i': 1 if f.i else 0,
                 'sz': float(f.sz or 11), 'color': col_hex(f.color) or '#000000', 'al': a.horizontal or '',
                 'va': a.vertical or 'bottom', 'wrap': 1 if a.wrap_text else 0, 'fill': fill, 'bd': bd}
            font_name = font_name or f.name
            if isinstance(v, str) and v.startswith('='):
                mm = re.fullmatch(r'=\$?([A-Z]{1,3})\$?(\d+)', v.strip())
                if mm and c0 <= CI(mm.group(1)) <= c1 and r0 <= int(mm.group(2)) <= r1:
                    d['ref'] = [int(mm.group(2)) - r0 + 1, CI(mm.group(1)) - c0 + 1]
                v = None
            if isinstance(v, dt.datetime):
                v = v.strftime('%d/%m/%Y')
            elif isinstance(v, dt.date):
                v = v.strftime('%d/%m/%Y')
            if v is not None:
                d['v'] = v if isinstance(v, (int, float)) else str(v)
            if cell.number_format and cell.number_format != 'General':
                d['nf'] = cell.number_format
            cells.append(d)
    merges = []
    for mg in ws.merged_cells.ranges:
        if mg.min_row >= r0 and mg.max_row <= r1 and mg.min_col >= c0 and mg.max_col <= c1:
            merges.append([mg.min_row - r0 + 1, mg.min_col - c0 + 1, mg.max_row - r0 + 1, mg.max_col - c0 + 1])
    images = []
    for xim in getattr(ws, '_images', []):
        try:
            a = xim.anchor
            fr = a._from
            ir0, ic0 = fr.row + 1, fr.col + 1
            to = getattr(a, 'to', None)
            ir1, ic1 = (to.row + 1, to.col + 1) if to is not None else (ir0, ic0)
            if not (r0 <= ir0 <= r1 and c0 <= ic0 <= c1):
                continue
            from PIL import Image
            data = xim._data()
            pil = Image.open(io.BytesIO(data))
            pil.load()
            fn = import_image(pil, sub='form_media', max_side=1100)
            images.append({'r0': ir0 - r0 + 1, 'c0': ic0 - c0 + 1, 'r1': min(ir1, r1) - r0 + 1,
                           'c1': min(ic1, c1) - c0 + 1, 'file': fn})
        except Exception:
            continue
    lists = {}
    for dv in ws.data_validations.dataValidation:
        if dv.type != 'list' or not dv.formula1 or dv.formula1.startswith('='):
            continue
        items = [t.strip() for t in dv.formula1.strip('"').split(',') if t.strip()]
        for part in str(dv.sqref).split():
            mm = re.fullmatch(r'([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?', part)
            if not mm:
                continue
            a0, b0 = CI(mm.group(1)), int(mm.group(2))
            a1 = CI(mm.group(3)) if mm.group(3) else a0
            b1 = int(mm.group(4)) if mm.group(4) else b0
            for rr in range(max(b0, r0), min(b1, r1) + 1):
                for cc in range(max(a0, c0), min(a1, c1) + 1):
                    lists[key(rr - r0 + 1, cc - c0 + 1)] = items
    return {'id': '', 'name': name, 'group': group, 'ncol': ncol, 'nrow': nrow,
            'cols': [round(ws.column_dimensions[L(c)].width or 8.43, 2) for c in range(c0, c1 + 1)],
            'rowh': [round(ws.row_dimensions[r].height or 15.0, 2) for r in range(r0, r1 + 1)],
            'merges': merges, 'cells': cells, 'images': images, 'lists': lists,
            'font': font_name or 'Times New Roman', 'name_cell': None, 'date_cell': None, 'concl_cell': None,
            'note': 'Nhập từ %s!%s' % (os.path.basename(path), rng or '')}
