# -*- coding: utf-8 -*-
"""
tem_ui.py - Tab "🏷 Tem BTW": nhan du lieu tem tu ANH / EXCEL / WORD / dan chu -> xuat file tem BarTender .btw
giong het mau (mac dinh tem_mau_AA06.btw - tem bao FILLER CALCIUM CARBONATE 50x60).

Luu danh sach tem dang lam trong tem_btw.json (canh phan mem). File xuat: TemBTW\\<nam-thang>\\
"""
import os
import sys
import json
import threading
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import btw_core as B

FONT = ('Segoe UI', 10)
FONT_B = ('Segoe UI', 10, 'bold')
IMG_EXT = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp')
DATA_EXT = ('.xlsx', '.xlsm', '.xls', '.csv', '.docx', '.docm', '.doc', '.txt')
DEFAULT_TPL = 'tem_mau_AA06.btw'


def _app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _resource(name):
    p = os.path.join(_app_dir(), name)
    if os.path.exists(p):
        return p
    return os.path.join(getattr(sys, '_MEIPASS', _app_dir()), name)


def _open_path(p):
    try:
        if sys.platform.startswith('win'):
            os.startfile(p)
        elif sys.platform == 'darwin':
            import subprocess
            subprocess.Popen(['open', p])
        else:
            import subprocess
            subprocess.Popen(['xdg-open', p])
    except Exception as ex:
        messagebox.showwarning('Mở file', str(ex))


_THUMB_CACHE = {}          # (duong dan, mtime) -> ket qua B.analyze_template


class FillTab(tk.Frame):
    """Dien du lieu vao mau tem .btw bat ky (mac dinh AA06) -> xuat .btw"""

    def __init__(self, master, app=None, base_dir=None):
        super().__init__(master, bg='#f4f6f9')
        self.app = app
        self.base = base_dir or _app_dir()
        self.cfg_path = os.path.join(self.base, 'tem_btw.json')
        self.cfg = self._load_cfg()
        self.records = self.cfg.get('records') or []
        self.cur = None                 # chi so tem dang sua
        self.vars = {}
        self._busy = False
        self._prev_job = None
        self._save_job = None
        self._tk_prev = None
        self.last_out = None
        self._build()
        self.load_template(self.cfg.get('template') or '', quiet=True)

    # ------------------------------------------------------------ cau hinh
    def _load_cfg(self):
        try:
            with open(self.cfg_path, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_cfg(self, later=False):
        if later:
            if self._save_job:
                self.after_cancel(self._save_job)
            self._save_job = self.after(800, self._save_cfg)
            return
        self._save_job = None
        self.cfg['records'] = self.records
        self.cfg['batch_mode'] = self.v_mode.get()
        self.cfg['outdir'] = self._store_out(self.v_out.get())
        self.cfg['months'] = self.v_months.get()
        try:
            tmp = self.cfg_path + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=1)
            os.replace(tmp, self.cfg_path)
            return True
        except Exception as ex:
            self._st('Không lưu được tem_btw.json: %s' % ex, True)
            return False

    # ------------------------------------------------------------ giao dien
    def _build(self):
        r1 = tk.Frame(self, bg='#f4f6f9')
        r1.pack(fill='x', padx=8, pady=(6, 2))
        tk.Label(r1, text='Mẫu tem:', bg='#f4f6f9', font=FONT_B).pack(side='left')
        self.lbl_tpl = tk.Label(r1, text='', bg='#f4f6f9', fg='#1f3a5f', font=FONT, anchor='w')
        self.lbl_tpl.pack(side='left', padx=6)
        ttk.Button(r1, text='↺ Mẫu gốc AA06', command=lambda: self.load_template('')).pack(side='right')
        ttk.Button(r1, text='📂 Chọn mẫu .btw khác…', command=self.pick_template).pack(side='right', padx=4)

        # --- vung nhan du lieu
        drop = tk.Frame(self, bg='#e3f2fd', highlightthickness=2, highlightbackground='#64b5f6')
        drop.pack(fill='x', padx=8, pady=4)
        self.drop_lbl = tk.Label(drop, text='📥  KÉO THẢ ảnh tem / file Excel / Word vào đây   ·   hoặc Ctrl+V dán ảnh '
                                            '(Win+Shift+S) / bảng Excel đã Copy',
                                 bg='#e3f2fd', fg='#0d47a1', font=FONT_B, pady=8)
        self.drop_lbl.pack(side='left', padx=8)
        bb = tk.Frame(drop, bg='#e3f2fd')
        bb.pack(side='right', padx=6)
        for t, c in (('📷 Ảnh…', lambda: self.pick_file('img')), ('📊 Excel…', lambda: self.pick_file('xls')),
                     ('📝 Word…', lambda: self.pick_file('doc')), ('📋 Dán (Ctrl+V)', self.paste)):
            ttk.Button(bb, text=t, command=c).pack(side='left', padx=2, pady=4)
        try:
            from recipe_ui import enable_file_drop
            for w in (drop, self.drop_lbl):
                enable_file_drop(w, self.on_drop)
        except Exception:
            pass
        self.bind_all('<Control-v>', self._ctrl_v, add='+')
        self.bind_all('<Control-V>', self._ctrl_v, add='+')

        body = tk.PanedWindow(self, orient='horizontal', sashwidth=6, bg='#cfd8e3')
        body.pack(fill='both', expand=True, padx=8, pady=4)

        # --- trai: danh sach tem
        lf = ttk.LabelFrame(body, text=' Danh sách tem ', padding=4)
        cols = (('n', '#', 32), ('grade', 'Grade', 90), ('batch', 'Batch', 120), ('mfg', 'NSX', 70),
                ('src', 'Nguồn', 110))
        self.tv = ttk.Treeview(lf, columns=[c for c, _, _ in cols], show='headings', selectmode='extended')
        for c, h, w in cols:
            self.tv.heading(c, text=h)
            self.tv.column(c, width=w, anchor='center' if c in ('n', 'mfg') else 'w')
        vs = ttk.Scrollbar(lf, orient='vertical', command=self.tv.yview)
        self.tv.configure(yscrollcommand=vs.set)
        lb = ttk.Frame(lf)
        lb.pack(side='bottom', fill='x', pady=(4, 0))
        ttk.Button(lb, text='➕ Tem mới', command=self.new_rec).pack(side='left')
        ttk.Button(lb, text='⧉ Nhân bản', command=self.dup_rec).pack(side='left', padx=3)
        ttk.Button(lb, text='🗑 Xoá', command=self.del_rec).pack(side='left')
        self.tv.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self.tv.bind('<<TreeviewSelect>>', lambda e: self._on_select())
        self.tv.bind('<Delete>', lambda e: self.del_rec())
        body.add(lf, minsize=300, width=440)

        # --- giua: form sua
        self.ff = ttk.LabelFrame(body, text=' Nội dung tem (sửa trực tiếp) ', padding=8)
        body.add(self.ff, minsize=380, width=500)

        # --- phai: xem truoc
        pf = ttk.LabelFrame(body, text=' Xem trước ', padding=4)
        self.cv = tk.Canvas(pf, bg='#eef1f5', highlightthickness=0, width=380, height=380)
        self.cv.pack(fill='both', expand=True)
        self.cv.bind('<Configure>', lambda e: self._schedule_preview())
        body.add(pf, minsize=260)

        # --- duoi: xuat
        bt = tk.Frame(self, bg='#f4f6f9')
        bt.pack(fill='x', padx=8, pady=(2, 2))
        tk.Label(bt, text='Lưu vào:', bg='#f4f6f9', font=FONT).pack(side='left')
        self.v_out = tk.StringVar(value=self._load_out(self.cfg.get('outdir')))
        ttk.Entry(bt, textvariable=self.v_out, width=46).pack(side='left', padx=4)
        ttk.Button(bt, text='…', width=3, command=self.pick_out).pack(side='left')
        ttk.Button(bt, text='📂 Mở thư mục', command=lambda: _open_path(self._outdir())).pack(side='left', padx=4)
        tk.Button(bt, text='📦 Xuất các tem đã chọn', bg='#455a64', fg='white', relief='flat', font=FONT_B,
                  padx=10, command=lambda: self.export(all_sel=True)).pack(side='right', padx=(4, 0))
        tk.Button(bt, text='💾 Xuất .btw', bg='#0068ff', fg='white', relief='flat', font=FONT_B, padx=14,
                  command=self.export).pack(side='right', padx=4)
        ttk.Button(bt, text='🖼 Lưu ảnh PNG', command=self.save_png).pack(side='right')
        tk.Button(bt, text='🖨 Khổ giấy A3/A4/A5', bg='#6a1b9a', fg='white', relief='flat', font=FONT_B, padx=10,
                  command=self.export_paper).pack(side='right', padx=4)
        tk.Button(bt, text='👁 Xem trước tem', bg='#ffb300', fg='#222', relief='flat', font=FONT_B, padx=10,
                  command=self.preview_big).pack(side='right', padx=4)
        tk.Button(bt, text='💾 Lưu tem', bg='#2e7d32', fg='white', relief='flat', font=FONT_B, padx=10,
                  command=self.save_now).pack(side='right', padx=4)
        self.btn_bt = ttk.Button(bt, text='🏷 Mở bằng BarTender', command=self.open_last, state='disabled')
        self.btn_bt.pack(side='right', padx=4)
        self.status = tk.Label(self, text='', bg='#f4f6f9', fg='#1f3a5f', font=FONT, anchor='w')
        self.status.pack(fill='x', padx=10, pady=(0, 6))

        self.v_mode = tk.StringVar(value=self.cfg.get('batch_mode') or 'data')
        self.v_months = tk.StringVar(value=str(self.cfg.get('months') or 12))

    def _st(self, s, err=False):
        try:
            self.status.configure(text=s, fg='#b00020' if err else '#1f3a5f')
        except tk.TclError:
            pass

    # ------------------------------------------------------------ mau tem
    def pick_template(self):
        p = filedialog.askopenfilename(parent=self, title='Chọn file tem BarTender mẫu',
                                       filetypes=[('BarTender', '*.btw'), ('Tất cả', '*.*')])
        if p:
            self.load_template(p)

    def load_template(self, path, quiet=False):
        path = path or _resource(DEFAULT_TPL)
        try:
            self.btw = B.BtwFile(path)
            self.layout = B.Layout(self.btw)
        except Exception as ex:
            if not quiet:
                messagebox.showerror('Mẫu tem', 'Không đọc được mẫu tem:\n%s\n\n%s' % (path, ex), parent=self)
            if path != _resource(DEFAULT_TPL):
                return self.load_template('', quiet=True)
            self.btw = self.layout = None
            self.lbl_tpl.configure(text='(thiếu file %s cạnh phần mềm)' % DEFAULT_TPL, fg='#b00020')
            return
        self.tpl_path = path
        self.cfg['template'] = '' if os.path.abspath(path) == os.path.abspath(_resource(DEFAULT_TPL)) else path
        self.defaults = self.layout.defaults()
        info = '%s   ·   %s   ·   máy in: %s' % (os.path.basename(path), self.btw.meta('TemplateSize') or '?',
                                                self.btw.meta('Printer') or '?')
        if not self.layout.known:
            info += '   (đang nhận diện các dòng / ô trên tem…)'
        self.lbl_tpl.configure(text=info, fg='#1f3a5f')
        self._tpl_info = info.split('   (')[0]
        if not self.layout.known:
            self._analyze_async(path)
        if not self.records:
            self.records = [dict(self.defaults, _src='Mẫu gốc')]
        self._build_form()
        self._fill_tree(select=0)
        self._save_cfg(later=True)

    def _analyze_async(self, path):
        """Mau la: doc anh xem truoc trong file .btw -> vi tri tung o (de hien dong ke + nhom theo dong)."""
        try:
            key = (os.path.abspath(path), os.path.getmtime(path))
        except OSError:
            key = (path, 0)
        lay = self.layout
        if key in _THUMB_CACHE:
            self._analyze_done(lay, _THUMB_CACHE[key])
            return
        res = {}

        def work():
            try:
                res['v'] = B.analyze_template(lay.btw)
            except Exception as ex:
                res['err'] = ex

        th = threading.Thread(target=work, daemon=True)
        th.start()

        def poll():
            if th.is_alive():
                self.after(150, poll)
                return
            v = res.get('v')
            if v is not None:
                _THUMB_CACHE[key] = v
            self._analyze_done(lay, v)
        self.after(150, poll)

    def _analyze_done(self, lay, th):
        if lay is not self.layout:              # da doi sang mau khac trong luc doc
            return
        lay.thumb = th
        if th:
            n = len(th['rows'])
            self.lbl_tpl.configure(text=self._tpl_info + '   (nhận diện %d dòng – sửa trực tiếp từng ô)' % n)
        else:
            self.lbl_tpl.configure(text=self._tpl_info + '   (chưa nhận diện được khung – cài OCR: '
                                                         'python -m pip install rapidocr onnxruntime)')
        cur = self.cur
        self._build_form()
        if cur is not None and cur < len(self.records):
            self._load_form(self.records[cur])
        self._schedule_preview()

    # ------------------------------------------------------------ form
    def _build_form_rows(self, f, th):
        """Mau la da nhan dien: moi DONG tem = 1 hang (o trai = nhan, o phai = gia tri)."""
        names = [k for k, _, _, _ in self.layout.fields]
        caps = {k: c for k, c, _, _ in self.layout.fields}
        ttk.Label(f, text='Ô trái / nhãn', foreground='#666').grid(row=0, column=1, sticky='w')
        ttk.Label(f, text='Ô phải / giá trị', foreground='#666').grid(row=0, column=2, sticky='w')
        r = 1
        shown = set()

        def entry(key, col, span=1, bold=False):
            v = tk.StringVar()
            self.vars[key] = v
            v.trace_add('write', lambda *a, k=key: self._changed(k))
            e = ttk.Entry(f, textvariable=v, font=FONT_B if bold else FONT)
            e.grid(row=r, column=col, columnspan=span, sticky='we', pady=3, padx=(0, 4))
            shown.add(key)
            return e

        for i, row in enumerate(th['rows']):
            row = [n for n in row if n in caps]
            if not row:
                continue
            ttk.Label(f, text='Dòng %d:' % (i + 1), font=FONT_B).grid(row=r, column=0, sticky='w', padx=(0, 6))
            if len(row) == 1:
                entry(row[0], 1, 2, True)
            else:
                entry(row[0], 1)
                entry(row[1], 2, 1, True)
                for extra in row[2:]:          # dong co > 2 o: them hang phu
                    r += 1
                    ttk.Label(f, text='   ↳ %s:' % caps[extra], foreground='#555').grid(row=r, column=0, sticky='w')
                    entry(extra, 1, 2)
            r += 1
        rest = [n for n in names if n not in shown]
        if rest:
            ttk.Label(f, text='Ô chữ khác (không thấy trên ảnh mẫu):', foreground='#8a5a00').grid(
                row=r, column=0, columnspan=3, sticky='w', pady=(8, 2))
            r += 1
            for n in rest:
                ttk.Label(f, text=caps[n] + ':', font=FONT_B).grid(row=r, column=0, sticky='w', padx=(0, 6))
                entry(n, 1, 2)
                r += 1

    def _build_form(self):
        for w in self.ff.winfo_children():
            w.destroy()
        self.vars = {}
        f = self.ff
        f.columnconfigure(1, weight=1)
        f.columnconfigure(2, weight=2)
        th = getattr(self.layout, 'thumb', None)
        if not self.layout.known and th and th.get('rows'):
            self._build_form_rows(f, th)
            return
        r = 0
        ttk.Label(f, text='Nhãn (cột trái)', foreground='#666').grid(row=r, column=1, sticky='w')
        ttk.Label(f, text='Giá trị (cột phải)', foreground='#666').grid(row=r, column=2, sticky='w')
        r += 1
        no = 0
        for key, cap, lo, vo in self.layout.fields:
            v = tk.StringVar()
            self.vars[key] = v
            v.trace_add('write', lambda *a, k=key: self._changed(k))
            if lo is None:
                ttk.Label(f, text=cap + ':', font=FONT_B).grid(row=r, column=0, sticky='w', pady=3, padx=(0, 6))
                ttk.Entry(f, textvariable=v, font=FONT).grid(row=r, column=1, columnspan=2, sticky='we', pady=3)
            else:
                lv = tk.StringVar()
                self.vars['lbl_' + key] = lv
                lv.trace_add('write', lambda *a, k='lbl_' + key: self._changed(k))
                no += 1
                ttk.Label(f, text='%d.' % no).grid(row=r, column=0, sticky='e', padx=(0, 6))
                ttk.Entry(f, textvariable=lv, font=FONT, foreground='#555').grid(row=r, column=1, sticky='we',
                                                                                  pady=3, padx=(0, 4))
                ttk.Entry(f, textvariable=v, font=FONT_B).grid(row=r, column=2, sticky='we', pady=3)
            r += 1
        if self.layout.known:
            hf = ttk.Frame(f)
            hf.grid(row=r, column=0, columnspan=3, sticky='w', pady=(8, 0))
            ttk.Label(hf, text='SHELF LIFE = DATE OF PRODUCTION +').pack(side='left')
            ttk.Spinbox(hf, from_=1, to=60, width=4, textvariable=self.v_months).pack(side='left', padx=3)
            ttk.Label(hf, text='tháng').pack(side='left')
            ttk.Button(hf, text='↻ Tính', command=self.calc_shelf).pack(side='left', padx=6)
            r += 1
            if self.layout.batch_db is not None:
                bf = ttk.LabelFrame(f, text=' Số BATCH ', padding=6)
                bf.grid(row=r, column=0, columnspan=3, sticky='we', pady=(8, 0))
                ttk.Label(bf, text='Mẫu tem lấy số mẻ từ Excel (cột %s). Nhiều tem liền nhau: gõ '
                                   '"VI. 1.1; 1.2; 1.3" (mỗi số = 1 tem khi in).' % (self.layout.batch_db.db_field or '?'),
                          foreground='#555', wraplength=440, justify='left').pack(anchor='w')
                ttk.Radiobutton(bf, text='Phần mềm tạo file Excel dữ liệu đi kèm (…_DATA.xlsx) – khuyên dùng',
                                value='data', variable=self.v_mode,
                                command=lambda: self._save_cfg(True)).pack(anchor='w')
                ttk.Radiobutton(bf, text='Giữ liên kết Excel gốc của mẫu (%s) – chỉ đổi tiền tố "%s"'
                                % (os.path.basename(self.btw.db_excel_path()) or '?',
                                   B.split_batch(self.defaults.get('batch', ''))[0]),
                                value='keep', variable=self.v_mode,
                                command=lambda: self._save_cfg(True)).pack(anchor='w')
                r += 1

    def _changed(self, key):
        if self._busy or self.cur is None or self.cur >= len(self.records):
            return
        self.records[self.cur][key] = self.vars[key].get()
        self._refresh_row(self.cur)
        self._schedule_preview()
        self._save_cfg(later=True)

    def _load_form(self, rec):
        self._busy = True
        try:
            for k, v in self.vars.items():
                v.set(rec.get(k, self.defaults.get(k, '')) or '')
        finally:
            self._busy = False
        self._schedule_preview()

    def calc_shelf(self):
        try:
            n = int(self.v_months.get())
        except ValueError:
            n = 12
        v = B.add_months(self.vars['mfg'].get(), n)
        if v:
            self.vars['shelf'].set(v)
            self._st('SHELF LIFE = %s (+%d tháng)' % (v, n))
        else:
            self._st('Không hiểu DATE OF PRODUCTION (dạng 08/2026 hoặc 15/08/2026).', True)

    # ------------------------------------------------------------ danh sach
    def _row_vals(self, i, rec):
        return (i + 1, rec.get('grade', ''), rec.get('batch', '').replace('\n', '; '), rec.get('mfg', ''),
                rec.get('_src', ''))

    def _fill_tree(self, select=None):
        self.tv.delete(*self.tv.get_children())
        for i, rec in enumerate(self.records):
            self.tv.insert('', 'end', iid=str(i), values=self._row_vals(i, rec))
        if select is not None and self.records:
            select = max(0, min(select, len(self.records) - 1))
            self.tv.selection_set(str(select))
            self.tv.see(str(select))
            self._on_select()
        elif not self.records:
            self.cur = None

    def _refresh_row(self, i):
        if self.tv.exists(str(i)):
            self.tv.item(str(i), values=self._row_vals(i, self.records[i]))

    def _on_select(self):
        sel = self.tv.selection()
        if not sel or self.layout is None:
            return
        self.cur = int(sel[0])
        self._load_form(self.records[self.cur])

    def _sel_idx(self):
        return sorted(int(x) for x in self.tv.selection())

    def new_rec(self):
        self.records.append(dict(self.defaults, _src='Nhập tay'))
        self._fill_tree(select=len(self.records) - 1)
        self._save_cfg(later=True)

    def dup_rec(self):
        if self.cur is None:
            return
        rec = dict(self.records[self.cur])
        rec['_src'] = 'Nhân bản'
        self.records.insert(self.cur + 1, rec)
        self._fill_tree(select=self.cur + 1)
        self._save_cfg(later=True)

    def del_rec(self):
        idx = self._sel_idx()
        if not idx or not messagebox.askyesno('Xoá tem', 'Xoá %d tem khỏi danh sách?' % len(idx), parent=self):
            return
        for i in reversed(idx):
            del self.records[i]
        self._fill_tree(select=min(idx[0], len(self.records) - 1) if self.records else None)
        if not self.records:
            self._load_form(self.defaults)
        self._save_cfg(later=True)

    # ------------------------------------------------------------ nhan du lieu
    def pick_file(self, kind):
        ft = {'img': [('Ảnh', '*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff')],
              'xls': [('Excel', '*.xlsx *.xlsm *.xls *.csv')],
              'doc': [('Word', '*.docx *.docm *.doc')]}[kind] + [('Tất cả', '*.*')]
        ps = filedialog.askopenfilenames(parent=self, title='Chọn file dữ liệu tem', filetypes=ft)
        if ps:
            self.on_drop(list(ps))

    def on_drop(self, files):
        files = [f for f in files if os.path.isfile(f)]
        if not files:
            return
        self._import_async([(os.path.basename(f), f) for f in files])

    def _ctrl_v(self, ev):
        try:
            if not self.winfo_viewable():
                return None
            w = self.focus_get()
            if w is None or w.winfo_toplevel() is not self.winfo_toplevel():
                return None                     # dang o cua so khac
            if isinstance(w, (tk.Entry, ttk.Entry, tk.Text, ttk.Spinbox, ttk.Combobox)):
                return None                     # dang go trong o -> dan chu binh thuong
        except (tk.TclError, KeyError):
            return None
        self.paste()
        return 'break'

    def paste(self):
        try:
            from paste_image import clipboard_images
            imgs = clipboard_images(self)
        except Exception:
            imgs = []
        if imgs:
            self._import_async(imgs)
            return
        try:
            txt = self.clipboard_get()
        except tk.TclError:
            txt = ''
        if not txt.strip():
            self._st('Clipboard trống – chụp ảnh tem (Win+Shift+S) hoặc Copy bảng Excel rồi bấm lại.', True)
            return
        recs = B.parse_rows(B.rows_from_text(txt))
        self._add_records(recs, 'Dán bảng', False)

    def _import_async(self, items):
        """items = [(ten, duong dan | PIL.Image)] - doc anh (OCR) chay nen de khong treo phan mem."""
        self._st('⏳ Đang đọc %d file / ảnh… (ảnh: lần đầu mất vài giây)' % len(items))
        self.drop_lbl.configure(text='⏳  Đang nhận diện…')

        def work():
            out, errs = [], []
            for name, src in items:
                try:
                    is_img = not isinstance(src, str) or src.lower().endswith(IMG_EXT)
                    if is_img:
                        recs = B.parse_rows(B.rows_from_image(src))
                    else:
                        recs = B.records_from_file(src)
                    out.append((name, recs, is_img))
                except Exception as ex:
                    errs.append('%s: %s' % (name, ex))
            self.after(0, lambda: self._import_done(out, errs))

        threading.Thread(target=work, daemon=True).start()

    def _import_done(self, out, errs):
        self.drop_lbl.configure(text='📥  KÉO THẢ ảnh tem / file Excel / Word vào đây   ·   hoặc Ctrl+V dán ảnh '
                                     '(Win+Shift+S) / bảng Excel đã Copy')
        n = 0
        bad = []
        for name, recs, is_img in out:
            if recs:
                n += self._add_records(recs, name, is_img, quiet=True)
            else:
                bad.append(name)
        msg = '✔ Đã nhận %d tem.' % n if n else ''
        if bad:
            msg += '  Không nhận ra bảng tem trong: %s' % ', '.join(bad)
        if errs:
            messagebox.showwarning('Nhận dữ liệu tem', '\n'.join(errs), parent=self)
        self._st(msg or 'Không nhận được tem nào.', not n)

    def _add_records(self, recs, src, is_img, quiet=False):
        if not recs:
            if not quiet:
                self._st('Không nhận ra bảng tem (cần các dòng NET WEIGHT / GRADE / BATCH NO…).', True)
            return 0
        start = len(self.records)
        # dang co 1 tem mau chua sua -> thay the luon
        if len(self.records) == 1 and self.records[0].get('_src') == 'Mẫu gốc':
            self.records, start = [], 0
        for r in recs:
            if is_img:
                r = B.clean_ocr_record(r, self.defaults)
            rec = dict(self.defaults)
            rec.update({k: v for k, v in r.items() if v not in (None, '')})
            if self.layout.known and r.get('mfg') and not r.get('shelf'):
                try:
                    rec['shelf'] = B.add_months(r['mfg'], int(self.v_months.get())) or rec.get('shelf', '')
                except ValueError:
                    pass
            rec['_src'] = src[:40]
            self.records.append(rec)
        self._fill_tree(select=start)
        self._save_cfg(later=True)
        if not quiet:
            self._st('✔ Đã nhận %d tem từ %s – kiểm tra lại nội dung rồi bấm 💾 Xuất .btw' % (len(recs), src))
        if is_img and not quiet:
            pass
        return len(recs)

    # ------------------------------------------------------------ xem truoc
    def _schedule_preview(self):
        if self._prev_job:
            self.after_cancel(self._prev_job)
        self._prev_job = self.after(250, self._preview)

    def _cur_rec(self):
        if self.cur is not None and self.cur < len(self.records):
            return self.records[self.cur]
        return dict(self.defaults) if self.layout else {}

    def _preview(self):
        self._prev_job = None
        if self.layout is None:
            return
        try:
            from PIL import ImageTk
            w, h = max(120, self.cv.winfo_width()), max(120, self.cv.winfo_height())
            sz = max(120, min(w, h) - 8)
            im = B.render_label(self._cur_rec(), self.layout, size=sz)
            self._tk_prev = ImageTk.PhotoImage(im)
            self.cv.delete('all')
            self.cv.create_image(w // 2, h // 2, image=self._tk_prev)
        except Exception as ex:
            self.cv.delete('all')
            self.cv.create_text(10, 10, anchor='nw', text='Lỗi xem trước: %s' % ex, fill='#b00020')

    # ------------------------------------------------------------ xuat
    # ---- thu muc luu: an toan khi chep phan mem sang may con
    def _default_out(self):
        return os.path.join(self.base, 'TemBTW')

    @staticmethod
    def _writable(d):
        """Tao duoc thu muc + ghi duoc file thu vao do?"""
        try:
            os.makedirs(d, exist_ok=True)
            t = os.path.join(d, '.htp_test_%d' % os.getpid())
            with open(t, 'w') as f:
                f.write('1')
            os.remove(t)
            return True
        except Exception:
            return False

    def _store_out(self, d):
        """Ghi vao tem_btw.json: thu muc nam trong thu muc phan mem -> luu duong dan TUONG DOI
        (chep sang may khac van dung). Thu muc mac dinh -> de trong."""
        d = (d or '').strip()
        if not d:
            return ''
        try:
            a = os.path.normcase(os.path.abspath(d))
            if a == os.path.normcase(os.path.abspath(self._default_out())):
                return ''
            rel = os.path.relpath(os.path.abspath(d), self.base)
            if not rel.startswith('..') and not os.path.isabs(rel):
                return rel
        except ValueError:          # khac o dia (C: / D:)
            pass
        return d

    def _load_out(self, d):
        """Doc tu tem_btw.json: duong dan cua MAY KHAC (vd C:\\Users\\admin\\... khong ton tai) -> ve mac dinh."""
        d = (d or '').strip()
        if not d:
            return self._default_out()
        if not os.path.isabs(d):
            return os.path.join(self.base, d)
        if os.path.isdir(d) or os.path.isdir(os.path.dirname(d.rstrip('\\/'))):
            return d
        return self._default_out()

    def pick_out(self):
        try:
            init = self._outdir(quiet=True)
        except Exception:
            init = os.path.expanduser('~')
        d = filedialog.askdirectory(parent=self, initialdir=init, title='Chọn thư mục lưu tem')
        if not d:
            return
        d = os.path.normpath(d)
        if not self._writable(d):
            messagebox.showerror('Lưu vào', 'Máy này KHÔNG có quyền ghi vào thư mục:\n%s\n\n'
                                 'Hãy chọn thư mục khác (vd Desktop / Documents) hoặc chạy phần mềm '
                                 'bằng quyền Administrator.' % d, parent=self)
            return
        self.v_out.set(d)
        if self._save_cfg():
            self._st('✔ Tem sẽ được lưu vào: %s' % d)
        else:
            messagebox.showwarning('Lưu vào', 'Đã chọn thư mục nhưng KHÔNG ghi được file cấu hình:\n%s\n\n'
                                   'Thư mục phần mềm đang bị chặn quyền ghi (thường do để trong '
                                   'C:\\Program Files). Lần mở sau sẽ mất lựa chọn này.\n'
                                   'Cách sửa: chép phần mềm ra D:\\ hoặc Desktop, hoặc chạy bằng quyền Administrator.'
                                   % self.cfg_path, parent=self)

    def _outdir(self, quiet=False):
        dflt = self._default_out()
        month = dt.date.today().strftime('%Y-%m')
        d = self.v_out.get().strip() or dflt
        if not os.path.isabs(d):
            d = os.path.join(self.base, d)
        if os.path.normcase(os.path.abspath(d)) == os.path.normcase(os.path.abspath(dflt)):
            d = os.path.join(d, month)
        if self._writable(d):
            return d
        # thu muc da chon khong dung duoc tren may nay -> lui ve thu muc du phong
        for fb in (os.path.join(dflt, month),
                   os.path.join(os.path.expanduser('~'), 'Documents', 'HitechPlast_TemBTW', month)):
            if self._writable(fb):
                if not quiet:
                    messagebox.showwarning('Lưu vào', 'Không ghi được vào thư mục:\n%s\n\nTạm lưu vào:\n%s\n\n'
                                           'Bấm nút … để chọn lại thư mục lưu.' % (d, fb), parent=self)
                self.v_out.set(os.path.dirname(fb))
                return fb
        raise OSError('Không có thư mục nào ghi được trên máy này: %s' % d)

    def _fname(self, rec, d, ext):
        p, nums = B.split_batches(rec.get('batch', ''))
        bat = ((p + ' ' + nums[0]) if nums else p).strip()
        if len(nums) > 1:
            bat += '-' + nums[-1]
        base = B.safe_name('TEM_%s_%s' % (rec.get('grade', ''), bat)).replace(' ', '')
        return os.path.join(d, base + ext)

    def export(self, all_sel=False):
        if self.layout is None:
            messagebox.showwarning('Xuất tem', 'Chưa có mẫu tem .btw.', parent=self)
            return
        idx = self._sel_idx() if all_sel else ([self.cur] if self.cur is not None else [])
        if not idx:
            messagebox.showinfo('Xuất tem', 'Chọn tem trong danh sách trước.', parent=self)
            return
        d = self._outdir()
        mode = self.v_mode.get()
        done, errs = [], []
        policy = [None]
        for i in idx:
            rec = self.records[i]
            path = ask_overwrite(self, self._fname(rec, d, '.btw'), policy)
            if path is None:
                continue
            try:
                out, data, n = B.export_record(self.tpl_path, rec, path, mode)
                done.append((out, data, n))
            except Exception as ex:
                errs.append('Tem %d: %s' % (i + 1, ex))
        self._save_cfg()
        if done:
            self.last_out = done[-1][0]
            self.btn_bt.configure(state='normal')
            lines = ['• %s%s' % (os.path.basename(o), ('  + %s (%d tem)' % (os.path.basename(x), n)) if x else '')
                     for o, x, n in done]
            self._st('✔ Đã xuất %d file .btw vào %s' % (len(done), d))
            note = ''
            if any(x for _, x, _ in done):
                note = ('\n\nĐể file …_DATA.xlsx CÙNG THƯ MỤC với file .btw (BarTender đọc số mẻ từ đó). '
                        'Chuyển sang máy khác: chép cả 2 file, mở .btw -> nếu BarTender báo không thấy Excel thì '
                        'vào Database Setup trỏ lại file _DATA.xlsx.')
            if messagebox.askyesno('Xuất tem', 'Đã xuất:\n%s%s\n\nMở thư mục chứa file?' % ('\n'.join(lines), note),
                                   parent=self):
                _open_path(d)
        if errs:
            messagebox.showerror('Xuất tem', '\n'.join(errs), parent=self)

    def save_now(self):
        """Luu danh sach tem + noi dung dang sua (tem_btw.json)."""
        self._save_cfg()
        self._st('💾 Đã lưu %d tem vào %s' % (len(self.records), os.path.basename(self.cfg_path)))

    def preview_big(self):
        """Xem truoc DUNG noi dung file .btw se xuat (tao file tam roi ve lai tu chinh file do)."""
        if self.layout is None or self.cur is None:
            return
        import tempfile
        rec = self.records[self.cur]
        if not self.layout.known and getattr(self.layout, 'thumb', None):
            img = B.render_thumb(rec, self.layout, 900)
            try:
                g = B.Geometry(self.btw)
                mm = (g.stock_w / B.MIL_PER_MM, g.stock_h / B.MIL_PER_MM)
            except Exception:
                mm = (0, 0)
            PreviewWindow(self, img, 'Tem %d – %s' % (self.cur + 1, os.path.basename(self.tpl_path)), mm,
                          on_export=lambda: self.export(), on_png=self.save_png)
            return
        tmp = os.path.join(tempfile.gettempdir(), 'hitech_tem_xemtruoc.btw')
        try:
            B.export_record(self.tpl_path, rec, tmp, self.v_mode.get())
            b = B.BtwFile(tmp)
            g = B.Geometry(b)
            img = render_btw_page(b, g, 10)
        except Exception as ex:
            messagebox.showerror('Xem trước', str(ex), parent=self)
            return
        mm = (g.stock_w / B.MIL_PER_MM, g.stock_h / B.MIL_PER_MM)
        PreviewWindow(self, img, 'Tem %d – %s' % (self.cur + 1, rec.get('grade', '')), mm,
                      on_export=lambda: self.export(), on_png=self.save_png)

    def save_png(self):
        if self.layout is None:
            return
        rec = self._cur_rec()
        p = filedialog.asksaveasfilename(parent=self, defaultextension='.png', filetypes=[('PNG', '*.png')],
                                         initialdir=self._outdir(),
                                         initialfile=os.path.basename(self._fname(rec, self._outdir(), '.png')))
        if p:
            B.render_label(rec, self.layout, size=1000).save(p)
            self._st('✔ Đã lưu ảnh %s' % p)

    def export_paper(self):
        """Xuất tem theo khổ giấy A3/A4/A5 (A5 tự xếp 2 tem + đường cắt giữa)."""
        if self.layout is None:
            messagebox.showinfo('Khổ giấy', 'Chưa có tem để xuất.', parent=self)
            return
        import tem_khogiay as TK
        rec = self._cur_rec()
        name = B.safe_name((rec.get('grade') if rec else '') or 'Tem')
        TK.choose_and_export(self, lambda: B.render_label(rec, self.layout, size=1400),
                             default_dir=self._outdir(), default_name=name)

    def open_last(self):
        if self.last_out and os.path.exists(self.last_out):
            _open_path(self.last_out)


# =====================================================================
#  DUNG CHUNG: hoi ghi de / anh trang giay / cua so xem truoc lon
# =====================================================================
def ask_overwrite(parent, path, policy=None):
    """File da co -> hoi: Co = ghi de, Khong = tao ban moi (_2, _3...), Huy = bo qua.
    policy = [None] dung chung cho nhieu file ('all_yes' / 'all_new' sau lan hoi dau)."""
    if not os.path.exists(path):
        return path
    pol = policy[0] if policy else None
    if pol is None:
        ans = messagebox.askyesnocancel(
            'File đã có', 'Đã có file:\n%s\n\nCó = GHI ĐÈ file cũ\nKhông = tạo bản mới (_2, _3…)\nHuỷ = bỏ qua'
            % os.path.basename(path), parent=parent)
        if ans is None:
            return None
        pol = 'all_yes' if ans else 'all_new'
        if policy is not None:
            policy[0] = pol
    if pol == 'all_yes':
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists('%s_%d%s' % (base, n, ext)):
        n += 1
    return '%s_%d%s' % (base, n, ext)


def render_btw_page(btw, geo, px_per_mm=10):
    """Ve lai tem tu CHINH file .btw (dung kho giay), co vien giay."""
    from PIL import Image, ImageDraw
    k = px_per_mm / B.MIL_PER_MM
    inner = B.render_btw(btw, k, geo)
    w, h = int(round(geo.stock_w * k)), int(round(geo.stock_h * k))
    page = Image.new('RGB', (w, h), 'white')
    page.paste(inner, (int(round(geo.mx * k)), 0))
    pad = 16
    out = Image.new('RGB', (w + 2 * pad, h + 2 * pad), (150, 150, 150))
    try:
        ImageDraw.Draw(out).rounded_rectangle([pad - 2, pad - 2, pad + w + 1, pad + h + 1], radius=10, fill='white')
    except AttributeError:
        pass
    out.paste(page, (pad, pad))
    return out


class PreviewWindow(tk.Toplevel):
    """Xem truoc tem LON truoc khi xuat .btw (phong to / thu nho, xuat ngay)."""

    def __init__(self, master, img, title, size_mm, on_export=None, on_png=None):
        super().__init__(master)
        self.title('👁 Xem trước tem – %s' % title)
        self.transient(master.winfo_toplevel())
        self.img = img
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.zoom = min(1.0, (sw * 0.8) / img.width, (sh * 0.72) / img.height)
        self.geometry('%dx%d' % (max(760, min(sw * 0.85, img.width * self.zoom + 40)),
                                 max(420, min(sh * 0.85, img.height * self.zoom + 130))))
        top = tk.Frame(self, bg='#1f3a5f')
        top.pack(fill='x')
        tk.Label(top, text='  Kích thước tem: %.0f x %.0f mm   (rộng x cao)' % size_mm, bg='#1f3a5f', fg='white',
                 font=FONT_B, pady=6).pack(side='left')
        for t, f in (('－', 1 / 1.25), ('＋', 1.25), ('Vừa khung', 0)):
            ttk.Button(top, text=t, command=lambda f=f: self._zoom(f)).pack(side='right', padx=2, pady=3)
        bt = ttk.Frame(self, padding=8)
        bt.pack(side='bottom', fill='x')
        ttk.Label(self, text='  Hình dưới vẽ lại đúng nội dung file .btw sẽ xuất (khung, chữ, cỡ chữ, khổ giấy).',
                  foreground='#555').pack(side='bottom', anchor='w')
        self.cv = tk.Canvas(self, bg='#dfe3e8', highlightthickness=0)
        self.cv.pack(fill='both', expand=True)
        ttk.Button(bt, text='Đóng', command=self.destroy).pack(side='right')
        if on_png:
            ttk.Button(bt, text='🖼 Lưu PNG', command=on_png).pack(side='right', padx=4)
        if on_export:
            tk.Button(bt, text='💾 Xuất .btw', bg='#0068ff', fg='white', relief='flat', font=FONT_B, padx=14,
                      command=lambda: (self.destroy(), on_export())).pack(side='right', padx=4)
        self.cv.bind('<Configure>', lambda e: self._show())
        self.bind('<Escape>', lambda e: self.destroy())
        self.focus_set()

    def _zoom(self, f):
        if f == 0:
            w, h = self.cv.winfo_width() - 20, self.cv.winfo_height() - 20
            self.zoom = min(w / self.img.width, h / self.img.height)
        else:
            self.zoom = max(0.1, min(6, self.zoom * f))
        self._show()

    def _show(self):
        from PIL import ImageTk, Image
        w = max(1, int(self.img.width * self.zoom))
        h = max(1, int(self.img.height * self.zoom))
        self._tk = ImageTk.PhotoImage(self.img.resize((w, h), Image.LANCZOS))
        self.cv.delete('all')
        cw, ch = self.cv.winfo_width(), self.cv.winfo_height()
        self.cv.create_image(max(w // 2 + 10, cw // 2), max(h // 2 + 10, ch // 2), image=self._tk)
        self.cv.configure(scrollregion=(0, 0, max(cw, w + 20), max(ch, h + 20)))


# =====================================================================
#  TAB "🎨 Tạo mẫu tem từ ảnh" - nhan dien khung + chon kich thuoc -> .btw moi
# =====================================================================
import btw_design as D


class DesignTab(tk.Frame):
    def __init__(self, master, fill_tab, base_dir=None, lib_name='tem_designs.json', kind=''):
        super().__init__(master, bg='#f4f6f9')
        self.fill = fill_tab                     # dung chung thu muc luu + mau phoi
        self.base = base_dir or _app_dir()
        self.kind = kind                         # '' = tem bao, 'pallet' = tem pallet
        self.lib_path = os.path.join(self.base, lib_name)
        self.lib = self._load_lib()
        self.d = D.normalize(self.lib[0] if self.lib else D.new_design())
        self.sel = None
        self._busy = False
        self._job = None
        self._build()
        self._to_ui()

    # ------------------------------------------------------------ thu vien mau
    def _load_lib(self):
        try:
            with open(self.lib_path, encoding='utf-8') as f:
                return json.load(f) or []
        except Exception:
            return []

    def _save_lib(self):
        tmp = self.lib_path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(self.lib, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.lib_path)

    # ------------------------------------------------------------ giao dien
    def _build(self):
        # --- nguon
        drop = tk.Frame(self, bg='#fff8e1', highlightthickness=2, highlightbackground='#ffb300')
        drop.pack(fill='x', padx=8, pady=(6, 3))
        self.drop_lbl = tk.Label(drop, text='🎨  KÉO THẢ / Ctrl+V ẢNH TEM MẪU → tự nhận diện khung bảng + chữ '
                                            '→ chọn kích thước → xuất .btw', bg='#fff8e1', fg='#6d4c00',
                                 font=FONT_B, pady=7)
        self.drop_lbl.pack(side='left', padx=8)
        bb = tk.Frame(drop, bg='#fff8e1')
        bb.pack(side='right', padx=6)
        ttk.Button(bb, text='📷 Chọn ảnh…', command=self.pick_image).pack(side='left', padx=2, pady=4)
        ttk.Button(bb, text='📋 Dán ảnh', command=self.paste).pack(side='left', padx=2)
        ttk.Button(bb, text='➕ Tem trống', command=self.blank).pack(side='left', padx=2)
        try:
            from recipe_ui import enable_file_drop
            for w in (drop, self.drop_lbl):
                enable_file_drop(w, lambda fs: fs and self.load_image(fs[0]))
        except Exception:
            pass
        self.bind_all('<Control-v>', self._ctrl_v, add='+')
        self.bind_all('<Control-V>', self._ctrl_v, add='+')

        # --- kich thuoc + tuy chon
        sz = ttk.LabelFrame(self, text=' Kích thước tem & kiểu ', padding=(8, 4))
        sz.pack(fill='x', padx=8, pady=3)
        ttk.Label(sz, text='Khổ có sẵn:').grid(row=0, column=0, sticky='w')
        self.v_pre = tk.StringVar()
        cb = ttk.Combobox(sz, textvariable=self.v_pre, state='readonly', width=18,
                          values=[p[0] for p in D.PRESETS] + ['Tự nhập…'])
        cb.grid(row=0, column=1, padx=(4, 12))
        cb.bind('<<ComboboxSelected>>', lambda e: self._preset())
        ttk.Label(sz, text='Rộng').grid(row=0, column=2)
        self.v_w = tk.StringVar()
        ttk.Spinbox(sz, from_=5, to=400, increment=1, width=6, textvariable=self.v_w).grid(row=0, column=3, padx=2)
        ttk.Label(sz, text='x Cao').grid(row=0, column=4)
        self.v_h = tk.StringVar()
        ttk.Spinbox(sz, from_=5, to=400, increment=1, width=6, textvariable=self.v_h).grid(row=0, column=5, padx=2)
        ttk.Label(sz, text='mm').grid(row=0, column=6)
        ttk.Button(sz, text='⇄ Xoay', command=self.swap).grid(row=0, column=7, padx=6)
        ttk.Label(sz, text='Lề').grid(row=0, column=8)
        self.v_m = tk.StringVar()
        ttk.Spinbox(sz, from_=0.5, to=20, increment=0.5, width=5, textvariable=self.v_m).grid(row=0, column=9, padx=2)
        ttk.Label(sz, text='mm   Cột trái').grid(row=0, column=10)
        self.v_split = tk.StringVar()
        ttk.Spinbox(sz, from_=10, to=90, increment=1, width=5, textvariable=self.v_split).grid(row=0, column=11, padx=2)
        ttk.Label(sz, text='%   Rộng chữ').grid(row=0, column=12)
        self.v_st = tk.StringVar()
        ttk.Spinbox(sz, from_=50, to=200, increment=5, width=5, textvariable=self.v_st).grid(row=0, column=13, padx=2)
        ttk.Label(sz, text='%').grid(row=0, column=14)
        of = ttk.Frame(sz)
        of.grid(row=1, column=0, columnspan=15, sticky='w', pady=(4, 0))
        self.v_border = tk.BooleanVar()
        self.v_bold = tk.BooleanVar()
        self.v_uni = tk.BooleanVar()
        self.v_wrap = tk.BooleanVar()
        for t, v in (('Kẻ khung bảng', self.v_border), ('Chữ đậm', self.v_bold),
                     ('Cỡ chữ đều các ô', self.v_uni), ('Tự xuống dòng chữ dài', self.v_wrap)):
            ttk.Checkbutton(of, text=t, variable=v, command=self._from_opts).pack(side='left', padx=(0, 12))
        for v in (self.v_w, self.v_h, self.v_m, self.v_split, self.v_st):
            v.trace_add('write', lambda *a: self._from_opts())

        body = tk.PanedWindow(self, orient='horizontal', sashwidth=6, bg='#cfd8e3')
        body.pack(fill='both', expand=True, padx=8, pady=3)
        # --- thu vien
        lf = ttk.LabelFrame(body, text=' Mẫu tem đã lưu ', padding=4)
        self.lb = tk.Listbox(lf, font=FONT, exportselection=False, width=24)
        self.lb.pack(fill='both', expand=True)
        self.lb.bind('<Double-1>', lambda e: self.open_lib())
        lbb = ttk.Frame(lf)
        lbb.pack(fill='x', pady=(4, 0))
        ttk.Button(lbb, text='📂 Mở', command=self.open_lib).pack(side='left')
        ttk.Button(lbb, text='🗑 Xoá', command=self.del_lib).pack(side='left', padx=3)
        body.add(lf, minsize=170, width=210)
        # --- cac dong
        mf = ttk.LabelFrame(body, text=' Các dòng của tem (bấm 1 dòng để sửa) ', padding=4)
        cols = (('n', '#', 30), ('a', 'Ô trái / ô gộp', 170), ('b', 'Ô phải', 150), ('h', 'Cao', 45),
                ('pt', 'Cỡ chữ', 55))
        self.tv = ttk.Treeview(mf, columns=[c for c, _, _ in cols], show='headings', height=9, selectmode='browse')
        for c, h, w in cols:
            self.tv.heading(c, text=h)
            self.tv.column(c, width=w, anchor='center' if c in ('n', 'h', 'pt') else 'w')
        self.tv.pack(fill='both', expand=True)
        self.tv.bind('<<TreeviewSelect>>', lambda e: self._pick_row())
        rb = ttk.Frame(mf)
        rb.pack(fill='x', pady=3)
        for t, c in (('➕ Thêm dòng', self.add_row), ('🗑 Xoá dòng', self.del_row), ('↑', lambda: self.move(-1)),
                     ('↓', lambda: self.move(1))):
            ttk.Button(rb, text=t, command=c).pack(side='left', padx=(0, 3))
        ed = ttk.LabelFrame(mf, text=' Sửa dòng đang chọn ', padding=6)
        ed.pack(fill='x')
        ed.columnconfigure(1, weight=1)
        ed.columnconfigure(3, weight=1)
        ttk.Label(ed, text='Ô trái:').grid(row=0, column=0, sticky='nw')
        self.t_a = tk.Text(ed, height=2, width=18, font=FONT, wrap='none', undo=True)
        self.t_a.grid(row=0, column=1, sticky='we', padx=(2, 8))
        ttk.Label(ed, text='Ô phải:').grid(row=0, column=2, sticky='nw')
        self.t_b = tk.Text(ed, height=2, width=18, font=FONT, wrap='none', undo=True)
        self.t_b.grid(row=0, column=3, sticky='we', padx=2)
        for t in (self.t_a, self.t_b):
            t.bind('<KeyRelease>', lambda e: self._from_row())
        o2 = ttk.Frame(ed)
        o2.grid(row=1, column=0, columnspan=4, sticky='w', pady=(4, 0))
        self.v_merge = tk.BooleanVar()
        ttk.Checkbutton(o2, text='Gộp 1 ô (cả hàng)', variable=self.v_merge, command=self._from_row).pack(side='left')
        ttk.Label(o2, text='   Cao').pack(side='left')
        self.v_rh = tk.StringVar()
        ttk.Spinbox(o2, from_=0.2, to=5, increment=0.1, width=5, textvariable=self.v_rh).pack(side='left', padx=2)
        ttk.Label(o2, text='  Cỡ chữ trái/phải (0 = tự động)').pack(side='left')
        self.v_pa = tk.StringVar()
        self.v_pb = tk.StringVar()
        ttk.Spinbox(o2, from_=0, to=72, increment=0.5, width=5, textvariable=self.v_pa).pack(side='left', padx=2)
        ttk.Spinbox(o2, from_=0, to=72, increment=0.5, width=5, textvariable=self.v_pb).pack(side='left', padx=2)
        self.v_line = tk.BooleanVar()
        ttk.Checkbutton(o2, text='Kẻ dưới', variable=self.v_line, command=self._from_row).pack(side='left', padx=8)
        for v in (self.v_rh, self.v_pa, self.v_pb):
            v.trace_add('write', lambda *a: self._from_row())
        ttk.Label(mf, text='Xuống dòng trong 1 ô: Enter.  Ô có nhiều dòng chữ dùng nhiều ô chữ của phôi (tối đa 17).',
                  foreground='#777').pack(anchor='w')
        body.add(mf, minsize=420, width=560)
        # --- xem truoc
        pf = ttk.LabelFrame(body, text=' Xem trước (đúng tỉ lệ) ', padding=4)
        self.cv = tk.Canvas(pf, bg='#eef1f5', highlightthickness=0, width=340, height=340)
        self.cv.pack(fill='both', expand=True)
        self.cv.bind('<Configure>', lambda e: self._schedule())
        self.lbl_use = ttk.Label(pf, text='', foreground='#1f3a5f')
        self.lbl_use.pack(anchor='w')
        body.add(pf, minsize=260)

        bt = tk.Frame(self, bg='#f4f6f9')
        bt.pack(fill='x', padx=8, pady=(2, 2))
        tk.Label(bt, text='Tên mẫu:', bg='#f4f6f9', font=FONT).pack(side='left')
        self.v_name = tk.StringVar()
        ttk.Entry(bt, textvariable=self.v_name, width=30).pack(side='left', padx=4)
        tk.Button(bt, text='💾 Lưu mẫu tem', bg='#2e7d32', fg='white', relief='flat', font=FONT_B, padx=10,
                  command=self.save_lib).pack(side='left', padx=4)
        tk.Button(bt, text='💾 Xuất .btw', bg='#0068ff', fg='white', relief='flat', font=FONT_B, padx=14,
                  command=self.export).pack(side='right', padx=4)
        tk.Button(bt, text='👁 Xem trước tem', bg='#ffb300', fg='#222', relief='flat', font=FONT_B, padx=10,
                  command=self.preview_big).pack(side='right', padx=4)
        tk.Button(bt, text='🖨 Khổ giấy A3/A4/A5', bg='#6a1b9a', fg='white', relief='flat', font=FONT_B, padx=10,
                  command=self.export_paper).pack(side='right', padx=4)
        ttk.Button(bt, text='🖼 Lưu PNG', command=self.save_png).pack(side='right')
        self.status = tk.Label(self, text='', bg='#f4f6f9', fg='#1f3a5f', font=FONT, anchor='w')
        self.status.pack(fill='x', padx=10, pady=(0, 6))
        self._fill_lib()

    def _st(self, s, err=False):
        self.status.configure(text=s, fg='#b00020' if err else '#1f3a5f')

    # ------------------------------------------------------------ dong bo UI <-> thiet ke
    def _to_ui(self):
        self._busy = True
        try:
            d = self.d
            self.v_w.set('%g' % d['w_mm'])
            self.v_h.set('%g' % d['h_mm'])
            self.v_m.set('%g' % d['margin'])
            self.v_split.set('%g' % round(d['split'] * 100))
            self.v_st.set('%g' % d['stretch'])
            self.v_border.set(d['border'])
            self.v_bold.set(d['bold'])
            self.v_uni.set(d.get('uniform', True))
            self.v_wrap.set(d.get('autowrap', True))
            self.v_name.set(d.get('name', ''))
            self._sync_preset()
        finally:
            self._busy = False
        self._fill_rows(0)

    def _sync_preset(self):
        key = next((p[0] for p in D.PRESETS if abs(p[1] - self.d['w_mm']) < 0.01 and abs(p[2] - self.d['h_mm']) < 0.01),
                   'Tự nhập…')
        self.v_pre.set(key)

    def _num(self, v, default):
        try:
            return float(str(v.get()).replace(',', '.'))
        except ValueError:
            return default

    def _from_opts(self):
        if self._busy:
            return
        d = self.d
        d['w_mm'] = max(5.0, self._num(self.v_w, d['w_mm']))
        d['h_mm'] = max(5.0, self._num(self.v_h, d['h_mm']))
        d['margin'] = max(0.3, self._num(self.v_m, d['margin']))
        d['split'] = max(0.1, min(0.9, self._num(self.v_split, d['split'] * 100) / 100.0))
        d['stretch'] = max(50.0, min(200.0, self._num(self.v_st, d['stretch'])))
        d['border'], d['bold'] = self.v_border.get(), self.v_bold.get()
        d['uniform'], d['autowrap'] = self.v_uni.get(), self.v_wrap.get()
        self._busy = True
        try:
            self._sync_preset()
        finally:
            self._busy = False
        self._refresh_pts()
        self._schedule()

    def _preset(self):
        v = self.v_pre.get()
        p = next((x for x in D.PRESETS if x[0] == v), None)
        if p:
            self.v_w.set('%g' % p[1])
            self.v_h.set('%g' % p[2])

    def swap(self):
        w, h = self.v_w.get(), self.v_h.get()
        self.v_w.set(h)
        self.v_h.set(w)

    def _fill_rows(self, select=None):
        self.tv.delete(*self.tv.get_children())
        for i, r in enumerate(self.d['rows']):
            self.tv.insert('', 'end', iid=str(i), values=self._row_vals(i, r))
        if select is not None and self.d['rows']:
            select = max(0, min(select, len(self.d['rows']) - 1))
            self.tv.selection_set(str(select))
            self.tv.see(str(select))
        self._schedule()

    def _row_vals(self, i, r):
        pts = self._pts().get(i, {})
        fmt = lambda c: ('%g' % r['pt'][c]) if r['pt'][c] else ('tự %g' % pts[c] if c in pts else '')
        return (i + 1, r['cells'][0].replace('\n', ' ⏎ '), r['cells'][1].replace('\n', ' ⏎ ') if len(r['cells']) > 1
                else '(gộp)', '%g' % round(r['h'], 2),
                ' / '.join(fmt(c) for c in range(len(r['cells']))))

    def _pts(self):
        out = {}
        try:
            for t in D.compute_layout(self.d)['texts']:
                out.setdefault(t['cell'][0], {})[t['cell'][1]] = t['pt']
        except Exception:
            pass
        return out

    def _refresh_pts(self):
        for i, r in enumerate(self.d['rows']):
            if self.tv.exists(str(i)):
                self.tv.item(str(i), values=self._row_vals(i, r))

    def _pick_row(self):
        s = self.tv.selection()
        if not s:
            return
        self.sel = int(s[0])
        r = self.d['rows'][self.sel]
        self._busy = True
        try:
            for w, txt in ((self.t_a, r['cells'][0]), (self.t_b, r['cells'][1] if len(r['cells']) > 1 else '')):
                w.delete('1.0', 'end')
                w.insert('1.0', txt)
            self.v_merge.set(len(r['cells']) == 1)
            self.t_b.configure(state='disabled' if len(r['cells']) == 1 else 'normal')
            self.v_rh.set('%g' % round(r['h'] / self._hbase(), 2))
            self.v_pa.set('%g' % r['pt'][0])
            self.v_pb.set('%g' % (r['pt'][1] if len(r['pt']) > 1 else 0))
            self.v_line.set(r.get('line', True))
        finally:
            self._busy = False

    def _hbase(self):
        hs = sorted(r['h'] for r in self.d['rows'])
        return hs[len(hs) // 2] or 1.0

    def _from_row(self):
        if self._busy or self.sel is None or self.sel >= len(self.d['rows']):
            return
        r = self.d['rows'][self.sel]
        a = self.t_a.get('1.0', 'end-1c')
        b = self.t_b.get('1.0', 'end-1c') if len(r['cells']) > 1 else (r.get('_b') or '')
        if self.v_merge.get():
            if len(r['cells']) > 1:
                r['_b'] = b
            r['cells'], r['pt'] = [a], [self._num(self.v_pa, 0)]
            self.t_b.configure(state='disabled')
        else:
            if len(r['cells']) == 1:
                self.t_b.configure(state='normal')
                b = r.pop('_b', '') or b
                self.t_b.delete('1.0', 'end')
                self.t_b.insert('1.0', b)
            r['cells'], r['pt'] = [a, b], [self._num(self.v_pa, 0), self._num(self.v_pb, 0)]
        r['h'] = max(0.05, self._num(self.v_rh, 1.0)) * self._hbase() if self.v_rh.get() else r['h']
        r['line'] = self.v_line.get()
        self._refresh_pts()
        self._schedule()

    def add_row(self):
        i = (self.sel + 1) if self.sel is not None else len(self.d['rows'])
        self.d['rows'].insert(i, {'cells': ['NHÃN', 'GIÁ TRỊ'], 'h': self._hbase(), 'pt': [0, 0], 'line': True})
        self._fill_rows(i)

    def del_row(self):
        if self.sel is None or len(self.d['rows']) <= 1:
            return
        del self.d['rows'][self.sel]
        self._fill_rows(min(self.sel, len(self.d['rows']) - 1))

    def move(self, k):
        i = self.sel
        if i is None or not (0 <= i + k < len(self.d['rows'])):
            return
        rs = self.d['rows']
        rs[i], rs[i + k] = rs[i + k], rs[i]
        self._fill_rows(i + k)

    # ------------------------------------------------------------ nguon
    def pick_image(self):
        p = filedialog.askopenfilename(parent=self, title='Chọn ảnh tem mẫu',
                                       filetypes=[('Ảnh', '*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff'),
                                                  ('Tất cả', '*.*')])
        if p:
            self.load_image(p)

    def _ctrl_v(self, ev):
        try:
            if not self.winfo_viewable():
                return None
            w = self.focus_get()
            if w is None or w.winfo_toplevel() is not self.winfo_toplevel():
                return None
            if isinstance(w, (tk.Entry, ttk.Entry, tk.Text, ttk.Spinbox, ttk.Combobox)):
                return None
        except (tk.TclError, KeyError):
            return None
        self.paste()
        return 'break'

    def paste(self):
        try:
            from paste_image import clipboard_images
            imgs = clipboard_images(self)
        except Exception:
            imgs = []
        if imgs:
            self.load_image(imgs[0][1])
        else:
            self._st('Clipboard chưa có ảnh – chụp tem (Win+Shift+S) hoặc Copy ảnh rồi bấm lại.', True)

    def load_image(self, src):
        """Nhan dien anh (chay nen). Giu kich thuoc dang chon."""
        self._st('⏳ Đang nhận diện khung + chữ… (lần đầu mất vài giây)')
        self.drop_lbl.configure(text='⏳  Đang nhận diện…')
        w, h = self.d['w_mm'], self.d['h_mm']

        def work():
            try:
                d = D.design_from_image(src, w, h)
                err = None
            except Exception as ex:
                d, err = None, ex
            self.after(0, lambda: self._loaded(d, err))

        threading.Thread(target=work, daemon=True).start()

    def _loaded(self, d, err):
        self.drop_lbl.configure(text='🎨  KÉO THẢ / Ctrl+V ẢNH TEM MẪU → tự nhận diện khung bảng + chữ '
                                     '→ chọn kích thước → xuất .btw')
        if err or not d:
            messagebox.showerror('Nhận diện tem', 'Không nhận diện được:\n%s' % err, parent=self)
            self._st('Không nhận diện được ảnh.', True)
            return
        for k in ('margin', 'stretch', 'bold'):
            d[k] = self.d.get(k, d[k])
        self.d = D.normalize(d)
        self._to_ui()
        self._st('✔ Nhận được %d dòng%s – KIỂM TRA LẠI CHỮ (máy đọc có thể sai), chọn kích thước rồi xuất.'
                 % (len(self.d['rows']), ', có kẻ khung' if self.d['border'] else ', không kẻ khung'))

    def blank(self):
        self.d = D.normalize(D.new_design(self.d['w_mm'], self.d['h_mm']))
        self._to_ui()

    # ------------------------------------------------------------ thu vien
    def _fill_lib(self):
        self.lb.delete(0, 'end')
        for d in self.lib:
            self.lb.insert('end', '%s  (%gx%g)' % (d.get('name') or '?', d.get('w_mm', 0), d.get('h_mm', 0)))

    def save_lib(self):
        name = self.v_name.get().strip() or self.d.get('name') or 'Mẫu tem'
        self.d['name'] = name
        self.v_name.set(name)
        old = next((i for i, x in enumerate(self.lib) if x.get('name') == name), None)
        if old is not None and not messagebox.askyesno('Lưu mẫu tem', 'Đã có mẫu "%s". Ghi đè?' % name, parent=self):
            return
        item = json.loads(json.dumps(self.d, ensure_ascii=False))
        if old is None:
            self.lib.append(item)
        else:
            self.lib[old] = item
        self._save_lib()
        self._fill_lib()
        self._st('💾 Đã lưu mẫu tem "%s" (%gx%g mm) – lần sau chọn ở danh sách bên trái.'
                 % (name, self.d['w_mm'], self.d['h_mm']))

    def open_lib(self):
        s = self.lb.curselection()
        if s:
            self.d = D.normalize(self.lib[s[0]])
            self._to_ui()
            self._st('Đã mở mẫu "%s".' % self.d.get('name'))

    def del_lib(self):
        s = self.lb.curselection()
        if s and messagebox.askyesno('Xoá mẫu', 'Xoá mẫu "%s"?' % self.lib[s[0]].get('name'), parent=self):
            del self.lib[s[0]]
            self._save_lib()
            self._fill_lib()

    # ------------------------------------------------------------ xem truoc / xuat
    def _schedule(self):
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(200, self._preview)

    def _preview(self):
        self._job = None
        try:
            from PIL import ImageTk
            w, h = max(100, self.cv.winfo_width() - 10), max(100, self.cv.winfo_height() - 10)
            ppm = min(w / (self.d['w_mm'] + 4), h / (self.d['h_mm'] + 4))
            im = D.render_design(self.d, px_per_mm=max(1.5, ppm))
            self._tk = ImageTk.PhotoImage(im)
            self.cv.delete('all')
            self.cv.create_image(self.cv.winfo_width() // 2, self.cv.winfo_height() // 2, image=self._tk)
            L = D.compute_layout(self.d)
            nl = len(L['hlines']) + len(L['vlines'])
            errs = D.check_limits(self.d)
            self.lbl_use.configure(text='%g x %g mm · ô chữ %d/17 · đường kẻ %d/11%s' % (
                self.d['w_mm'], self.d['h_mm'], len(L['texts']), nl, ('  ⚠ ' + errs[0]) if errs else ''),
                foreground='#b00020' if errs else '#1f3a5f')
        except Exception as ex:
            self.cv.delete('all')
            self.cv.create_text(10, 10, anchor='nw', text='Lỗi xem trước: %s' % ex, fill='#b00020')

    def preview_big(self):
        errs = D.check_limits(self.d)
        if errs:
            messagebox.showwarning('Xem trước', '\n'.join(errs), parent=self)
        img = D.render_design(self.d, px_per_mm=12)
        PreviewWindow(self, img, self.d.get('name') or 'mẫu mới', (self.d['w_mm'], self.d['h_mm']),
                      on_export=self.export, on_png=self.save_png)

    def _donor(self):
        p = _resource(DEFAULT_TPL)
        if not os.path.exists(p):
            raise B.BtwError('Thiếu file phôi %s cạnh phần mềm.' % DEFAULT_TPL)
        return p

    def export(self):
        errs = D.check_limits(self.d)
        if errs:
            messagebox.showwarning('Xuất .btw', '\n'.join(errs), parent=self)
            return
        d = self.fill._outdir()
        nm = self.v_name.get().strip() or self.d.get('name') or 'MAU'
        size = '%gx%g' % (self.d['w_mm'], self.d['h_mm'])
        name = B.safe_name('TEM_%s' % nm if size in nm.replace(' ', '') else 'TEM_%s_%s' % (nm, size)).replace(' ', '_')
        path = ask_overwrite(self, os.path.join(d, name + '.btw'))
        if not path:
            return
        try:
            out, data = D.export_design(self.d, self._donor(), path)
        except Exception as ex:
            messagebox.showerror('Xuất .btw', str(ex), parent=self)
            return
        self.fill.last_out = out
        self.fill.btn_bt.configure(state='normal')
        self._st('✔ Đã xuất %s (%g x %g mm)' % (os.path.basename(out), self.d['w_mm'], self.d['h_mm']))
        if messagebox.askyesno('Xuất .btw', 'Đã xuất:\n• %s\n• %s\n\nĐể 2 file CÙNG THƯ MỤC.\n\n'
                               'Mở bằng BarTender ngay?' % (os.path.basename(out),
                                                             os.path.basename(data) if data else ''), parent=self):
            _open_path(out)

    def save_png(self):
        p = filedialog.asksaveasfilename(parent=self, defaultextension='.png', filetypes=[('PNG', '*.png')],
                                         initialdir=self.fill._outdir(),
                                         initialfile=B.safe_name('TEM_%s' % (self.d.get('name') or 'MAU')) + '.png')
        if p:
            D.render_design(self.d, px_per_mm=12, frame=False).save(p, dpi=(305, 305))
            self._st('✔ Đã lưu ảnh %s' % p)

    def export_paper(self):
        """Xuất mẫu tem theo khổ giấy A3/A4/A5 (A5 tự xếp 2 tem + đường cắt giữa)."""
        import tem_khogiay as TK
        TK.choose_and_export(self, lambda: D.render_design(self.d, px_per_mm=12, frame=False),
                             default_dir=self.fill._outdir(),
                             default_name=B.safe_name('TEM_%s' % (self.d.get('name') or 'MAU')))


# =====================================================================
#  TAB CHINH "🏷 Tem BTW" = 2 tab con
# =====================================================================
class TemTab(tk.Frame):
    def __init__(self, master, app=None, base_dir=None):
        super().__init__(master, bg='#f4f6f9')
        hd = tk.Frame(self, bg='#1f3a5f')
        hd.pack(fill='x')
        tk.Label(hd, text='🏷  TEM BTW – nhận dữ liệu tem / tạo mẫu tem mới → xuất file BarTender (.btw)',
                 bg='#1f3a5f', fg='white', font=('Segoe UI', 11, 'bold'), padx=10, pady=6).pack(side='left')
        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=4, pady=4)
        self.fill_tab = FillTab(nb, app=app, base_dir=base_dir)
        nb.add(self.fill_tab, text=' 📋 Điền tem theo mẫu (.btw) ')
        # Tem BAO (thư viện tem_designs.json)
        self.design_tab = DesignTab(nb, self.fill_tab, base_dir=base_dir,
                                    lib_name='tem_designs.json', kind='bao')
        nb.add(self.design_tab, text=' 🎨 Tạo Mẫu Tem Bao Từ Ảnh ')
        # Tem PALLET (thư viện riêng tem_designs_pallet.json)
        self.design_pallet = DesignTab(nb, self.fill_tab, base_dir=base_dir,
                                       lib_name='tem_designs_pallet.json', kind='pallet')
        nb.add(self.design_pallet, text=' 📦 Tạo Mẫu Tem Pallet Từ Ảnh ')
        self.nb = nb

    def __getattr__(self, k):
        if k in ('fill_tab', 'design_tab', 'design_pallet', 'nb'):
            raise AttributeError(k)
        return getattr(self.fill_tab, k)
