# -*- coding: utf-8 -*-
"""
presence.py - MAY NAO DANG ONLINE / DANG DUNG PHAN MEM - HitechPlast
====================================================================
  - Che do MAY CHU: moi may dang mo phan mem cu 20 giay ghi "nhip tim" vao bang htp_presence
    tren may chu (ten may, tai khoan dang nhap, IP LAN, IP Tailscale, duong ket noi...).
    May nao > 60 giay khong ghi = offline. Dung GIO CUA MAY CHU (now()) nen khong lech gio.
  - Them: danh sach may trong mang TAILSCALE (lenh 'tailscale status') - biet may nao dang bat.
  - Mo cua so:  presence.open_window(master)      (nut "🟢 N máy online" / trong "Máy chủ dữ liệu")

Nhip tim chay tren LUONG RIENG voi KET NOI RIENG -> khong anh huong giao dien / giao dich.
"""
import os
import platform
import socket
import sys
import threading
import time
import uuid

HEARTBEAT_SEC = 20
ONLINE_SEC = 60
APP_VERSION = '2026.10'
_SEC = '%d seconds' % ONLINE_SEC

SCHEMA = [
    "CREATE TABLE IF NOT EXISTS htp_presence(machine TEXT PRIMARY KEY, pc_name TEXT, os_user TEXT,"
    " app_user TEXT, ip_lan TEXT, ip_ts TEXT, route TEXT, os TEXT, app_ver TEXT,"
    " started TIMESTAMP DEFAULT now(), last_seen TIMESTAMP DEFAULT now(), status TEXT)",
    # lenh quan tri: cmd='kick' (ngat ket noi 1 lan) | blocked=true (khoa, khong mo duoc phan mem)
    "ALTER TABLE htp_presence ADD COLUMN IF NOT EXISTS cmd TEXT, ADD COLUMN IF NOT EXISTS blocked BOOLEAN"
    " DEFAULT false, ADD COLUMN IF NOT EXISTS cmd_by TEXT, ADD COLUMN IF NOT EXISTS cmd_msg TEXT,"
    " ADD COLUMN IF NOT EXISTS cmd_at TIMESTAMP, ADD COLUMN IF NOT EXISTS username TEXT",
    # netlock=true: KHOA MANG - may con khong ket noi duoc may chu (giong may chu dang tat)
    # ip_rm: IP Radmin VPN cua may
    "ALTER TABLE htp_presence ADD COLUMN IF NOT EXISTS netlock BOOLEAN DEFAULT false,"
    " ADD COLUMN IF NOT EXISTS ip_rm TEXT",
]
_COLS = ['cmd', 'blocked', 'cmd_by', 'cmd_msg', 'cmd_at', 'username', 'netlock', 'ip_rm']

_THREAD = None
_STOP = threading.Event()
_STATE = {'online': None, 'err': '', 'user': '', 'kicked': None, 'username': ''}


# =====================================================================
#  Thong tin may nay
# =====================================================================
def machine_id():
    """Ma may co dinh (luu trong %APPDATA%/HitechPlast/machine_id.txt)."""
    base = os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), '.config')
    d = os.path.join(base, 'HitechPlast')
    p = os.path.join(d, 'machine_id.txt')
    try:
        return open(p, encoding='utf-8').read().strip() or _new_id(d, p)
    except Exception:
        return _new_id(d, p)


def _new_id(d, p):
    mid = f'{socket.gethostname()}-{uuid.uuid4().hex[:8]}'
    try:
        os.makedirs(d, exist_ok=True)
        with open(p, 'w', encoding='utf-8') as f:
            f.write(mid)
    except Exception:
        pass
    return mid


def lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('192.168.1.1', 9))           # chi de hoi bang dinh tuyen, khong gui goi tin
        ip = s.getsockname()[0]
        return '' if (ip.startswith('100.') or ip.startswith('26.')) else ip
    except OSError:
        return ''
    finally:
        s.close()


def ts_ip():
    try:
        import tailscale_util as TS
        return TS.local_ts_ip() or ''
    except Exception:
        return ''


def rm_ip():
    try:
        import tailscale_util as TS
        return TS.local_radmin_ip() or ''
    except Exception:
        return ''


def os_user():
    return os.environ.get('USERNAME') or os.environ.get('USER') or ''


# =====================================================================
#  Nhip tim
# =====================================================================
_SCHEMA_OK = [False]


def _ensure_schema(c):
    """Tao bang / them cot 1 lan moi phien (ALTER TABLE khoa bang -> khong chay lai moi nhip tim).
    TOC DO: bang va du cot roi (hau het cac lan) -> chi 1 truy van kiem tra."""
    if not _SCHEMA_OK[0]:
        try:
            r = c.one("SELECT count(*) FROM information_schema.columns WHERE table_name='htp_presence'"
                      " AND column_name = ANY(%s)", (_COLS,))
            if r and int(r[0]) >= len(_COLS):
                _SCHEMA_OK[0] = True
                return
        except Exception:
            pass
        for q in SCHEMA:
            c.execute(q)
        _SCHEMA_OK[0] = True


def _connect(cfg):
    import datastore
    c = datastore.Conn(cfg)
    _ensure_schema(c)
    return c


def _app_ver():
    """Phien ban dang cai tren may nay (updater.py / version.json) -> admin xem may nao chua cap nhat."""
    try:
        import updater
        v = updater.current_version()
        if v and v != '0':
            return v
    except Exception:
        pass
    return APP_VERSION


def _beat(c, mid, info):
    c.execute("INSERT INTO htp_presence(machine, pc_name, os_user, app_user, ip_lan, ip_ts, route, os, app_ver,"
              " started, last_seen, status, username, ip_rm)"
              " VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s, now(), now(), 'online', %s, %s)"
              " ON CONFLICT (machine) DO UPDATE SET pc_name=EXCLUDED.pc_name, os_user=EXCLUDED.os_user,"
              " app_user=EXCLUDED.app_user, ip_lan=EXCLUDED.ip_lan, ip_ts=EXCLUDED.ip_ts, route=EXCLUDED.route,"
              " username=EXCLUDED.username, ip_rm=EXCLUDED.ip_rm,"
              " os=EXCLUDED.os, app_ver=EXCLUDED.app_ver, last_seen=now(), status='online',"
              " started=CASE WHEN htp_presence.status='online' AND htp_presence.last_seen > now() - interval '" + _SEC + "'"
              " THEN htp_presence.started ELSE now() END",
              (mid, info['pc_name'], info['os_user'], info['app_user'], lan_ip(), ts_ip(),
               info['route'], info['os'], _app_ver(), info.get('username', ''), rm_ip()))
    r = c.one("SELECT COUNT(*) FROM htp_presence WHERE status='online' AND last_seen > now() - interval '"
              + _SEC + "'")
    return int(r[0]) if r else 0


def _poll_cmd(c, mid):
    """Lenh quan tri cho may nay -> None | {'kind': 'kick'|'block', 'msg', 'by'}.
    Khoa theo ma may HOAC ten may (xoa file machine_id cung khong thoat duoc)."""
    r = c.one("SELECT cmd_msg, cmd_by FROM htp_presence WHERE blocked IS TRUE AND"
              " (machine=%s OR lower(pc_name)=lower(%s)) LIMIT 1", (mid, socket.gethostname()))
    if r:
        return {'kind': 'block', 'msg': r[0] or '', 'by': r[1] or ''}
    if is_netlocked(c, mid):
        return {'kind': 'netlock', 'msg': '', 'by': ''}
    r = c.one("SELECT cmd, cmd_msg, cmd_by FROM htp_presence WHERE machine=%s", (mid,))
    if r and r[0] == 'kick':
        c.execute("UPDATE htp_presence SET cmd=NULL WHERE machine=%s", (mid,))     # ngat 1 lan
        return {'kind': 'kick', 'msg': r[1] or '', 'by': r[2] or ''}
    if r and str(r[0] or '').startswith('logout:'):
        c.execute("UPDATE htp_presence SET cmd=NULL WHERE machine=%s", (mid,))
        target = str(r[0])[7:]
        if target.lower() == str(_STATE.get('username') or '').lower():    # dung tai khoan dang dung
            return {'kind': 'logout', 'msg': r[1] or '', 'by': r[2] or ''}
    return None


def is_netlocked(c, mid=None):
    """May nay co bi KHOA MANG (admin cat khoi may chu) khong. Loi / chua co cot -> False."""
    try:
        r = c.one("SELECT 1 FROM htp_presence WHERE netlock IS TRUE AND"
                  " (machine=%s OR lower(pc_name)=lower(%s)) LIMIT 1", (mid or machine_id(), socket.gethostname()))
        return bool(r)
    except Exception:
        return False


def user_sessions(username, cfg=None):
    """Cac MAY KHAC dang online voi tai khoan `username` -> [{'machine','pc_name','ip_lan','ip_ts'}]."""
    import datastore
    own = cfg is not None
    if not own and not datastore.is_server():
        return []
    c = _connect(cfg) if own else datastore.server().c
    try:
        _ensure_schema(c)
        rows = c.query("SELECT machine, pc_name, ip_lan, ip_ts FROM htp_presence WHERE lower(username)=lower(%s)"
                       " AND machine<>%s AND status='online' AND last_seen > now() - interval '" + _SEC + "'",
                       (str(username or ''), machine_id()))
    finally:
        if own:
            c.close()
    return [{'machine': r[0], 'pc_name': r[1] or '', 'ip_lan': r[2] or '', 'ip_ts': r[3] or ''} for r in rows]


def logout_user(username, by='', msg='', exclude_me=True, cfg=None):
    """Dang xuat tai khoan `username` tren cac may dang online (tru may nay neu exclude_me). -> so may."""
    import datastore
    own = cfg is not None
    if not own and not datastore.is_server():
        return 0
    c = _connect(cfg) if own else datastore.server().c
    try:
        _ensure_schema(c)
        q = ("SELECT machine FROM htp_presence WHERE lower(username)=lower(%s) AND status='online'"
             " AND last_seen > now() - interval '" + _SEC + "'")
        args = [str(username or '')]
        if exclude_me:
            q += " AND machine<>%s"
            args.append(machine_id())
        ms = [r[0] for r in c.query(q, tuple(args))]
        for m in ms:
            c.execute("UPDATE htp_presence SET cmd=%s, cmd_by=%s, cmd_msg=%s, cmd_at=now() WHERE machine=%s",
                      ('logout:' + str(username), by, msg, m))
    finally:
        if own:
            c.close()
    return len(ms)


def check_blocked(cfg=None):
    """Truoc khi dang nhap (che do may chu): may nay co bi KHOA khong -> None | dict nhu _poll_cmd."""
    import datastore
    own = cfg is not None
    try:
        if cfg is None:
            if not datastore.is_server():
                return None
            c = datastore.server().c            # TOC DO: dung ket noi da co (truoc mo ket noi moi)
            _ensure_schema(c)
        else:
            c = _connect(cfg)
    except Exception:
        return None
    try:
        r = c.one("SELECT cmd_msg, cmd_by FROM htp_presence WHERE blocked IS TRUE AND"
                  " (machine=%s OR lower(pc_name)=lower(%s)) LIMIT 1", (machine_id(), socket.gethostname()))
        return {'kind': 'block', 'msg': r[0] or '', 'by': r[1] or ''} if r else None
    except Exception:
        return None
    finally:
        if own:
            try:
                c.close()
            except Exception:
                pass


def kicked():
    """Phan mem chinh hoi dinh ky: may nay vua bi admin ngat / khoa chua."""
    return _STATE.get('kicked')


def send_cmd(machine, kind, by='', msg='', cfg=None):
    """ADMIN: kind = 'kick' (ngat ket noi) | 'block' (khoa + ngat) | 'unblock' (mo khoa)
    | 'netlock' (khoa mang: may con nhu mat ket noi may chu) | 'netunlock' (mo mang)."""
    import datastore
    own = cfg is not None
    c = _connect(cfg) if own else datastore.server().c
    try:
        _ensure_schema(c)
        if kind == 'kick':
            c.execute("UPDATE htp_presence SET cmd='kick', cmd_by=%s, cmd_msg=%s, cmd_at=now()"
                      " WHERE machine=%s", (by, msg, machine))
        elif kind == 'block':
            c.execute("UPDATE htp_presence SET blocked=true, cmd='kick', cmd_by=%s, cmd_msg=%s, cmd_at=now()"
                      " WHERE machine=%s", (by, msg, machine))
        elif kind == 'netlock':
            # KHOA MANG: may con thay may chu nhu dang tat (khong ket noi, khong chay offline)
            c.execute("UPDATE htp_presence SET netlock=true, cmd_by=%s, cmd_msg=%s, cmd_at=now()"
                      " WHERE machine=%s", (by, msg, machine))
        elif kind == 'netunlock':
            r = c.one("SELECT pc_name FROM htp_presence WHERE machine=%s", (machine,))
            c.execute("UPDATE htp_presence SET netlock=false, cmd_by=%s, cmd_at=now()"
                      " WHERE machine=%s OR lower(pc_name)=lower(%s)", (by, machine, (r[0] if r else '') or ''))
        elif kind == 'unblock':
            # mo khoa ca cac dong cung ten may (khoa theo ten may)
            r = c.one("SELECT pc_name FROM htp_presence WHERE machine=%s", (machine,))
            c.execute("UPDATE htp_presence SET blocked=false, cmd=NULL, cmd_by=%s, cmd_at=now()"
                      " WHERE machine=%s OR lower(pc_name)=lower(%s)", (by, machine, (r[0] if r else '') or ''))
    finally:
        if own:
            c.close()


def start(app_user='', username=''):
    """Goi sau khi dang nhap (che do may chu). Chay nen, loi mang -> tu thu lai."""
    global _THREAD
    import datastore
    if not datastore.is_server():
        return False
    _STATE['user'] = app_user
    _STATE['username'] = username or app_user
    if _THREAD is not None and _THREAD.is_alive():
        return True
    cfg = dict(datastore.server().cfg)
    route = getattr(datastore.server().c, 'route', 'LAN')
    _STOP.clear()
    _STATE['kicked'] = None
    mid = machine_id()

    def _cur_cfg():
        """Luon lay cau hinh MAY CHU DANG DUNG (chinh hoac du phong) -> nhip tim TU BAM THEO khi
        failover sang may chu du phong, khong can khoi dong lai phan mem."""
        try:
            s = datastore.server()
            if s is not None and getattr(s, 'cfg', None):
                return dict(s.cfg)
        except Exception:
            pass
        return dict(cfg)

    def run():
        c = None
        while not _STOP.is_set():
            try:
                if c is None:
                    c = _connect(_cur_cfg())         # ket noi toi may chu DANG DUNG (theo datastore)
                info = {'pc_name': socket.gethostname(), 'os_user': os_user(), 'app_user': _STATE['user'],
                        'username': _STATE['username'],
                        'route': getattr(c, 'route', route), 'os': f'{platform.system()} {platform.release()}'}
                _STATE['online'] = _beat(c, mid, info)
                _STATE['err'] = ''
                k = _poll_cmd(c, mid)
                if k:
                    _STATE['kicked'] = k
                    break                                # dung nhip tim -> bao offline ben duoi
            except Exception as e:
                _STATE['err'] = str(e)[:200]
                _STATE['online'] = None                 # mat ket noi -> KHONG giu so cu (tranh tieu de bao sai)
                try:
                    if c is not None:
                        c.close()
                except Exception:
                    pass
                c = None
            _STOP.wait(5 if _STATE.get('err') else HEARTBEAT_SEC)   # loi -> thu lai nhanh (bam theo may chu moi)
        try:                                     # dong phan mem -> bao offline ngay
            if c is None:
                c = _connect(_cur_cfg())
            c.execute("UPDATE htp_presence SET status='offline', last_seen=now() - interval '1 hour'"
                      " WHERE machine=%s", (mid,))
            c.close()
        except Exception:
            pass

    _THREAD = threading.Thread(target=run, name='htp-presence', daemon=True)
    _THREAD.start()
    return True


def set_user(app_user):
    _STATE['user'] = app_user


def stop(wait=1.5):
    _STOP.set()
    if _THREAD is not None:
        _THREAD.join(wait)


def online_count():
    """So may online (lan nhip tim gan nhat) hoac None (chua biet / che do file)."""
    return _STATE['online']


def list_machines(cfg=None):
    """-> [dict] tat ca may da tung ket noi, moi nhat truoc. 'online' = True/False, 'ago' = so giay."""
    import datastore
    own = False
    if cfg is not None:
        c = _connect(cfg)
        own = True
    else:
        if not datastore.is_server():
            return []
        c = datastore.server().c
        _ensure_schema(c)
    try:
        rows = c.query("SELECT machine, pc_name, os_user, app_user, ip_lan, ip_ts, route, os, app_ver,"
                       " to_char(started, 'DD/MM HH24:MI'), EXTRACT(EPOCH FROM (now() - last_seen))::int, status,"
                       " COALESCE(blocked, false), cmd_by, COALESCE(netlock, false), ip_rm, username"
                       " FROM htp_presence ORDER BY last_seen DESC")
    finally:
        if own:
            c.close()
    me = machine_id()
    out = []
    for r in rows:
        ago = int(r[10] or 0)
        out.append({'machine': r[0], 'pc_name': r[1] or '', 'os_user': r[2] or '', 'app_user': r[3] or '',
                    'ip_lan': r[4] or '', 'ip_ts': r[5] or '', 'route': r[6] or '', 'os': r[7] or '',
                    'app_ver': r[8] or '', 'started': r[9] or '', 'ago': ago,
                    'online': ago <= ONLINE_SEC and (r[11] or 'online') == 'online', 'me': r[0] == me,
                    'blocked': bool(r[12]), 'cmd_by': r[13] or '', 'netlock': bool(r[14]),
                    'ip_rm': r[15] or '', 'username': r[16] or ''})
    return out


def forget_offline(cfg=None, days=30):
    """Xoa cac may khong ket noi qua N ngay."""
    import datastore
    c = _connect(cfg) if cfg is not None else datastore.server().c
    try:
        c.execute("DELETE FROM htp_presence WHERE last_seen < now() - interval '" + str(int(days)) + " days'")
    finally:
        if cfg is not None:
            c.close()


def _ago_text(s):
    if s < 60:
        return f'{s} giây trước'
    if s < 3600:
        return f'{s // 60} phút trước'
    if s < 86400:
        return f'{s // 3600} giờ trước'
    return f'{s // 86400} ngày trước'


# =====================================================================
#  Cua so danh sach may
# =====================================================================
def server_hosts(cfg=None):
    """Dia chi may chu CSDL trong cau hinh (LAN + Tailscale), chu thuong."""
    if cfg is None:
        try:
            import datastore
            cfg = datastore.load_config() or {}
        except Exception:
            cfg = {}
    return [str(cfg.get(k) or '').strip().lower() for k in ('host', 'host_ts', 'host_rm')
            if str(cfg.get(k) or '').strip()]


def is_server_machine(r, hosts):
    """May (dong presence) co phai MAY CHU khong: IP LAN / IP Tailscale / ten may trung dia chi may chu.
    May chu cau hinh 'localhost' -> chinh may co dia chi do tu ket noi toi."""
    if not hosts:
        return False
    mine = {str(r.get(k) or '').strip().lower() for k in ('ip_lan', 'ip_ts', 'ip_rm', 'pc_name')} - {''}
    if mine & set(hosts):
        return True
    return bool(r.get('me')) and any(h in ('localhost', '127.0.0.1', '::1') for h in hosts)


def open_window(master, cfg=None, admin=False, by=''):
    """cfg=None: dung ket noi may chu dang chay. cfg=dict: dung thong tin nhap o cua so May chu du lieu.
    admin=True: hien nut Ngat ket noi / Khoa may / Mo khoa may con."""
    import tkinter as tk
    from tkinter import ttk, messagebox, simpledialog
    win = tk.Toplevel(master)
    win.title('Máy đang sử dụng phần mềm - HitechPlast')
    win.configure(bg='#f4f6f9')
    win.geometry('1280x600')
    F, FB = ('Segoe UI', 10), ('Segoe UI', 10, 'bold')
    top = tk.Frame(win, bg='#f4f6f9')
    top.pack(fill='x', padx=12, pady=(10, 4))
    tk.Label(top, text='👥 Máy tính đang dùng phần mềm', bg='#f4f6f9', font=('Segoe UI', 13, 'bold')).pack(side='left')
    lbl = tk.Label(top, text='', bg='#f4f6f9', font=FB, fg='#2d6a4f')
    lbl.pack(side='left', padx=14)
    v_all = tk.BooleanVar(value=False)
    tk.Checkbutton(top, text='Hiện cả máy offline', variable=v_all, bg='#f4f6f9', font=F,
                   command=lambda: refresh()).pack(side='right')

    cols = [('st', 'Trạng thái', 85), ('role', 'Vai trò', 105), ('pc', 'Tên máy', 150), ('au', 'Tài khoản phần mềm', 150),
            ('ou', 'User Windows', 110), ('lan', 'IP LAN', 110), ('ts', 'IP Tailscale', 110),
            ('rm', 'IP Radmin', 105),
            ('rt', 'Kết nối qua', 85), ('since', 'Mở lúc', 90), ('seen', 'Hoạt động cuối', 110)]
    fr = tk.Frame(win)
    fr.pack(fill='both', expand=True, padx=12)
    tv = ttk.Treeview(fr, columns=[c for c, _, _ in cols], show='headings', height=12)
    for c, t, w in cols:
        tv.heading(c, text=t)
        tv.column(c, width=w, anchor='w')
    sb = ttk.Scrollbar(fr, orient='vertical', command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    tv.pack(side='left', fill='both', expand=True)
    sb.pack(side='right', fill='y')
    tv.tag_configure('on', foreground='#1b5e20')
    tv.tag_configure('off', foreground='#999999')
    tv.tag_configure('me', background='#e3f2fd')
    tv.tag_configure('srv', background='#fff3cd', font=('Segoe UI', 10, 'bold'))
    tv.tag_configure('blk', background='#ffd6d6', foreground='#b00020')
    tv.tag_configure('nlk', background='#ffe8cc', foreground='#8a4b00')
    row_of = {}                                     # iid -> dong may

    tk.Label(win, text='🌐 Thiết bị trong mạng Tailscale (máy nào đang bật Tailscale):', bg='#f4f6f9',
             font=FB).pack(anchor='w', padx=12, pady=(8, 2))
    fr2 = tk.Frame(win)
    fr2.pack(fill='x', padx=12)
    tv2 = ttk.Treeview(fr2, columns=('st', 'role', 'name', 'ip', 'os'), show='headings', height=5)
    for c, t, w in (('st', 'Trạng thái', 90), ('role', 'Vai trò', 110), ('name', 'Tên thiết bị', 220),
                    ('ip', 'IP Tailscale', 140), ('os', 'Hệ điều hành', 120)):
        tv2.heading(c, text=t)
        tv2.column(c, width=w, anchor='w')
    tv2.pack(fill='x')
    tv2.tag_configure('on', foreground='#1b5e20')
    tv2.tag_configure('off', foreground='#999999')
    tv2.tag_configure('srv', background='#fff3cd', font=('Segoe UI', 10, 'bold'))

    note = tk.Label(win, text='', bg='#f4f6f9', fg='#666', font=('Segoe UI', 9), justify='left', wraplength=1000)
    note.pack(anchor='w', padx=12, pady=4)
    bf = tk.Frame(win, bg='#f4f6f9')
    bf.pack(fill='x', padx=12, pady=(0, 10))
    tk.Button(bf, text='🔄 Làm mới', font=FB, bg='#1d3557', fg='white', relief='flat', padx=10,
              command=lambda: refresh()).pack(side='left')
    tk.Button(bf, text='🧹 Xoá máy không dùng > 30 ngày', font=F, relief='flat', bg='#dfe6ee', padx=10,
              command=lambda: (forget_offline(cfg), refresh())).pack(side='left', padx=6)
    tk.Button(bf, text='Đóng', font=F, relief='flat', padx=12, command=win.destroy).pack(side='right')

    def _picked():
        sel = tv.selection()
        r = row_of.get(sel[0]) if sel else None
        if r is None:
            messagebox.showwarning('Chưa chọn máy', 'Chọn 1 máy trong bảng trên.', parent=win)
        return r

    def _act(kind):
        r = _picked()
        if r is None:
            return
        name = r['pc_name'] or r['machine']
        if kind != 'unblock' and r['me']:
            messagebox.showwarning('Không thực hiện', 'Đây là máy bạn đang dùng.', parent=win)
            return
        if kind in ('netlock', 'netunlock'):
            if kind == 'netunlock' and not r.get('netlock'):
                messagebox.showinfo('Mở mạng', f'Máy {name} không bị khoá mạng.', parent=win)
                return
            if kind == 'netlock' and r.get('_srv'):
                messagebox.showwarning('Không thực hiện', 'Không khoá mạng MÁY CHỦ.', parent=win)
                return
            q = (f'KHOÁ MẠNG máy {name}?\n\nMáy này sẽ KHÔNG kết nối được máy chủ – giống như máy chủ đang tắt '
                 '(phần mềm tự đóng trong ~20 giây, mở lại sẽ báo "Không kết nối được máy chủ dữ liệu").\n'
                 'Bấm "🌐 Mở mạng" để cho kết nối lại.' if kind == 'netlock' else
                 f'Mở mạng cho máy {name}? Máy này sẽ kết nối lại máy chủ được.')
            if not messagebox.askyesno('Khoá mạng' if kind == 'netlock' else 'Mở mạng', q,
                                       icon='warning' if kind == 'netlock' else 'question', parent=win):
                return
            try:
                send_cmd(r['machine'], kind, by=by, cfg=cfg)
            except Exception as e:
                messagebox.showerror('Lỗi', f'Không gửi được lệnh tới máy chủ:\n{e}', parent=win)
                return
            refresh(again=False)
            return
        if kind == 'unblock':
            if not r.get('blocked'):
                messagebox.showinfo('Mở khoá', f'Máy {name} không bị khoá.', parent=win)
                return
            if not messagebox.askyesno('Mở khoá máy', f'Mở khoá máy {name}?\nMáy này sẽ mở lại được phần mềm.',
                                       parent=win):
                return
            msg = ''
        else:
            warn = ('\n\n⚠ ĐÂY LÀ MÁY CHỦ: chỉ đóng phần mềm trên máy chủ, CSDL vẫn chạy.' if r.get('_srv') else '')
            title = 'Ngắt kết nối' if kind == 'kick' else 'KHOÁ máy'
            what = ('Phần mềm trên máy này sẽ tự đóng trong ~20 giây (mở lại được).' if kind == 'kick' else
                    'Phần mềm trên máy này sẽ tự đóng trong ~20 giây và KHÔNG mở lại được cho đến khi '
                    'quản trị viên bấm "Mở khoá".')
            if not messagebox.askyesno(title, f'{title} máy {name} ({r["ip_lan"]} / {r["ip_ts"]}) – '
                                              f'người dùng: {r["app_user"]} / {r["os_user"]}?\n\n{what}{warn}',
                                       icon='warning', parent=win):
                return
            msg = simpledialog.askstring(title, 'Lời nhắn hiện trên máy đó (có thể để trống):', parent=win) or ''
        try:
            send_cmd(r['machine'], kind, by=by, msg=msg, cfg=cfg)
        except Exception as e:
            messagebox.showerror('Lỗi', f'Không gửi được lệnh tới máy chủ:\n{e}', parent=win)
            return
        refresh(again=False)

    if admin:
        tk.Button(bf, text='🌐 Mở mạng', font=F, relief='flat', bg='#d0ebff', padx=10,
                  command=lambda: _act('netunlock')).pack(side='right', padx=4)
        tk.Button(bf, text='🚫 Khoá mạng', font=FB, relief='flat', bg='#8a4b00', fg='white', padx=10,
                  command=lambda: _act('netlock')).pack(side='right', padx=4)
        tk.Button(bf, text='🔓 Mở khoá', font=F, relief='flat', bg='#d8f3dc', padx=10,
                  command=lambda: _act('unblock')).pack(side='right', padx=4)
        tk.Button(bf, text='🔒 Khoá máy', font=FB, relief='flat', bg='#b00020', fg='white', padx=10,
                  command=lambda: _act('block')).pack(side='right', padx=4)
        tk.Button(bf, text='⛔ Ngắt kết nối', font=FB, relief='flat', bg='#e07a00', fg='white', padx=10,
                  command=lambda: _act('kick')).pack(side='right', padx=4)

    def refresh(again=True):
        try:
            if not win.winfo_exists():
                return
        except Exception:
            return
        tv.delete(*tv.get_children())
        row_of.clear()
        tv2.delete(*tv2.get_children())
        err = ''
        try:
            rows = list_machines(cfg)
        except Exception as e:
            rows, err = [], str(e)
        hosts = server_hosts(cfg)
        for r in rows:
            r['_srv'] = is_server_machine(r, hosts)
        rows.sort(key=lambda r: (not r['_srv'], not r['online']))      # may chu len dau
        n_on = n_srv_on = 0
        for r in rows:
            if r['online']:
                n_on += 1
                n_srv_on += r['_srv']
            elif not v_all.get() and not r.get('blocked') and not r.get('netlock'):
                continue
            tags = ('on' if r['online'] else 'off',) + (('srv',) if r['_srv'] else (('me',) if r['me'] else ()))
            if r.get('blocked'):
                tags += ('blk',)
            elif r.get('netlock'):
                tags += ('nlk',)
            st_txt = ('🔒 Bị khoá' if r.get('blocked') else '🚫 Khoá mạng' if r.get('netlock') else
                      ('🟢 Online' if r['online'] else '⚪ Offline'))
            iid = tv.insert('', 'end', tags=tags, values=(
                st_txt,
                '🖥 MÁY CHỦ' if r['_srv'] else '💻 Máy con',
                r['pc_name'] + ('  (máy này)' if r['me'] else ''), r['app_user'], r['os_user'],
                r['ip_lan'], r['ip_ts'], r.get('ip_rm', ''), r['route'], r['started'], _ago_text(r['ago'])))
            row_of[iid] = r
        if not err:
            lbl.config(text=f'{n_on} máy đang online ({n_srv_on} máy chủ + {n_on - n_srv_on} máy con)'
                            f' / {len(rows)} máy đã từng kết nối')
        else:
            lbl.config(text='')
        try:
            import tailscale_util as TS
            st = TS.status()
            peers = list(st['peers'])
            if st.get('ip'):                                   # them chinh may nay vao danh sach
                peers.insert(0, {'name': (st.get('name') or socket.gethostname()) + '  (máy này)',
                                 'ip': st['ip'], 'online': bool(st['running']), 'os': platform.system().lower(),
                                 '_me': True})
            for p in peers:
                srv = is_server_machine({'ip_ts': p['ip'], 'pc_name': p['name'].replace('  (máy này)', ''),
                                         'me': p.get('_me')}, hosts)
                tv2.insert('', 'end' if not srv else 0,
                           tags=('on' if p['online'] else 'off',) + (('srv',) if srv else ()),
                           values=('🟢 Bật' if p['online'] else '⚪ Tắt', '🖥 MÁY CHỦ' if srv else '💻 Máy con',
                                   p['name'], p['ip'], p['os']))
            ts_txt = (f"Tailscale máy này: {'BẬT' if st['running'] else 'TẮT'} {st['ip'] or ''}. "
                      if st['installed'] or st['running'] else 'Máy này chưa cài Tailscale. ')
        except Exception as e:
            ts_txt = f'Không đọc được Tailscale: {e}. '
        srv_txt = ('Máy chủ dữ liệu (theo cấu hình): %s. ' % ' / '.join(hosts)) if hosts else \
            'Máy này chưa cấu hình máy chủ dữ liệu. '
        if hosts and not any(r['_srv'] and r['online'] for r in rows):
            srv_txt += '(Máy chủ hiện không mở phần mềm – CSDL vẫn có thể đang chạy.) '
        note.config(text=(('✖ ' + err + '\n') if err else '') + srv_txt + ts_txt +
                    f'Mỗi máy mở phần mềm (chế độ máy chủ) tự báo mỗi {HEARTBEAT_SEC} giây; '
                    f'quá {ONLINE_SEC} giây không báo = offline.')
        if admin:
            note.config(text=note.cget('text') + '\nQuản trị: chọn 1 máy → ⛔ Ngắt kết nối (đóng phần mềm, mở lại '
                                                 'được) | 🔒 Khoá máy (đóng và không mở lại được) | 🔓 Mở khoá |'
                                                 ' 🚫 Khoá mạng (máy đó thấy máy chủ như đang tắt) | 🌐 Mở mạng.')
        if again:
            win.after(10000, refresh)

    win.after(50, refresh)
    return win
