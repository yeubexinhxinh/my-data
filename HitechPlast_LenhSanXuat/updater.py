# -*- coding: utf-8 -*-
"""
CAP NHAT PHAN MEM (UPDATE) - HitechPlast
========================================
Y tuong: moi PHIEN BAN = anh chup TOAN BO thu muc ban build (LenhSanXuat_HitechPlast.exe + _internal\\
+ file mau...). Moi file luu tren may chu 1 lan theo ma SHA-256 (file giong nhau giua cac ban KHONG luu lai).

  📤 DANG BAN (admin, may nao cung duoc):
       chon thu muc ban build (vd GOI_CAI_DAT\\HitechPlast_LenhSanXuat, hoac chinh thu muc phan mem
       dang chay neu la ban EXE) -> phan mem quet, chi TAI LEN cac file may chu chua co.
  ⬆ CAP NHAT (moi may):
       phan mem tu KIEM TRA tung file tren may minh so voi phien ban da chon ->
       CHI TAI cac file THIEU / KHAC -> sao luu file cu -> dong phan mem -> chep -> mo lai.

Ban EXE chi nhan ban build EXE; may chay ma nguon Python chi nhan ban ma nguon (khong lan lon).
KHONG BAO GIO dong goi / ghi de du lieu rieng cua tung may (data.json, sheets.json, qc_data.db,
users.dat, settings.json, server.json, Du_lieu_offline...) - xem EXCLUDE_*.
Khong co may chu: "💾 Luu ra file goi" (.htpupd) -> may kia "📥 Cai tu file goi" (cung chi chep file khac).
"""
import fnmatch
import hashlib
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import threading
import traceback
import zipfile
import zlib
from datetime import datetime

MAIN_EXE = 'LenhSanXuat_HitechPlast.exe'
MAIN_PY = 'app_hitechplast.py'
TMP_DIR = '_update_tmp'
BACKUP_DIR = '_update_backup'
VERSION_FILE = 'version.json'
STATE_FILE = '_update_state.json'
PKG_EXT = '.htpupd'
CHUNK = 1024 * 1024                  # 1 MB / dong tren may chu
CHECK_EVERY_MS = 5 * 60 * 1000       # 5 phut kiem tra ban moi 1 lan
KEEP_BACKUPS = 3
FORMAT = 2

# thu muc KHONG BAO GIO dong goi (so sanh chu thuong) - ngoai _internal
EXCLUDE_DIRS = {'__pycache__', 'du_lieu_offline', 'du_lieu_theo_nam', TMP_DIR.lower(), BACKUP_DIR.lower(),
                'app_patch', 'build', 'dist', 'goi_cai_dat', '.git', '.idea', '.vscode', 'venv',
                '.venv', 'sao_luu', 'backup'}
# file KHONG BAO GIO dong goi (so theo TEN file, chu thuong) - du lieu rieng cua tung may
EXCLUDE_FILES = ['data.json', 'sheets.json', '*.db', '*.db-journal', '*.db-wal', '*.db-shm',
                 'users.dat', 'users.json', 'users_base.dat', 'settings.json', 'server.json',
                 'server.json.*', 'server_link.dat', VERSION_FILE, STATE_FILE, 'doc_params.json',
                 'claim_khach_hang.json', '*.jsonl', '*.log', '*.pyc', '*.zip', '*.rar', '*.7z',
                 '*.bak', '*.tmp', 'lang_missing.json', 'lang_auto.json', 'rmqr_settings.json',
                 'tem_btw.json', 'tem_designs.json', 'machine_id*', '*' + PKG_EXT, '~$*',
                 'thumbs.db', 'desktop.ini', 'meta.json', 'base.json', 'apply_update.bat']

FONT = ('Segoe UI', 10)
FONT_B = ('Segoe UI', 10, 'bold')
FONT_T = ('Segoe UI', 13, 'bold')
FONT_S = ('Segoe UI', 9)

TARGET_LABEL = {'exe': 'Bản EXE', 'py': 'Mã nguồn Python'}


# =====================================================================
#  Duong dan / phien ban
# =====================================================================
def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0].endswith('.py')
                                           else __file__))


def run_mode():
    """'exe' (ban dong goi PyInstaller) | 'py' (chay ma nguon python)."""
    return 'exe' if getattr(sys, 'frozen', False) else 'py'


def folder_target(root):
    """Thu muc co file .exe chinh -> 'exe' (ban build), nguoc lai 'py' (ma nguon)."""
    return 'exe' if os.path.isfile(os.path.join(root, MAIN_EXE)) else 'py'


def default_publish_dir():
    """Thu muc nen dang: ban EXE -> chinh thu muc dang chay; ma nguon -> ban build vua tao (neu co)."""
    root = app_dir()
    if run_mode() == 'exe':
        return root
    for sub in (os.path.join('GOI_CAI_DAT', 'HitechPlast_LenhSanXuat'), os.path.join('dist', 'LenhSanXuat_HitechPlast')):
        p = os.path.join(root, sub)
        if os.path.isfile(os.path.join(p, MAIN_EXE)):
            return p
    return root


def ver_key(v):
    """'2026.10.02.3' -> (2026, 10, 2, 3) de so sanh. '' / None -> ()."""
    return tuple(int(x) for x in re.findall(r'\d+', str(v or '')))


def _rjson(p, default=None):
    try:
        with open(p, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def _wjson(p, obj):
    tmp = p + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, p)


def read_local():
    return _rjson(os.path.join(app_dir(), VERSION_FILE), {}) or {}


def write_local(d):
    _wjson(os.path.join(app_dir(), VERSION_FILE), d)


def current_version():
    return read_local().get('version') or '0'


def suggest_version(existing=()):
    """Phien ban de xuat: ngay hom nay (trung thi them .2 .3 ...), luon LON HON moi ban da co
    (dong ho may dang lech ngay cung khong tao ra ban 'cu hon')."""
    have = {str(v) for v in existing}
    top = max(have, key=ver_key) if have else ''
    base = datetime.now().strftime('%Y.%m.%d')
    if top and ver_key(base) <= ver_key(top):
        base = '.'.join(top.split('.')[:3])
    cand, n = base, 1
    while cand in have or (top and ver_key(cand) <= ver_key(top)):
        n += 1
        cand = f'{base}.{n}'
    return cand


def is_excluded(rel):
    parts = rel.replace('\\', '/').split('/')
    if parts[0].lower() == '_internal':           # ban build: lay het (tru __pycache__)
        return any(p.lower() == '__pycache__' for p in parts[:-1])
    for d in parts[:-1]:
        dl = d.lower()
        if dl in EXCLUDE_DIRS or dl.startswith('app_patch') or dl.startswith('_update'):
            return True
    name = parts[-1].lower()
    return any(fnmatch.fnmatch(name, p) for p in EXCLUDE_FILES)


def sha256_file(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def sha256_bytes(b):
    return hashlib.sha256(b).hexdigest()


def fmt_size(n):
    n = float(n or 0)
    for u in ('B', 'KB', 'MB', 'GB'):
        if n < 1024 or u == 'GB':
            return f'{n:.0f} {u}' if u == 'B' else f'{n:.1f} {u}'
        n /= 1024


def _state():
    return _rjson(os.path.join(app_dir(), STATE_FILE), {}) or {}


def _save_state(st):
    try:
        _wjson(os.path.join(app_dir(), STATE_FILE), st)
    except Exception:
        pass


def classify(rel):
    return 'exe' if rel.replace('\\', '/').lower().startswith('_internal/') or rel.lower() == MAIN_EXE.lower() \
        else 'py' if rel.lower().endswith('.py') else 'any'


# =====================================================================
#  So sanh: may nay THIEU / KHAC file nao so voi phien ban
# =====================================================================
def missing_files(manifest, local):
    """manifest: phien ban. local: ket qua scan_folder(app_dir()).
    -> danh sach file (entry trong manifest) can tai: chua co hoac khac SHA-256."""
    out = []
    for f in manifest.get('files', []):
        lv = local.get(f['path'])
        if lv is None or lv['sha'] != f['sha256']:
            out.append(f)
    return out


def diff_status(scan, prev_manifest):
    """Trang thai tung file khi DANG: 'Mới' | 'Đã đổi' | '' (giong ban truoc) | '—' (chua co ban truoc)."""
    prev = {f['path']: f['sha256'] for f in (prev_manifest or {}).get('files', [])}
    st = {}
    for r, v in scan.items():
        if not prev_manifest:
            st[r] = '—'
        elif r not in prev:
            st[r] = 'Mới'
        elif prev[r] != v['sha']:
            st[r] = 'Đã đổi'
        else:
            st[r] = ''
    return st


# =====================================================================
#  May chu (PostgreSQL): phien ban + kho file theo SHA-256
# =====================================================================
SCHEMA = [
    """CREATE TABLE IF NOT EXISTS htp_updates(
        version TEXT PRIMARY KEY,
        created_at TIMESTAMPTZ DEFAULT now(),
        author TEXT, pc_name TEXT, title TEXT, notes TEXT, kind TEXT,
        manifest TEXT, size BIGINT, sha256 TEXT, nchunks INT,
        active BOOLEAN DEFAULT TRUE)""",
    "ALTER TABLE htp_updates ADD COLUMN IF NOT EXISTS target TEXT",
    "ALTER TABLE htp_updates ADD COLUMN IF NOT EXISTS fmt INT",
    """CREATE TABLE IF NOT EXISTS htp_upd_blob(
        sha TEXT NOT NULL, idx INT NOT NULL, data BYTEA, PRIMARY KEY(sha, idx))""",
    """CREATE TABLE IF NOT EXISTS htp_upd_blobok(
        sha TEXT PRIMARY KEY, size BIGINT, csize BIGINT, nchunks INT, created_at TIMESTAMPTZ DEFAULT now())""",
]
_SCHEMA_OK = [False]


def server_ok():
    try:
        import datastore
        return datastore.is_server() and not datastore.is_offline()
    except Exception:
        return False


def _shared_conn():
    import datastore
    return datastore.server().c


def _own_conn():
    """Ket noi RIENG cho viec nang (tai / dang) -> khong giu khoa ket noi chung cua giao dien."""
    import datastore
    return datastore.Conn(dict(datastore.server().cfg))


def _ensure(c):
    if not _SCHEMA_OK[0]:
        for s in SCHEMA:
            c.execute(s)
        _SCHEMA_OK[0] = True


def remote_list(c=None):
    """-> [manifest] moi nhat truoc (chi phien ban dinh dang moi)."""
    c = c or _shared_conn()
    _ensure(c)
    rows = c.query("SELECT version, to_char(created_at, 'YYYY-MM-DD HH24:MI'), author, pc_name, title, notes,"
                   " manifest, size, active, target FROM htp_updates WHERE fmt = %s", (FORMAT,))
    out = []
    for r in rows:
        try:
            man = json.loads(r[6] or '{}')
        except Exception:
            continue
        man.update({'version': r[0], 'created': r[1] or man.get('created', ''), 'author': r[2] or '',
                    'pc': r[3] or '', 'title': r[4] or '', 'notes': r[5] or '', 'size': int(r[7] or 0),
                    'active': bool(r[8]), 'target': r[9] or man.get('target', 'exe')})
        out.append(man)
    out.sort(key=lambda m: ver_key(m['version']), reverse=True)
    return out


def remote_latest(target, c=None):
    """Ban moi nhat (dang bat) danh cho loai may target -> version | None. Truy van nhe."""
    c = c or _shared_conn()
    _ensure(c)
    vs = [r[0] for r in c.query("SELECT version FROM htp_updates WHERE active IS TRUE AND fmt=%s AND target=%s",
                                (FORMAT, target))]
    return max(vs, key=ver_key) if vs else None


def remote_have(shas, c):
    """-> tap SHA da co DAY DU tren may chu."""
    shas = list(shas)
    have = set()
    for i in range(0, len(shas), 500):
        have |= {r[0] for r in c.query("SELECT sha FROM htp_upd_blobok WHERE sha = ANY(%s)", (shas[i:i + 500],))}
    return have


def _upload_blob(c, sha, path):
    with open(path, 'rb') as f:
        raw = f.read()
    if sha256_bytes(raw) != sha:
        raise RuntimeError(f'File {path} vừa bị thay đổi trong lúc đăng. Hãy bấm Quét lại.')
    comp = zlib.compress(raw, 6)
    n = max(1, (len(comp) + CHUNK - 1) // CHUNK)
    c.execute("DELETE FROM htp_upd_blob WHERE sha=%s", (sha,))          # sot lai tu lan dang loi truoc
    for i in range(n):
        c.execute("INSERT INTO htp_upd_blob(sha, idx, data) VALUES(%s, %s, %s)", (sha, i, comp[i * CHUNK:(i + 1) * CHUNK]))
    c.execute("INSERT INTO htp_upd_blobok(sha, size, csize, nchunks) VALUES(%s, %s, %s, %s)"
              " ON CONFLICT (sha) DO NOTHING", (sha, len(raw), len(comp), n))
    return len(raw)


def download_blob(c, sha):
    r = c.one("SELECT nchunks FROM htp_upd_blobok WHERE sha=%s", (sha,))
    if not r:
        raise RuntimeError(f'Máy chủ thiếu file có mã {sha[:12]}… – hãy đăng lại phiên bản.')
    parts = []
    for i in range(int(r[0])):
        x = c.one("SELECT data FROM htp_upd_blob WHERE sha=%s AND idx=%s", (sha, i))
        if x is None:
            raise RuntimeError(f'File {sha[:12]}… trên máy chủ bị thiếu phần {i + 1}.')
        parts.append(bytes(x[0]))
    raw = zlib.decompress(b''.join(parts))
    if sha256_bytes(raw) != sha:
        raise RuntimeError(f'File {sha[:12]}… tải về bị lỗi (sai mã kiểm tra SHA-256). Hãy thử lại.')
    return raw


def publish(root, rels, scan, version, title, notes, author, progress=None):
    """Dang phien ban = anh chup cac file rels trong root. Chi tai len file may chu chua co.
    -> (manifest, so file da tai len, dung luong da tai len)"""
    root = os.path.abspath(root)
    c = _own_conn()
    try:
        _ensure(c)
        if c.one("SELECT 1 FROM htp_updates WHERE version=%s", (version,)):
            raise RuntimeError(f'Phiên bản {version} đã có trên máy chủ. Hãy đổi số phiên bản.')
        files = [{'path': r, 'sha256': scan[r]['sha'], 'size': scan[r]['size']} for r in sorted(rels)]
        have = remote_have({f['sha256'] for f in files}, c)
        todo, seen = [], set()
        for f in files:
            if f['sha256'] not in have and f['sha256'] not in seen:
                todo.append(f)
                seen.add(f['sha256'])
        total = sum(f['size'] for f in todo) or 1
        done = 0
        for f in todo:
            if progress:
                progress(done, total, f'Đang tải lên {f["path"]}')
            done += _upload_blob(c, f['sha256'], os.path.join(root, f['path']))
            if progress:
                progress('__file__', f['sha256'])     # bao UI: file nay da tai len xong
        if progress:
            progress(total, total, 'Đang ghi phiên bản...')
        target = folder_target(root)
        man = {'format': FORMAT, 'version': version, 'target': target, 'title': title,
               'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'author': author,
               'pc': socket.gethostname(), 'files': files}
        tsize = sum(f['size'] for f in files)
        c.execute("INSERT INTO htp_updates(version, author, pc_name, title, notes, kind, manifest, size, active,"
                  " target, fmt) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,TRUE,%s,%s)",
                  (version, author, socket.gethostname(), title, notes, 'full',
                   json.dumps(man, ensure_ascii=False), tsize, target, FORMAT))
        man['notes'] = notes
        return man, len(todo), sum(f['size'] for f in todo)
    finally:
        c.close()


def remote_set_active(version, flag):
    _shared_conn().execute("UPDATE htp_updates SET active=%s WHERE version=%s", (bool(flag), version))


def remote_delete(version):
    """Xoa phien ban + cac file KHONG con phien ban nao dung."""
    c = _shared_conn()
    c.execute("DELETE FROM htp_updates WHERE version=%s", (version,))
    used = set()
    for (m,) in c.query("SELECT manifest FROM htp_updates WHERE fmt=%s", (FORMAT,)):
        try:
            used |= {f['sha256'] for f in json.loads(m).get('files', [])}
        except Exception:
            pass
    orphan = [r[0] for r in c.query("SELECT sha FROM htp_upd_blobok") if r[0] not in used]
    for i in range(0, len(orphan), 500):
        part = orphan[i:i + 500]
        c.execute("DELETE FROM htp_upd_blob WHERE sha = ANY(%s)", (part,))
        c.execute("DELETE FROM htp_upd_blobok WHERE sha = ANY(%s)", (part,))
    return len(orphan)


# =====================================================================
#  Goi file (khong co may chu)
# =====================================================================
def export_package(man, out_path, source, progress=None):
    """Ghi goi .htpupd day du (moi file cua phien ban). source(sha, path) -> bytes."""
    files = man['files']
    total = sum(f['size'] for f in files) or 1
    done = 0
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        light = {k: v for k, v in man.items()}
        z.writestr('_manifest.json', json.dumps(light, ensure_ascii=False, indent=1))
        for f in files:
            if progress:
                progress(done, total, f'Đang ghi {f["path"]}')
            z.writestr('files/' + f['path'], source(f['sha256'], f['path']))
            done += f['size']


def read_package_file(path):
    z = zipfile.ZipFile(path)
    man = json.loads(z.read('_manifest.json').decode('utf-8'))
    return man, z


# =====================================================================
#  Chuan bi cai dat: chi lay file thieu / khac -> giai nen -> sao luu
# =====================================================================
def prepare_install(man, need, fetch, progress=None):
    """need: cac file (entry manifest) can lay. fetch(entry) -> bytes (da kiem tra SHA).
    -> job cho write_apply_script / apply_now_python."""
    root = app_dir()
    if not _check_writable(root):
        raise RuntimeError('Không ghi được vào thư mục phần mềm:\n' + root +
                           '\n\nHãy chạy phần mềm bằng quyền Administrator hoặc đặt thư mục phần mềm ở ổ D:\\')
    target = man['version']
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    stage = os.path.join(root, TMP_DIR, 'stage_' + ts)
    os.makedirs(stage, exist_ok=True)
    total = sum(f['size'] for f in need) or 1
    done = 0
    for f in need:
        if progress:
            progress(done, total, f'Đang tải {f["path"]}')
        data = fetch(f)
        if sha256_bytes(data) != f['sha256']:
            raise RuntimeError(f"File {f['path']} tải về bị lỗi (sai SHA-256). Hãy thử lại.")
        out = os.path.join(stage, f['path'])
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with open(out, 'wb') as fo:
            fo.write(data)
        done += f['size']
    if progress:
        progress(total, total, 'Đang sao lưu file cũ...')

    # ---- sao luu file sap bi ghi de ----
    local = read_local()
    cur = local.get('version') or '0'
    bdir = os.path.join(root, BACKUP_DIR, f'{cur}__{ts}')
    bfiles = os.path.join(bdir, 'files')
    os.makedirs(bfiles, exist_ok=True)
    new_files = []
    for f in need:
        src = os.path.join(root, f['path'])
        if os.path.isfile(src):
            dst = os.path.join(bfiles, f['path'])
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
        else:
            new_files.append(f['path'])
    vf = os.path.join(root, VERSION_FILE)
    if os.path.isfile(vf):
        shutil.copy2(vf, os.path.join(bfiles, VERSION_FILE))
    _wjson(os.path.join(bdir, '_backup.json'), {
        'from_version': cur, 'to_version': target, 'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'new_files': new_files, 'had_version_file': os.path.isfile(vf)})
    _prune_backups(root)

    # ---- version.json moi (chep cung cac file) ----
    hist = list(local.get('history') or [])[-30:]
    hist.append({'version': target, 'at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'from': cur,
                 'files': len(need)})
    newv = dict(local)
    newv.update({'version': target, 'installed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                 'mode': run_mode(), 'history': hist,
                 'notice': {'version': target, 'title': man.get('title', ''), 'notes': man.get('notes', ''),
                            'files': len(need)}})
    _wjson(os.path.join(stage, VERSION_FILE), newv)
    return {'stage': stage, 'delete': [], 'delete_files': [], 'backup': bdir, 'target': target,
            'count': len(need)}


def mark_installed(version):
    """May da giong het phien ban (khong can chep file) -> chi ghi so phien ban."""
    loc = read_local()
    loc['version'] = version
    loc['installed_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    loc['mode'] = run_mode()
    write_local(loc)


# =====================================================================
#  Quet thu muc (co bo nho dem SHA theo kich thuoc + gio sua)
# =====================================================================
def scan_folder(root, progress=None):
    """-> {rel: {'size','mtime','sha','target'}} (bo qua file du lieu). Dung cache (size, mtime)."""
    root = os.path.abspath(root)
    st = _state()
    cache = (st.get('hash_cache') or {}).get(root, {})
    out = {}
    all_files = []
    for dp, dns, fns in os.walk(root):
        in_internal = os.path.relpath(dp, root).replace('\\', '/').lower().split('/')[0] == '_internal'
        if in_internal:
            dns[:] = [d for d in dns if d.lower() != '__pycache__']
        else:
            dns[:] = [d for d in dns if d.lower() not in EXCLUDE_DIRS
                      and not d.lower().startswith('app_patch') and not d.lower().startswith('_update')]
        for fn in fns:
            ap = os.path.join(dp, fn)
            rel = os.path.relpath(ap, root).replace('\\', '/')
            if not is_excluded(rel):
                all_files.append((ap, rel))
    n = len(all_files)
    for i, (ap, rel) in enumerate(all_files):
        try:
            s = os.stat(ap)
        except OSError:
            continue
        c = cache.get(rel)
        if c and c[0] == s.st_size and abs(c[1] - s.st_mtime) < 0.001:
            sha = c[2]
        else:
            try:
                sha = sha256_file(ap)
            except OSError:
                continue
        out[rel] = {'size': s.st_size, 'mtime': s.st_mtime, 'sha': sha, 'target': classify(rel)}
        if progress and (i % 25 == 0 or i == n - 1):
            progress(i + 1, n, rel)
    st = _state()
    st.setdefault('hash_cache', {})[root] = {r: [v['size'], v['mtime'], v['sha']] for r, v in out.items()}
    _save_state(st)
    return out



def _check_writable(root):
    p = os.path.join(root, TMP_DIR)
    try:
        os.makedirs(p, exist_ok=True)
        t = os.path.join(p, '.w')
        with open(t, 'w') as f:
            f.write('1')
        os.remove(t)
        return True
    except OSError:
        return False



def _prune_backups(root):
    b = os.path.join(root, BACKUP_DIR)
    try:
        ds = sorted((d for d in os.listdir(b) if os.path.isdir(os.path.join(b, d))),
                    key=lambda d: os.path.getmtime(os.path.join(b, d)))
        for d in ds[:-KEEP_BACKUPS]:
            shutil.rmtree(os.path.join(b, d), ignore_errors=True)
    except OSError:
        pass


def list_backups():
    b = os.path.join(app_dir(), BACKUP_DIR)
    out = []
    try:
        for d in os.listdir(b):
            info = _rjson(os.path.join(b, d, '_backup.json'))
            if info:
                info['dir'] = os.path.join(b, d)
                out.append(info)
    except OSError:
        pass
    out.sort(key=lambda x: x.get('created', ''), reverse=True)
    return out


def prepare_restore(bk):
    """Chep lai ban sao luu bk (tu list_backups)."""
    files = os.path.join(bk['dir'], 'files')
    dels = list(bk.get('new_files') or [])
    if not bk.get('had_version_file'):
        dels.append(VERSION_FILE)
    else:
        v = _rjson(os.path.join(files, VERSION_FILE), {}) or {}
        v['notice'] = {'version': v.get('version', '0'), 'title': 'Đã khôi phục bản trước', 'notes': ''}
        _wjson(os.path.join(files, VERSION_FILE), v)
    return {'stage': files, 'delete': [], 'delete_files': dels, 'backup': None,
            'target': bk.get('from_version', '0'), 'count': -1, 'keep_stage': True}


def _restart_cmd():
    root = app_dir()
    if run_mode() == 'exe':
        return f'start "" "{sys.executable}"'
    exe = sys.executable
    w = os.path.join(os.path.dirname(exe), 'pythonw.exe')
    if os.path.isfile(w):
        exe = w
    main = os.path.join(root, MAIN_PY)
    return f'start "" "{exe}" "{main}"'


def write_apply_script(job):
    """Tao file .bat: doi phan mem dong -> xoa / chep file -> mo lai."""
    root = app_dir()
    pid = os.getpid()
    tmp = os.path.join(root, TMP_DIR)
    os.makedirs(tmp, exist_ok=True)
    bat = os.path.join(tmp, 'apply_update.bat')
    err = os.path.join(tmp, 'last_error.txt')
    L = ['@echo off', 'chcp 65001 >nul', 'setlocal', f'cd /d "{root}"',
         'set /a N=0',
         ':wait',
         f'tasklist /FI "PID eq {pid}" /NH 2>nul | find "{pid}" >nul',
         'if errorlevel 1 goto closed',
         'set /a N+=1',
         'if %N% GEQ 90 goto closed',
         'timeout /t 1 /nobreak >nul',
         'goto wait',
         ':closed',
         'timeout /t 1 /nobreak >nul']
    def W(rel):                                   # duong dan kieu Windows (dau / bi del hieu la tham so)
        return os.path.join(root, rel).replace('/', '\\') if os.name == 'nt' else os.path.join(root, rel)
    for d in job.get('delete') or []:
        L.append(f'if exist "{W(d)}" rmdir /s /q "{W(d)}"')
    for f in job.get('delete_files') or []:
        L.append(f'if exist "{W(f)}" del /f /q "{W(f)}"')
    L += [f'robocopy "{job["stage"]}" "{root}" /E /R:15 /W:1 /NFL /NDL /NJH /NJS /NP >nul',
          'if %ERRORLEVEL% GEQ 8 (',
          f'  echo Loi chep file khi cap nhat len ban {job["target"]} - ma loi robocopy %ERRORLEVEL% > "{err}"',
          ')']
    if not job.get('keep_stage'):
        L.append(f'rmdir /s /q "{job["stage"]}"')
    L += [_restart_cmd(), 'endlocal', '(goto) 2>nul & del "%~f0"']
    with open(bat, 'w', encoding='utf-8', newline='\r\n') as f:
        f.write('\n'.join(L) + '\n')
    return bat


def apply_now_python(job):
    """Chep file NGAY (khong qua .bat) - dung khi khong phai Windows (kiem thu)."""
    root = app_dir()
    for d in job.get('delete') or []:
        shutil.rmtree(os.path.join(root, d), ignore_errors=True)
    for f in job.get('delete_files') or []:
        try:
            os.remove(os.path.join(root, f))
        except OSError:
            pass
    for dp, _d, fns in os.walk(job['stage']):
        for fn in fns:
            src = os.path.join(dp, fn)
            rel = os.path.relpath(src, job['stage'])
            dst = os.path.join(root, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
    if not job.get('keep_stage'):
        shutil.rmtree(job['stage'], ignore_errors=True)


def launch_and_quit(app, job):
    """Chay .bat roi dong phan mem (bat se doi phan mem tat han moi chep)."""
    if os.name != 'nt':
        apply_now_python(job)
        _quit_app(app)
        return
    bat = write_apply_script(job)
    flags = 0x08000000 | 0x00000200          # CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(['cmd', '/c', bat], cwd=app_dir(), creationflags=flags, close_fds=True)
    _quit_app(app)


def _quit_app(app):
    try:
        import presence
        presence.stop()
    except Exception:
        pass
    try:
        app.restart_flag = False
        app.logout_flag = False
        app.destroy()
    except Exception:
        pass



def run_bg(widget, fn, done, progress=None):
    q = queue.Queue()

    def prog(*a):
        q.put(('p', a))

    def work():
        try:
            q.put(('ok', fn(prog)))
        except Exception as e:
            q.put(('err', e, traceback.format_exc()))

    threading.Thread(target=work, daemon=True).start()

    def poll():
        try:
            while True:
                item = q.get_nowait()
                if item[0] == 'p':
                    if progress:
                        try:
                            progress(*item[1])
                        except Exception:
                            pass
                else:
                    done(item)
                    return
        except queue.Empty:
            pass
        try:
            if widget.winfo_exists():
                widget.after(120, poll)
        except Exception:
            pass
    widget.after(120, poll)


# =====================================================================
#  GIAO DIEN
# =====================================================================
def _is_admin(app):
    try:
        return (app.user or {}).get('tabs') == 'all'
    except Exception:
        return False


def _who(app):
    try:
        u = app.user or {}
        return u.get('hoten') or u.get('username') or ''
    except Exception:
        return ''


def _btn_text():
    cur = current_version()
    return f'⬆ Cập nhật  v{cur}' if cur != '0' else '⬆ Cập nhật'


def install_button(parent, app, bg='#1f3a5f'):
    """Nut '⬆ Cap nhat' tren thanh tieu de + tu kiem tra ban moi (chay nen)."""
    import tkinter as tk
    btn = tk.Button(parent, text=_btn_text(), font=FONT, bg='#34495e', fg='white', activebackground='#2c3e50',
                    activeforeground='white', relief='flat', padx=10, pady=4, cursor='hand2',
                    command=lambda: UpdateWindow(app))
    app._upd_btn = btn
    app.after(4000, lambda: _check_loop(app))
    return btn


def _set_badge(app, newest):
    b = getattr(app, '_upd_btn', None)
    if b is None:
        return
    try:
        if not b.winfo_exists():
            return
        if newest and ver_key(newest) > ver_key(current_version()):
            b.config(text=f'🆕 Có bản mới {newest}', bg='#e67e22', activebackground='#d35400')
        else:
            b.config(text=_btn_text(), bg='#34495e', activebackground='#2c3e50')
    except Exception:
        pass


def _check_loop(app):
    try:
        if not app.winfo_exists():
            return
    except Exception:
        return
    if server_ok():
        def done(item):
            if item[0] == 'ok':
                _set_badge(app, item[1])
        run_bg(app, lambda prog: remote_latest(run_mode()), done)
    try:
        app.after(CHECK_EVERY_MS, lambda: _check_loop(app))
    except Exception:
        pass


def after_start(app):
    """Goi 1 lan sau khi mo phan mem: bao 'Da cap nhat len ban X' / loi cap nhat lan truoc; don rac."""
    from tkinter import messagebox
    root = app_dir()
    err = os.path.join(root, TMP_DIR, 'last_error.txt')
    if os.path.isfile(err):
        try:
            with open(err, 'r', encoding='utf-8', errors='replace') as f:
                txt = f.read().strip()
            os.remove(err)
        except OSError:
            txt = ''
        messagebox.showwarning('Cập nhật chưa hoàn tất',
                               'Lần cập nhật trước có file chưa chép được (có thể do đang bị mở / '
                               'phần mềm diệt virus chặn):\n\n' + txt +
                               '\n\nHãy bấm ⬆ Cập nhật và cài lại – phần mềm chỉ tải lại file còn thiếu.',
                               parent=app)
    loc = read_local()
    n = loc.get('notice')
    if n:
        loc.pop('notice', None)
        try:
            write_local(loc)
        except Exception:
            pass
        msg = f"Phần mềm đã được cập nhật lên bản {n.get('version')}"
        if n.get('files'):
            msg += f" ({n['files']} file được thay)"
        msg += '.'
        if n.get('title'):
            msg += f"\n\n{n.get('title')}"
        if n.get('notes'):
            msg += '\n\nNội dung thay đổi:\n' + n['notes'][:1500]
        messagebox.showinfo('✔ Đã cập nhật', msg, parent=app)
    try:                                       # don thu muc giai nen cu
        t = os.path.join(root, TMP_DIR)
        for d in os.listdir(t):
            if d.startswith('stage_'):
                shutil.rmtree(os.path.join(t, d), ignore_errors=True)
    except OSError:
        pass


def _progress_box(parent):
    import tkinter as tk
    from tkinter import ttk
    fr = tk.Frame(parent, bg='#f4f6f9')
    pb = ttk.Progressbar(fr, mode='determinate', length=300)
    lb = tk.Label(fr, text='', bg='#f4f6f9', font=FONT_S, fg='#555', anchor='w')
    pb.pack(side='left', padx=(0, 8))
    lb.pack(side='left', fill='x', expand=True)

    def set_(i, n, t=''):
        pb['maximum'] = max(1, n)
        pb['value'] = i
        lb.config(text=t if t else f'{i}/{n}')
    fr.set = set_
    fr.text = lambda t: lb.config(text=t)
    return fr


def _button_rows(parent):
    """2 hang nut (nguoi dung / quan tri) -> khong bi che khi cua so hep."""
    import tkinter as tk

    def make(row, text, cmd, bg, side='left'):
        b = tk.Button(row, text=text, font=FONT_B, bg=bg, fg='white', relief='flat', padx=10, pady=5,
                      command=cmd, activebackground=bg, activeforeground='white')
        b.pack(side=side, padx=3, pady=2)
        return b
    return make


class UpdateWindow:
    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk
        old = getattr(app, '_upd_win', None)
        try:
            if old is not None and old.winfo_exists():
                old.lift()
                old.focus_force()
                return
        except Exception:
            pass
        self.app = app
        self.admin = _is_admin(app)
        self.items = []
        self.mode = run_mode()
        w = self.w = tk.Toplevel(app)
        app._upd_win = w
        w.title('Cập nhật phần mềm - HitechPlast')
        w.geometry('1100x660')
        w.minsize(900, 520)
        w.configure(bg='#f4f6f9')
        w.transient(app)

        loc = read_local()
        top = tk.Frame(w, bg='#1f3a5f')
        top.pack(fill='x')
        tk.Label(top, text='⬆ CẬP NHẬT PHẦN MỀM', bg='#1f3a5f', fg='white', font=FONT_T).pack(side='left',
                                                                                             padx=14, pady=8)
        info = f"Máy này: {socket.gethostname()}  •  Phiên bản: {loc.get('version') or 'chưa có'}"
        if loc.get('installed_at'):
            info += f"  (cài {loc['installed_at']})"
        info += f'  •  {TARGET_LABEL[self.mode]}'
        tk.Label(top, text=info, bg='#1f3a5f', fg='#9fc0e8', font=FONT).pack(side='right', padx=14)

        src = ('Bấm cài: phần mềm tự KIỂM TRA từng file trên máy này và CHỈ TẢI file thiếu / khác.'
               if server_ok() else
               'Đang ở chế độ FILE / OFFLINE: không có máy chủ -> dùng "📥 Cài từ file gói" '
               '(file .htpupd nhận qua Zalo / USB).')
        tk.Label(w, text=src, bg='#f4f6f9', fg='#1f3a5f', font=FONT_B).pack(anchor='w', padx=12, pady=(8, 2))

        # nut o DUOI CUNG (pack truoc de luon hien du, phan danh sach co gian phia tren)
        make = _button_rows(w)
        if self.admin:
            r2 = tk.Frame(w, bg='#e9edf2')
            r2.pack(side='bottom', fill='x', padx=12, pady=(0, 8))
            tk.Label(r2, text='Quản trị:', bg='#e9edf2', font=FONT_B).pack(side='left', padx=(6, 4))
            make(r2, '📤 Đăng bản cập nhật…', lambda: PublishWindow(self.app, on_done=self.reload), '#b0413e')
            make(r2, '🚫 Thu hồi / bật lại', self._toggle_active, '#8e44ad')
            make(r2, '🗑 Xoá bản', self._delete, '#c0392b')
            make(r2, '💾 Lưu gói ra file', self._export_file, '#16a085')
        r1 = tk.Frame(w, bg='#f4f6f9')
        r1.pack(side='bottom', fill='x', padx=12, pady=(4, 4))
        self.b_latest = make(r1, '⬇ Cài bản mới nhất', self._install_latest, '#e67e22')
        make(r1, '⬇ Cài bản đã chọn', self._install_selected, '#2d6a4f')
        make(r1, '↩ Khôi phục bản trước', self._restore, '#7f8c8d')
        make(r1, '📥 Cài từ file gói…', self._install_file, '#1d3557')
        make(r1, '🔄 Làm mới', self.reload, '#34495e')
        self.prog = _progress_box(w)
        self.prog.pack(side='bottom', fill='x', padx=12, pady=(2, 0))

        body = tk.PanedWindow(w, orient='vertical', bg='#f4f6f9', sashwidth=6)
        body.pack(fill='both', expand=True, padx=12, pady=4)
        tf = tk.Frame(body, bg='#f4f6f9')
        cols = ('ver', 'date', 'kind', 'title', 'author', 'size', 'state')
        self.tv = ttk.Treeview(tf, columns=cols, show='headings', height=9, selectmode='browse')
        for c, t, wd, an in (('ver', 'Phiên bản', 105, 'center'), ('date', 'Ngày đăng', 135, 'center'),
                             ('kind', 'Loại', 110, 'center'), ('title', 'Tiêu đề', 225, 'w'),
                             ('author', 'Người đăng / máy', 170, 'w'), ('size', 'Dung lượng', 85, 'e'),
                             ('state', 'Trạng thái', 150, 'center')):
            self.tv.heading(c, text=t)
            self.tv.column(c, width=wd, minwidth=60, anchor=an, stretch=(c == 'title'))
        self.tv.tag_configure('new', background='#fff3e0')
        self.tv.tag_configure('cur', background='#e8f5e9', font=FONT_B)
        self.tv.tag_configure('off', foreground='#999')
        sb = ttk.Scrollbar(tf, orient='vertical', command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.pack(side='left', fill='both', expand=True)
        sb.pack(side='left', fill='y')
        self.tv.bind('<<TreeviewSelect>>', lambda e: self._show_sel())
        self.tv.bind('<Double-1>', lambda e: self._install_selected())
        body.add(tf, minsize=120)
        df = tk.Frame(body, bg='#f4f6f9')
        tk.Label(df, text='Nội dung bản đã chọn:', bg='#f4f6f9', font=FONT_B).pack(anchor='w')
        self.txt = tk.Text(df, height=7, font=FONT, wrap='word', relief='solid', bd=1)
        self.txt.pack(fill='both', expand=True)
        body.add(df, minsize=80)
        self.reload()

    # ---------------- danh sach ----------------
    def reload(self):
        if not server_ok():
            self._fill([])
            return
        self.prog.text('Đang đọc danh sách phiên bản trên máy chủ...')

        def done(item):
            if item[0] == 'ok':
                self.prog.text('')
                self._fill(item[1])
            else:
                self.prog.text('Lỗi đọc máy chủ: ' + str(item[1]))
        run_bg(self.w, lambda prog: remote_list(), done)

    def _mine(self, m):
        return m.get('target') == self.mode

    def _fill(self, items):
        self.items = items
        self.tv.delete(*self.tv.get_children())
        cur = current_version()
        newest = None
        for m in items:
            v = m['version']
            if not self._mine(m):
                st, tag = 'Không dành cho máy này', 'off'
            elif not m.get('active', True):
                st, tag = 'Đã thu hồi', 'off'
            elif v == cur:
                st, tag = '● Đang dùng', 'cur'
            elif ver_key(v) > ver_key(cur):
                st, tag = '🆕 Mới', 'new'
                newest = newest or v
            else:
                st, tag = 'Cũ hơn', ''
            who = m.get('author', '') + (f" ({m.get('pc')})" if m.get('pc') else '')
            self.tv.insert('', 'end', iid=v, values=(v, m.get('created', ''), TARGET_LABEL.get(m.get('target'), '?'),
                                                     m.get('title', ''), who, fmt_size(m.get('size')), st),
                           tags=(tag,))
        self.b_latest.config(text=f'⬇ Cài bản mới nhất ({newest})' if newest else '⬇ Cài bản mới nhất',
                             state='normal' if newest else 'disabled')
        _set_badge(self.app, newest)
        if newest:
            self.tv.selection_set(newest)
            self.tv.see(newest)

    def _sel(self):
        s = self.tv.selection()
        if not s:
            return None
        return next((m for m in self.items if m['version'] == s[0]), None)

    def _show_sel(self):
        m = self._sel()
        self.txt.delete('1.0', 'end')
        if not m:
            return
        lines = [f"Phiên bản {m['version']}  •  {m.get('created', '')}  •  {m.get('author', '')} ({m.get('pc', '')})"
                 f"  •  {TARGET_LABEL.get(m.get('target'), '?')}  •  {len(m.get('files', []))} file, "
                 f"{fmt_size(m.get('size'))}", '']
        if m.get('title'):
            lines += [m['title'], '']
        if m.get('notes'):
            lines += [m['notes'], '']
        if not self._mine(m):
            lines.append(f"⚠ Bản này dành cho máy chạy {TARGET_LABEL.get(m.get('target'))}, máy này đang chạy "
                         f"{TARGET_LABEL[self.mode]} -> không cài được.")
        else:
            lines.append('Bấm cài: phần mềm kiểm tra máy này và chỉ tải các file thiếu / khác.')
        self.txt.insert('1.0', '\n'.join(lines))

    # ---------------- cai dat ----------------
    def _install_latest(self):
        act = [m for m in self.items if m.get('active', True) and self._mine(m)]
        if act:
            self._install(max(act, key=lambda m: ver_key(m['version'])))

    def _install_selected(self):
        from tkinter import messagebox
        m = self._sel()
        if not m:
            messagebox.showinfo('Cập nhật', 'Chọn 1 phiên bản trong danh sách.', parent=self.w)
            return
        if not self._mine(m):
            messagebox.showwarning('Cập nhật', f"Bản {m['version']} dành cho máy chạy "
                                               f"{TARGET_LABEL.get(m.get('target'))}, không cài được trên máy này.",
                                   parent=self.w)
            return
        if not m.get('active', True):
            messagebox.showwarning('Cập nhật', f"Bản {m['version']} đã bị THU HỒI, không cài được.", parent=self.w)
            return
        self._install(m)

    def _check_then_install(self, man, fetch_factory, close=None):
        """Buoc 1 (nen): kiem tra file tren may -> hoi -> Buoc 2 (nen): tai file thieu + chuan bi -> cai."""
        from tkinter import messagebox
        self.w.config(cursor='watch')

        def check(prog):
            local = scan_folder(app_dir(), lambda i, n, r: prog(i, n, 'Đang kiểm tra file trên máy này...'))
            return missing_files(man, local)

        def checked(item):
            self.w.config(cursor='')
            if item[0] == 'err':
                self.prog.text('Lỗi: ' + str(item[1]))
                messagebox.showerror('Cập nhật lỗi', str(item[1]), parent=self.w)
                if close:
                    close()
                return
            need = item[1]
            v = man['version']
            cur = current_version()
            if not need:
                self.prog.text(f'Máy này đã giống bản {v} – không cần tải file nào.')
                mark_installed(v)
                messagebox.showinfo('Cập nhật', f'Máy này đã có ĐỦ và ĐÚNG mọi file của bản {v}.\n'
                                                'Không cần tải gì – đã ghi nhận phiên bản.', parent=self.w)
                if close:
                    close()
                self._fill(self.items)
                return
            size = sum(f['size'] for f in need)
            show = '\n'.join(f"  • {f['path']}  ({fmt_size(f['size'])})" for f in need[:12])
            if len(need) > 12:
                show += f'\n  … và {len(need) - 12} file khác'
            down = ver_key(v) <= ver_key(cur)
            if not messagebox.askyesno(
                    'Cập nhật phần mềm',
                    f"{'VỀ BẢN CŨ / CÀI LẠI' if down else 'CẬP NHẬT'}: {cur} → {v}\n\n"
                    f'Đã kiểm tra máy này: thiếu / khác {len(need)} file, cần tải {fmt_size(size)}:\n{show}\n\n'
                    'Phần mềm sẽ ĐÓNG và tự MỞ LẠI sau khi chép file.\n'
                    'Hãy LƯU công việc đang làm trước khi tiếp tục.\n'
                    'File cũ được sao lưu – có thể "↩ Khôi phục bản trước".\n\nTiếp tục?', parent=self.w):
                self.prog.text('')
                if close:
                    close()
                return
            self.w.config(cursor='watch')

            def work(prog):
                fetch, fin = fetch_factory()
                try:
                    return prepare_install(man, need, fetch, prog)
                finally:
                    fin()
            run_bg(self.w, work, self._prepared, self.prog.set)
        run_bg(self.w, check, checked, self.prog.set)

    def _prepared(self, item):
        from tkinter import messagebox
        try:
            self.w.config(cursor='')
        except Exception:
            pass
        if item[0] == 'err':
            self.prog.text('Lỗi: ' + str(item[1]))
            messagebox.showerror('Cập nhật lỗi', str(item[1]), parent=self.w)
            return
        job = item[1]
        self.prog.text('Đã chuẩn bị xong – đang đóng phần mềm để chép file...')
        try:
            self.w.destroy()
        except Exception:
            pass
        launch_and_quit(self.app, job)

    def _install(self, m):
        def factory():
            c = _own_conn()
            return (lambda f: download_blob(c, f['sha256'])), c.close
        self._check_then_install(m, factory)

    def _install_file(self):
        from tkinter import filedialog, messagebox
        p = filedialog.askopenfilename(parent=self.w, title='Chọn file gói cập nhật',
                                       filetypes=[('Gói cập nhật HitechPlast', '*' + PKG_EXT), ('Tất cả', '*.*')])
        if not p:
            return
        try:
            man, z = read_package_file(p)
        except Exception as e:
            messagebox.showerror('Gói lỗi', f'Không đọc được gói cập nhật:\n{e}', parent=self.w)
            return
        if man.get('target') != self.mode:
            z.close()
            messagebox.showwarning('Cài từ file', f"Gói {man.get('version')} dành cho máy chạy "
                                                  f"{TARGET_LABEL.get(man.get('target'))}, máy này chạy "
                                                  f"{TARGET_LABEL[self.mode]}.", parent=self.w)
            return

        def factory():
            return (lambda f: z.read('files/' + f['path'])), (lambda: None)
        self._check_then_install(man, factory, close=z.close)

    def _restore(self):
        from tkinter import messagebox
        bks = list_backups()
        if not bks:
            messagebox.showinfo('Khôi phục', 'Chưa có bản sao lưu nào trên máy này.', parent=self.w)
            return
        bk = bks[0]
        if not messagebox.askyesno('Khôi phục bản trước',
                                   f"Khôi phục về bản {bk.get('from_version')} (sao lưu lúc {bk.get('created')}, "
                                   f"trước khi cài bản {bk.get('to_version')})?\n\n"
                                   'Phần mềm sẽ đóng và tự mở lại.', parent=self.w):
            return
        try:
            job = prepare_restore(bk)
        except Exception as e:
            messagebox.showerror('Khôi phục lỗi', str(e), parent=self.w)
            return
        self.w.destroy()
        launch_and_quit(self.app, job)

    # ---------------- admin ----------------
    def _toggle_active(self):
        from tkinter import messagebox
        m = self._sel()
        if not m:
            messagebox.showinfo('Thu hồi', 'Chọn 1 phiên bản trong danh sách.', parent=self.w)
            return
        on = not m.get('active', True)
        if not messagebox.askyesno('Thu hồi / bật lại',
                                   f"{'BẬT LẠI' if on else 'THU HỒI'} bản {m['version']}?\n"
                                   + ('' if on else 'Các máy sẽ không cài được bản này nữa '
                                                    '(máy đã cài vẫn giữ – dùng ↩ Khôi phục nếu cần).'),
                                   parent=self.w):
            return
        try:
            remote_set_active(m['version'], on)
        except Exception as e:
            messagebox.showerror('Lỗi', str(e), parent=self.w)
        self.reload()

    def _delete(self):
        from tkinter import messagebox
        m = self._sel()
        if not m:
            messagebox.showinfo('Xoá', 'Chọn 1 phiên bản trong danh sách.', parent=self.w)
            return
        if not messagebox.askyesno('Xoá phiên bản', f"XOÁ HẲN bản {m['version']} khỏi máy chủ?\n"
                                                    'Các file không còn bản nào dùng cũng bị xoá để giải phóng '
                                                    'dung lượng.\nNên dùng "Thu hồi" nếu chỉ muốn tạm khoá.',
                                   parent=self.w):
            return
        try:
            n = remote_delete(m['version'])
            self.prog.text(f"Đã xoá bản {m['version']} (giải phóng {n} file không dùng).")
        except Exception as e:
            messagebox.showerror('Lỗi', str(e), parent=self.w)
        self.reload()

    def _export_file(self):
        from tkinter import filedialog, messagebox
        m = self._sel()
        if not m:
            messagebox.showinfo('Lưu gói', 'Chọn 1 phiên bản trong danh sách.', parent=self.w)
            return
        p = filedialog.asksaveasfilename(parent=self.w, defaultextension=PKG_EXT,
                                         initialfile=f"HitechPlast_{m['version']}{PKG_EXT}",
                                         filetypes=[('Gói cập nhật', '*' + PKG_EXT)])
        if not p:
            return

        def work(prog):
            c = _own_conn()
            try:
                export_package(m, p, lambda sha, path: download_blob(c, sha), prog)
            finally:
                c.close()
            return p

        def done(item):
            if item[0] == 'ok':
                self.prog.text('Đã lưu: ' + item[1])
                messagebox.showinfo('Đã lưu', f'Đã lưu gói ra file:\n{item[1]}\n\nGửi file này cho máy không '
                                              'kết nối máy chủ -> "📥 Cài từ file gói" (máy đó cũng chỉ chép '
                                              'các file khác).', parent=self.w)
            else:
                messagebox.showerror('Lỗi', str(item[1]), parent=self.w)
        run_bg(self.w, work, done, self.prog.set)


class PublishWindow:
    """ADMIN: chon thu muc ban build -> quet -> dang (chi tai len file may chu chua co)."""

    def __init__(self, app, on_done=None):
        import tkinter as tk
        from tkinter import ttk
        self.app = app
        self.on_done = on_done
        self.root = default_publish_dir()
        self.scan = {}
        self.status = {}
        self.have = set()
        self.prev = None
        self.versions = []
        self.excluded = set()           # file admin bo tick (khong dua vao phien ban)
        w = self.w = tk.Toplevel(app)
        w.title('Đăng bản cập nhật')
        w.geometry('1100x720')
        w.minsize(900, 560)
        w.configure(bg='#f4f6f9')
        w.transient(app)

        f1 = tk.Frame(w, bg='#f4f6f9')
        f1.pack(fill='x', padx=12, pady=(10, 4))
        tk.Label(f1, text='Phiên bản:', bg='#f4f6f9', font=FONT_B).grid(row=0, column=0, sticky='w')
        self.e_ver = tk.Entry(f1, font=FONT_B, width=16)
        self.e_ver.grid(row=0, column=1, sticky='w', padx=6)
        self.e_ver.insert(0, suggest_version())
        tk.Label(f1, text='Tiêu đề:', bg='#f4f6f9', font=FONT_B).grid(row=0, column=2, sticky='w', padx=(14, 0))
        self.e_title = tk.Entry(f1, font=FONT, width=60)
        self.e_title.grid(row=0, column=3, sticky='we', padx=6)
        f1.columnconfigure(3, weight=1)
        tk.Label(f1, text='Ghi chú\n(tính năng mới):', bg='#f4f6f9', font=FONT_B, justify='left').grid(
            row=1, column=0, sticky='nw', pady=4)
        self.t_notes = tk.Text(f1, font=FONT, height=3, wrap='word', relief='solid', bd=1)
        self.t_notes.grid(row=1, column=1, columnspan=3, sticky='we', padx=6, pady=4)

        f2 = tk.Frame(w, bg='#f4f6f9')
        f2.pack(fill='x', padx=12)
        tk.Label(f2, text='Thư mục bản build:', bg='#f4f6f9', font=FONT_B).pack(side='left')
        self.l_root = tk.Label(f2, text=self.root, bg='#f4f6f9', font=FONT, fg='#1f3a5f')
        self.l_root.pack(side='left', padx=6)
        tk.Button(f2, text='📁 Chọn thư mục khác…', font=FONT, relief='groove',
                  command=self._pick_root).pack(side='left', padx=4)
        self.v_only = tk.BooleanVar(value=True)
        tk.Checkbutton(f2, text='Chỉ hiện file khác bản trước', variable=self.v_only, bg='#f4f6f9', font=FONT,
                       command=self._fill).pack(side='right')
        self.l_kind = tk.Label(w, text='', bg='#f4f6f9', font=FONT_B, justify='left', anchor='w')
        self.l_kind.pack(fill='x', padx=12, pady=(4, 0))

        # nut duoi cung (pack truoc)
        make = _button_rows(w)
        bf = tk.Frame(w, bg='#f4f6f9')
        bf.pack(side='bottom', fill='x', padx=12, pady=8)
        make(bf, '☑ Chọn hết', lambda: self._set_all(True), '#34495e')
        make(bf, '☐ Bỏ hết', lambda: self._set_all(False), '#34495e')
        make(bf, '🔄 Quét lại', self._rescan, '#34495e')
        self.b_pub = make(bf, '📤 Đăng lên máy chủ', lambda: self._publish(False), '#b0413e', 'right')
        make(bf, '💾 Lưu ra file gói', lambda: self._publish(True), '#16a085', 'right')
        if not server_ok():
            self.b_pub.config(state='disabled')
        self.l_sum = tk.Label(w, text='', bg='#f4f6f9', font=FONT_B, anchor='w', justify='left')
        self.l_sum.pack(side='bottom', fill='x', padx=12)
        self.prog = _progress_box(w)
        self.prog.pack(side='bottom', fill='x', padx=12)

        tf = tk.Frame(w, bg='#f4f6f9')
        tf.pack(fill='both', expand=True, padx=12, pady=4)
        cols = ('chk', 'path', 'size', 'mtime', 'st', 'srv')
        self.tv = ttk.Treeview(tf, columns=cols, show='headings', selectmode='extended')
        for c, t, wd, an in (('chk', '✔', 40, 'center'), ('path', 'File', 430, 'w'), ('size', 'Dung lượng', 90, 'e'),
                             ('mtime', 'Sửa lúc', 135, 'center'), ('st', 'So với bản trước', 130, 'center'),
                             ('srv', 'Máy chủ', 120, 'center')):
            self.tv.heading(c, text=t)
            self.tv.column(c, width=wd, minwidth=40, anchor=an, stretch=(c == 'path'))
        self.tv.tag_configure('chg', background='#fff3e0')
        self.tv.tag_configure('new', background='#e3f2fd')
        self.tv.tag_configure('off', foreground='#999')
        sb = ttk.Scrollbar(tf, orient='vertical', command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.pack(side='left', fill='both', expand=True)
        sb.pack(side='left', fill='y')
        self.tv.bind('<Button-1>', self._click)
        self.tv.bind('<space>', lambda e: self._toggle(self.tv.selection()))
        self._rescan()

    def _pick_root(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(parent=self.w, title='Chọn thư mục bản build (có LenhSanXuat_HitechPlast.exe)',
                                    initialdir=self.root)
        if d:
            self.root = os.path.abspath(d)
            self.l_root.config(text=self.root)
            self.excluded.clear()
            self._rescan()

    def _rescan(self):
        root = self.root
        tgt = folder_target(root)
        if tgt == 'exe':
            self.l_kind.config(fg='#2d6a4f', text='✔ Thư mục BẢN EXE (có LenhSanXuat_HitechPlast.exe) – các máy '
                                                  'chạy EXE sẽ nhận bản này.')
        else:
            self.l_kind.config(fg='#c0392b', text='⚠ Thư mục này là MÃ NGUỒN (không có LenhSanXuat_HitechPlast.exe) '
                                                  '– chỉ máy chạy Python nhận được.\n   Các máy chạy EXE: chạy '
                                                  'build_exe.bat rồi bấm "📁 Chọn thư mục khác" -> '
                                                  'GOI_CAI_DAT\\HitechPlast_LenhSanXuat.')
        self.prog.text('Đang quét thư mục...')

        def work(prog):
            scan = scan_folder(root, lambda i, n, r: prog(i, n, f'Đang quét {i}/{n} file...'))
            prev, have, vers = None, set(), []
            if server_ok():
                c = _own_conn()
                try:
                    lst = remote_list(c)
                    vers = [m['version'] for m in lst]
                    same = [m for m in lst if m.get('target') == tgt]
                    prev = same[0] if same else None
                    have = remote_have({v['sha'] for v in scan.values()}, c)
                finally:
                    c.close()
            return scan, prev, have, vers

        def done(item):
            if item[0] != 'ok':
                self.prog.text('Lỗi quét: ' + str(item[1]))
                return
            self.scan, self.prev, self.have, self.versions = item[1]
            self.e_ver.delete(0, 'end')
            self.e_ver.insert(0, suggest_version(self.versions))
            self.status = diff_status(self.scan, self.prev)
            self.prog.text(f'Đã quét {len(self.scan)} file.' +
                           (f"  So sánh với bản trước: {self.prev['version']}" if self.prev else
                            '  Chưa có bản nào trước đó trên máy chủ.'))
            self._fill()
        run_bg(self.w, work, done, self.prog.set)

    def _fill(self):
        self.tv.delete(*self.tv.get_children())
        only = self.v_only.get() and self.prev is not None
        rows = sorted(self.scan, key=lambda r: (self.status.get(r, '') == '', r))
        for r in rows:
            st = self.status.get(r, '')
            if only and not st and r not in self.excluded:
                continue
            v = self.scan[r]
            srv = 'Đã có' if v['sha'] in self.have else 'Cần tải lên'
            inc = r not in self.excluded
            tag = 'off' if not inc else 'new' if st == 'Mới' else 'chg' if st == 'Đã đổi' else ''
            mt = datetime.fromtimestamp(v['mtime']).strftime('%Y-%m-%d %H:%M')
            self.tv.insert('', 'end', iid=r, values=('☑' if inc else '☐', r, fmt_size(v['size']), mt, st, srv),
                           tags=(tag,))
        if only and not self.tv.get_children():
            self.tv.insert('', 'end', iid='__none__', values=(
                '', f"✔ Thư mục này GIỐNG HỆT bản {self.prev['version']} trên máy chủ – không có file nào mới / "
                    "đã đổi. (Bỏ tick 'Chỉ hiện file khác bản trước' để xem tất cả.)", '', '', '', ''), tags=('off',))
        self._summary()

    def _summary(self):
        inc = [r for r in self.scan if r not in self.excluded]
        tot = sum(self.scan[r]['size'] for r in inc)
        chg = [r for r in inc if self.status.get(r) in ('Mới', 'Đã đổi')]
        up = {self.scan[r]['sha']: self.scan[r]['size'] for r in inc if self.scan[r]['sha'] not in self.have}
        txt = f'Bản này gồm {len(inc)} file ({fmt_size(tot)})'
        if self.prev:
            txt += f'  •  khác bản {self.prev["version"]}: {len(chg)} file'
        if server_ok():
            txt += f'  •  cần tải lên máy chủ: {len(up)} file ({fmt_size(sum(up.values()))})'
        if self.excluded:
            txt += f'  •  bỏ ra: {len(self.excluded)} file'
        if folder_target(self.root) == 'exe' and MAIN_EXE in self.excluded:
            txt += '   ⚠ đang bỏ file .exe!'
        self.l_sum.config(text=txt)

    def _click(self, ev):
        if self.tv.identify_region(ev.x, ev.y) != 'cell' or self.tv.identify_column(ev.x) != '#1':
            return
        iid = self.tv.identify_row(ev.y)
        if iid and iid in self.scan:
            self._toggle([iid])
            return 'break'

    def _toggle(self, iids):
        for i in [x for x in iids if x in self.scan]:
            (self.excluded.discard if i in self.excluded else self.excluded.add)(i)
            self.tv.set(i, 'chk', '☐' if i in self.excluded else '☑')
        self._summary()

    def _set_all(self, on):
        for i in [x for x in self.tv.get_children() if x in self.scan]:
            (self.excluded.discard if on else self.excluded.add)(i)
            self.tv.set(i, 'chk', '☑' if on else '☐')
        self._summary()

    def _publish(self, to_file):
        from tkinter import messagebox, filedialog
        ver = self.e_ver.get().strip()
        title = self.e_title.get().strip()
        notes = self.t_notes.get('1.0', 'end').strip()
        rels = sorted(r for r in self.scan if r not in self.excluded)
        if not ver_key(ver):
            messagebox.showwarning('Đăng bản', 'Nhập số phiên bản (vd 2026.10.02).', parent=self.w)
            return
        if not rels:
            messagebox.showwarning('Đăng bản', 'Chưa có file nào (hãy Quét lại / chọn đúng thư mục).', parent=self.w)
            return
        if not title:
            messagebox.showwarning('Đăng bản', 'Nhập TIÊU ĐỀ (tính năng mới là gì) để các máy khác biết.',
                                   parent=self.w)
            return
        top = max(self.versions, key=ver_key) if self.versions else None
        if top and ver_key(ver) <= ver_key(top):
            messagebox.showwarning('Đăng bản', f'Số phiên bản {ver} phải LỚN HƠN bản mới nhất trên máy chủ ({top}),\n'
                                               'nếu không các máy khác sẽ không nhận là bản mới.\n'
                                               f'Gợi ý: {suggest_version(self.versions)}', parent=self.w)
            return
        tgt = folder_target(self.root)
        if tgt == 'py' and not messagebox.askyesno(
                'Đăng bản MÃ NGUỒN', 'Thư mục này là MÃ NGUỒN, các máy chạy EXE sẽ KHÔNG nhận được.\n'
                                     'Vẫn đăng (chỉ cho máy chạy Python)?', parent=self.w):
            return
        if self.prev and not any(self.status.get(r) in ('Mới', 'Đã đổi') for r in rels) and not to_file:
            if not messagebox.askyesno('Đăng bản', f"Không có file nào khác bản {self.prev['version']}. Vẫn đăng?",
                                       parent=self.w):
                return
        out = None
        if to_file:
            out = filedialog.asksaveasfilename(parent=self.w, defaultextension=PKG_EXT,
                                               initialfile=f'HitechPlast_{ver}{PKG_EXT}',
                                               filetypes=[('Gói cập nhật', '*' + PKG_EXT)])
            if not out:
                return
        elif not messagebox.askyesno('Đăng bản cập nhật',
                                     f'Đăng bản {ver} ({TARGET_LABEL[tgt]}, {len(rels)} file) lên máy chủ?\n'
                                     'Chỉ tải lên các file máy chủ chưa có.\n'
                                     'Các máy khác sẽ thấy nút "🆕 Có bản mới" trong vòng 5 phút.', parent=self.w):
            return
        root, scan = self.root, dict(self.scan)
        author = _who(self.app)

        def work(prog):
            if to_file:
                files = [{'path': r, 'sha256': scan[r]['sha'], 'size': scan[r]['size']} for r in rels]
                man = {'format': FORMAT, 'version': ver, 'target': tgt, 'title': title, 'notes': notes,
                       'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'author': author,
                       'pc': socket.gethostname(), 'files': files}

                def src(sha, path):
                    with open(os.path.join(root, path), 'rb') as f:
                        return f.read()
                export_package(man, out, src, prog)
                return man, len(files), sum(f['size'] for f in files)
            return publish(root, rels, scan, ver, title, notes, author, prog)

        def done(item):
            if item[0] == 'err':
                self.prog.text('Lỗi: ' + str(item[1]))
                messagebox.showerror('Đăng bản lỗi', str(item[1]), parent=self.w)
                return
            man, nup, sup = item[1]
            if not to_file and os.path.abspath(root) == os.path.abspath(app_dir()):
                mark_installed(ver)                   # may dang = ban vua dang
            msg = (f'Đã lưu gói {ver} ({len(man["files"])} file) ra:\n{out}' if to_file else
                   f'Đã đăng bản {ver}: {len(man["files"])} file, tải lên {nup} file mới ({fmt_size(sup)}).\n'
                   'Các máy khác bấm "⬆ Cập nhật" – mỗi máy chỉ tải file nó thiếu / khác.')
            messagebox.showinfo('Xong', msg, parent=self.w)
            self.w.destroy()
            if self.on_done:
                try:
                    self.on_done()
                except Exception:
                    pass

        def prog_ui(*a):
            # publish() gui tin '__file__' moi khi 1 file tai len xong -> danh dau "Da hoan thanh"
            if a and a[0] == '__file__':
                self._mark_done(a[1])
            else:
                try:
                    self.prog.set(*a)
                except Exception:
                    pass
        run_bg(self.w, work, done, prog_ui)

    def _mark_done(self, sha):
        """Doi cot 'May chu' cua dong co sha nay sang 'Da hoan thanh'."""
        for iid in self.tv.get_children():
            if iid in self.scan and self.scan[iid].get('sha') == sha:
                try:
                    self.tv.set(iid, 'srv', '✔ Đã hoàn thành')
                except Exception:
                    pass
