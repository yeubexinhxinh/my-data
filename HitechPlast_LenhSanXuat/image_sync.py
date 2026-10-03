# -*- coding: utf-8 -*-
"""
ANH TEM / BAO BI TREN MAY CHU - HitechPlast
===========================================
Hai nguon anh (chon trong "📁 Đường dẫn ảnh (Tem/Bao bì)", luu settings.json cua tung may):

  'local'  : THEO MAY NAY  - doc thu muc tren may (stamp_dir / pkg_dir) nhu truoc.
  'server' : THEO MAY CHU  - anh luu trong bang htp_images tren may chu PostgreSQL,
             moi may dung chung. May con:
               - "⬆ Tải lên máy chủ": gui anh tu thu muc / file tren may con len may chu
                 (chi gui anh MOI hoac DA SUA, so theo ma SHA1).
               - Khi lam lenh: tu tai anh cua dung khach hang ve bo nho dem
                 <thu muc phan mem>\\Anh_tem_may_chu\\<Stamp images|Packaging Picture>\\<Ma KH>\\...
               - Mat ket noi / offline: van dung anh da tai ve truoc do.

Cau truc giu nguyen: <goc>\\<Ma KH>\\<Ma KH>-<Ma hang>.png  (image_finder.find_image_ex).
"""
import hashlib
import json
import os
import sys
import threading
import time

TABLE = 'htp_images'
KINDS = {'stamp': 'Stamp images', 'pkg': 'Packaging Picture'}
KIND_LABEL = {'stamp': 'Tem', 'pkg': 'Bao bì'}
IMG_EXTS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
MAX_BYTES = 15 * 1024 * 1024          # bo qua anh > 15 MB
SYNC_EVERY = 30                        # giay: khong hoi lai may chu cung 1 thu muc qua day

_SCHEMA = (
    "CREATE TABLE IF NOT EXISTS htp_images(kind TEXT NOT NULL, folder TEXT NOT NULL, name TEXT NOT NULL,"
    " data BYTEA NOT NULL, size INT NOT NULL DEFAULT 0, sha1 TEXT NOT NULL,"
    " updated_at TIMESTAMP DEFAULT now(), updated_by TEXT, host TEXT,"
    " PRIMARY KEY(kind, folder, name))",
    "CREATE INDEX IF NOT EXISTS ix_htp_images_folder ON htp_images(kind, upper(folder))",
)
_schema_ok = False
_last_sync = {}
_lock = threading.RLock()


# ---------------------------------------------------------------------
#  Ket noi
# ---------------------------------------------------------------------
def server_available():
    """True neu dang o che do may chu va dang ket noi duoc (khong offline)."""
    try:
        import datastore
        if datastore.is_offline():
            return False
        return datastore.is_server() and datastore.server() is not None
    except Exception:
        return False


def _store():
    import datastore
    if datastore.is_offline() or not datastore.is_server():
        raise RuntimeError('Chưa kết nối máy chủ dữ liệu (đang ở chế độ file hoặc offline).')
    st = datastore.server()
    if st is None:
        raise RuntimeError('Chưa kết nối máy chủ dữ liệu.')
    return st


def _conn():
    global _schema_ok
    st = _store()
    if not _schema_ok:
        for q in _SCHEMA:
            st.c.execute(q)
        _schema_ok = True
    return st


def _who(st):
    return getattr(st, 'who', '') or ''


def _host():
    try:
        import socket
        return socket.gethostname()
    except Exception:
        return ''


# ---------------------------------------------------------------------
#  Che do / duong dan
# ---------------------------------------------------------------------
def mode(settings):
    """'server' | 'local'. Mac dinh: may dang noi may chu -> 'server', khong -> 'local'."""
    m = (settings or {}).get('img_source')
    if m in ('server', 'local'):
        return m
    try:
        import datastore
        return 'server' if datastore.load_config() is not None else 'local'
    except Exception:
        return 'local'


def _app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def cache_root():
    d = os.path.join(_app_dir(), 'Anh_tem_may_chu')
    try:
        os.makedirs(d, exist_ok=True)
        t = os.path.join(d, '.w')
        open(t, 'w').close()
        os.remove(t)
        return d
    except OSError:
        d = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'HitechPlast', 'Anh_tem_may_chu')
        os.makedirs(d, exist_ok=True)
        return d


def cache_dir(kind):
    d = os.path.join(cache_root(), KINDS[kind])
    os.makedirs(d, exist_ok=True)
    return d


def roots(settings, kind):
    """Danh sach thu muc goc de tim anh, theo thu tu uu tien."""
    st = settings or {}
    local = st.get('stamp_dir' if kind == 'stamp' else 'pkg_dir') or ''
    if mode(st) == 'server':
        out = [cache_dir(kind)]
        if st.get('img_local_fallback', True) and local:
            out.append(local)            # khong co tren may chu -> thu thu muc tren may
        return out
    return [local]


def _sha1(data):
    return hashlib.sha1(data).hexdigest()


def _safe(name):
    bad = '<>:"/\\|?*'
    return ''.join('_' if c in bad else c for c in str(name)).strip().strip('.') or '_'


def _manifest_path(kind):
    return os.path.join(cache_dir(kind), '.manifest.json')


def _manifest_load(kind):
    try:
        return json.load(open(_manifest_path(kind), encoding='utf-8'))
    except Exception:
        return {}


def _manifest_save(kind, m):
    try:
        json.dump(m, open(_manifest_path(kind), 'w', encoding='utf-8'), ensure_ascii=False)
    except Exception:
        pass


# ---------------------------------------------------------------------
#  May chu -> may nay
# ---------------------------------------------------------------------
def sync_folder(kind, folder, force=False):
    """Tai anh cua 1 thu muc khach hang (Ma KH) tu may chu ve bo nho dem.
    Tra ve so anh vua tai. Loi ket noi -> tra ve 0 (van dung anh cu)."""
    folder = str(folder or '').strip()
    if not folder:
        return 0
    key = (kind, folder.upper())
    now = time.time()
    if not force and now - _last_sync.get(key, 0) < SYNC_EVERY:
        return 0
    try:
        st = _conn()
        rows = st.c.query(f"SELECT folder, name, sha1 FROM {TABLE} WHERE kind=%s AND upper(folder)=upper(%s)",
                          (kind, folder))
    except Exception:
        return 0
    _last_sync[key] = now
    return _apply(st, kind, rows, only_folder=folder)


def sync_all(kind, progress=None):
    """Tai TOAN BO anh (Tem hoac Bao bi) tu may chu ve may nay."""
    st = _conn()
    rows = st.c.query(f"SELECT folder, name, sha1 FROM {TABLE} WHERE kind=%s ORDER BY folder, name", (kind,))
    return _apply(st, kind, rows, progress=progress)


def _apply(st, kind, rows, only_folder=None, progress=None):
    with _lock:
        base = cache_dir(kind)
        man = _manifest_load(kind)
        want = {}
        for folder, name, sha in rows:
            rel = os.path.join(_safe(folder), _safe(name))
            want[rel] = (folder, name, sha)
        n = 0
        total = len(want)
        for i, (rel, (folder, name, sha)) in enumerate(want.items(), 1):
            p = os.path.join(base, rel)
            if man.get(rel) == sha and os.path.isfile(p):
                continue
            r = st.c.one(f"SELECT data FROM {TABLE} WHERE kind=%s AND folder=%s AND name=%s",
                         (kind, folder, name))
            if not r:
                continue
            os.makedirs(os.path.dirname(p), exist_ok=True)
            tmp = p + '.tmp'
            with open(tmp, 'wb') as f:
                f.write(bytes(r[0]))
            os.replace(tmp, p)
            man[rel] = sha
            n += 1
            if progress:
                progress(i, total, rel)
        # xoa anh da bi xoa tren may chu (chi trong pham vi vua dong bo)
        for rel in list(man):
            fol = rel.split(os.sep)[0]
            in_scope = (only_folder is None) or fol.upper() == _safe(only_folder).upper()
            if in_scope and rel not in want:
                try:
                    os.remove(os.path.join(base, rel))
                except OSError:
                    pass
                man.pop(rel, None)
        _manifest_save(kind, man)
        return n


# ---------------------------------------------------------------------
#  May nay -> may chu
# ---------------------------------------------------------------------
def _server_index(st, kind):
    return {(f.upper(), n.upper()): (f, n, s)
            for f, n, s in st.c.query(f"SELECT folder, name, sha1 FROM {TABLE} WHERE kind=%s", (kind,))}


def _put(st, kind, folder, name, data, index):
    sha = _sha1(data)
    old = index.get((folder.upper(), name.upper()))
    if old and old[2] == sha:
        return False                                   # giong het -> bo qua
    if old:                                            # trung ten khac hoa/thuong -> ghi de ban cu
        folder, name = old[0], old[1]
    st.c.execute(f"INSERT INTO {TABLE}(kind, folder, name, data, size, sha1, updated_at, updated_by, host)"
                 " VALUES(%s, %s, %s, %s, %s, %s, now(), %s, %s)"
                 " ON CONFLICT (kind, folder, name) DO UPDATE SET data=EXCLUDED.data, size=EXCLUDED.size,"
                 " sha1=EXCLUDED.sha1, updated_at=now(), updated_by=EXCLUDED.updated_by, host=EXCLUDED.host",
                 (kind, folder, name, data, len(data), sha, _who(st), _host()))
    index[(folder.upper(), name.upper())] = (folder, name, sha)
    return True


def scan_folder(root):
    """Liet ke anh de tai len may chu -> [(Ma KH, ten file, duong dan)].
    Nhan CA 2 kieu sap xep:
      1) <root>\\<Ma KH>\\*.png     (anh trong thu muc con theo Ma KH)
      2) <root>\\*.png              (anh de thang trong thu muc goc -> suy Ma KH tu ten file)
    """
    out = []
    if not root or not os.path.isdir(root):
        return out
    try:
        entries = sorted(os.listdir(root))
    except OSError:
        return out
    for d in entries:
        p0 = os.path.join(root, d)
        if os.path.isdir(p0) and not d.startswith('.'):
            for f in sorted(os.listdir(p0)):
                p = os.path.join(p0, f)
                if f.lower().endswith(IMG_EXTS) and os.path.isfile(p):
                    out.append((d, f, p))
        elif os.path.isfile(p0) and d.lower().endswith(IMG_EXTS):
            out.append((guess_folder(d), d, p0))       # anh de thang trong goc
    return out


def upload_items(kind, items, progress=None):
    """items = [(folder, name, path)]. Tra ve (so_anh_gui, so_bo_qua, [loi])."""
    st = _conn()
    index = _server_index(st, kind)
    sent = skipped = 0
    errs = []
    total = len(items)
    for i, (folder, name, path) in enumerate(items, 1):
        try:
            if os.path.getsize(path) > MAX_BYTES:
                errs.append(f'{folder}\\{name}: lớn hơn {MAX_BYTES // (1024 * 1024)} MB - bỏ qua')
                continue
            with open(path, 'rb') as f:
                data = f.read()
            if _put(st, kind, str(folder).strip(), str(name).strip(), data, index):
                sent += 1
            else:
                skipped += 1
        except Exception as e:
            errs.append(f'{folder}\\{name}: {e}')
        if progress:
            progress(i, total, f'{folder}\\{name}')
    if sent:
        try:
            st.log(f'Tải {sent} ảnh {KIND_LABEL.get(kind, kind)} lên máy chủ')
        except Exception:
            pass
        _last_sync.clear()
    return sent, skipped, errs


def upload_folder(kind, root, progress=None):
    return upload_items(kind, scan_folder(root), progress)


def guess_folder(filename):
    """'KSP-HTP-P1014BF-41.png' -> 'KSP'; 'TBC.png' -> 'TBC'."""
    stem = os.path.splitext(os.path.basename(filename))[0].strip()
    for sep in ('-', '.', '_', ' '):
        if sep in stem:
            return stem.split(sep)[0].strip()
    return stem


def delete_image(kind, folder, name):
    st = _conn()
    st.c.execute(f"DELETE FROM {TABLE} WHERE kind=%s AND folder=%s AND name=%s", (kind, folder, name))
    _last_sync.clear()


def server_list(kind):
    """[(folder, name, size, updated_at, updated_by)]"""
    st = _conn()
    return st.c.query(f"SELECT folder, name, size, updated_at, updated_by FROM {TABLE}"
                      " WHERE kind=%s ORDER BY upper(folder), upper(name)", (kind,))


def server_counts():
    """{'stamp': n, 'pkg': m}"""
    st = _conn()
    out = {k: 0 for k in KINDS}
    for k, n in st.c.query(f"SELECT kind, count(*) FROM {TABLE} GROUP BY kind"):
        out[k] = int(n)
    return out
