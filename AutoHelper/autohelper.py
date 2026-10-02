#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ============================================================================
#  AutoHelper (ban Python) - Tool theo doi mau pixel (HP/MP) va tu dong nhan phim
#  Nen tang: Windows (dung ctypes goi WinAPI truc tiep, KHONG can thu vien ngoai)
#
#  Day la ban port tu AutoHelper (C++) - code cua rieng ban, viet moi tu dau.
#  Doc CUNG mot file config.ini nhu ban C++.
#
#  3 khoi hanh dong chay song song:
#    [monitor] : doc mau pixel (vd thanh HP/MP) -> nhan phim khi mau thay doi
#    [timer]   : nhan 1 phim lap theo chu ky (auto danh / auto ky nang)
#    [combo]   : nhan nhieu phim lan luot theo thu tu (1->2->3->4)
#
#  Phim dieu khien khi chay:
#    F9  : Bat / tat che do tu dong (toggle)
#    F10 : In ra mau pixel hien tai tai cac diem theo doi (de do mau)
#    F12 : Thoat chuong trinh
# ============================================================================

import sys
import time
import math
import random
import ctypes
from ctypes import wintypes

# ---------------------------------------------------------------------------
# Nap cac thu vien he thong cua Windows
# ---------------------------------------------------------------------------
user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

# --- Ma Virtual-Key ---
VK_F1 = 0x70
VK_NUMPAD0 = 0x60
VK_SPACE = 0x20
VK_RETURN = 0x0D
VK_TAB = 0x09
VK_ESCAPE = 0x1B
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_F9 = 0x78
VK_F10 = 0x79
VK_F12 = 0x7B

# --- Hang so cho SendInput ---
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

ULONG_PTR = wintypes.WPARAM


# --- Cau truc du lieu cho SendInput ---
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class _INPUTunion(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTunion)]


# --- Khai bao prototype ham WinAPI (giup ctypes truyen tham so dung) ---
user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
user32.SendInput.restype = wintypes.UINT

user32.GetDC.argtypes = (wintypes.HWND,)
user32.GetDC.restype = wintypes.HDC
user32.ReleaseDC.argtypes = (wintypes.HWND, wintypes.HDC)
user32.ReleaseDC.restype = ctypes.c_int

gdi32.GetPixel.argtypes = (wintypes.HDC, ctypes.c_int, ctypes.c_int)
gdi32.GetPixel.restype = wintypes.COLORREF  # tra -1 (CLR_INVALID) neu loi

user32.ClientToScreen.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.POINT))
user32.ClientToScreen.restype = wintypes.BOOL

user32.GetAsyncKeyState.argtypes = (ctypes.c_int,)
user32.GetAsyncKeyState.restype = ctypes.c_short

user32.IsWindow.argtypes = (wintypes.HWND,)
user32.IsWindow.restype = wintypes.BOOL
user32.IsWindowVisible.argtypes = (wintypes.HWND,)
user32.IsWindowVisible.restype = wintypes.BOOL
user32.GetWindowTextW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
user32.GetWindowTextW.restype = ctypes.c_int

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
user32.EnumWindows.argtypes = (WNDENUMPROC, wintypes.LPARAM)
user32.EnumWindows.restype = wintypes.BOOL


# ---------------------------------------------------------------------------
# Tien ich chung
# ---------------------------------------------------------------------------
def now_ms():
    """Thoi diem hien tai tinh bang mili-giay (tang deu, khong lui)."""
    return int(time.monotonic() * 1000)


def key_name_to_vk(raw_name):
    """Chuyen ten phim (trong config) sang ma Virtual-Key. 0 = khong hop le."""
    name = raw_name.strip().upper()
    if not name:
        return 0
    if len(name) == 1 and "A" <= name <= "Z":
        return ord(name)
    if len(name) == 1 and "0" <= name <= "9":
        return ord(name)
    if len(name) >= 2 and name[0] == "F" and name[1:].isdigit():
        n = int(name[1:])
        if 1 <= n <= 12:
            return VK_F1 + (n - 1)
    if name.startswith("NUM") and len(name) == 4 and name[3].isdigit():
        return VK_NUMPAD0 + (ord(name[3]) - ord("0"))
    special = {
        "SPACE": VK_SPACE, "ENTER": VK_RETURN, "TAB": VK_TAB,
        "ESC": VK_ESCAPE, "ESCAPE": VK_ESCAPE, "SHIFT": VK_SHIFT,
        "CTRL": VK_CONTROL, "CONTROL": VK_CONTROL, "ALT": VK_MENU,
    }
    return special.get(name, 0)


def jittered_delay(base, jitter):
    """Mot khoang cho (ms) co dao ngau nhien +/- jitter, toi thieu 10ms."""
    j = random.randint(-jitter, jitter) if jitter > 0 else 0
    return max(10, base + j)


# ---------------------------------------------------------------------------
# Gui 1 lan nhan phim (down + up) toi cua so dang active
# ---------------------------------------------------------------------------
def send_key(vk):
    down = INPUT(type=INPUT_KEYBOARD)
    down.u.ki = KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=0)
    up = INPUT(type=INPUT_KEYBOARD)
    up.u.ki = KEYBDINPUT(wVk=vk, wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)
    arr = (INPUT * 2)(down, up)
    user32.SendInput(2, arr, ctypes.sizeof(INPUT))


# ---------------------------------------------------------------------------
# Tim cua so game theo chuoi con trong tieu de (khong phan biet hoa/thuong)
# ---------------------------------------------------------------------------
def find_game_window(needle_upper):
    found = [None]

    def _cb(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        buf = ctypes.create_unicode_buffer(512)
        if user32.GetWindowTextW(hwnd, buf, 512) <= 0:
            return True
        if needle_upper in buf.value.upper():
            found[0] = hwnd
            return False  # dung duyet khi da tim thay
        return True

    user32.EnumWindows(WNDENUMPROC(_cb), 0)
    return found[0]


# ---------------------------------------------------------------------------
# Doc mau 1 pixel.
#   hwnd None               : (x,y) la toa do man hinh tuyet doi.
#   hwnd + use_window_dc=F  : (x,y) tuong doi client -> doi sang toa do man hinh.
#   hwnd + use_window_dc=T  : doc truc tiep tu DC cua so tai (x,y) client.
# Tra ve (r, g, b).
# ---------------------------------------------------------------------------
def get_pixel_color(hwnd, use_window_dc, x, y):
    if hwnd and use_window_dc:
        hdc = user32.GetDC(hwnd)
        c = gdi32.GetPixel(hdc, x, y)
        user32.ReleaseDC(hwnd, hdc)
    else:
        if hwnd:
            p = wintypes.POINT(x, y)
            user32.ClientToScreen(hwnd, ctypes.byref(p))
            x, y = p.x, p.y
        hdc = user32.GetDC(None)
        c = gdi32.GetPixel(hdc, x, y)
        user32.ReleaseDC(None, hdc)
    if c == 0xFFFFFFFF:  # CLR_INVALID
        return (-1, -1, -1)
    return (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF)


def color_distance(r1, g1, b1, r2, g2, b2):
    return math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2)


# ---------------------------------------------------------------------------
# Cac cau truc cau hinh
# ---------------------------------------------------------------------------
class PixelMonitor:
    def __init__(self):
        self.name = "monitor"
        self.x = 0
        self.y = 0
        self.r = 0
        self.g = 0
        self.b = 0
        self.tolerance = 30
        self.key_name = ""
        self.vk = 0
        self.trigger = "absent"   # "absent" hoac "present"
        self.cooldown_ms = 800
        self.last_fire = 0


class TimerAction:
    def __init__(self):
        self.name = "timer"
        self.key_name = ""
        self.vk = 0
        self.interval_ms = 1000
        self.jitter_ms = 0
        self.next_fire = 0


class ComboAction:
    def __init__(self):
        self.name = "combo"
        self.vks = []
        self.key_names = []
        self.step_delay_ms = 300
        self.loop_delay_ms = 1000
        self.jitter_ms = 0
        self.index = 0
        self.next_fire = 0


class AppConfig:
    def __init__(self):
        self.poll_interval_ms = 100
        self.window_title = ""
        self.use_window_dc = False
        self.monitors = []
        self.timers = []
        self.combos = []


# ---------------------------------------------------------------------------
# Doc file config.ini (cung dinh dang voi ban C++)
# ---------------------------------------------------------------------------
def load_config(path):
    cfg = AppConfig()
    blk = "general"
    cur_m = cur_t = cur_c = None

    def flush():
        if cur_m is not None and blk == "monitor":
            cur_m.vk = key_name_to_vk(cur_m.key_name)
            cfg.monitors.append(cur_m)
        elif cur_t is not None and blk == "timer":
            cur_t.vk = key_name_to_vk(cur_t.key_name)
            cfg.timers.append(cur_t)
        elif cur_c is not None and blk == "combo":
            cfg.combos.append(cur_c)

    with open(path, "r", encoding="utf-8", errors="replace") as fin:
        for raw in fin:
            line = raw.strip()
            if not line or line[0] in "#;":
                continue
            up = line.upper()
            if up == "[MONITOR]":
                flush(); cur_m = PixelMonitor(); blk = "monitor"; continue
            if up == "[TIMER]":
                flush(); cur_t = TimerAction(); blk = "timer"; continue
            if up == "[COMBO]":
                flush(); cur_c = ComboAction(); blk = "combo"; continue
            if up == "[GENERAL]":
                flush(); blk = "general"; continue

            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            key = k.strip().upper()
            val = v.strip()

            if blk == "general":
                if key == "POLL_INTERVAL_MS":
                    cfg.poll_interval_ms = int_or(val, cfg.poll_interval_ms)
                elif key == "WINDOW_TITLE":
                    cfg.window_title = val
                elif key == "CAPTURE_MODE":
                    cfg.use_window_dc = (val.upper() == "WINDOW")
            elif blk == "monitor":
                if key == "NAME": cur_m.name = val
                elif key == "X": cur_m.x = int_or(val, cur_m.x)
                elif key == "Y": cur_m.y = int_or(val, cur_m.y)
                elif key == "R": cur_m.r = int_or(val, cur_m.r)
                elif key == "G": cur_m.g = int_or(val, cur_m.g)
                elif key == "B": cur_m.b = int_or(val, cur_m.b)
                elif key == "TOLERANCE": cur_m.tolerance = int_or(val, cur_m.tolerance)
                elif key == "KEY": cur_m.key_name = val
                elif key == "TRIGGER":
                    cur_m.trigger = "present" if val.upper() == "PRESENT" else "absent"
                elif key == "COOLDOWN_MS": cur_m.cooldown_ms = int_or(val, cur_m.cooldown_ms)
            elif blk == "timer":
                if key == "NAME": cur_t.name = val
                elif key == "KEY": cur_t.key_name = val
                elif key == "INTERVAL_MS": cur_t.interval_ms = int_or(val, cur_t.interval_ms)
                elif key == "JITTER_MS": cur_t.jitter_ms = int_or(val, cur_t.jitter_ms)
            elif blk == "combo":
                if key == "NAME": cur_c.name = val
                elif key == "KEYS":
                    cur_c.vks = []
                    cur_c.key_names = []
                    for item in val.split(","):
                        item = item.strip()
                        if not item:
                            continue
                        cur_c.vks.append(key_name_to_vk(item))
                        cur_c.key_names.append(item)
                elif key == "STEP_DELAY_MS": cur_c.step_delay_ms = int_or(val, cur_c.step_delay_ms)
                elif key == "LOOP_DELAY_MS": cur_c.loop_delay_ms = int_or(val, cur_c.loop_delay_ms)
                elif key == "JITTER_MS": cur_c.jitter_ms = int_or(val, cur_c.jitter_ms)
    flush()

    if not cfg.monitors and not cfg.timers and not cfg.combos:
        raise ValueError("Config khong co [monitor], [timer] hay [combo] nao.")

    cfg.poll_interval_ms = max(10, cfg.poll_interval_ms)
    for t in cfg.timers:
        t.interval_ms = max(10, t.interval_ms)
        t.jitter_ms = max(0, t.jitter_ms)
    for c in cfg.combos:
        c.step_delay_ms = max(10, c.step_delay_ms)
        c.loop_delay_ms = max(10, c.loop_delay_ms)
        c.jitter_ms = max(0, c.jitter_ms)
    return cfg


def int_or(s, default):
    """atoi kieu C: lay phan so dau chuoi, loi thi tra default."""
    s = s.strip()
    i = 0
    if i < len(s) and s[i] in "+-":
        i += 1
    while i < len(s) and s[i].isdigit():
        i += 1
    try:
        return int(s[:i])
    except ValueError:
        return default


# ---------------------------------------------------------------------------
# In mau hien tai tai cac diem (de do mau cho config)
# ---------------------------------------------------------------------------
def dump_colors(cfg, hwnd):
    print("\n--- Mau pixel hien tai ---")
    for m in cfg.monitors:
        r, g, b = get_pixel_color(hwnd, cfg.use_window_dc, m.x, m.y)
        print("  [%s] (%d,%d) = R:%d G:%d B:%d" % (m.name, m.x, m.y, r, g, b))
    print("--------------------------")


def key_down(vk):
    return (user32.GetAsyncKeyState(vk) & 0x8000) != 0


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    if sys.platform != "win32":
        print("[LOI] Tool nay chi chay tren Windows.")
        return 1

    # Giup toa do GetPixel/ClientToScreen khop nhau khi Windows scale DPI
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

    try:
        cfg = load_config("config.ini")
    except Exception as e:
        print("[LOI] %s" % e)
        input("Nhan Enter de thoat...")
        return 1

    needle_upper = cfg.window_title.upper()
    hwnd = find_game_window(needle_upper) if cfg.window_title else None

    print("============================================")
    print(" AutoHelper (Python) da nap cau hinh thanh cong.")
    print(" So diem theo doi : %d" % len(cfg.monitors))
    print(" So phim chu ky   : %d" % len(cfg.timers))
    print(" So chuoi combo   : %d" % len(cfg.combos))
    print(" Chu ky quet      : %d ms" % cfg.poll_interval_ms)
    if not cfg.window_title:
        print(" Che do toa do    : MAN HINH TUYET DOI")
    else:
        print(' Bam cua so       : "%s" %s' % (
            cfg.window_title,
            "(doc tu DC cua so)" if cfg.use_window_dc else "(doc tu man hinh)"))
        print(" Trang thai cua so: %s" % (
            "DA TIM THAY" if hwnd else "CHUA THAY (se tu tim lai khi chay)"))
    print("--------------------------------------------")
    for m in cfg.monitors:
        print("  * %s @(%d,%d) mau(%d,%d,%d) sai lech<=%d phim=%s%s trigger=%s" % (
            m.name, m.x, m.y, m.r, m.g, m.b, m.tolerance, m.key_name,
            "" if m.vk else " [!! phim khong hop le]", m.trigger))
    for t in cfg.timers:
        extra = (" +/-%dms" % t.jitter_ms) if t.jitter_ms else ""
        print("  ~ %s phim=%s%s moi %dms%s" % (
            t.name, t.key_name, "" if t.vk else " [!! phim khong hop le]",
            t.interval_ms, extra))
    for c in cfg.combos:
        seq = "->".join(kn + ("(!)" if vk == 0 else "")
                        for kn, vk in zip(c.key_names, c.vks))
        extra = (" +/-%dms" % c.jitter_ms) if c.jitter_ms else ""
        print("  > %s [%s] step=%dms loop=%dms%s" % (
            c.name, seq, c.step_delay_ms, c.loop_delay_ms, extra))
    print("============================================")
    print(" F9  = Bat/Tat tu dong")
    print(" F10 = In mau pixel hien tai (do mau)")
    print(" F12 = Thoat")
    print("============================================")
    print(" Trang thai: TAT (nhan F9 de bat)")

    running = False
    prev_f9 = prev_f10 = prev_f12 = False
    need_window = bool(cfg.window_title)

    while True:
        f9, f10, f12 = key_down(VK_F9), key_down(VK_F10), key_down(VK_F12)

        if f9 and not prev_f9:
            running = not running
            print(" Trang thai: %s" % ("BAT" if running else "TAT"))
            if running:
                t0 = now_ms()
                for t in cfg.timers:
                    t.next_fire = t0
                for c in cfg.combos:
                    c.index = 0
                    c.next_fire = t0
        if f10 and not prev_f10:
            dump_colors(cfg, hwnd)
        if f12 and not prev_f12:
            print(" Thoat chuong trinh.")
            break
        prev_f9, prev_f10, prev_f12 = f9, f10, f12

        if running:
            # Neu bam theo cua so: dam bao handle con hop le, neu khong thi tim lai
            if need_window and (hwnd is None or not user32.IsWindow(hwnd)):
                re = find_game_window(needle_upper)
                if re != hwnd:
                    hwnd = re
                    print(" [cua so] da tim thay lai." if hwnd
                          else " [cua so] khong thay, dang cho...")

            t_now = now_ms()

            # --- [monitor]: theo mau pixel ---
            for m in cfg.monitors:
                if m.vk == 0:
                    continue
                if need_window and hwnd is None:
                    continue
                r, g, b = get_pixel_color(hwnd, cfg.use_window_dc, m.x, m.y)
                dist = color_distance(r, g, b, m.r, m.g, m.b)
                matches = dist <= m.tolerance
                should_fire = ((m.trigger == "absent" and not matches) or
                               (m.trigger == "present" and matches))
                if should_fire and (t_now - m.last_fire) >= m.cooldown_ms:
                    send_key(m.vk)
                    m.last_fire = t_now
                    print(" [%s] -> nhan phim %s (dist=%d)" % (m.name, m.key_name, int(dist)))

            # --- [timer]: phim lap theo chu ky ---
            for t in cfg.timers:
                if t.vk == 0:
                    continue
                if need_window and hwnd is None:
                    continue
                if t_now - t.next_fire >= 0:
                    send_key(t.vk)
                    t.next_fire = t_now + jittered_delay(t.interval_ms, t.jitter_ms)
                    print(" [%s] ~ nhan phim %s" % (t.name, t.key_name))

            # --- [combo]: chuoi phim theo thu tu ---
            for c in cfg.combos:
                if not c.vks:
                    continue
                if need_window and hwnd is None:
                    continue
                if t_now - c.next_fire < 0:
                    continue
                vk = c.vks[c.index]
                kn = c.key_names[c.index]
                if vk != 0:
                    send_key(vk)
                    print(" [%s] > nhan phim %s (%d/%d)" % (c.name, kn, c.index + 1, len(c.vks)))
                c.index += 1
                if c.index >= len(c.vks):
                    c.index = 0
                    c.next_fire = t_now + jittered_delay(c.loop_delay_ms, c.jitter_ms)
                else:
                    c.next_fire = t_now + jittered_delay(c.step_delay_ms, c.jitter_ms)

        time.sleep(cfg.poll_interval_ms / 1000.0)

    return 0


if __name__ == "__main__":
    sys.exit(main())
