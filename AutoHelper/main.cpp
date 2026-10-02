// ============================================================================
//  AutoHelper - Tool theo doi mau pixel (HP/MP) va tu dong nhan phim
//  Nen tang: Windows (WinAPI thuan, khong phu thuoc thu vien ngoai)
//  Ngon ngu: C++ (chuan C++11 tro len)
//
//  Y tuong:
//    - Ban chon mot (hoac nhieu) diem toa do tren man hinh (vi du mot diem
//      nam tren thanh HP khi mau con "du").
//    - Tool doc mau pixel tai diem do lien tuc.
//    - Khi mau KHONG con khop mau tham chieu (thanh mau da lui qua diem do)
//      -> tool nhan phim ban da gan (vi du phim dung thuoc hoi mau).
//
//  Toan bo cau hinh nam trong file config.ini de ban tu chinh, khong can
//  sua lai code.
//
//  BAM THEO CUA SO GAME:
//    - Neu dat window_title trong config, toa do x,y cua moi monitor duoc hieu
//      la TUONG DOI so voi goc tren-trai vung client cua cua so do.
//    - Khi ban keo/di chuyen cua so, tool tu tinh lai vi tri -> khong bi lech.
//    - De trong window_title => quay ve che do toa do man hinh tuyet doi (cu).
//
//  Phim dieu khien khi chay:
//    F9  : Bat / tat che do tu dong (toggle)
//    F10 : In ra mau pixel hien tai tai cac diem dang theo doi (de do mau)
//    F12 : Thoat chuong trinh
// ============================================================================

#include <windows.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <iostream>
#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cmath>
#include <cstdlib>

// ---------------------------------------------------------------------------
// Tien ich xu ly chuoi
// ---------------------------------------------------------------------------
static std::string trim(const std::string& s) {
    size_t a = s.find_first_not_of(" \t\r\n");
    if (a == std::string::npos) return "";
    size_t b = s.find_last_not_of(" \t\r\n");
    return s.substr(a, b - a + 1);
}

static std::string toUpper(std::string s) {
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char c) { return (char)std::toupper(c); });
    return s;
}

// ---------------------------------------------------------------------------
// Chuyen ten phim (trong config) sang ma Virtual-Key cua Windows
// Ho tro: A-Z, 0-9, F1-F12, va mot so phim dac biet pho bien.
// ---------------------------------------------------------------------------
static WORD keyNameToVk(const std::string& rawName) {
    std::string name = toUpper(trim(rawName));
    if (name.empty()) return 0;

    // Phim chu cai don A-Z
    if (name.size() == 1 && name[0] >= 'A' && name[0] <= 'Z')
        return (WORD)name[0];

    // Phim so don 0-9 (hang phim so tren cung)
    if (name.size() == 1 && name[0] >= '0' && name[0] <= '9')
        return (WORD)name[0];

    // Phim chuc nang F1 - F12
    if (name.size() >= 2 && name[0] == 'F') {
        int n = atoi(name.c_str() + 1);
        if (n >= 1 && n <= 12) return (WORD)(VK_F1 + (n - 1));
    }

    // Phim so ban phim numpad: NUM0 - NUM9
    if (name.rfind("NUM", 0) == 0 && name.size() == 4 &&
        name[3] >= '0' && name[3] <= '9')
        return (WORD)(VK_NUMPAD0 + (name[3] - '0'));

    // Mot so phim dac biet
    if (name == "SPACE")  return VK_SPACE;
    if (name == "ENTER")  return VK_RETURN;
    if (name == "TAB")    return VK_TAB;
    if (name == "ESC" || name == "ESCAPE") return VK_ESCAPE;
    if (name == "SHIFT")  return VK_SHIFT;
    if (name == "CTRL" || name == "CONTROL") return VK_CONTROL;
    if (name == "ALT")    return VK_MENU;

    return 0; // khong nhan dang duoc
}

// ---------------------------------------------------------------------------
// Tim cua so game theo chuoi con trong tieu de (khong phan biet hoa/thuong)
// EnumWindows duyet moi cua so top-level dang hien thi.
// ---------------------------------------------------------------------------
static HWND        g_foundHwnd = NULL;
static std::string g_titleNeedleUpper;

static BOOL CALLBACK enumWindowProc(HWND h, LPARAM) {
    if (!IsWindowVisible(h)) return TRUE;
    char buf[512] = {0};
    int len = GetWindowTextA(h, buf, (int)sizeof(buf) - 1);
    if (len <= 0) return TRUE;
    std::string title(buf);
    // so khop chuoi con (title viet hoa)
    std::string up = title;
    std::transform(up.begin(), up.end(), up.begin(),
                   [](unsigned char c) { return (char)std::toupper(c); });
    if (up.find(g_titleNeedleUpper) != std::string::npos) {
        g_foundHwnd = h;
        return FALSE;   // dung duyet khi da tim thay
    }
    return TRUE;
}

static HWND findGameWindow(const std::string& needleUpper) {
    g_foundHwnd = NULL;
    g_titleNeedleUpper = needleUpper;
    EnumWindows(enumWindowProc, 0);
    return g_foundHwnd;
}

// ---------------------------------------------------------------------------
// Cau truc mot "diem theo doi"
// ---------------------------------------------------------------------------
struct PixelMonitor {
    std::string name = "monitor";  // ten de hien thi log
    int x = 0;                     // toa do X tren man hinh
    int y = 0;                     // toa do Y tren man hinh
    int r = 0, g = 0, b = 0;       // mau tham chieu (mau "binh thuong")
    int tolerance = 30;            // do sai lech mau cho phep (0-441)
    WORD vk = 0;                   // phim se nhan khi dieu kien thoa man
    std::string keyName = "";      // ten phim (de log)

    // trigger = "absent": nhan phim khi pixel KHONG con khop mau tham chieu
    //                     (dung cho HP/MP: thanh mau tut -> doi mau)
    // trigger = "present": nhan phim khi pixel KHOP mau tham chieu
    //                     (dung khi doi mot thong bao/icon xuat hien)
    std::string trigger = "absent";

    int cooldownMs = 800;          // thoi gian cho toi thieu giua 2 lan nhan
    DWORD lastFire = 0;            // thoi diem nhan phim gan nhat (noi bo)
};

// ---------------------------------------------------------------------------
// Cau truc mot "phim lap theo chu ky" (auto danh / auto ky nang)
//   Cu moi interval_ms (co the +/- jitter_ms ngau nhien) thi nhan 1 lan phim.
// ---------------------------------------------------------------------------
struct TimerAction {
    std::string name = "timer";    // ten de hien thi log
    WORD vk = 0;                    // phim se nhan
    std::string keyName = "";      // ten phim (de log)
    int intervalMs = 1000;         // chu ky nhan (ms)
    int jitterMs = 0;              // dao ngau nhien +/- (ms) cho tu nhien hon
    DWORD nextFire = 0;           // thoi diem nhan ke tiep (noi bo)
};

// ---------------------------------------------------------------------------
// Doc mau 1 pixel.
//   hwnd == NULL                : (x,y) la toa do man hinh tuyet doi.
//   hwnd hop le, useWindowDC=0  : (x,y) tuong doi vung client -> doi sang
//                                 toa do man hinh roi doc tu DC man hinh.
//                                 (Cua so can hien ra, khong bi che.)
//   hwnd hop le, useWindowDC=1  : doc truc tiep tu DC cua so tai (x,y) client.
//                                 (Chay duoc ca khi cua so khong focus; nhung
//                                  mot so game DirectX co the tra ve mau den.)
// ---------------------------------------------------------------------------
static COLORREF getPixelColor(HWND hwnd, bool useWindowDC, int x, int y) {
    if (hwnd != NULL && useWindowDC) {
        HDC hdc = GetDC(hwnd);
        COLORREF c = GetPixel(hdc, x, y);
        ReleaseDC(hwnd, hdc);
        return c;
    }
    if (hwnd != NULL) {
        POINT p = { x, y };
        ClientToScreen(hwnd, &p);   // tuong doi client -> man hinh
        x = p.x; y = p.y;
    }
    HDC hdc = GetDC(NULL);          // DC cua toan man hinh
    COLORREF c = GetPixel(hdc, x, y);
    ReleaseDC(NULL, hdc);
    return c;
}

// Khoang cach mau theo khong gian RGB (Euclid). Cang nho cang giong nhau.
static int colorDistance(int r1, int g1, int b1, int r2, int g2, int b2) {
    int dr = r1 - r2, dg = g1 - g2, db = b1 - b2;
    // tra ve binh phuong khoang cach de so sanh (tranh sqrt cho nhanh)
    long sq = (long)dr * dr + (long)dg * dg + (long)db * db;
    // nhung de nguong "tolerance" de hieu, ta tra ve khoang cach that
    return (int)(std::sqrt((double)sq) + 0.5);
}

// ---------------------------------------------------------------------------
// Gui mot lan nhan phim (down + up) toi cua so dang active
// ---------------------------------------------------------------------------
static void sendKey(WORD vk) {
    INPUT in[2] = {};

    in[0].type = INPUT_KEYBOARD;
    in[0].ki.wVk = vk;

    in[1].type = INPUT_KEYBOARD;
    in[1].ki.wVk = vk;
    in[1].ki.dwFlags = KEYEVENTF_KEYUP;

    SendInput(2, in, sizeof(INPUT));
}

// Tinh chu ky ke tiep cua mot timer (co dao ngau nhien +/- jitter)
static DWORD nextInterval(const TimerAction& t) {
    int j = 0;
    if (t.jitterMs > 0) j = (rand() % (2 * t.jitterMs + 1)) - t.jitterMs;
    int iv = t.intervalMs + j;
    if (iv < 10) iv = 10;
    return (DWORD)iv;
}

// ---------------------------------------------------------------------------
// Doc file config.ini
// Dinh dang moi dong trong 1 [monitor] la key=value.
// Cac block duoc ngan cach boi dong [monitor].
// ---------------------------------------------------------------------------
struct AppConfig {
    int pollIntervalMs = 100;          // chu ky quet (ms)
    std::string windowTitle = "";      // chuoi con trong tieu de cua so game
                                       // (de trong = dung toa do man hinh tuyet doi)
    bool useWindowDC = false;          // true = doc truc tiep tu DC cua so
    std::vector<PixelMonitor> monitors;
    std::vector<TimerAction>  timers;
};

static bool loadConfig(const std::string& path, AppConfig& cfg, std::string& err) {
    std::ifstream fin(path.c_str());
    if (!fin) {
        err = "Khong mo duoc file config: " + path;
        return false;
    }

    enum Block { B_GENERAL, B_MONITOR, B_TIMER };
    Block blk = B_GENERAL;
    PixelMonitor curM;
    TimerAction  curT;

    // Luu lai block vua xong truoc khi chuyen sang block khac
    auto flush = [&]() {
        if (blk == B_MONITOR) {
            curM.vk = keyNameToVk(curM.keyName);
            cfg.monitors.push_back(curM);
        } else if (blk == B_TIMER) {
            curT.vk = keyNameToVk(curT.keyName);
            cfg.timers.push_back(curT);
        }
    };

    std::string line;
    while (std::getline(fin, line)) {
        line = trim(line);
        if (line.empty() || line[0] == '#' || line[0] == ';') continue;

        std::string up = toUpper(line);
        if (up == "[MONITOR]") {
            flush(); curM = PixelMonitor(); blk = B_MONITOR; continue;
        }
        if (up == "[TIMER]") {
            flush(); curT = TimerAction(); blk = B_TIMER; continue;
        }
        if (up == "[GENERAL]") {
            flush(); blk = B_GENERAL; continue;
        }

        size_t eq = line.find('=');
        if (eq == std::string::npos) continue;
        std::string key = toUpper(trim(line.substr(0, eq)));
        std::string val = trim(line.substr(eq + 1));

        if (blk == B_GENERAL) {
            if      (key == "POLL_INTERVAL_MS") cfg.pollIntervalMs = atoi(val.c_str());
            else if (key == "WINDOW_TITLE")     cfg.windowTitle = val;
            else if (key == "CAPTURE_MODE")     cfg.useWindowDC = (toUpper(val) == "WINDOW");
        }
        else if (blk == B_MONITOR) {
            if      (key == "NAME")        curM.name = val;
            else if (key == "X")           curM.x = atoi(val.c_str());
            else if (key == "Y")           curM.y = atoi(val.c_str());
            else if (key == "R")           curM.r = atoi(val.c_str());
            else if (key == "G")           curM.g = atoi(val.c_str());
            else if (key == "B")           curM.b = atoi(val.c_str());
            else if (key == "TOLERANCE")   curM.tolerance = atoi(val.c_str());
            else if (key == "KEY")         curM.keyName = val;
            else if (key == "TRIGGER")     curM.trigger = toUpper(val) == "PRESENT" ? "present" : "absent";
            else if (key == "COOLDOWN_MS") curM.cooldownMs = atoi(val.c_str());
        }
        else if (blk == B_TIMER) {
            if      (key == "NAME")        curT.name = val;
            else if (key == "KEY")         curT.keyName = val;
            else if (key == "INTERVAL_MS") curT.intervalMs = atoi(val.c_str());
            else if (key == "JITTER_MS")   curT.jitterMs = atoi(val.c_str());
        }
    }
    flush();

    if (cfg.monitors.empty() && cfg.timers.empty()) {
        err = "Config khong co [monitor] hay [timer] nao.";
        return false;
    }
    if (cfg.pollIntervalMs < 10) cfg.pollIntervalMs = 10;
    for (auto& t : cfg.timers) {
        if (t.intervalMs < 10) t.intervalMs = 10;
        if (t.jitterMs < 0)    t.jitterMs = 0;
    }
    return true;
}

// ---------------------------------------------------------------------------
// In ra mau hien tai tai cac diem (de ban do mau de dua vao config)
// ---------------------------------------------------------------------------
static void dumpColors(const AppConfig& cfg, HWND hwnd) {
    std::cout << "\n--- Mau pixel hien tai ---\n";
    for (const auto& m : cfg.monitors) {
        COLORREF c = getPixelColor(hwnd, cfg.useWindowDC, m.x, m.y);
        std::cout << "  [" << m.name << "] ("
                  << m.x << "," << m.y << ") = R:"
                  << (int)GetRValue(c) << " G:"
                  << (int)GetGValue(c) << " B:"
                  << (int)GetBValue(c) << "\n";
    }
    std::cout << "--------------------------\n";
}

// ---------------------------------------------------------------------------
// MAIN
// ---------------------------------------------------------------------------
int main() {
    SetConsoleTitleA("AutoHelper - HP/MP pixel watcher");

    // Giup toa do GetPixel/ClientToScreen khop nhau khi Windows dang scale DPI.
    SetProcessDPIAware();
    srand((unsigned)GetTickCount());   // seed cho jitter ngau nhien

    AppConfig cfg;
    std::string err;
    if (!loadConfig("config.ini", cfg, err)) {
        std::cerr << "[LOI] " << err << "\n";
        std::cout << "Nhan Enter de thoat...";
        std::cin.get();
        return 1;
    }

    // Tim cua so game (neu co cau hinh window_title)
    std::string needleUpper = toUpper(cfg.windowTitle);
    HWND hwnd = NULL;
    if (!cfg.windowTitle.empty()) {
        hwnd = findGameWindow(needleUpper);
    }

    std::cout << "============================================\n";
    std::cout << " AutoHelper da nap cau hinh thanh cong.\n";
    std::cout << " So diem theo doi : " << cfg.monitors.size() << "\n";
    std::cout << " So phim chu ky   : " << cfg.timers.size() << "\n";
    std::cout << " Chu ky quet      : " << cfg.pollIntervalMs << " ms\n";
    if (cfg.windowTitle.empty()) {
        std::cout << " Che do toa do    : MAN HINH TUYET DOI\n";
    } else {
        std::cout << " Bam cua so       : \"" << cfg.windowTitle << "\""
                  << (cfg.useWindowDC ? " (doc tu DC cua so)" : " (doc tu man hinh)") << "\n";
        std::cout << " Trang thai cua so: "
                  << (hwnd ? "DA TIM THAY" : "CHUA THAY (se tu tim lai khi chay)") << "\n";
    }
    std::cout << "--------------------------------------------\n";
    for (const auto& m : cfg.monitors) {
        std::cout << "  * " << m.name << " @(" << m.x << "," << m.y << ") "
                  << "mau(" << m.r << "," << m.g << "," << m.b << ") "
                  << "sai lech<=" << m.tolerance << " "
                  << "phim=" << m.keyName
                  << (m.vk == 0 ? " [!! phim khong hop le]" : "")
                  << " trigger=" << m.trigger << "\n";
    }
    for (const auto& t : cfg.timers) {
        std::cout << "  ~ " << t.name << " phim=" << t.keyName
                  << (t.vk == 0 ? " [!! phim khong hop le]" : "")
                  << " moi " << t.intervalMs << "ms"
                  << (t.jitterMs ? (" +/-" + std::to_string(t.jitterMs) + "ms") : "")
                  << "\n";
    }
    std::cout << "============================================\n";
    std::cout << " F9  = Bat/Tat tu dong\n";
    std::cout << " F10 = In mau pixel hien tai (do mau)\n";
    std::cout << " F12 = Thoat\n";
    std::cout << "============================================\n";
    std::cout << " Trang thai: TAT (nhan F9 de bat)\n";

    bool running = false;      // che do tu dong dang bat?
    bool prevF9 = false, prevF10 = false, prevF12 = false;

    while (true) {
        // --- Doc cac phim dieu khien (phat hien suon len de tranh lap) ---
        bool f9  = (GetAsyncKeyState(VK_F9)  & 0x8000) != 0;
        bool f10 = (GetAsyncKeyState(VK_F10) & 0x8000) != 0;
        bool f12 = (GetAsyncKeyState(VK_F12) & 0x8000) != 0;

        if (f9 && !prevF9) {
            running = !running;
            std::cout << " Trang thai: " << (running ? "BAT" : "TAT") << "\n";
            if (running) {
                // khi vua bat: lap lich cho cac timer (nhan ngay o chu ky dau)
                DWORD now = GetTickCount();
                for (auto& t : cfg.timers) t.nextFire = now;
            }
        }
        if (f10 && !prevF10) {
            dumpColors(cfg, hwnd);
        }
        if (f12 && !prevF12) {
            std::cout << " Thoat chuong trinh.\n";
            break;
        }
        prevF9 = f9; prevF10 = f10; prevF12 = f12;

        // --- Neu dang bat, quet cac diem theo doi ---
        if (running) {
            // Neu bam theo cua so: dam bao handle con hop le, neu khong thi tim lai.
            if (!cfg.windowTitle.empty() && (hwnd == NULL || !IsWindow(hwnd))) {
                HWND re = findGameWindow(needleUpper);
                if (re != hwnd) {
                    hwnd = re;
                    std::cout << (hwnd ? " [cua so] da tim thay lai.\n"
                                       : " [cua so] khong thay, dang cho...\n");
                }
            }

            DWORD now = GetTickCount();
            for (auto& m : cfg.monitors) {
                if (m.vk == 0) continue;  // bo qua phim khong hop le
                if (!cfg.windowTitle.empty() && hwnd == NULL) continue; // chua co cua so

                COLORREF c = getPixelColor(hwnd, cfg.useWindowDC, m.x, m.y);
                int dist = colorDistance(GetRValue(c), GetGValue(c), GetBValue(c),
                                         m.r, m.g, m.b);
                bool matches = (dist <= m.tolerance);

                bool shouldFire =
                    (m.trigger == "absent"  && !matches) ||
                    (m.trigger == "present" &&  matches);

                if (shouldFire && (now - m.lastFire) >= (DWORD)m.cooldownMs) {
                    sendKey(m.vk);
                    m.lastFire = now;
                    std::cout << " [" << m.name << "] -> nhan phim "
                              << m.keyName << " (dist=" << dist << ")\n";
                }
            }

            // --- Phim lap theo chu ky (auto danh / auto ky nang) ---
            bool needWindow = !cfg.windowTitle.empty();
            for (auto& t : cfg.timers) {
                if (t.vk == 0) continue;               // phim khong hop le
                if (needWindow && hwnd == NULL) continue; // chua co cua so
                if ((long)(now - t.nextFire) >= 0) {
                    sendKey(t.vk);
                    t.nextFire = now + nextInterval(t);
                    std::cout << " [" << t.name << "] ~ nhan phim "
                              << t.keyName << "\n";
                }
            }
        }

        Sleep(cfg.pollIntervalMs);
    }

    return 0;
}
