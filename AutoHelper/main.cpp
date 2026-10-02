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
// Doc mau 1 pixel tren toan man hinh
// ---------------------------------------------------------------------------
static COLORREF getPixelColor(int x, int y) {
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

// ---------------------------------------------------------------------------
// Doc file config.ini
// Dinh dang moi dong trong 1 [monitor] la key=value.
// Cac block duoc ngan cach boi dong [monitor].
// ---------------------------------------------------------------------------
struct AppConfig {
    int pollIntervalMs = 100;          // chu ky quet (ms)
    std::vector<PixelMonitor> monitors;
};

static bool loadConfig(const std::string& path, AppConfig& cfg, std::string& err) {
    std::ifstream fin(path.c_str());
    if (!fin) {
        err = "Khong mo duoc file config: " + path;
        return false;
    }

    PixelMonitor cur;
    bool inMonitor = false;

    auto flush = [&]() {
        if (inMonitor) {
            cur.vk = keyNameToVk(cur.keyName);
            cfg.monitors.push_back(cur);
        }
    };

    std::string line;
    while (std::getline(fin, line)) {
        line = trim(line);
        if (line.empty() || line[0] == '#' || line[0] == ';') continue;

        // Bat dau mot block monitor moi
        if (toUpper(line) == "[MONITOR]") {
            flush();
            cur = PixelMonitor();   // reset ve mac dinh
            inMonitor = true;
            continue;
        }
        // Block cai dat chung
        if (toUpper(line) == "[GENERAL]") {
            flush();
            inMonitor = false;
            continue;
        }

        size_t eq = line.find('=');
        if (eq == std::string::npos) continue;
        std::string key = toUpper(trim(line.substr(0, eq)));
        std::string val = trim(line.substr(eq + 1));

        if (!inMonitor) {
            // thuoc [GENERAL]
            if (key == "POLL_INTERVAL_MS") cfg.pollIntervalMs = atoi(val.c_str());
            continue;
        }

        // thuoc [MONITOR]
        if      (key == "NAME")        cur.name = val;
        else if (key == "X")           cur.x = atoi(val.c_str());
        else if (key == "Y")           cur.y = atoi(val.c_str());
        else if (key == "R")           cur.r = atoi(val.c_str());
        else if (key == "G")           cur.g = atoi(val.c_str());
        else if (key == "B")           cur.b = atoi(val.c_str());
        else if (key == "TOLERANCE")   cur.tolerance = atoi(val.c_str());
        else if (key == "KEY")         cur.keyName = val;
        else if (key == "TRIGGER")     cur.trigger = toUpper(val) == "PRESENT" ? "present" : "absent";
        else if (key == "COOLDOWN_MS") cur.cooldownMs = atoi(val.c_str());
    }
    flush();

    if (cfg.monitors.empty()) {
        err = "Config khong co diem theo doi nao ([monitor]).";
        return false;
    }
    if (cfg.pollIntervalMs < 10) cfg.pollIntervalMs = 10;
    return true;
}

// ---------------------------------------------------------------------------
// In ra mau hien tai tai cac diem (de ban do mau de dua vao config)
// ---------------------------------------------------------------------------
static void dumpColors(const AppConfig& cfg) {
    std::cout << "\n--- Mau pixel hien tai ---\n";
    for (const auto& m : cfg.monitors) {
        COLORREF c = getPixelColor(m.x, m.y);
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

    AppConfig cfg;
    std::string err;
    if (!loadConfig("config.ini", cfg, err)) {
        std::cerr << "[LOI] " << err << "\n";
        std::cout << "Nhan Enter de thoat...";
        std::cin.get();
        return 1;
    }

    std::cout << "============================================\n";
    std::cout << " AutoHelper da nap cau hinh thanh cong.\n";
    std::cout << " So diem theo doi : " << cfg.monitors.size() << "\n";
    std::cout << " Chu ky quet      : " << cfg.pollIntervalMs << " ms\n";
    std::cout << "--------------------------------------------\n";
    for (const auto& m : cfg.monitors) {
        std::cout << "  * " << m.name << " @(" << m.x << "," << m.y << ") "
                  << "mau(" << m.r << "," << m.g << "," << m.b << ") "
                  << "sai lech<=" << m.tolerance << " "
                  << "phim=" << m.keyName
                  << (m.vk == 0 ? " [!! phim khong hop le]" : "")
                  << " trigger=" << m.trigger << "\n";
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
        }
        if (f10 && !prevF10) {
            dumpColors(cfg);
        }
        if (f12 && !prevF12) {
            std::cout << " Thoat chuong trinh.\n";
            break;
        }
        prevF9 = f9; prevF10 = f10; prevF12 = f12;

        // --- Neu dang bat, quet cac diem theo doi ---
        if (running) {
            DWORD now = GetTickCount();
            for (auto& m : cfg.monitors) {
                if (m.vk == 0) continue;  // bo qua phim khong hop le

                COLORREF c = getPixelColor(m.x, m.y);
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
        }

        Sleep(cfg.pollIntervalMs);
    }

    return 0;
}
