# AutoHelper

Tool cho Windows: **theo dõi màu pixel tại tọa độ** (ví dụ thanh HP/MP) và
**tự động nhấn phím** (dùng thuốc) khi màu cho thấy máu/mana đã tụt dưới ngưỡng.

Toàn bộ code do bạn sở hữu — chỉ dùng WinAPI chuẩn, không phụ thuộc thư viện ngoài.

Có **2 bản, cùng tính năng, cùng đọc một file `config.ini`**:

- **Bản C++** (`main.cpp`) — build ra `.exe`, chạy nhanh. Xem mục *Build* bên dưới.
- **Bản Python** (`autohelper.py`) — chạy ngay không cần build, tiện để test nhanh.

## Chạy bản Python (nhanh để test)

Cần **Python 3** trên Windows (không cần cài thêm thư viện — dùng `ctypes` gọi
WinAPI trực tiếp). Đặt `autohelper.py` cạnh `config.ini` rồi chạy:

```bat
python autohelper.py
```

> Nên chạy Command Prompt **với quyền Administrator** để phím gửi vào được cửa sổ
> game. Phím điều khiển giống hệt bản C++: **F9** bật/tắt, **F10** dò màu,
> **F12** thoát. Tool này chỉ chạy trên Windows.

## Cách hoạt động

1. Bạn chọn 1 điểm tọa độ nằm trên thanh HP (lúc máu còn "đủ" thì điểm đó có
   màu đỏ máu).
2. Tool đọc màu pixel tại điểm đó liên tục.
3. Khi thanh máu tụt qua điểm đó, pixel đổi sang màu nền → tool phát hiện
   "không còn khớp màu máu" (`trigger = absent`) → nhấn phím thuốc bạn đã gán.

Ngược lại, `trigger = present` sẽ nhấn phím khi màu **xuất hiện** (ví dụ chờ một
icon/thông báo hiện ra).

### Phím lặp theo chu kỳ (auto đánh / auto kỹ năng)

Ngoài phần theo màu, tool còn có khối `[timer]`: **tự nhấn một phím lặp lại theo
chu kỳ** — chính là kiểu "auto đánh thường" hoặc "tung kỹ năng" mà bot hay có.
Mỗi `[timer]` có thể thêm `jitter_ms` để dao động chu kỳ ngẫu nhiên cho nhịp bấm
tự nhiên hơn.

```ini
[timer]
name = Attack
key = 1
interval_ms = 600     # nhấn phím "1" mỗi ~0.6 giây
jitter_ms = 150       # dao động ngẫu nhiên +/- 0.15 giây
```

Bạn có thể khai báo nhiều `[timer]` cùng lúc (mỗi skill một timer) và chúng chạy
song song với các `[monitor]` HP/MP. Tất cả đều bật/tắt chung bằng **F9**.

### Chuỗi combo theo thứ tự (1→2→3→4)

Khối `[combo]` nhấn **một chuỗi phím lần lượt theo thứ tự**, thay vì mỗi phím một
timer riêng. Sau khi hết chuỗi, nó nghỉ `loop_delay_ms` rồi lặp lại từ đầu.

```ini
[combo]
name = Combo1
keys = 1,2,3,4        # nhấn lần lượt 1 -> 2 -> 3 -> 4
step_delay_ms = 300   # cách nhau 0.3s giữa mỗi phím
loop_delay_ms = 1500  # xong chuỗi, nghỉ 1.5s rồi lặp lại
jitter_ms = 80        # dao động ngẫu nhiên +/- cho các khoảng chờ
```

`keys` nhận cùng danh sách tên phím như trên (`Q,W,E,R,SPACE`, `F1,F2,...`,
`NUM1,...`). Có thể khai báo nhiều `[combo]`; chúng chạy song song với `[timer]`
và `[monitor]`, cùng bật/tắt bằng **F9**.

### Bám theo cửa sổ game (chống lệch khi di chuyển cửa sổ)

Nếu bạn đặt `window_title` trong `[general]`, tool sẽ **tìm cửa sổ game theo tiêu
đề** và hiểu `x`, `y` của mỗi `[monitor]` là **tọa độ tương đối so với góc
trên-trái vùng client của cửa sổ**. Khi bạn kéo/di chuyển cửa sổ, tool tự tính
lại vị trí thật nên **không bị lệch**. Nếu cửa sổ tạm biến mất (minimize/đóng),
tool tự tìm lại khi cửa sổ xuất hiện trở lại.

Để trống `window_title` → quay về chế độ tọa độ màn hình tuyệt đối như cũ.

Hai chế độ đọc màu (`capture_mode`):

| Giá trị | Mô tả |
|---------|-------|
| `screen` | Tính tọa độ client → đổi sang tọa độ màn hình rồi đọc. **Cần cửa sổ hiện ra, không bị che.** (mặc định, hợp nhất) |
| `window` | Đọc trực tiếp từ DC của cửa sổ. Chạy được cả khi cửa sổ không focus, nhưng một số game DirectX có thể trả về màu đen — khi đó hãy dùng `screen`. |

> Mẹo lấy tọa độ tương đối cửa sổ: đặt `window_title`, chạy tool, rồi dùng `F10`
> để in màu. Nếu màu không đúng điểm mong muốn, chỉnh `x`, `y` (tính từ mép trong
> trái-trên của cửa sổ game, không phải mép màn hình).

## Build

Cần **Visual Studio Build Tools** (có `cl`) hoặc **MinGW-w64** (có `g++`).

```bat
build.bat
```

Hoặc build thủ công:

```bat
REM MSVC
cl /EHsc /O2 /Fe:AutoHelper.exe main.cpp user32.lib gdi32.lib

REM MinGW
g++ -O2 -std=c++11 main.cpp -o AutoHelper.exe -lgdi32 -luser32 -static
```

## Chạy

Đặt `AutoHelper.exe` cạnh `config.ini` rồi chạy `AutoHelper.exe`.

Phím điều khiển khi chạy:

| Phím | Chức năng |
|------|-----------|
| `F9`  | Bật / tắt chế độ tự động |
| `F10` | In ra màu pixel hiện tại tại các điểm theo dõi (để dò màu) |
| `F12` | Thoát |

## Cách lấy tọa độ và màu

1. Build & chạy tool.
2. Mở game, canh con trỏ/quan sát điểm trên thanh HP lúc máu còn ~50%.
3. Điền `x`, `y` vào `config.ini`.
4. Nhấn `F10` trong tool để in màu R/G/B hiện tại tại điểm đó → điền vào
   `r`, `g`, `b`.
5. Chỉnh `tolerance` (độ sai lệch cho phép). Màu nhấp nháy/gradient thì tăng lên.

## Cấu hình (`config.ini`)

Mỗi `[monitor]` là một điểm theo dõi:

| Khóa | Ý nghĩa |
|------|---------|
| `name` | Tên hiển thị trong log |
| `x`, `y` | Tọa độ pixel trên màn hình |
| `r`, `g`, `b` | Màu tham chiếu (màu "bình thường") |
| `tolerance` | Sai lệch màu cho phép (0–441) |
| `key` | Phím sẽ nhấn: `A`–`Z`, `0`–`9`, `F1`–`F12`, `NUM0`–`NUM9`, `SPACE`, `ENTER`, `TAB`, `ESC`, `SHIFT`, `CTRL`, `ALT` |
| `trigger` | `absent` = nhấn khi màu **mất** (HP/MP tụt); `present` = nhấn khi màu **xuất hiện** |
| `cooldown_ms` | Thời gian chờ tối thiểu giữa 2 lần nhấn (ms) |

`[general]`:

| Khóa | Ý nghĩa |
|------|---------|
| `poll_interval_ms` | Chu kỳ quét (ms). Nhỏ hơn = nhạy hơn, tốn CPU hơn |
| `window_title` | Chuỗi con trong tiêu đề cửa sổ game. Có giá trị → `x,y` là tọa độ tương đối cửa sổ. Để trống → tọa độ màn hình tuyệt đối |
| `capture_mode` | `screen` (mặc định) hoặc `window` — xem bảng ở trên |

`[timer]` (phím lặp theo chu kỳ):

| Khóa | Ý nghĩa |
|------|---------|
| `name` | Tên hiển thị trong log |
| `key` | Phím sẽ nhấn (cùng danh sách phím như trên) |
| `interval_ms` | Chu kỳ nhấn (ms) |
| `jitter_ms` | Dao động ngẫu nhiên +/- quanh chu kỳ (ms). 0 = đều tăm tắp |

`[combo]` (chuỗi phím theo thứ tự):

| Khóa | Ý nghĩa |
|------|---------|
| `name` | Tên hiển thị trong log |
| `keys` | Danh sách phím cách nhau bằng dấu phẩy, nhấn theo thứ tự (vd `1,2,3,4`) |
| `step_delay_ms` | Thời gian chờ giữa mỗi phím trong chuỗi (ms) |
| `loop_delay_ms` | Thời gian chờ sau khi xong chuỗi, trước khi lặp lại (ms) |
| `jitter_ms` | Dao động ngẫu nhiên +/- cho các khoảng chờ (ms) |

## Lưu ý

- Đây là công cụ mô phỏng nhấn phím ở mức hệ điều hành (`SendInput`). Hãy tự chịu
  trách nhiệm khi dùng và tôn trọng điều khoản của phần mềm/dịch vụ bạn tương tác.
- `GetPixel` đọc theo tọa độ màn hình thật; nếu bạn để DPI scaling ≠ 100% thì tọa
  độ có thể lệch — nên dò lại màu bằng `F10` sau khi đổi độ phân giải.
