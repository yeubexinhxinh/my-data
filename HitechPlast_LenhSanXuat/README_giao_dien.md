# Báo cáo CL NVL — bản giao diện đã làm đẹp

File `report_forms.py` ở đây là bản đã chỉnh **giao diện** của tab
"📑 Báo cáo Chất Lượng Nguyên Vật Liệu" (thay cho file cùng tên trong gói).

## Đã thay đổi (chỉ giao diện, không đổi logic/dữ liệu)
- **Bảng màu & font thống nhất**: thêm `CLR{}` + bộ font (`FONT_*`) ở đầu file,
  để chỉnh 1 chỗ là đổi toàn bộ. Giữ nguyên `BG`/`SEL` nên phần còn lại chạy y cũ.
- **Header công ty**: dải tiêu đề trên cùng có **logo + tên công ty + tên báo cáo**
  (logo lấy từ `logo.png`/icon có sẵn; tên công ty đọc `company_name` trong
  `rmqr_settings.json`, mặc định "HITECH PLAST").
- **Danh sách/cây đẹp hơn**: style riêng `RF.Treeview` (hàng cao thoáng, chữ rõ,
  tiêu đề cột có nền, dòng được chọn màu xanh chính) — đặt tên riêng nên KHÔNG
  ảnh hưởng các tab khác của app.
- **Zebra**: danh sách "Báo cáo đã lưu" xen kẽ màu hàng cho dễ đọc.
- **Thanh công cụ gọn**: nút bo phẳng, khoảng cách đều, có vạch ngăn nhóm; màu
  Lưu/PDF/Excel/In lấy theo bảng màu.

## Cách dùng
Chép `report_forms.py` này **đè** lên file cũ trong thư mục phần mềm
(cạnh `app_hitechplast.py`). Cần: `pip install openpyxl reportlab Pillow`.

Tùy chỉnh tên công ty: mở `rmqr_settings.json`, thêm `"company_name": "TÊN CỦA BẠN"`.

---

## Cập nhật thêm (đợt 2)

### 1) Phần KẾT LUẬN (trên dòng "Người lập báo cáo")
- Thêm nút thiết kế **"🧾 Đặt ô Kết luận"** (Admin, trong "✏ Thiết kế form"):
  chọn ô nằm **phía trên dòng Người lập báo cáo** rồi bấm để đánh dấu ô Kết luận.
- Thêm nút **"⚖ Kết luận"** trên thanh công cụ: tự **so sánh kết quả với tiêu chuẩn**
  (đọc được các dạng `2-3`, `≤0.5`, `≥1`, `min/max`, `5±0.5`) rồi ghi **gợi ý**
  vào ô Kết luận — **người lập tự kiểm tra và sửa lại**. Nếu có chỉ tiêu vượt/thấp
  hơn tiêu chuẩn sẽ liệt kê rõ; nếu không đọc được dạng tiêu chuẩn thì ghi mẫu
  trống để người lập tự ghi.
- Logic so sánh đặt trong `report_forms_io.py` (`concl_range`, `suggest_conclusion`,
  `conclusion_text`) — đã viết test, chạy đúng các trường hợp.
- Ô Kết luận là ô bình thường nên **tự hiển thị trong PDF/Excel** khi xuất.

> Lưu ý: so sánh tự động theo kiểu "cùng một hàng có ô tiêu chuẩn + ô kết quả".
> Với form 2 phần (Tiêu chuẩn ở phần I, kết quả ở phần II) có thể không tự ghép
> được — khi đó nút sẽ ghi mẫu trống để người lập tự kết luận. Nếu bạn muốn so
> sánh đúng theo layout riêng, gửi tôi bố cục cột để tôi chỉnh quy tắc.

### 2) Ô dữ liệu rộng & giãn chữ (chỉ khi xem trên app)
- Giãn cột/hàng khi hiển thị: thêm `grid_col_scale` (mặc định 1.18) và
  `grid_row_scale` (1.15) trong `rmqr_settings.json` — ô dữ liệu rộng rãi hơn,
  dễ nhập. Tăng đệm chữ trái/phải trong ô cho thoáng.
- **Không ảnh hưởng PDF/Excel xuất ra** (giữ kích thước gốc). Muốn file xuất ra
  cũng rộng hơn thì báo tôi.

---

## Cập nhật thêm (đợt 3)

### Kết luận: full-width, đúng chỗ, TỰ ĐỘNG ghi
- Ô Kết luận giờ **gộp full chiều ngang** và đặt ở **hàng trống ngay trên dòng
  "Người lập báo cáo"** (đã dò đúng trên cả 147 form).
- **Tự động ghi kết luận khi nhập dữ liệu**: mỗi khi nhập 1 chỉ tiêu SỐ mà có
  tiêu chuẩn tương ứng, app tự so sánh và cập nhật ô Kết luận (ĐẠT / liệt kê chỉ
  tiêu vượt–thấp hơn). **Không đè** nếu người lập đã tự sửa nội dung kết luận.
- So sánh đã sửa cho **đúng cấu trúc 2 phần** (Tiêu chuẩn ở phần I, kết quả ở
  phần II): chỉ đánh giá ô **người dùng thực nhập**, tìm tiêu chuẩn theo **cùng
  cột (phía trên)** hoặc **cùng hàng** → hết tình trạng báo nhầm khi chưa nhập.
- Admin vẫn có thể tự đặt: "✏ Thiết kế form" → chọn hàng trống trên chữ ký →
  "🧾 Đặt ô Kết luận" (tự gộp full-width).

### Chuyển ảnh bị add nhầm ô
- Chuột phải trong ô → **"✂ Cắt ảnh trong ô (để chuyển)"**, rồi chọn ô đúng →
  **"📌 Dán ảnh đã cắt vào ô này"**. Ảnh chuyển sang đúng ô, vừa vùng đã chọn.

### Nút Copy / Paste ô khi sửa form (Admin)
- Trong "✏ Thiết kế form", thanh cam có thêm **"⧉ Copy ô"** và **"📋 Paste ô"**:
  copy toàn bộ định dạng + nội dung mẫu + danh sách chọn của ô, dán sang ô/vùng
  khác (áp cho cả vùng đang chọn) → dựng form nhanh hơn nhiều.

---

## Cập nhật thêm (đợt 4) — SỔ THEO DÕI (sheet tổng hợp)

Thêm nút **"📋 Sổ theo dõi"** trên thanh công cụ. Mở ra bảng tổng hợp **mọi báo
cáo đã lưu**, mỗi dòng gồm:

| Cột | Nguồn dữ liệu |
|-----|----------------|
| Tên nguyên vật liệu | ô Tên NVL (name_cell) |
| Nhà cung cấp | dò nhãn "Nhà cung cấp" → lấy ô bên phải |
| Địa chỉ | dò nhãn "Địa chỉ / Address" |
| Ngày nhận mẫu | dò nhãn "Ngày nhận mẫu" |
| Ngày trả KQ | ô Ngày trả KQ (date_cell) |
| Đánh giá kết quả | ĐẠT / KHÔNG ĐẠT (tự so sánh) — tô xanh/đỏ |
| Tiêu chuẩn gốc | tiêu chuẩn quy định của từng chỉ tiêu (vd `MFI: 2–3`) |
| Kết quả mẫu | giá trị đo được người lập nhập (vd `MFI: 3.5`) |
| Người lập | người lập báo cáo |

- Nút **"📊 Xuất Excel"** trong cửa sổ: xuất ra **1 sheet Excel** ("Sổ theo dõi
  NVL") đúng các cột trên, có tô màu Đạt/Không đạt, in/nộp được ngay.
- Logic trích xuất + xuất Excel nằm trong `report_forms_io.py`
  (`extract_fields`, `compare_all`, `report_summary`, `export_log_excel`) — đã
  test với form thật, xuất Excel đọc lại đúng.

> "Tiêu chuẩn gốc" = tiêu chuẩn quy định (phần I). "Kết quả mẫu" = giá trị người
> lập nhập (phần II). Nếu bạn muốn hiểu "tiêu chuẩn của mẫu" theo nghĩa khác
> (vd tiêu chuẩn do NCC công bố), báo tôi để chỉnh cột cho đúng.

---

## Cập nhật thêm (đợt 5) — Ảnh: Copy / Chèn / Fit vừa ô (cho người dùng thường)

Tất cả thao tác ảnh dưới đây **người dùng thường đều làm được** (không cần Admin):

- **🖼 Chèn ảnh** (nút sẵn có) + **📋 Dán ảnh** (Ctrl+V): chèn ảnh vào ô đang chọn.
- **⧉ Copy ảnh** (chuột phải): chép ảnh trong ô (giữ ảnh cũ) → chọn ô khác →
  **📌 Dán ảnh**. (Khác với **✂ Cắt ảnh** = chuyển ảnh đi.)
- **↔ Fit ảnh** (nút trên thanh công cụ + chuột phải "Fit ảnh vừa ô khung"):
  ảnh **phủ kín ô khung, độ rộng = ô**. Bấm **"⤢ Ảnh giữ tỉ lệ"** để trả về mặc định.
- Chế độ fit được **lưu theo ảnh** và áp dụng đồng nhất cả **trên màn hình, PDF và
  Excel** (thêm `fit_box2` + cờ `fit` cho mỗi ảnh).

---

## Cập nhật thêm (đợt 6)

### 1) Admin fit được ảnh MẪU
Trong "✏ Thiết kế form": thanh cam có nút **"↔ Fit ảnh mẫu"**, và chuột phải có
**"↔ Fit ảnh mẫu vừa ô khung"** / **"⤢ Ảnh mẫu giữ tỉ lệ"**. Admin chỉnh ảnh mẫu
(ảnh cố định của form) phủ kín ô hoặc giữ tỉ lệ — áp dụng cả màn hình/PDF/Excel.

### 2) Thêm ô nhập tiêu chuẩn Độ trắng / Độ sáng (6 form Chất Độn Khác)
File `rmqr_forms.json` đã cập nhật cho 6 form: Titan Dioxit, Cao Lanh, Bari
Sunfat, Natri Sunfat, Talc, Bột Đá NaNo. Ở **phần I (Tiêu chuẩn)**:
- Nhãn đổi sang tiếng Việt: **"Độ trắng (Whiteness)"**, **"Độ sáng (Brightness)"**.
- Ô nhập tiêu chuẩn bên cạnh (cột E) được tô **nền vàng nhạt** để thấy rõ chỗ nhập.

Chép `rmqr_forms.json` này đè lên file cùng tên trong thư mục phần mềm. Lưu ý: chỉ
áp cho **báo cáo tạo MỚI**; báo cáo đã lưu trước đó giữ mẫu cũ. Nếu đã từng sửa
form (có `BaoCaoNVL/rmqr_forms.json`) thì vào ⚙ Form → "↺ Khôi phục bộ form gốc"
hoặc sửa tay theo mẫu mới.

---

## Cập nhật thêm (đợt 7) — Nút Xem trước (ai cũng dùng)

Thêm nút **"👁 Xem trước"** trên thanh công cụ (cạnh PDF/Excel). Bấm để xem báo
cáo **đúng như khi xuất ra** — tạo PDF A4 (có giấy tiêu đề, ô gộp, viền, ảnh,
phân trang) vào thư mục tạm rồi mở bằng trình xem PDF của máy, **không** ghi vào
thư mục báo cáo. Dùng được cho **mọi tài khoản** (không cần quyền admin), và xem
được cả khi đang ở chế độ thiết kế form.

> Báo cáo xuất ra ở dạng **PDF (A4)** và **Excel (.xlsx)**. Nút Xem trước hiển thị
> đúng bản PDF sẽ xuất. (Cần `pip install reportlab pillow`.)

---

## Cập nhật thêm (đợt 8) — Thêm cột Độ trắng / Độ sáng vào sheet "Đơn giá vật tư"

"Đơn giá vật tư" là **sheet dữ liệu** (16 cột, ~389 dòng), lưu trong `sheets.json`
(đã **mã hoá**) hoặc trên **máy chủ** — không nằm trong mã Python. Nên 2 cột mới
phải thêm vào DỮ LIỆU, không phải sửa code.

Hai cột **"Độ trắng"** và **"Độ sáng"** được chèn **ngay sau cột "Cỡ hạt µm"**
(các cột % Tro, Tạp chất... dịch sang phải).

**Cách 1 — chạy máy chủ/admin (khuyên dùng, hợp mọi chế độ):**
  1. Chép `them_cot_do_trang_do_sang.py` cạnh `app_hitechplast.py`.
  2. **Sao lưu** dữ liệu trước.
  3. Chạy: `python them_cot_do_trang_do_sang.py` (tự nhận máy chủ/offline/cục bộ,
     tự mã hoá lại, chạy lại nhiều lần vẫn an toàn).

**Cách 2 — chỉ máy chạy LƯU CỤC BỘ (không máy chủ):**
  Chép đè `sheets_da_them_cot_DoTrang_DoSang.json` thành `sheets.json` trong thư
  mục phần mềm (file này đã mã hoá sẵn, đã chèn 2 cột). KHÔNG dùng cách này nếu
  đang chạy chế độ máy chủ (sẽ bị dữ liệu máy chủ ghi đè).

> Hiện 2 cột để **nhập dữ liệu** (giá trị độ trắng/sáng của vật tư). Nếu muốn
> phần mềm tự dùng 2 cột này làm **tiêu chuẩn tự đối chiếu** trong báo cáo QC
> (như MFI/% Tro), báo tôi để thêm vào bộ nhận diện chỉ tiêu (RM_KEYS/BB_UP_COLS
> trong qc_app.py).

---

## Cập nhật thêm (đợt 9) — Xuất tem theo khổ giấy A3/A4/A5 (A5 = 2 tem + đường cắt)

Thêm nút **"🖨 Khổ giấy A3/A4/A5"** ở cả 2 tab của "Làm Tem Nhanh":
"📋 Điền tem theo mẫu" và "🎨 Tạo mẫu tem mới từ ảnh".

Bấm nút → hộp thoại chọn **khổ giấy A3 / A4 / A5** (chuẩn ISO), định dạng **PDF**
(in trực tiếp) hoặc **PNG**:
- **A3 / A4**: 1 tem/trang, căn giữa, phóng vừa khổ (giữ tỉ lệ, tự xoay nếu vừa hơn).
- **A5** (mặc định bật **"2 tem/trang"**): xếp **2 tem** trên 1 trang và **kẻ 1
  đường (nét đứt) ở giữa** để cắt → ra 2 tem bằng nhau. Có thể bật/tắt 2 tem và
  đường cắt tuỳ ý cho mọi khổ.

File mới: `tem_khogiay.py` (lõi ghép khổ giấy, không phụ thuộc giao diện — đã test
xuất PDF A3/A4/A5 và kiểm tra A5 có đủ 2 tem + đường giữa). Chép file này cùng
`tem_ui.py` vào thư mục phần mềm. Cần `pip install pillow`.

---

## Cập nhật thêm (đợt 10) — Tách 2 mục tạo mẫu tem: Bao / Pallet

Trong "Làm Tem Nhanh" giờ có 3 tab:
- **📋 Điền tem theo mẫu (.btw)** (như cũ)
- **🎨 Tạo Mẫu Tem Bao Từ Ảnh** (đổi tên từ "Tạo mẫu tem mới từ ảnh") — thư viện
  mẫu lưu ở `tem_designs.json`.
- **📦 Tạo Mẫu Tem Pallet Từ Ảnh** (MỚI) — giống hệt cách làm, nhưng **thư viện
  mẫu riêng** ở `tem_designs_pallet.json` để mẫu tem bao và tem pallet không lẫn.

Cả 2 tab tạo mẫu đều có đủ: nhận diện khung từ ảnh, chọn kích thước, Xem trước,
Xuất .btw, Lưu PNG và **🖨 Khổ giấy A3/A4/A5** (A5 = 2 tem + đường cắt giữa).

Chép đè `tem_ui.py`. File `tem_designs_pallet.json` sẽ tự tạo khi bạn lưu mẫu
pallet đầu tiên.

---

## Cập nhật thêm (đợt 11)

### A) Tab "Tạo Mẫu Tem Pallet Từ Ảnh" dùng CHỌN KHỔ GIẤY (không copy y tem bao)
- Bỏ danh sách preset khổ tem bao (60x90, 60x50...); thay bằng **"Khổ giấy:" A5 / A4 / A3**
  (chuẩn ISO). Chọn khổ nào thì Rộng×Cao tự đặt theo khổ đó.
- Mẫu pallet **mặc định khổ A4**; cho phép tới 500mm (A3 = 297×420 vừa đủ).
- Nhãn hướng dẫn đổi thành "…chọn KHỔ GIẤY (A3/A4/A5)…".
- Tab "Tạo Mẫu Tem Bao Từ Ảnh" giữ nguyên preset tem bao như cũ.

### B) Cửa sổ "Ảnh theo mốc / DC" không còn "trống trơn" khó hiểu
Khi mở mà bạn chỉ có lịch **"Chọn DC/mốc"** (không có lịch "Mẫu mới nhất từng dây
chuyền"), cửa sổ nay **ghi rõ**: cửa sổ này chỉ dùng cho lịch "Mẫu mới nhất từng
dây chuyền"; còn lịch "Chọn DC/mốc" thì đặt ảnh bằng **Sửa → 🎯 Chọn DC/mốc… →
bấm cột 🎞/⚪ của từng mốc**. (Đây là đúng thiết kế, không phải lỗi — chỉ thiếu
hướng dẫn nên nhìn như trống.)

---

## Cập nhật thêm (đợt 12) — Lịch "Chọn DC/mốc" thêm được ẢNH THEO TỪNG MỐC

Trước đây cửa sổ "🖼 Ảnh theo mốc / DC" chỉ nhận lịch "Mẫu mới nhất từng dây
chuyền". Nay nhỗ trợ cả lịch **"Chọn DC / mốc"** và **"Mẫu đang chọn"**:

- Mở "🖼 Ảnh theo mốc / DC" → bảng hiện các lịch pick/selected; mỗi ô
  (mốc × dây chuyền) có nút **🎞 (màng film)** / **⚪ (hạt Filler)** để thêm ảnh
  RIÊNG cho mốc đó.
- Khi gửi: mốc nào ĐÃ thêm ảnh riêng → dùng ảnh đó; mốc nào CHƯA → **tự dùng ảnh
  của mẫu** (không để trống). Không làm thay đổi cách chạy của các lịch cũ.

Thay đổi code (zalo_send.py):
- Thêm `job_can_slot()` (latest + pick + selected) cho bảng ảnh và nút mở.
- `_items` truyền `job` vào `_cards` cho pick/selected; `_cards` áp `slot_info`
  (fallback ảnh mẫu, ép `slot_img_only=False` để không bao giờ để trống).
- Nhãn hàng & thông báo trống cập nhật cho đúng.

Chép đè `zalo_send.py`.

---

## Cập nhật thêm (đợt 13) — Xem trước tem: thêm Xuất PDF + ×2 tem A5 (kẻ giữa)

Cửa sổ **"👁 Xem trước tem"** nay có thêm:
- **📄 Xuất PDF**: xuất 1 tem ra PDF **đúng kích thước tem** (in ra đúng khổ), mở
  luôn sau khi xuất.
- **🖨 ×2 tem A5 (kẻ giữa)**: CHỈ hiện khi kích thước tem là **A5** (148×210, kể cả
  xoay ngang) — xuất 1 trang A5 có **2 tem + đường cắt nét đứt ở giữa**.

Dùng chung `tem_khogiay.py` (thêm `save_label_pdf`, `is_a5`). Chép đè `tem_ui.py`
và `tem_khogiay.py`.

---

## Cập nhật thêm (đợt 14) — Xem trước tem: thêm Xuất Excel + Xuất Word

Cửa sổ "👁 Xem trước tem" nay có đủ: **📄 Xuất PDF · 📊 Xuất Excel · 📝 Xuất Word**
(và 🖨 ×2 tem A5 khi khổ A5). Ảnh tem được nhúng ĐÚNG kích thước tem.

- Excel (.xlsx): dùng openpyxl (có sẵn).
- Word (.docx): dùng python-docx — **nếu chưa cài** sẽ báo "pip install python-docx".

Hàm trong `tem_khogiay.py`: `save_label_excel`, `save_label_word`.
Chép đè `tem_ui.py` và `tem_khogiay.py`. Word cần: `pip install python-docx`.

---

## Cập nhật thêm (đợt 15)

### 1) Thêm "Lưu vào" (thư mục) cho CẢ 2 tab tạo mẫu (Bao & Pallet)
Mỗi tab có hàng **"Lưu vào: [đường dẫn] [...] [📂 Mở thư mục]"** riêng (lưu trong
`tem_design_cfg.json` theo từng loại). Các nút Xuất .btw / PNG / Khổ giấy dùng
đúng thư mục này. Để trống thì về thư mục mặc định như cũ.

### 2) Vì sao xuất tem Pallet ra .btw không được — và cách khắc phục
Nguyên nhân: file .btw được dựng bằng cách ĐIỀN NỘI DUNG vào **phôi có sẵn**
(`tem_mau_AA06.btw` = tem BAO, chỉ có **17 ô chữ, 11 đường kẻ**). Tem pallet
thường 9 hàng × 2 cột = **18 ô chữ** → vượt số ô của phôi → báo lỗi "phôi không
đủ đối tượng", không xuất được.

Khắc phục (đã thêm):
- Mỗi tab có nút **"Chọn phôi .btw…"** — với tab Pallet, hãy chọn **phôi tem
  pallet có đủ số dòng** (vd `TEM PALLET PA.btw`). Khi đó xuất .btw chạy đúng.
- Nếu vẫn dùng phôi mặc định mà tem quá nhiều dòng, thông báo lỗi nay **nói rõ**
  cách xử lý (chọn phôi pallet, hoặc dùng 📄 PDF / 📊 Excel / 📝 Word — không giới
  hạn số dòng).

Chép đè `tem_ui.py`.

---

## Cập nhật thêm (đợt 16) — Xuất .btw vượt giới hạn phôi → tự đề nghị xuất PDF

Vì sao tem pallet (19 dòng) báo lỗi khi Xuất .btw: file .btw tạo ra bằng cách
ĐIỀN vào phôi có sẵn; BarTender .btw **không thể tự thêm ô chữ/đối tượng** qua
công cụ này, nên số dòng tối đa = số ô chữ của phôi (phôi mặc định tem bao = 17 ô).
Tem pallet 9 hàng × 2 cột = 18–19 ô → vượt.

Nay khi bấm **Xuất .btw** mà vượt giới hạn phôi, phần mềm **hỏi và tự chuyển sang
Xuất PDF** (A3/A4/A5, không giới hạn số dòng) để luôn ra được file. Vẫn giữ nút
"Chọn phôi .btw…" nếu bạn có phôi tem pallet đủ ô.

Chép đè `tem_ui.py`.
