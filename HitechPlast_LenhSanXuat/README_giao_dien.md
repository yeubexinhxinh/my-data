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
