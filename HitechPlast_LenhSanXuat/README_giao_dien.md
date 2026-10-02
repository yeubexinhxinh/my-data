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
