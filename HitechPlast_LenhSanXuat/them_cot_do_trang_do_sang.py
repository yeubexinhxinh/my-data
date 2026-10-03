# -*- coding: utf-8 -*-
"""
them_cot_do_trang_do_sang.py
Thêm 2 cột "Độ trắng" và "Độ sáng" vào sheet "Đơn giá vật tư",
ngay SAU cột "Cỡ hạt µm". Mỗi dòng được chèn 1 ô trống tương ứng.

CÁCH DÙNG:
  - Đặt file này CẠNH app_hitechplast.py (cùng thư mục phần mềm).
  - Trên MÁY CHỦ (hoặc máy admin đang kết nối máy chủ), chạy:
        python them_cot_do_trang_do_sang.py
  - Script tự nhận chế độ (máy chủ / offline / lưu cục bộ) và tự mã hoá lại.
  - Chạy lại nhiều lần cũng an toàn: nếu đã có cột thì bỏ qua (idempotent).

LƯU Ý: nên SAO LƯU sheets.json (hoặc dữ liệu máy chủ) trước khi chạy.
"""
import sys

# Bật mã hoá file dữ liệu giống phần mềm (để đọc/ghi sheets.json mã hoá)
try:
    import file_crypt
    file_crypt.install(migrate=False)
except Exception:
    pass

import datastore

SHEET = 'Đơn giá vật tư'
AFTER = 'Cỡ hạt µm'
NEW_COLS = ['Độ trắng', 'Độ sáng']


def main():
    sheets = datastore.load_sheets()
    up = next((s for s in sheets if s.get('title') == SHEET), None)
    if not up:
        print('[LỖI] Không tìm thấy sheet:', SHEET)
        return 1

    hd = up.setdefault('headers', [])
    to_add = [c for c in NEW_COLS if c not in hd]
    if not to_add:
        print('Sheet đã có đủ cột Độ trắng / Độ sáng — không cần thêm.')
        return 0
    if AFTER not in hd:
        print('[LỖI] Không thấy cột "%s" trong sheet.' % AFTER)
        return 1

    at = hd.index(AFTER) + 1
    rows = up.get('rows') or []
    for off, col in enumerate(to_add):
        pos = at + off
        hd.insert(pos, col)
        for r in rows:
            while len(r) < pos:          # đệm nếu dòng ngắn
                r.append('')
            r.insert(pos, '')            # ô trống cho cột mới

    datastore.save_sheets(sheets, force_titles=(SHEET,))
    print('Đã thêm cột %s vào "%s" (sau cột "%s").' % (to_add, SHEET, AFTER))
    print('Tổng số cột hiện tại:', len(hd))
    return 0


if __name__ == '__main__':
    sys.exit(main())
