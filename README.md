# Tra cứu BOM quạt trần

Công cụ tra cứu BOM quạt trần chạy **offline**, gói trong **một file HTML duy nhất** (`BOM_Lookup.html`).
Mở bằng trình duyệt (Chrome/Edge, máy tính hoặc điện thoại), không cần cài gì, không cần mạng.

## Cách dùng
1. Tải `BOM_Lookup.html` về máy (nút **Download raw file** trên GitHub, hoặc clone repo).
2. Mở file bằng trình duyệt.
3. Gõ tên quạt hoặc model ở ô tìm trên cùng, chọn model để xem tổng quan (dây nguồn, mô tơ, tụ điện, tem, miếng đệm…).
4. Sang tab **Linh kiện BOM** để tra từng linh kiện: gõ tiếng Việt (có hoặc không dấu), tiếng Trung hoặc mã liệu.

Phím tắt: `/` để nhảy vào ô tìm; trong tab BOM có thể lọc theo cấp (1, 2, 3…) và theo nhóm linh kiện.

## Cập nhật dữ liệu (người giữ BOM)
- Đặt các file BOM `.xlsx` vào thư mục BOM gốc, ảnh quạt vào `PICTURE/` (tên file = model khách hàng).
- Chạy `tools\CAP_NHAT_DU_LIEU.bat` → file `BOM_Lookup.html` được dựng lại.
- Sửa bản dịch: mở `tools\tu_dien.csv` (UTF-8), sửa cột `vi`, rồi chạy lại file `.bat` ở trên.
  Từ nào chưa dịch được nằm trong `tools\tu_chua_dich.csv`.

## Lưu ý
- Dữ liệu trong file là **BOM nội bộ** (mã liệu, nhà cung cấp). Chỉ chia sẻ trong công ty.
- Bản dịch Trung → Việt là từ điển của xưởng, các từ ghi "cần xác nhận" trong `tu_dien.csv` có thể cần chỉnh.
