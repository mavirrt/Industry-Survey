# Industry Survey – Live Dashboard

Dashboard HTML tĩnh cho các câu A7, B1, B4, C5, D4, E4.

- Dữ liệu đọc trực tiếp từ Google Sheet `RMIT_Industry Survey` (sheet `Raw_Data`: XML SurveyToGo được sync và chia chunk; `Sync_Log`: lịch sử sync).
- Trình duyệt tự tải lại dữ liệu mỗi 5 phút; nút **↻ Làm mới** để tải ngay.
- Chỉ phiếu `Completed = True` được tính, gồm cả phiếu đang chờ duyệt.
- Không lưu dữ liệu đáp viên trong repo.

## Chạy
Mở `index.html` qua GitHub Pages (Settings → Pages → Deploy from branch `main`, thư mục `/`). Mở file trực tiếp từ máy cũng chạy được nếu trình duyệt cho phép fetch.

## Cấu hình
Trong `index.html`:
- `SHEET_ID`, `GID_RAW`, `GID_LOG` – nguồn dữ liệu
- `REFRESH_MS` – chu kỳ cập nhật (mặc định 5 phút)
- `Q`, `B4_LBL`, `C5_LBL`, `D4_SCALE`, `CARD`, `E4_LBL`, `RF_VI` – tên câu và nhãn hiển thị
