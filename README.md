# Industry Survey – Live Dashboard

Dashboard HTML tĩnh cho các câu A7, B1, B4, C5, D4, E4.

- Dữ liệu đọc trực tiếp từ Google Sheet `RMIT_Industry Survey`:
  - `Dash_Data` (ưu tiên): mỗi phiếu 1 dòng JSON gọn, do script sync tạo ra – nhẹ, tải nhanh.
  - `Raw_Data` (dự phòng): XML SurveyToGo chia chunk; chỉ dùng khi chưa có `Dash_Data`.
  - `Sync_Log`: lịch sử sync.
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

## Script sync (Apps Script)
Phần đọc XML (`stgSplitSubjects_`, `stgExtractSubject_`…) nằm ở `apps-script/extract_core.js` và được chép y hệt vào cuối `syncSTGSurvey.gs` và trong `index.html`. Sửa ở một chỗ thì chép lại sang hai chỗ kia.

`apps-script/syncSTGSurvey.gs` – tải XML từ SurveyToGo API, chia chunk 45.000 ký tự và ghi vào `Raw_Data` (mỗi chunk có dấu `'` ở đầu để Sheets không hiểu nhầm thành công thức). API key/password lưu trong Script Properties (`STG_API_KEY`, `STG_PASSWORD`), không nằm trong repo.


## Sheets written by the script
- `Dash_Data` – one JSON row per interview, read by this dashboard.
- `QC_Data` – one compact JSON row per interview (`{m,a,times}`), read by the QC page. Written by `syncSTGSurvey` (5-minute trigger).
- `Raw_Data` – full XML chunks. Only written by `syncSTGRaw` (manual/slow); the QC page uses it only as a fallback.
