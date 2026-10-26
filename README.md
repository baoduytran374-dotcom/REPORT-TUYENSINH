# BETMA ENGLISH — Dashboard tuyển sinh & tài chính 2026

Website tĩnh (HTML + JS), giao diện và tính năng giữ nguyên bản dashboard gốc, **không cần build, không cần API, không cần Cloudflare Worker, không cần biến môi trường** (không có `VITE_API_URL`). Dữ liệu nằm trong `data/data.json` ngay trong repo GitHub.

## Cấu trúc

```
index.html          Trang chính
config.js           Cấu hình (mặc định để trống, không cần sửa)
assets/style.css    Giao diện (CSS gốc + vài dòng cho hộp thoại cập nhật)
assets/app.js       Dashboard (mã gốc, đã bổ sung nạp dữ liệu động)
assets/parser.js    Đọc file Excel → dữ liệu dashboard
assets/updater.js   Tải Excel, xem trước, xác nhận, đăng lên GitHub
data/data.json      DỮ LIỆU ĐANG HIỂN THỊ (file này được cập nhật khi bạn đăng Excel mới)
data/data.js        Dữ liệu dự phòng (chỉ dùng khi mở file trực tiếp/không đọc được data.json)
vendor/             (tùy chọn) đặt chart.umd.min.js và xlsx.full.min.js để chạy offline
.nojekyll           Báo GitHub Pages phục vụ file nguyên trạng
```

Thư viện Chart.js và SheetJS được nạp từ CDN (cdnjs → jsdelivr → `vendor/`), giống bản dashboard gốc.

## A. Publish lên GitHub Pages (làm một lần, ~5 phút)

1. Đăng nhập GitHub → **New repository** → đặt tên (ví dụ `betma-dashboard`) → **Create repository**.
2. Giải nén file ZIP. Trong repo vừa tạo bấm **Add file → Upload files**, kéo **toàn bộ nội dung bên trong thư mục `betma-dashboard`** (gồm `index.html`, `assets`, `data`, `vendor`, `config.js`, `.nojekyll`…) vào, rồi **Commit changes**.
   *Lưu ý: `index.html` phải nằm ở gốc repo, không nằm trong thư mục con.* Nếu GitHub không cho kéo thư mục ẩn `.nojekyll` cũng không sao.
3. Vào **Settings → Pages**. Mục *Build and deployment* → *Source*: **Deploy from a branch** → Branch: **main** / thư mục **/ (root)** → **Save**.
4. Đợi 1–2 phút, trang hiện địa chỉ `https://<tên-tài-khoản>.github.io/<tên-repo>/`. Mở địa chỉ đó là thấy dashboard.

### Lưu ý quan trọng về quyền riêng tư
GitHub Pages là **website công khai**: ai có đường link đều xem được, và `data/data.json` chứa **họ tên, học phí, thanh toán** của học viên (bản dashboard gốc cũng hiển thị các thông tin này; không có số điện thoại/CCCD). Nếu cần hạn chế:
- không chia sẻ link rộng rãi; hoặc
- đổi cột "Họ và tên" thành mã/tên viết tắt trước khi tải Excel lên; hoặc
- dùng gói GitHub có Pages riêng tư (Enterprise Cloud) / hosting có đăng nhập.
Dữ liệu đã commit vẫn nằm trong lịch sử Git của repo.

## B. Cập nhật dữ liệu bằng Excel

### Chuẩn bị file
1. Mở dashboard → bấm **Cập nhật dữ liệu** → bấm **file mẫu** để tải `BETMA_mau_du_lieu.xlsx` (đã chứa dữ liệu hiện tại; có sheet *Hướng dẫn*).
2. Sửa/thêm học viên trong sheet **TỔNG 2026**, mỗi dòng một học viên:

| Cột | Ghi chú |
|---|---|
| Mã HV, Họ và tên | Họ và tên là bắt buộc |
| Ngày đăng ký | dd/mm/yyyy hoặc kiểu Ngày của Excel |
| Tháng khai giảng | số (8, 9, 10…) |
| Mã lớp, Khóa học, Ưu đãi, Tư vấn viên, Nguồn tuyển sinh, Ghi chú | văn bản |
| Học phí áp dụng | số tiền |
| Tình trạng | Đã hoàn tất / Chưa hoàn tất / Cần đối soát / Chưa đủ dữ liệu — để trống thì tự tính (học phí − đã thu) |
| Ngày TT 1, Số tiền TT 1, Ngày TT 2, Số tiền TT 2, … | mỗi lần đóng tiền một cặp cột (mẫu có sẵn 6 cặp; thêm TT 7, TT 8… theo cùng quy tắc tên) |

   Sheet **LỚP** (tùy chọn): Mã lớp, Lịch học, Ca học, Giáo viên, Khai giảng. Bỏ sheet này thì dashboard giữ thông tin lớp hiện có.
3. File mới **thay thế toàn bộ** dữ liệu hiện tại, vì vậy luôn giữ đủ danh sách học viên.
4. Tên cột nhận diện linh hoạt (không phân biệt hoa/thường, có/không dấu). Dòng tiêu đề có thể nằm trong 30 dòng đầu.

### Tải lên và xác nhận
1. **Cập nhật dữ liệu → Chọn file Excel**.
2. Màn hình **Xem trước** so sánh số học viên, doanh thu, thực thu, công nợ… giữa dữ liệu hiện tại và file mới, kèm các điểm cần lưu ý (thiếu ngày, chưa xếp lớp, mã trùng…).
3. Chọn một trong hai:
   - **Chỉ xem trên máy này**: toàn bộ KPI, biểu đồ, bảng cập nhật ngay, nhưng mất khi tải lại trang (dùng để kiểm tra trước).
   - **Xác nhận & đăng lên GitHub**: lưu vào `data/data.json` để **mọi người** truy cập đều thấy. Sau khi đăng, GitHub Pages cần 1–2 phút để phát hành (nếu còn thấy bản cũ, nhấn Ctrl+F5).

### Cách đăng lên GitHub an toàn (không để lộ token)
Source code **không chứa token nào**. Người có quyền ghi vào repo tự tạo token và dán khi đăng:

1. GitHub → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
2. *Repository access*: **Only select repositories** → chọn repo dashboard.
3. *Permissions → Repository permissions → **Contents: Read and write*** (không cần quyền nào khác). *Expiration*: 30–90 ngày.
4. Sao chép token (`github_pat_…`) và dán vào ô trong hộp thoại đăng. Token chỉ nằm trong bộ nhớ của tab trình duyệt (chọn "Nhớ trong tab này" nếu muốn giữ đến khi đóng tab), chỉ được gửi tới `api.github.com`, không được lưu vào repo hay trình duyệt lâu dài.
5. Hết hạn hoặc nghi lộ token → vào GitHub **Revoke** token là xong; website vẫn chạy bình thường.

Ai không có token/quyền ghi chỉ xem được hoặc dùng "Chỉ xem trên máy này", không thể thay đổi dữ liệu chung.

### Cách thủ công (không cần token)
Trên GitHub mở `data/data.json` → bút chì (Edit) hoặc **Add file → Upload files** để thay file. Dữ liệu phải đúng định dạng; cách dễ nhất là tải file mẫu, sửa, đăng bằng nút trên dashboard.

## C. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Trang trắng / "Không thể khởi động dashboard" | Kiểm tra mạng (cần tải Chart.js từ CDN) hoặc đặt file vào `vendor/` |
| Báo "dữ liệu dự phòng" | Không đọc được `data/data.json`: kiểm tra file có trong repo và Pages đã bật; mở bằng link GitHub Pages thay vì bấm đúp file |
| Đăng lỗi 401 | Token sai/hết hạn → tạo token mới |
| Đăng lỗi 403/404 | Token chưa cấp *Contents: Read and write* cho đúng repo, hoặc nhập sai tên repo/nhánh |
| Đăng xong vẫn thấy dữ liệu cũ | Đợi 1–2 phút rồi Ctrl+F5 |
| "Không tìm thấy sheet học viên" | File cần dòng tiêu đề có cột **Họ và tên**; dùng file mẫu |
| Dùng tên miền riêng, ô repo để trống | Điền `owner`/`repo` trong `config.js` hoặc nhập trực tiếp ở hộp thoại đăng |

## D. Chạy thử trên máy
Không cần cài gì. Mở thư mục bằng máy chủ tĩnh bất kỳ, ví dụ `python3 -m http.server 8000` rồi vào `http://localhost:8000`. (Bấm đúp `index.html` vẫn xem được nhờ `data/data.js`, nhưng nên dùng máy chủ để thấy đúng dữ liệu mới nhất.)

## E. Offline (tùy chọn)
Tải `chart.umd.min.js` (Chart.js 4.4.1) và `xlsx.full.min.js` (SheetJS 0.18.5) bỏ vào thư mục `vendor/`; trang tự dùng khi CDN không truy cập được.
