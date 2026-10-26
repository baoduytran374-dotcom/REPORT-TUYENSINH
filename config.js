/* Cấu hình dashboard BETMA ENGLISH.
 * Mặc định KHÔNG cần chỉnh gì: không cần API, không cần Cloudflare Worker, không cần biến môi trường (VITE_API_URL…).
 * Dữ liệu đọc từ data/data.json cùng thư mục với trang web.
 * Chỉ điền owner/repo nếu dùng tên miền riêng (không phải <tên>.github.io/<repo>) để nút "Đăng lên GitHub" tự điền sẵn repo. */
window.BETMA_CONFIG = {
  owner: '',                    // ví dụ: 'betma-english'
  repo: '',                     // ví dụ: 'dashboard'
  branch: '',                   // để trống = nhánh mặc định của repo
  dataPath: 'data/data.json'    // vị trí file dữ liệu trong repo
};
