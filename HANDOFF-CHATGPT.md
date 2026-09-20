# Handoff — Funnel "21 Ngày Hiểu Mình™"

Tài liệu này tổng hợp mọi thông tin cần thiết để tiếp tục chỉnh sửa landing page + backend này bằng công cụ khác (ChatGPT hoặc bất kỳ AI nào). Copy toàn bộ nội dung file này dán vào đầu 1 conversation mới để AI có đủ context.

⚠️ **File này chứa API key/thông tin nhạy cảm — không paste công khai, không commit lên GitHub, không chia sẻ ngoài mục đích làm việc với AI của riêng bạn.**

---

## 1. Tổng quan dự án

- **Sản phẩm:** "21 Ngày Hiểu Mình™" — khoá học online 21 ngày, giá 499.000đ, thuộc thương hiệu "Chuyển Mình · Shift Within" của Quinn Nguyễn (HKD QuinnFit).
- **Mô hình:** Landing page tĩnh (HTML/CSS/JS thuần, không framework, không build step) + backend Google Apps Script xử lý đơn hàng/thanh toán/email tự động qua Sepay.
- **Vận hành:** Evergreen (tự động, không theo đợt/cohort cố định).

## 2. Vị trí file trên máy

```
/Users/ap24h.vn/Documents/QUINNFIT/Landing page- Chuyển mình/SKILL CLAUDE CODE/Skill Claude Builder 2/landing-page-builder/funnel/
```

File trong đó:
| File | Vai trò |
|---|---|
| `index.html` | Sale page chính |
| `dang-ky.html` | Form đăng ký (bước 1) |
| `checkout.html` | Trang thanh toán QR (bước 2) |
| `xac-nhan.html` | Trang xác nhận sau thanh toán (bước 3) |
| `khoahoc.html` | Khu vực học viên (đăng nhập bằng email + password, xem nội dung khoá học) |
| `tdee.html` | Công cụ tính TDEE (nhúng iframe từ thehinh.com) |
| `privacy.html` | Chính sách bảo mật |
| `refund.html` | Chính sách hoàn tiền |
| `vercel.json` | Config routing cho Vercel (clean URLs) |
| `Code.gs` | Backend Google Apps Script — **CHỈ lưu local, KHÔNG push lên GitHub** (chứa API key) |
| `user-context.md` | Ghi chú cấu hình/lịch sử — **CHỈ lưu local** |
| `HANDOFF-CHATGPT.md` | Chính là file này |

## 3. Git / GitHub

- Repo: `https://github.com/quinnfit/chuyen-minh-funnel.git`
- Branch: `main` (Vercel auto-deploy mỗi khi push lên `main`)
- **Chỉ commit + push 4 trang HTML chính + `vercel.json` + các trang phụ (`privacy.html`, `refund.html`, `tdee.html`, `khoahoc.html`).** KHÔNG commit `Code.gs` và `user-context.md` (chứa secret).
- Lệnh thường dùng:
  ```bash
  cd "/Users/ap24h.vn/Documents/QUINNFIT/Landing page- Chuyển mình/SKILL CLAUDE CODE/Skill Claude Builder 2/landing-page-builder/funnel"
  git add index.html dang-ky.html checkout.html xac-nhan.html   # chỉ file đã sửa
  git commit -m "Mô tả thay đổi"
  git push origin main
  ```
- **Lưu ý xác thực:** GitHub không nhận password thường cho git, phải dùng Personal Access Token (PAT). Nếu push báo lỗi "Invalid username or token":
  1. Vào https://github.com/settings/tokens/new
  2. Tick scope "repo" → Generate token → copy
  3. Khi Terminal hỏi Username: `quinnfit`, Password: paste token đó (không phải mật khẩu tài khoản)

## 4. Deploy — Vercel

- Project name: `21-ngay-hieu-minh` (đã đổi tên từ `chuyen-minh-funnel` cũ)
- **URL live:** https://21ngay-hieuminh.vercel.app
- Auto-deploy khi push lên `main` của repo trên. Không cần chạy lệnh `vercel` thủ công.
- Domain riêng: chưa mua/gắn (đã mua `quinnfit.vn` ở tenten.vn nhưng CHƯA trỏ — xem mục 9).

## 5. Backend — Google Apps Script (Code.gs)

- Google Sheet lưu đơn hàng: Sheet ID `1vld_cg-b28w3a9wkzl6Ie831L8EETn7gArRCaI__XsA`, tab "Orders"
- GAS Web App URL (đã deploy, KHÔNG đổi khi sửa code + deploy "New version"):
  ```
  https://script.google.com/macros/s/AKfycbyc1gwhRXHeLljXby5PVqA5zkbT8LuJcdORLwELQG3Wu0lYUknTwrv5JUV_nYFQmZT6/exec
  ```
- URL này đã được hard-code sẵn trong `dang-ky.html` và `checkout.html` (biến `GAS_URL` trong thẻ `<script>`).
- Mỗi khi sửa `Code.gs`: phải copy paste vào Apps Script editor (Extensions → Apps Script trên Google Sheet) rồi **Deploy → Manage deployments → Edit (bút chì) → Version: New version → Deploy**. URL giữ nguyên, không cần sửa lại frontend.
- Config quan trọng trong `Code.gs`:
  ```js
  SEPAY_API_KEY: 'CG0VVKTTYKTUNABBL47OERHIERYIWE2ZIJ3YM2H4WPDQOFJS65JZ5LSGDU19ASNM'
  CK_PREFIX: 'HM'   // ⚠️ XEM MỤC 6 — không được đổi tuỳ tiện
  EMAIL_REPLY_TO: 'quinnfit.training@gmail.com'
  ZALO_GROUP_URL: 'https://zalo.me/g/hu37f2kanss5wtvzpvfx'
  COURSE_URL: 'https://21ngay-hieuminh.vercel.app/khoahoc'
  PRODUCT_PRICE: 499000
  ```

## 6. Thanh toán — Sepay (⚠️ đọc kỹ, đã tốn nhiều thời gian debug)

- Ngân hàng: **VPBank**, chủ TK **HO KINH DOANH QUINNFIT**
- Số tài khoản gốc: `606804407`
- **Tài khoản ảo (VA) dùng để nhận thanh toán: `AGBSPQUINNFIT`**
  - ⚠️ QUAN TRỌNG: STK gốc `606804407` KHÔNG detect được giao dịch real-time (đã test thật và fail). Bắt buộc phải dùng VA `AGBSPQUINNFIT` trong QR code + hiển thị trên `checkout.html`. Không tạo VA mới được nữa vì tài khoản là hộ kinh doanh (Sepay chỉ cho tạo VA chính thức với tài khoản cá nhân).
  - Webhook Sepay vẫn gắn vào tài khoản gốc `606804407` (không phải VA) — giao dịch qua VA vẫn ghi nhận đúng trên tài khoản gốc nên không cần đổi.
- Webhook Sepay tên: **"Hiểu Mình Auto Confirm"** (Sepay dashboard → Tích hợp WebHooks), trỏ tới GAS_URL ở mục 5 kèm `?key=<SEPAY_API_KEY>`.
- ⚠️ **QUAN TRỌNG NHẤT:** Sepay có tính năng "Cấu hình chung công ty → Cấu trúc mã thanh toán" định nghĩa tiền tố mã đơn hợp lệ (mặc định là "DH"). Đã đổi tiền tố mặc định thành **"HM"** (3-10 ký tự số) để khớp với `CK_PREFIX: 'HM'` trong Code.gs và filter "Lọc theo mã thanh toán" của webhook. **Nếu sau này đổi `CK_PREFIX` trong code, PHẢI đổi luôn cấu hình này bên Sepay, nếu không webhook sẽ không bao giờ bắn về** (đây là lỗi đã từng xảy ra và mất nhiều bước debug mới tìm ra).

## 7. Email tự động

- Gửi qua Gmail (MailApp trong Apps Script), sender: `quinnfit.training@gmail.com`
- Điểm chạm hằng ngày (21 ngày) gửi qua **email**, KHÔNG phải Zalo (đã đổi copy trên landing page để phản ánh đúng điều này — chỉ còn "Cộng đồng Zalo kín" là nhóm cộng đồng dùng Zalo, tách biệt với điểm chạm hằng ngày).
- Nội dung khoá học đầy đủ giờ nằm trong khu vực học viên `khoahoc.html` (đăng nhập bằng email + mật khẩu được tạo sau khi thanh toán), không còn gửi 1 video chào mừng đơn lẻ qua email như thiết kế ban đầu.

## 8. Meta Pixel (Facebook Ads)

- Pixel ID: `1403596911940428`
- Đã cài trên tất cả các trang, sự kiện đã setup:
  - `PageView` — tất cả trang
  - `ViewContent` — sale page (`index.html`)
  - `InitiateCheckout` — sau khi tạo đơn thành công ở `dang-ky.html`
  - `Purchase` — ở `xac-nhan.html`, kèm value/currency thật, có chống đếm trùng khi refresh (dùng `sessionStorage`)

## 9. Nội dung nguồn (Google Docs)

- Outline khoá học đầy đủ (3 Vòng, script từng ngày, tên sách bài tập): 
  `https://docs.google.com/document/d/1FqrB7JyjhzanK6W_JkGBBhiMU1DehdyEfw31RyQNVbw/edit`
- Ebook quà tặng "Cơm Nhà Vẫn Vào Form" (50 công thức + hướng dẫn):
  `https://docs.google.com/document/d/1meiCBNt9fgXD60TvbVyKmA-WL4zDrJQC/edit`
- **Lưu ý:** nội dung sale page phải luôn khớp với outline thật (số lượng video, tên sách bài tập...) — mỗi khi outline đổi, cần so sánh lại và cập nhật `index.html` phần Deliverables/Offer Stack cho khớp.

## 10. Domain riêng (chưa hoàn tất)

- Đã mua `quinnfit.vn` tại tenten.vn nhưng CHƯA trỏ về Vercel (đang chờ kết quả chạy ads thử trước khi đầu tư domain riêng).
- Kế hoạch đã thống nhất: dùng **subdomain riêng cho từng sản phẩm** (vd `21ngay.quinnfit.vn`), KHÔNG dùng subpath (vd `quinnfit.vn/21ngay`) — vì code hiện tại dùng link tuyệt đối kiểu `/dang-ky`, `/checkout` sẽ vỡ nếu deploy dưới subpath.
- Khi cần làm: vào Vercel project → Settings → Domains → thêm subdomain → lấy bản ghi CNAME → thêm vào DNS của domain trên tenten.vn.

## 11. Thương hiệu / Style

- Font: Be Vietnam Pro (Google Fonts)
- Bảng màu (CSS variables dùng xuyên suốt mọi trang):
  ```css
  --bg:#060D1A; --bg2:#0A1628; --bg3:#0D1E3A;
  --gold:#D4B896; --gold-dim:rgba(212,184,150,0.12); --gold-border:rgba(212,184,150,0.2);
  --blue:#1B9FE8;
  --white:#FAFAFA; --w8:rgba(250,250,250,0.8); --w6:rgba(250,250,250,0.6); --w4:rgba(250,250,250,0.4);
  ```
- Không dùng dấu ™ sau "21 Ngày Hiểu Mình" trên sale page (đã bỏ theo yêu cầu), nhưng vẫn giữ ở "90 Ngày Chuyển Mình™" (sản phẩm khác).
- Tránh dùng em-dash (—) trong copy mới nếu user yêu cầu (đã có tiền lệ 1 lần).

## 12. Lưu ý khi làm việc với AI khác (ChatGPT...)

ChatGPT bản chat thường (không phải Claude Code/Cursor) **không có quyền truy cập trực tiếp vào file trên máy, terminal, hay trình duyệt** như Claude Code đang có. Nghĩa là:
- Bạn sẽ cần tự copy nội dung file (vd `index.html`) dán vào ChatGPT để nó đọc/sửa, rồi copy kết quả dán ngược lại vào file.
- Các thao tác như `git commit/push`, deploy, mở Google Sheet, sửa Apps Script... bạn sẽ phải tự làm theo hướng dẫn của ChatGPT (nó chỉ đưa lệnh, không tự chạy được).
- Nếu muốn AI có thể tự thao tác file/terminal/browser như Claude Code đang làm, cần dùng ChatGPT phiên bản có "Agent"/Code Interpreter kết nối máy thật (hoặc công cụ tương tự Claude Code/Cursor), không phải chatgpt.com thường.
