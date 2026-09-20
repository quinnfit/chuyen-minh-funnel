# Sepay Funnel — User Context (21 Ngày Hiểu Mình™)

- Brand: 21 Ngày Hiểu Mình™ (Chuyển Mình · Shift Within)
- Sản phẩm chính: 21 Ngày Hiểu Mình™
- Giá: 499000 (499.000đ)
- Bumps: không có
- Email sender (Gmail): quinnfit.training@gmail.com
- Ngân hàng: VPBank
- Số tài khoản: 606804407
- Chủ tài khoản: HO KINH DOANH QUINNFIT
- (Lưu ý: VA "AGBSPQUINNFIT" của dự án cũ không tạo VA mới được vì tài khoản là hộ kinh doanh, SePay chỉ hỗ trợ tạo VA chính thức cho cá nhân)
- QUAN TRỌNG: STK gốc 606804407 KHÔNG detect real-time transaction (đã confirm qua test CK thật) — PHẢI dùng VA AGBSPQUINNFIT cho QR/checkout.html. Webhook vẫn tied vào account 606804407 nên không cần đổi webhook config, vì giao dịch qua VA vẫn ghi nhận trên account gốc.
- Video chào mừng (WELCOME_VIDEO_URL): CHƯA CÓ — cần bổ sung trước Phase 7 (test end-to-end thật). Email chỉ gửi video chào mừng, nội dung 3 Vòng unlock dần qua email/Zalo, không gửi hết 1 lần.
- Zalo group URL: CHƯA CÓ — skip, cần bổ sung sau
- Domain: mặc định Vercel (*.vercel.app)

- Webhook Sepay: dùng chung webhook cũ, đổi tên "Hiểu Mình Auto Confirm" (thay thế automation "Chuyển Mình", không chạy song song)
- Sepay API key: CG0VVKTTYKTUNABBL47OERHIERYIWE2ZIJ3YM2H4WPDQOFJS65JZ5LSGDU19ASNM (lấy lại từ Code.gs cũ, gắn với webhook trên)
- CK_PREFIX: HM (Sepay webhook "Hiểu Mình Auto Confirm" filter "Lọc theo mã thanh toán" đã được user tự sửa để nhận "HM" — Code.gs khớp theo, KHÔNG đổi lại DH nữa)
- Email tự động: viết lại toàn bộ nội dung riêng cho 21 Ngày Hiểu Mình™ (không dùng lại nội dung Chuyển Mình)

## Trạng thái
- Phase 0: done
- Phase 1: done
- Phase 2: done (API key lấy lại từ Code.gs cũ)
- Phase 3: done (Sheet + Code.gs + trigger + format date, test pass)
- Google Sheet ID: 1vld_cg-b28w3a9wkzl6Ie831L8EETn7gArRCaI__XsA
- GAS_URL: https://script.google.com/macros/s/AKfycbyc1gwhRXHeLljXby5PVqA5zkbT8LuJcdORLwELQG3Wu0lYUknTwrv5JUV_nYFQmZT6/exec
- Vercel URL: https://21ngay-hieuminh.vercel.app (domain đổi tên sau deploy, verified live)
- Phase 5: done (deployed, /dang-ky verified live, cleanUrls OK)
- Phase 6: skipped (GAS_URL + bank info đã điền trước khi deploy)
- Phase 7: done — test end-to-end CK thật 1.000đ pass (mã HM0005): checkout tự chuyển /xac-nhan, email đến, Sheet PAID/paidAt/emailSent đều update
- Root cause đã fix: Sepay "Cấu trúc mã thanh toán" (Cấu hình chung công ty) đổi Mẫu mặc định tiền tố DH → HM (3-10 ký tự số) để khớp CK_PREFIX code + webhook filter
- Phase 4: done (webhook "Hiểu Mình Auto Confirm" trỏ GAS_URL mới, no test button trên Sepay UI này — verify ở Phase 7)
- dang-ky.html + checkout.html: đã điền GAS_URL + bank info (VPBank/606804407/HO KINH DOANH QUINNFIT) trực tiếp, không còn placeholder __GAS_URL__/__BANK*__/__BUMP*_PRICE__
