/**
 * ============================================================
 * 21 Ngày Hiểu Mình™ — Funnel Backend · Google Apps Script
 * ============================================================
 */

const CONFIG = {
  SHEET_ID: '1vld_cg-b28w3a9wkzl6Ie831L8EETn7gArRCaI__XsA',
  SHEET_TAB_NAME: 'Orders',

  SEPAY_API_KEY: 'CG0VVKTTYKTUNABBL47OERHIERYIWE2ZIJ3YM2H4WPDQOFJS65JZ5LSGDU19ASNM',

  EMAIL_SENDER_NAME: '21 Ngày Hiểu Mình™',
  EMAIL_REPLY_TO: 'quinnfit.training@gmail.com',

  // Điền link Zalo group sau khi có
  ZALO_GROUP_URL: '__ZALO_GROUP_URL__',

  PRODUCT_NAME: '21 Ngày Hiểu Mình™',
  PRODUCT_PRICE: 499000,
  // Video chào mừng — gửi ngay trong email xác nhận. Nội dung 3 Vòng sẽ mở khoá dần, KHÔNG gửi hết trong email này.
  WELCOME_VIDEO_URL: '__WELCOME_VIDEO_URL__',

  BUMP1_NAME: '',
  BUMP1_PRICE: 0,
  BUMP1_URL: '',

  BUMP2_NAME: '',
  BUMP2_PRICE: 0,
  BUMP2_URL: '',

  // Sepay webhook filter đã sửa để nhận mã bắt đầu bằng "HM" — phải giữ nguyên prefix này
  CK_PREFIX: 'HM',

  HOTLINE: '__HOTLINE__',
  ADDRESS: 'Việt Nam',
  SITE_URL: 'https://21ngay-hieuminh.vercel.app'
};

// ============================================================
// SHEET HELPERS
// ============================================================
function getSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_TAB_NAME);
    sheet.appendRow([
      'orderId','createdAt','name','phone','email','goal',
      'bump1','bump2','totalAmount','ckContent','status',
      'paidAt','sepayTxId','emailSent','rawSepay','password'
    ]);
    sheet.getRange('A1:P1').setFontWeight('bold').setBackground('#060D1A').setFontColor('#D4B896');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 16, 130);
    const dateFormat = 'dd/MM/yyyy HH:mm:ss';
    sheet.getRange('B:B').setNumberFormat(dateFormat);
    sheet.getRange('L:L').setNumberFormat(dateFormat);
    sheet.getRange('N:N').setNumberFormat(dateFormat);
  }
  return sheet;
}

// Chạy 1 lần để thêm cột "password" (cột P) vào Sheet đang có sẵn.
function addPasswordColumn() {
  const sheet = getSheet();
  if (sheet.getRange(1, 16).getValue()) {
    Logger.log('Cột password (P) đã có sẵn.');
    return;
  }
  sheet.getRange(1, 16).setValue('password')
    .setFontWeight('bold').setBackground('#060D1A').setFontColor('#D4B896');
  sheet.setColumnWidth(16, 130);
  Logger.log('Đã thêm cột password ở vị trí P.');
}

function formatDateColumns() {
  const sheet = getSheet();
  const dateFormat = 'dd/MM/yyyy HH:mm:ss';
  sheet.getRange('B:B').setNumberFormat(dateFormat);
  sheet.getRange('L:L').setNumberFormat(dateFormat);
  sheet.getRange('N:N').setNumberFormat(dateFormat);
  Logger.log('Date columns formatted');
}

function findOrderRowByContent(content) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const upper = (content || '').toUpperCase().trim();
  for (let i = 1; i < data.length; i++) {
    const orderCK = String(data[i][9] || '').toUpperCase().trim();
    const orderStatus = data[i][10];
    if (orderStatus === 'PENDING' && orderCK && upper.indexOf(orderCK) !== -1) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

function findOrderRowById(orderId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) return { row: i + 1, data: data[i] };
  }
  return null;
}

// ============================================================
// CORE ACTIONS
// ============================================================
function nextCkCode() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let max = 0;
  const pattern = new RegExp('^' + CONFIG.CK_PREFIX + '(\\d{4})$', 'i');
  for (let i = 1; i < data.length; i++) {
    const ck = String(data[i][9] || '').trim();
    const m = ck.match(pattern);
    if (m) {
      const num = parseInt(m[1], 10);
      if (num > max) max = num;
    }
  }
  return CONFIG.CK_PREFIX + String(max + 1).padStart(4, '0');
}

function createOrder(data) {
  const sheet = getSheet();
  const orderId = CONFIG.CK_PREFIX + '-' + Date.now().toString().slice(-9);
  const ckContent = nextCkCode();

  sheet.appendRow([
    orderId,
    new Date(),
    data.name || '',
    data.phone || '',
    data.email || '',
    data.goal || '',
    data.bump1 ? CONFIG.BUMP1_NAME : '',
    data.bump2 ? CONFIG.BUMP2_NAME : '',
    Number(data.totalAmount) || CONFIG.PRODUCT_PRICE,
    ckContent,
    'PENDING',
    '', '', '', '', ''
  ]);

  return {
    success: true,
    orderId: orderId,
    ckContent: ckContent,
    totalAmount: Number(data.totalAmount) || CONFIG.PRODUCT_PRICE
  };
}

function updateOrder(orderId, updates) {
  const row = findOrderRowById(orderId);
  if (!row) return { success: false, error: 'Order not found' };
  const sheet = getSheet();
  if (row.data[10] !== 'PENDING') {
    return { success: false, error: 'Order already ' + row.data[10] };
  }
  if (updates.bump1 !== undefined) {
    sheet.getRange(row.row, 7).setValue(updates.bump1 ? CONFIG.BUMP1_NAME : '');
  }
  if (updates.bump2 !== undefined) {
    sheet.getRange(row.row, 8).setValue(updates.bump2 ? CONFIG.BUMP2_NAME : '');
  }
  if (updates.totalAmount !== undefined) {
    sheet.getRange(row.row, 9).setValue(Number(updates.totalAmount));
  }
  return { success: true, orderId: orderId };
}

function getStatus(orderId) {
  const row = findOrderRowById(orderId);
  if (!row) return { success: false, error: 'Order not found' };
  return {
    success: true,
    orderId: orderId,
    status: row.data[10],
    paidAt: row.data[11] ? new Date(row.data[11]).toISOString() : null
  };
}

// Xác thực đăng nhập khu vực học viên: email + password + đơn đã PAID.
function verifyLogin(email, password) {
  if (!email || !password) return { success: false, error: 'Missing credentials' };
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const emailLower = String(email).toLowerCase().trim();
  const pwd = String(password).trim();
  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][4] || '').toLowerCase().trim();
    const rowPwd = String(data[i][15] || '').trim();
    const status = data[i][10];
    if (rowEmail === emailLower && rowPwd && rowPwd === pwd && status === 'PAID') {
      return { success: true, name: data[i][2], orderId: data[i][0] };
    }
  }
  return { success: false, error: 'Invalid credentials' };
}

function detectBumpsFromAmount(amount) {
  const base = CONFIG.PRODUCT_PRICE;
  const p1 = CONFIG.BUMP1_PRICE;
  const p2 = CONFIG.BUMP2_PRICE;
  function near(target) { return Math.abs(amount - target) <= 500; }
  if (p1 && p2 && near(base + p1 + p2)) return { bump1: true,  bump2: true };
  if (p1 && near(base + p1))            return { bump1: true,  bump2: false };
  if (p2 && near(base + p2))            return { bump1: false, bump2: true };
  return { bump1: false, bump2: false };
}

function handleSepayWebhook(payload) {
  const content = payload.content || '';
  const amount = Number(payload.transferAmount || 0);
  const sheet = getSheet();

  if (payload.transferType !== 'in') {
    return { success: true, message: 'Outgoing transaction, skipped' };
  }

  const match = findOrderRowByContent(content);

  if (!match) {
    sheet.appendRow([
      'UNMATCHED-' + Date.now(),
      new Date(),
      '', '', '', '',
      '', '', amount, content,
      'UNMATCHED', '', String(payload.id || ''), '',
      JSON.stringify(payload), ''
    ]);
    return { success: true, message: 'No matching order, logged for manual review' };
  }

  let bump1Set = !!match.data[6];
  let bump2Set = !!match.data[7];
  if (!bump1Set && !bump2Set) {
    const detected = detectBumpsFromAmount(amount);
    bump1Set = detected.bump1;
    bump2Set = detected.bump2;
    sheet.getRange(match.row, 7).setValue(bump1Set ? CONFIG.BUMP1_NAME : '');
    sheet.getRange(match.row, 8).setValue(bump2Set ? CONFIG.BUMP2_NAME : '');
  }

  sheet.getRange(match.row, 9).setValue(amount);
  sheet.getRange(match.row, 11).setValue('PAID');
  sheet.getRange(match.row, 12).setValue(new Date());
  sheet.getRange(match.row, 13).setValue(String(payload.id || ''));
  sheet.getRange(match.row, 15).setValue(JSON.stringify(payload));

  match.data[6] = bump1Set ? CONFIG.BUMP1_NAME : '';
  match.data[7] = bump2Set ? CONFIG.BUMP2_NAME : '';
  match.data[8] = amount;

  triggerPaidActions(match.row);

  return { success: true, message: 'Order paid', orderId: match.data[0] };
}

function triggerPaidActions(rowIndex) {
  const sheet = getSheet();
  const data = sheet.getRange(rowIndex, 1, 1, 16).getValues()[0];

  if (data[13] && String(data[13]).indexOf('ERROR') !== 0) {
    Logger.log('Order ' + data[0] + ' already processed, skipping');
    return;
  }

  // Sinh mật khẩu đăng nhập khu vực học viên nếu chưa có (cột P = 16).
  let password = data[15];
  if (!password) {
    password = generatePassword(8);
    sheet.getRange(rowIndex, 16).setValue(password);
    data[15] = password;
  }

  try {
    sendConfirmationEmail({
      orderId: data[0],
      name: data[2],
      email: data[4],
      goal: data[5],
      bump1: !!data[6],
      bump2: !!data[7],
      totalAmount: data[8],
      password: password
    });
    sheet.getRange(rowIndex, 14).setValue(new Date());
  } catch (err) {
    sheet.getRange(rowIndex, 14).setValue('ERROR: ' + err.toString());
  }
}

function onSheetEdit(e) {
  if (!e || !e.range || !e.source) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_TAB_NAME) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row === 1 || col !== 11) return;
  const newValue = String(e.range.getValue() || '').toUpperCase().trim();
  if (newValue === 'PAID') {
    if (!sheet.getRange(row, 12).getValue()) {
      sheet.getRange(row, 12).setValue(new Date());
    }
    triggerPaidActions(row);
    Logger.log('Manual trigger PAID for row ' + row);
  }
}

// ============================================================
// EMAIL
// ============================================================
function sendConfirmationEmail(order) {
  const v = {
    name: order.name,
    orderId: order.orderId,
    email: order.email,
    password: order.password,
    totalAmount: formatVnd(order.totalAmount),
    goal: order.goal,
    hasZalo: !!CONFIG.ZALO_GROUP_URL
  };

  const subject = '21 Ngày Hiểu Mình™ — Xác nhận đơn hàng · Mã ' + v.orderId;

  MailApp.sendEmail({
    to: order.email,
    subject: subject,
    htmlBody: buildHtmlEmail(v),
    body: buildPlainEmail(v),
    name: CONFIG.EMAIL_SENDER_NAME,
    replyTo: CONFIG.EMAIL_REPLY_TO
  });
}

function buildHtmlEmail(v) {
  const zaloRow = v.hasZalo
    ? '<p>3. Tham gia <a href="' + CONFIG.ZALO_GROUP_URL + '" style="color:#1B9FE8">cộng đồng Zalo kín</a> — nơi bạn duy trì nhịp thực hành cùng mọi người.</p>'
    : '<p>3. Cộng đồng Zalo kín — link sẽ được gửi riêng cho bạn trong 24 giờ.</p>';

  return [
    '<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"></head>',
    '<body style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.7;color:#222;max-width:600px;margin:0 auto;padding:24px;background:#fff">',

    '<div style="background:#060D1A;padding:20px 24px;margin-bottom:24px">',
    '<p style="color:#D4B896;font-size:13px;letter-spacing:0.2em;font-weight:700;margin:0">21 NGÀY HIỂU MÌNH™</p>',
    '</div>',

    '<h2 style="color:#222;font-size:20px;margin:0 0 16px">Chào ' + esc(v.name) + ',</h2>',
    '<p>Thanh toán đã được xác nhận. Chào mừng bạn bắt đầu hành trình <strong>21 Ngày Hiểu Mình™</strong> — hiểu bản thân, hiểu cơ thể để chuyển hóa vóc dáng một cách tự nhiên, không kỷ luật cưỡng ép.</p>',

    '<div style="background:#f9f6f2;padding:16px 20px;border-left:3px solid #D4B896;margin:20px 0">',
    '<p style="margin:0;font-size:14px">',
    'Mã đơn: <strong>' + esc(v.orderId) + '</strong><br>',
    'Số tiền: <strong>' + v.totalAmount + 'đ</strong>',
    v.goal ? '<br>Điều bạn đang gặp phải: <strong>' + esc(v.goal) + '</strong>' : '',
    '</p>',
    '</div>',

    '<h3 style="font-size:16px;color:#222;margin:24px 0 8px">Thông tin đăng nhập khu vực học viên</h3>',
    '<div style="background:#fafafa;padding:16px 20px;border:2px dashed #D4B896;border-radius:8px;margin:12px 0;font-family:monospace;font-size:14px">',
    'Trang học: <strong>' + CONFIG.SITE_URL + '/khoahoc</strong><br>',
    'Email: <strong>' + esc(v.email) + '</strong><br>',
    'Mật khẩu: <strong style="color:#5C1A1B;font-size:16px">' + esc(v.password) + '</strong>',
    '</div>',
    '<p style="margin:12px 0"><a href="' + CONFIG.SITE_URL + '/khoahoc" style="display:inline-block;background:#D4B896;color:#0A0A0A;padding:12px 26px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">Vào khu vực học viên →</a></p>',
    '<p style="font-size:13px;color:#666">Lưu lại email và mật khẩu này để đăng nhập lại bất cứ lúc nào. Nội dung mở dần theo từng ngày.</p>',

    '<h3 style="font-size:16px;color:#222;margin:24px 0 8px">Video chào mừng từ Quinn</h3>',
    '<p><a href="' + CONFIG.WELCOME_VIDEO_URL + '" style="color:#1B9FE8;font-weight:600">' + CONFIG.WELCOME_VIDEO_URL + '</a></p>',

    '<div style="background:#f9f6f2;padding:16px 20px;border-left:3px solid #1B9FE8;margin:24px 0;font-size:14px">',
    '<p style="margin:0"><strong>Về cách nội dung được gửi:</strong><br>',
    'Nội dung hành trình (video + sách bài tập từng Vòng) sẽ được <strong>mở khoá dần theo từng ngày</strong> trong khu vực học viên — không mở hết một lần.</p>',
    '</div>',

    '<div style="background:#f9f6f2;padding:16px 20px;margin:24px 0;font-size:14px">',
    '<strong>Bước tiếp theo:</strong><br>',
    '<p style="margin:8px 0 0">1. Đăng nhập khu vực học viên bằng email và mật khẩu ở trên.</p>',
    '<p>2. Xem video chào mừng, rồi bắt đầu Ngày 1.</p>',
    zaloRow,
    '<p>4. Mỗi ngày hoàn thành một bài; ngày tiếp theo mở vào sáng hôm sau.</p>',
    '</div>',

    '<p style="font-size:14px;color:#555">Có thắc mắc? Reply email này hoặc liên hệ <strong>' + esc(CONFIG.EMAIL_REPLY_TO) + '</strong></p>',

    '<p style="font-size:13px;color:#888;margin-top:32px;border-top:1px solid #eee;padding-top:16px">',
    '21 Ngày Hiểu Mình™ · Quinn Nguyễn<br>',
    esc(CONFIG.SITE_URL),
    '</p>',

    '</body></html>'
  ].join('\n');
}

function buildPlainEmail(v) {
  let txt = '21 NGÀY HIỂU MÌNH\n\n';
  txt += 'Chào ' + v.name + ',\n\n';
  txt += 'Thanh toán đã xác nhận. Chào mừng bạn bắt đầu hành trình 21 Ngày Hiểu Mình!\n\n';
  txt += 'Mã đơn: ' + v.orderId + '\n';
  txt += 'Số tiền: ' + v.totalAmount + 'đ\n';
  if (v.goal) txt += 'Điều bạn đang gặp phải: ' + v.goal + '\n';
  txt += '\nTHONG TIN DANG NHAP KHU VUC HOC VIEN\n';
  txt += 'Trang hoc: ' + CONFIG.SITE_URL + '/khoahoc\n';
  txt += 'Email: ' + v.email + '\n';
  txt += 'Mat khau: ' + v.password + '\n';
  txt += 'Luu lai email va mat khau nay de dang nhap lai bat cu luc nao.\n\n';
  txt += 'VIDEO CHAO MUNG TU QUINN\n' + CONFIG.WELCOME_VIDEO_URL + '\n\n';
  txt += 'VE CACH NOI DUNG DUOC GUI\n';
  txt += 'Noi dung hanh trinh (video + sach bai tap tung Vong) se duoc mo khoa dan theo tung ngay trong khu vuc hoc vien.\n\n';
  txt += 'BUOC TIEP THEO\n';
  txt += '1. Dang nhap khu vuc hoc vien bang email va mat khau o tren.\n';
  txt += '2. Xem video chao mung, roi bat dau Ngay 1.\n';
  txt += v.hasZalo
    ? '3. Tham gia cong dong Zalo kin: ' + CONFIG.ZALO_GROUP_URL + '\n'
    : '3. Link cong dong Zalo kin se duoc gui rieng trong 24 gio.\n';
  txt += '4. Moi ngay hoan thanh mot bai; ngay tiep theo mo vao sang hom sau.\n\n';
  txt += 'Lien he: ' + CONFIG.EMAIL_REPLY_TO + '\n\n';
  txt += '--\n21 Ngay Hieu Minh - ' + CONFIG.SITE_URL + '\n';
  return txt;
}

// ============================================================
// UTILITIES
// ============================================================
function formatVnd(n) {
  return Number(n).toLocaleString('en-US').replace(/,/g, '.');
}

function esc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Sinh mật khẩu 8 ký tự, bỏ các ký tự dễ nhầm (0/O, 1/I/L).
function generatePassword(length) {
  length = length || 8;
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let pwd = '';
  for (let i = 0; i < length; i++) {
    pwd += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return pwd;
}

function corsJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HTTP ENTRY POINTS
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'createOrder') return corsJson(createOrder(body.data || {}));
    if (body.action === 'updateOrder') return corsJson(updateOrder(body.orderId, body.data || {}));
    if (body.action === 'verifyLogin') return corsJson(verifyLogin(body.email, body.password));
    if (body.content !== undefined && body.transferAmount !== undefined) {
      const key = (e.parameter && e.parameter.key) || '';
      if (key !== CONFIG.SEPAY_API_KEY) return corsJson({ success: false, error: 'Unauthorized' });
      return corsJson(handleSepayWebhook(body));
    }
    return corsJson({ success: false, error: 'Unknown request' });
  } catch (err) {
    return corsJson({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    if (action === 'getStatus') return corsJson(getStatus(e.parameter.orderId));
    if (action === 'health') return corsJson({ success: true, message: '21 Ngay Hieu Minh backend running' });
    return corsJson({ success: false, error: 'Unknown action' });
  } catch (err) {
    return corsJson({ success: false, error: err.toString() });
  }
}

// ============================================================
// TEST FUNCTIONS
// ============================================================
function testCreateOrder() {
  const result = createOrder({
    name: 'Test User',
    phone: '0901234567',
    email: CONFIG.EMAIL_REPLY_TO,
    goal: 'Đã bắt đầu nhiều lần nhưng vẫn bỏ cuộc',
    bump1: false,
    bump2: false,
    totalAmount: CONFIG.PRODUCT_PRICE
  });
  Logger.log(JSON.stringify(result, null, 2));
}

function testSendEmail() {
  sendConfirmationEmail({
    orderId: 'HM-TEST001',
    name: 'Quinn Test',
    email: CONFIG.EMAIL_REPLY_TO,
    goal: 'Đã bắt đầu nhiều lần nhưng vẫn bỏ cuộc',
    bump1: false,
    bump2: false,
    totalAmount: CONFIG.PRODUCT_PRICE,
    password: 'TEST1234'
  });
  Logger.log('Test email sent to ' + CONFIG.EMAIL_REPLY_TO);
}

function testVerifyLogin() {
  Logger.log(JSON.stringify(verifyLogin(CONFIG.EMAIL_REPLY_TO, 'PASTE_PASSWORD_HERE'), null, 2));
}

function manualTriggerRow() {
  const rowNumber = 2; // ← Đổi số row trước khi Run
  triggerPaidActions(rowNumber);
  Logger.log('Manual trigger fired for row ' + rowNumber);
}
