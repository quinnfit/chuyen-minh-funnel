/**
 * ============================================================
 * Chuyển Mình · Shift Within — Funnel Backend · Google Apps Script
 * ============================================================
 */

const CONFIG = {
  SHEET_ID: '1CNJXTKqbip9-AZydO5n5cT9nUAp9OyLmzKzfkNiuRr8',
  SHEET_TAB_NAME: 'Orders',

  // Điền API key Sepay sau khi liên kết VPBank xong
  SEPAY_API_KEY: 'CG0VVKTTYKTUNABBL47OERHIERYIWE2ZIJ3YM2H4WPDQOFJS65JZ5LSGDU19ASNM',

  EMAIL_SENDER_NAME: 'Chuyển Mình · Shift Within',
  EMAIL_REPLY_TO: 'quinnfit.training@gmail.com',

  // Không dùng Zalo — để trống
  ZALO_GROUP_URL: '',

  PRODUCT_NAME: 'Hành Trình 8 Tuần Chuyển Mình',
  PRODUCT_PRICE: 699000,
  // Điền link Drive tài liệu khoá học sau
  EBOOK_URL: '__EBOOK_URL__',

  BUMP1_NAME: '1-1 Call Định Hướng Lộ Trình với Quinn (30 phút)',
  BUMP1_PRICE: 500000,
  BUMP1_URL: '__BUMP1_URL__',

  BUMP2_NAME: 'Meal Plan Mẫu 1 Tuần Cá Nhân Hoá',
  BUMP2_PRICE: 200000,
  BUMP2_URL: '__BUMP2_URL__',

  CK_PREFIX: 'CM',

  HOTLINE: '__HOTLINE__',
  ADDRESS: 'Việt Nam',
  SITE_URL: 'https://chuyen-minh-funnel.vercel.app'
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
      'paidAt','sepayTxId','emailSent','rawSepay'
    ]);
    sheet.getRange('A1:O1').setFontWeight('bold').setBackground('#060D1A').setFontColor('#D4B896');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 15, 130);
    const dateFormat = 'dd/MM/yyyy HH:mm:ss';
    sheet.getRange('B:B').setNumberFormat(dateFormat);
    sheet.getRange('L:L').setNumberFormat(dateFormat);
    sheet.getRange('N:N').setNumberFormat(dateFormat);
  }
  return sheet;
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
    '', '', '', ''
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
      JSON.stringify(payload)
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
  const data = sheet.getRange(rowIndex, 1, 1, 15).getValues()[0];

  if (data[13] && String(data[13]).indexOf('ERROR') !== 0) {
    Logger.log('Order ' + data[0] + ' already processed, skipping');
    return;
  }

  try {
    sendConfirmationEmail({
      orderId: data[0],
      name: data[2],
      email: data[4],
      goal: data[5],
      bump1: !!data[6],
      bump2: !!data[7],
      totalAmount: data[8]
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
    totalAmount: formatVnd(order.totalAmount),
    goal: order.goal,
    hasBumps: !!(order.bump1 || order.bump2),
    bump1: !!order.bump1,
    bump2: !!order.bump2
  };

  const subject = 'Chuyển Mình · Shift Within — Xác nhận đơn hàng · Mã ' + v.orderId;

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
  let bumpSection = '';
  if (v.hasBumps) {
    bumpSection = '<h3 style="font-size:16px;color:#222;margin:24px 0 8px">Dịch vụ đi kèm bạn đã đăng ký</h3>';
    if (v.bump1) {
      bumpSection += '<p>• <strong>' + CONFIG.BUMP1_NAME + '</strong><br>'
        + 'Quinn sẽ liên hệ sắp xếp lịch call qua email này trong vòng 24 giờ.</p>';
    }
    if (v.bump2) {
      bumpSection += '<p>• <strong>' + CONFIG.BUMP2_NAME + '</strong><br>'
        + 'Tải tại: <a href="' + CONFIG.BUMP2_URL + '" style="color:#1B9FE8">' + CONFIG.BUMP2_URL + '</a></p>';
    }
  }

  return [
    '<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"></head>',
    '<body style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.7;color:#222;max-width:600px;margin:0 auto;padding:24px;background:#fff">',

    '<div style="background:#060D1A;padding:20px 24px;margin-bottom:24px">',
    '<p style="color:#D4B896;font-size:13px;letter-spacing:0.2em;font-weight:700;margin:0">CHUYỂN MÌNH · SHIFT WITHIN</p>',
    '</div>',

    '<h2 style="color:#222;font-size:20px;margin:0 0 16px">Chào ' + esc(v.name) + ',</h2>',
    '<p>Thanh toán đã được xác nhận. Chào mừng bạn đến với <strong>Hành Trình 8 Tuần Chuyển Mình</strong>!</p>',

    '<div style="background:#f9f6f2;padding:16px 20px;border-left:3px solid #D4B896;margin:20px 0">',
    '<p style="margin:0;font-size:14px">',
    'Mã đơn: <strong>' + esc(v.orderId) + '</strong><br>',
    'Số tiền: <strong>' + v.totalAmount + 'đ</strong>',
    v.goal ? '<br>Tình trạng: <strong>' + esc(v.goal) + '</strong>' : '',
    '</p>',
    '</div>',

    '<h3 style="font-size:16px;color:#222;margin:24px 0 8px">Tài liệu khoá học</h3>',
    '<p><a href="' + CONFIG.EBOOK_URL + '" style="color:#1B9FE8;font-weight:600">' + CONFIG.EBOOK_URL + '</a></p>',

    bumpSection,

    '<div style="background:#f9f6f2;padding:16px 20px;margin:24px 0;font-size:14px">',
    '<strong>Bước tiếp theo:</strong><br>',
    '1. Mở tài liệu và đọc phần Tuần 1<br>',
    '2. Điền form check-in đầu tiên Quinn sẽ gửi riêng cho bạn<br>',
    '3. Tham gia cộng đồng Telegram — link có trong tài liệu',
    '</div>',

    '<p style="font-size:14px;color:#555">Có thắc mắc? Reply email này hoặc liên hệ <strong>' + esc(CONFIG.EMAIL_REPLY_TO) + '</strong></p>',

    '<p style="font-size:13px;color:#888;margin-top:32px;border-top:1px solid #eee;padding-top:16px">',
    'Chuyển Mình · Shift Within · Quinn Nguyễn<br>',
    esc(CONFIG.SITE_URL),
    '</p>',

    '</body></html>'
  ].join('\n');
}

function buildPlainEmail(v) {
  let txt = 'CHUYỂN MÌNH · SHIFT WITHIN\n\n';
  txt += 'Chào ' + v.name + ',\n\n';
  txt += 'Thanh toán đã xác nhận. Chào mừng bạn đến với Hành Trình 8 Tuần Chuyển Mình!\n\n';
  txt += 'Mã đơn: ' + v.orderId + '\n';
  txt += 'Số tiền: ' + v.totalAmount + 'đ\n';
  if (v.goal) txt += 'Tình trạng: ' + v.goal + '\n';
  txt += '\nTÀI LIỆU KHOÁ HỌC\n' + CONFIG.EBOOK_URL + '\n\n';
  if (v.hasBumps) {
    txt += 'DỊCH VỤ ĐI KÈM\n';
    if (v.bump1) txt += '• ' + CONFIG.BUMP1_NAME + ': Quinn sẽ liên hệ sắp lịch call trong 24h\n';
    if (v.bump2) txt += '• ' + CONFIG.BUMP2_NAME + ': ' + CONFIG.BUMP2_URL + '\n';
    txt += '\n';
  }
  txt += 'Liên hệ: ' + CONFIG.EMAIL_REPLY_TO + '\n\n';
  txt += '—\nChuyển Mình · Shift Within · ' + CONFIG.SITE_URL + '\n';
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
    if (action === 'health') return corsJson({ success: true, message: 'Chuyen Minh backend running' });
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
    goal: 'Đã thử nhiều lần nhưng cứ bỏ cuộc',
    bump1: false,
    bump2: false,
    totalAmount: CONFIG.PRODUCT_PRICE
  });
  Logger.log(JSON.stringify(result, null, 2));
}

function testSendEmail() {
  sendConfirmationEmail({
    orderId: 'CM-TEST001',
    name: 'Quinn Test',
    email: CONFIG.EMAIL_REPLY_TO,
    goal: 'Đã thử nhiều lần nhưng cứ bỏ cuộc',
    bump1: false,
    bump2: false,
    totalAmount: CONFIG.PRODUCT_PRICE
  });
  Logger.log('Test email sent to ' + CONFIG.EMAIL_REPLY_TO);
}

function manualTriggerRow() {
  const rowNumber = 2; // ← Đổi số row trước khi Run
  triggerPaidActions(rowNumber);
  Logger.log('Manual trigger fired for row ' + rowNumber);
}
