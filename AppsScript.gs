/**
 * نظام سكانر أدوية هيئة رصد - الباك إند الكامل (Google Apps Script)
 * ---------------------------------------------------------------
 * يشمل: تسجيل مستخدمين + تحقق OTP بالإيميل + موافقة أدمن + تفعيل/إيقاف حسابات
 *       + تسجيل سكانات الأدوية + تصدير CSV بصيغة رصد مع منع تصدير نفس الصنف مرتين
 *
 * طريقة التركيب:
 * 1) افتح جوجل شيت جديد فاضي (أو استخدم شيت موجود).
 * 2) من القائمة: Extensions > Apps Script
 * 3) امسح أي كود موجود، والصق هذا الكود كامل، احفظ (Ctrl+S).
 * 4) روح لدالة setupAdmin() تحت في آخر الملف:
 *    - غيّر قيمة ADMIN_PASSWORD_PLACEHOLDER لكلمة السر الحقيقية بتاعتك (اكتبها هنا بس، متبعتهاش لحد).
 *    - من شريط الأدوات فوق اختار الدالة setupAdmin من القائمة المنسدلة، ودوس ▶ Run.
 *    - هيطلب منك صلاحيات (Authorize) - وافق.
 *    - بعد ما تشتغل بنجاح، امسح القيمة اللي كتبتها في ADMIN_PASSWORD_PLACEHOLDER (رجّعها زي ما كانت) واحفظ تاني، عشان الباسورد متفضلش مكتوبة في الكود.
 * 5) Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 6) هيديك رابط ينتهي بـ /exec -> ده رابط الربط اللي هتحطه في صفحة portal.html
 * 7) سجّل دخول كأدمن بإيميلك وباسوردك، وأول حاجة غيّر الباسورد من داخل النظام (زرار "تغيير كلمة السر").
 */

var DATA_SHEET_NAME = 'Sheet1';
var USERS_SHEET_NAME = 'Users';
var REPORT_SHEET_NAME = 'ExportReport';
var ADMIN_EMAIL = 'ahmedkamal1587@gmail.com';
var OTP_VALID_MINUTES = 10;
var TOKEN_VALID_HOURS = 12;

// ============ نقطة الدخول الرئيسية ============
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'scan';

    switch (action) {
      case 'register':        return handleRegister(body);
      case 'verifyOtp':        return handleVerifyOtp(body);
      case 'resendOtp':        return handleResendOtp(body);
      case 'login':             return handleLogin(body);
      case 'changePassword':   return handleChangePassword(body);
      case 'scan':              return handleScan(body);
      case 'adminListUsers':   return handleAdminListUsers(body);
      case 'adminApprove':     return handleAdminApprove(body);
      case 'adminReject':      return handleAdminReject(body);
      case 'adminToggleStatus':return handleAdminToggleStatus(body);
      case 'exportRegulator':  return handleExportRegulator(body);
      case 'getExportLog':     return handleGetExportLog(body);
      default:
        return jsonOutput({ status: 'error', message: 'إجراء غير معروف' });
    }
  } catch (err) {
    return jsonOutput({ status: 'error', message: 'خطأ: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ أدوات مساعدة للشيتات ============
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function getUsersSheet() {
  // Email | PasswordHash | Salt | Role | Status | OTP | OTPExpiry | Token | TokenExpiry | CreatedAt
  return getOrCreateSheet(USERS_SHEET_NAME,
    ['Email', 'PasswordHash', 'Salt', 'Role', 'Status', 'OTP', 'OTPExpiry', 'Token', 'TokenExpiry', 'CreatedAt']);
}

function getDataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['GTIN', 'SN', 'BN', 'XD', 'ScannedBy', 'ScannedAt', 'Exported']);
  }
  // فرض تنسيق "نص" على الأعمدة دي عشان جوجل شيت مايشيلش الأصفار البادئة
  // أو يحوّل تاريخ الصلاحية لصيغة تانية غير DD/MM/YYYY
  sheet.getRange('A:D').setNumberFormat('@');
  return sheet;
}

// يضمن إن GTIN يفضل 14 رقم بالظبط حتى لو جوجل شيت شال صفر بادئ من بيانات قديمة
function padGtin(gtin) {
  var digits = String(gtin || '').replace(/[^0-9]/g, '');
  while (digits.length < 14 && digits.length > 0) digits = '0' + digits;
  return digits;
}

function getReportSheet() {
  var sheet = getOrCreateSheet(REPORT_SHEET_NAME,
    ['GTIN', 'SN', 'BN', 'XD', 'ExportedAt', 'ExportedBy', 'FileName']);
  sheet.getRange('A:D').setNumberFormat('@');
  return sheet;
}

function findUserRow(email) {
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  var target = String(email || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) {
      return { rowIndex: i + 1, row: data[i], sheet: sheet };
    }
  }
  return null;
}

// ============ تشفير كلمة السر ============
function generateSalt() { return Utilities.getUuid(); }

function hashPassword(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + '::' + salt);
  return digest.map(function (b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function generateToken() { return Utilities.getUuid() + '-' + Utilities.getUuid(); }
function generateOtp() { return String(Math.floor(100000 + Math.random() * 900000)); }

// ============ التسجيل + OTP ============
function handleRegister(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var password = String(body.password || '');
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return jsonOutput({ status: 'error', message: 'بريد إلكتروني غير صحيح' });
  }
  if (password.length < 6) {
    return jsonOutput({ status: 'error', message: 'كلمة السر لازم تكون 6 أحرف على الأقل' });
  }

  var salt = generateSalt();
  var hash = hashPassword(password, salt);
  var otp = generateOtp();
  var otpExpiry = new Date(Date.now() + OTP_VALID_MINUTES * 60000);

  var existing = findUserRow(email);
  if (existing) {
    var status = existing.row[4];
    if (status === 'active') {
      return jsonOutput({ status: 'error', message: 'الحساب موجود بالفعل، استخدم تسجيل الدخول' });
    }
    if (status === 'approval_pending') {
      return jsonOutput({ status: 'error', message: 'طلبك مُرسل بالفعل وبانتظار موافقة الأدمن' });
    }
    // otp_pending أو rejected: اسمح بإعادة المحاولة وإرسال OTP جديد
    existing.sheet.getRange(existing.rowIndex, 2, 1, 7).setValues([[hash, salt, existing.row[3] || 'user', 'otp_pending', otp, otpExpiry, existing.row[9] || new Date()]]);
  } else {
    var sheet = getUsersSheet();
    sheet.appendRow([email, hash, salt, 'user', 'otp_pending', otp, otpExpiry, '', '', new Date()]);
  }

  sendOtpEmail(email, otp);
  return jsonOutput({ status: 'ok', message: 'تم إرسال رمز التحقق (OTP) إلى بريدك الإلكتروني' });
}

function sendOtpEmail(email, otp) {
  MailApp.sendEmail({
    to: email,
    subject: 'رمز التحقق - نظام سكانر أدوية رصد',
    body: 'رمز التحقق الخاص بك هو: ' + otp + '\nصالح لمدة ' + OTP_VALID_MINUTES + ' دقائق.\n\nلو معملتش الطلب ده، تجاهل الرسالة.'
  });
}

function handleResendOtp(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var existing = findUserRow(email);
  if (!existing || existing.row[4] !== 'otp_pending') {
    return jsonOutput({ status: 'error', message: 'لا يوجد طلب تحقق بانتظار لهذا البريد' });
  }
  var otp = generateOtp();
  var otpExpiry = new Date(Date.now() + OTP_VALID_MINUTES * 60000);
  existing.sheet.getRange(existing.rowIndex, 6, 1, 2).setValues([[otp, otpExpiry]]);
  sendOtpEmail(email, otp);
  return jsonOutput({ status: 'ok', message: 'تم إرسال رمز جديد' });
}

function handleVerifyOtp(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var otp = String(body.otp || '').trim();
  var existing = findUserRow(email);
  if (!existing) return jsonOutput({ status: 'error', message: 'الحساب غير موجود' });
  if (existing.row[4] !== 'otp_pending') {
    return jsonOutput({ status: 'error', message: 'لا يوجد طلب تحقق بانتظار لهذا البريد' });
  }
  if (String(existing.row[5]) !== otp) {
    return jsonOutput({ status: 'error', message: 'رمز التحقق غير صحيح' });
  }
  if (new Date(existing.row[6]) < new Date()) {
    return jsonOutput({ status: 'error', message: 'انتهت صلاحية الرمز، اطلب رمز جديد' });
  }
  existing.sheet.getRange(existing.rowIndex, 5, 1, 3).setValues([['approval_pending', '', '']]);
  notifyAdminNewRequest(email);
  return jsonOutput({ status: 'ok', message: 'تم تأكيد بريدك الإلكتروني، طلبك الآن بانتظار موافقة الأدمن' });
}

function notifyAdminNewRequest(email) {
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: 'طلب تسجيل جديد - نظام سكانر رصد',
      body: 'فيه طلب تسجيل جديد بانتظار موافقتك من: ' + email + '\nادخل على لوحة الإدارة في النظام للموافقة أو الرفض.'
    });
  } catch (err) { /* تجاهل لو فشل الإرسال */ }
}

function notifyUserApproved(email) {
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'تم تفعيل حسابك - نظام سكانر رصد',
      body: 'تم قبول طلب تسجيلك، تقدر تسجّل دخولك دلوقتي.'
    });
  } catch (err) { /* تجاهل */ }
}

// ============ تسجيل الدخول ============
function handleLogin(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var password = String(body.password || '');
  var existing = findUserRow(email);
  if (!existing) return jsonOutput({ status: 'error', message: 'الحساب غير موجود' });

  var hash = hashPassword(password, existing.row[2]);
  if (hash !== existing.row[1]) {
    return jsonOutput({ status: 'error', message: 'كلمة السر غير صحيحة' });
  }

  var status = existing.row[4];
  if (status === 'otp_pending') return jsonOutput({ status: 'error', message: 'من فضلك أكمل تأكيد رمز OTP أولاً' });
  if (status === 'approval_pending') return jsonOutput({ status: 'error', message: 'حسابك بانتظار موافقة الأدمن' });
  if (status === 'disabled') return jsonOutput({ status: 'error', message: 'تم إيقاف هذا الحساب، تواصل مع الأدمن' });
  if (status === 'rejected') return jsonOutput({ status: 'error', message: 'تم رفض طلب التسجيل، تواصل مع الأدمن' });
  if (status !== 'active') return jsonOutput({ status: 'error', message: 'حالة الحساب غير معروفة' });

  var token = generateToken();
  var tokenExpiry = new Date(Date.now() + TOKEN_VALID_HOURS * 3600000);
  existing.sheet.getRange(existing.rowIndex, 8, 1, 2).setValues([[token, tokenExpiry]]);

  return jsonOutput({ status: 'ok', token: token, email: email, role: existing.row[3] });
}

function validateToken(email, token) {
  var existing = findUserRow(email);
  if (!existing) return { valid: false };
  if (existing.row[4] !== 'active') return { valid: false };
  if (String(existing.row[7]) !== String(token) || !token) return { valid: false };
  if (new Date(existing.row[8]) < new Date()) return { valid: false };
  return { valid: true, role: existing.row[3], rowIndex: existing.rowIndex, sheet: existing.sheet };
}

function handleChangePassword(body) {
  var auth = validateToken(body.email, body.token);
  if (!auth.valid) return jsonOutput({ status: 'error', message: 'يجب تسجيل الدخول أولاً' });
  var newPassword = String(body.newPassword || '');
  if (newPassword.length < 6) return jsonOutput({ status: 'error', message: 'كلمة السر الجديدة قصيرة جدًا (6 أحرف على الأقل)' });
  var salt = generateSalt();
  var hash = hashPassword(newPassword, salt);
  auth.sheet.getRange(auth.rowIndex, 2, 1, 2).setValues([[hash, salt]]);
  return jsonOutput({ status: 'ok', message: 'تم تغيير كلمة السر بنجاح' });
}

// ============ تسجيل السكان ============
function handleScan(body) {
  var auth = validateToken(body.email, body.token);
  if (!auth.valid) return jsonOutput({ status: 'error', message: 'يجب تسجيل الدخول أولاً' });
  if (!body.gtin || !body.sn || !body.bn || !body.xd) {
    return jsonOutput({ status: 'error', message: 'بيانات ناقصة (GTIN/SN/BN/XD)' });
  }
  var sheet = getDataSheet();
  var gtin = padGtin(body.gtin);
  sheet.appendRow([gtin, body.sn, body.bn, body.xd, body.email, new Date(), '']);
  return jsonOutput({ status: 'ok' });
}

// ============ لوحة الإدارة ============
function requireAdmin(body) {
  var auth = validateToken(body.email, body.token);
  if (!auth.valid || auth.role !== 'admin') return null;
  return auth;
}

function handleAdminListUsers(body) {
  if (!requireAdmin(body)) return jsonOutput({ status: 'error', message: 'غير مصرح' });
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      email: data[i][0],
      role: data[i][3],
      status: data[i][4],
      createdAt: data[i][9]
    });
  }
  return jsonOutput({ status: 'ok', users: list });
}

function handleAdminApprove(body) {
  if (!requireAdmin(body)) return jsonOutput({ status: 'error', message: 'غير مصرح' });
  var target = findUserRow(body.targetEmail);
  if (!target) return jsonOutput({ status: 'error', message: 'المستخدم غير موجود' });
  target.sheet.getRange(target.rowIndex, 5).setValue('active');
  notifyUserApproved(target.row[0]);
  return jsonOutput({ status: 'ok' });
}

function handleAdminReject(body) {
  if (!requireAdmin(body)) return jsonOutput({ status: 'error', message: 'غير مصرح' });
  var target = findUserRow(body.targetEmail);
  if (!target) return jsonOutput({ status: 'error', message: 'المستخدم غير موجود' });
  target.sheet.getRange(target.rowIndex, 5).setValue('rejected');
  return jsonOutput({ status: 'ok' });
}

function handleAdminToggleStatus(body) {
  if (!requireAdmin(body)) return jsonOutput({ status: 'error', message: 'غير مصرح' });
  var target = findUserRow(body.targetEmail);
  if (!target) return jsonOutput({ status: 'error', message: 'المستخدم غير موجود' });
  if (String(target.row[0]).toLowerCase() === ADMIN_EMAIL.toLowerCase()) {
    return jsonOutput({ status: 'error', message: 'لا يمكن إيقاف حساب الأدمن الرئيسي' });
  }
  var current = target.row[4];
  var next = current === 'active' ? 'disabled' : 'active';
  target.sheet.getRange(target.rowIndex, 5).setValue(next);
  return jsonOutput({ status: 'ok', newStatus: next });
}

// ============ التصدير لهيئة رصد (مرة واحدة فقط لكل صنف) ============
function performExport(triggeredByEmail) {
  var dataSheet = getDataSheet();
  var values = dataSheet.getDataRange().getValues();
  if (values.length < 2) return { count: 0 };

  var header = values[0];
  var exportedCol = header.indexOf('Exported');
  var rowsToExport = [];
  var rowIndexes = [];

  for (var i = 1; i < values.length; i++) {
    if (!values[i][exportedCol]) {
      rowsToExport.push(values[i]);
      rowIndexes.push(i + 1);
    }
  }
  if (rowsToExport.length === 0) return { count: 0 };

  var lines = ['GTIN;SN;BN;XD'];
  rowsToExport.forEach(function (r) { lines.push([padGtin(r[0]), r[1], r[2], r[3]].join(';')); });
  // بدون سطر فارغ في آخر الملف — لازم يطابق شكل نموذج هيئة رصد الرسمي بالظبط
  var csvContent = lines.join('\r\n').replace(/[\r\n]+$/, '');

  // الصيغة الثانية المقبولة من رصد: تجميع حسب الدفعة (GTIN;QUANTITY;BN;XD)
  var groupedCsvContent = buildGroupedCsv(rowsToExport.map(function (r) {
    return { gtin: padGtin(r[0]), bn: r[2], xd: r[3] };
  }));

  var fileName = 'rasd_export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.csv';
  var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);

  var now = new Date();
  var reportSheet = getReportSheet();
  rowsToExport.forEach(function (r) {
    reportSheet.appendRow([padGtin(r[0]), r[1], r[2], r[3], now, triggeredByEmail || '-', fileName]);
  });
  rowIndexes.forEach(function (rowIdx) {
    dataSheet.getRange(rowIdx, exportedCol + 1).setValue(now);
  });

  return {
    count: rowsToExport.length,
    fileUrl: file.getUrl(),
    fileName: fileName,
    csvContent: csvContent,
    groupedCsvContent: groupedCsvContent
  };
}

// يبني صيغة "مجمّع حسب الدفعة" (GTIN;QUANTITY;BN;XD) — كل صنف+دفعة+تاريخ صلاحية يتحسب مرة واحدة مع عدّاد الكمية
function buildGroupedCsv(items) {
  var groups = {};
  var order = [];
  items.forEach(function (item) {
    var key = item.gtin + '|' + item.bn + '|' + item.xd;
    if (!groups[key]) { groups[key] = { gtin: item.gtin, bn: item.bn, xd: item.xd, qty: 0 }; order.push(key); }
    groups[key].qty++;
  });
  var lines = ['GTIN;QUANTITY;BN;XD'];
  order.forEach(function (key) {
    var g = groups[key];
    lines.push([g.gtin, g.qty, g.bn, g.xd].join(';'));
  });
  return lines.join('\r\n').replace(/[\r\n]+$/, '');
}

function handleExportRegulator(body) {
  var auth = validateToken(body.email, body.token);
  if (!auth.valid) return jsonOutput({ status: 'error', message: 'يجب تسجيل الدخول أولاً' });

  var result = performExport(body.email);
  if (result.count === 0) {
    return jsonOutput({ status: 'error', message: 'كل الأصناف تم تصديرها بالفعل، لا يوجد جديد لتصديره' });
  }
  return jsonOutput({
    status: 'ok',
    count: result.count,
    fileUrl: result.fileUrl,
    fileName: result.fileName,
    csvContent: result.csvContent,
    groupedCsvContent: result.groupedCsvContent
  });
}

function handleGetExportLog(body) {
  var auth = validateToken(body.email, body.token);
  if (!auth.valid) return jsonOutput({ status: 'error', message: 'يجب تسجيل الدخول أولاً' });

  var sheet = getReportSheet();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      gtin: padGtin(data[i][0]),
      sn: data[i][1],
      bn: data[i][2],
      xd: data[i][3],
      exportedAt: data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(data[i][4]),
      exportedBy: data[i][5],
      fileName: data[i][6]
    });
  }
  rows.reverse(); // الأحدث أولاً
  return jsonOutput({ status: 'ok', rows: rows });
}

// ============ قائمة الشيت (تصدير يدوي بديل من داخل جوجل شيت) ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('تصدير هيئة رصد')
    .addItem('تصدير CSV بصيغة رصد (الأصناف الجديدة فقط)', 'exportRegulatorCsvFromMenu')
    .addToUi();
}

function exportRegulatorCsvFromMenu() {
  var result = performExport(Session.getActiveUser().getEmail() || 'sheet-menu');
  if (result.count === 0) {
    SpreadsheetApp.getUi().alert('كل الأصناف تم تصديرها بالفعل، مفيش أصناف جديدة.');
    return;
  }
  SpreadsheetApp.getUi().alert(
    'تم تصدير ' + result.count + ' صنف بنجاح:\n' + result.fileName +
    '\n\nهتلاقيه في Google Drive (My Drive) وجاهز يترفع لهيئة رصد.\n' +
    'التفاصيل مسجلة كمان في شيت "' + REPORT_SHEET_NAME + '".'
  );
}

// ============ إعداد حساب الأدمن (يتشغّل مرة واحدة يدويًا من محرر الأكواد) ============
function setupAdmin() {
  // 1) غيّر السطر اللي تحت ده وحط باسوردك الحقيقي بدل النص ده
  var ADMIN_PASSWORD_PLACEHOLDER = 'اكتب_الباسورد_هنا_مؤقتًا';

  if (ADMIN_PASSWORD_PLACEHOLDER === 'اكتب_الباسورد_هنا_مؤقتًا') {
    Logger.log('لازم تغيّر قيمة ADMIN_PASSWORD_PLACEHOLDER الأول قبل التشغيل.');
    return;
  }

  var existing = findUserRow(ADMIN_EMAIL);
  if (existing) {
    Logger.log('حساب الأدمن موجود بالفعل: ' + ADMIN_EMAIL);
    return;
  }

  var salt = generateSalt();
  var hash = hashPassword(ADMIN_PASSWORD_PLACEHOLDER, salt);
  var sheet = getUsersSheet();
  sheet.appendRow([ADMIN_EMAIL, hash, salt, 'admin', 'active', '', '', '', '', new Date()]);
  Logger.log('تم إنشاء حساب الأدمن بنجاح: ' + ADMIN_EMAIL);
  Logger.log('دلوقتي امسح الباسورد اللي كتبته فوق في ADMIN_PASSWORD_PLACEHOLDER وارجّعها زي ما كانت، واحفظ الملف.');
}
