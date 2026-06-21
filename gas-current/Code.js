const CHECKIN_SHEET = 'CheckinData';
const USERS_SHEET = 'Users';
const SETTINGS_SHEET = 'Settings';
const POSITION_SETTINGS_SHEET = 'PositionSettings';
const PDF_LOGS_SHEET = 'PdfLogs';
const PERMISSION_REQUESTS_SHEET = 'PermissionRequests';
const CHANNEL_ACCESS_TOKEN = 'gTwyyibbdpRljPfvdL0pZDK+WrUWpZXPHIC8uCoOutkRafIqOEYxTcSBv4MuOKaSALYfALhggNY8Lz6HFucbg5byKFqvGi2KdGYpHX6gEPU7zvbrTSKjSdXS6eqYaWFN4baFCvIgNBu5w3QjIlu5EQdB04t89/1O/w1cDnyilFU=';
const GROUP_ID = 'C3ef221ea7a454eca09401e333eb1590a';
const QR_TOKEN = 'school2026';
const INDEX_FILE = 'Index';
const DASHBOARD_FILE = 'Dashboard';
const REPORT_TEMPLATE_DOC_ID = '1XFEiaz3xRKVVXqQFkXpGk_Ts7oxEajElQe4VRkKvql0';
const LEAVE_TEMPLATE_DOC_ID = '1ZRMTGxVzPg4EtDn39lUaVDGin0I_xAuhTX9STH1bXqI';
const REPORT_FOLDER_NAME = 'รายงานลงเวลาปฏิบัติงาน';

const CHECKIN_HEADERS = [
  'Timestamp',
  'ชื่อ',
  'เบอร์โทร',
  'ตำแหน่ง',
  'Latitude',
  'Longitude',
  'ระยะเมตร',
  'เวลาเข้า',
  'เวลาออกอัตโนมัติ',
  'สถานะเวลา',
  'สถานะอนุมัติ',
  'เหตุผลมาสาย',
  'ผู้อนุมัติ',
  'เวลาอนุมัติ',
  'หมายเหตุอนุมัติ',
  'TimeStatusCode',
  'ApprovalStatusCode'
];

const USERS_HEADERS = [
  'ชื่อ',
  'เบอร์โทร',
  'ตำแหน่ง',
  'Role',
  'สถานะบัญชี',
  'ProfilePhotoUrl',
  'SignatureUrl',
  'ProfilePhotoFileId',
  'SignatureFileId',
  'PinHash',
  'PinSalt',
  'LastLoginAt'
];

const SETTINGS_HEADERS = [
  'Key',
  'Value',
  'คำอธิบาย'
];

const POSITION_HEADERS = [
  'ตำแหน่ง',
  'เวลาออกอัตโนมัติ'
];

const PDF_LOG_HEADERS = [
  'วันที่รายงาน',
  'เวลาทำงาน',
  'สถานะ',
  'ชื่อไฟล์',
  'ลิงก์ไฟล์',
  'ข้อความ'
];


const PERMISSION_REQUEST_HEADERS = [
  'RequestId',
  'วันที่ส่งคำขอ',
  'ชื่อ',
  'เบอร์โทร',
  'ตำแหน่ง',
  'ประเภทคำขอ',
  'วันที่มีผล',
  'เวลาที่คาดว่าจะถึง',
  'เหตุผล',
  'สถานะอนุมัติ',
  'ผู้อนุมัติ',
  'เวลาอนุมัติ',
  'หมายเหตุ',
  'ApprovalStatusCode',
  'วันที่สิ้นสุด',
  'จำนวนวัน',
  'ลิงก์ PDF ใบลา',
  'หลักฐาน URL',
  'ชื่อไฟล์หลักฐาน',
  'ครั้งที่ลา',
  'เลขที่หนังสือ',
  'ปีการศึกษา',
  'ภาคเรียน',
  'ครั้งลาสะสมภาคเรียน',
  'วันลาสะสมภาคเรียน'
];
const DEFAULT_SETTINGS = {
  SCHOOL_NAME: 'ระบบลงเวลา',
  SCHOOL_LOGO_URL: '',
  SCHOOL_LOGO_DATA: '',
  SCHOOL_LOGO_SIZE_PERCENT: '35',
  APP_VERSION: '1.0.67',
  REPORT_TEMPLATE_DOC_ID: REPORT_TEMPLATE_DOC_ID,
  LEAVE_TEMPLATE_DOC_ID: LEAVE_TEMPLATE_DOC_ID,
  REPORT_FOLDER_ID: '1wObdl_rdUGaQz_9IQT6NX54ln--yB0-r',
  SCHOOL_LAT: '14.473635',
  SCHOOL_LNG: '100.018044',
  ALLOW_RADIUS: '500',
START_TIME: '08:00',
  ACADEMIC_YEAR: '2569',
  CURRENT_TERM: '1',
  TERM_START_DATE: '',
  TERM_END_DATE: '',
  LEAVE_DOC_PREFIX: 'ผอ',
  LEAVE_DOC_START_NO: '1'
};

function doGet(e) {
  initializeSheets();

  const params = e && e.parameter ? e.parameter : {};
  const token = params.token;
  const page = params.page || 'index';

  if (token !== QR_TOKEN) {
    return HtmlService
      .createHtmlOutput('QR Code ไม่ถูกต้อง');
  }

  if (page === 'dashboard') {
    const pageUser = params.user || '';

    if (!canOpenDashboard(pageUser)) {
      return HtmlService
        .createHtmlOutput('ไม่มีสิทธิ์เข้า Dashboard')
        .setTitle('Access denied');
    }

    if (params.view === 'settings' && !canOpenSettings(pageUser)) {
      return HtmlService
        .createHtmlOutput('ไม่มีสิทธิ์เข้าหน้าตั้งค่าระบบ')
        .setTitle('Access denied');
    }

    return HtmlService
      .createTemplateFromFile(DASHBOARD_FILE)
      .evaluate()
      .setTitle('Dashboard ผู้บริหาร')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService
    .createTemplateFromFile(INDEX_FILE)
    .evaluate()
    .setTitle('ระบบลงเวลาปฏิบัติงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let checkSheet = ss.getSheetByName(CHECKIN_SHEET);
  if (!checkSheet) {
    checkSheet = ss.insertSheet(CHECKIN_SHEET);
    checkSheet.appendRow(CHECKIN_HEADERS);
  }
  ensureSheetHeaders(checkSheet, CHECKIN_HEADERS);
  checkSheet.getRange('C:C').setNumberFormat('@');

  let userSheet = ss.getSheetByName(USERS_SHEET);
  if (!userSheet) {
    userSheet = ss.insertSheet(USERS_SHEET);
    userSheet.appendRow(USERS_HEADERS);
  }
  ensureSheetHeaders(userSheet, USERS_HEADERS);
  userSheet.getRange('B:B').setNumberFormat('@');
  userSheet.getRange('J:K').setNumberFormat('@');

  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET);
    settingsSheet.appendRow(SETTINGS_HEADERS);
    settingsSheet.appendRow(['SCHOOL_NAME', DEFAULT_SETTINGS.SCHOOL_NAME, 'ชื่อโรงเรียน']);
    settingsSheet.appendRow(['SCHOOL_LOGO_URL', DEFAULT_SETTINGS.SCHOOL_LOGO_URL, 'URL โลโก้โรงเรียน']);
    settingsSheet.appendRow(['SCHOOL_LOGO_DATA', DEFAULT_SETTINGS.SCHOOL_LOGO_DATA, 'ข้อมูลรูปโลโก้โรงเรียน']);
    settingsSheet.appendRow(['SCHOOL_LOGO_SIZE_PERCENT', DEFAULT_SETTINGS.SCHOOL_LOGO_SIZE_PERCENT, 'ขนาดโลโก้ หน่วยเปอร์เซ็นต์']);
    settingsSheet.appendRow(['APP_VERSION', DEFAULT_SETTINGS.APP_VERSION, 'Version ของแอพ']);
    settingsSheet.appendRow(['REPORT_TEMPLATE_DOC_ID', DEFAULT_SETTINGS.REPORT_TEMPLATE_DOC_ID, 'รหัส Google Docs Template รายงาน']);
    settingsSheet.appendRow(['LEAVE_TEMPLATE_DOC_ID', DEFAULT_SETTINGS.LEAVE_TEMPLATE_DOC_ID, 'รหัส Google Docs Template ใบลา']);
    settingsSheet.appendRow(['REPORT_FOLDER_ID', DEFAULT_SETTINGS.REPORT_FOLDER_ID, 'รหัสโฟลเดอร์เก็บ PDF รายงาน']);
    settingsSheet.appendRow(['SCHOOL_LAT', DEFAULT_SETTINGS.SCHOOL_LAT, 'ละติจูดโรงเรียน']);
    settingsSheet.appendRow(['SCHOOL_LNG', DEFAULT_SETTINGS.SCHOOL_LNG, 'ลองจิจูดโรงเรียน']);
    settingsSheet.appendRow(['ALLOW_RADIUS', DEFAULT_SETTINGS.ALLOW_RADIUS, 'รัศมีที่อนุญาต หน่วยเมตร']);
    settingsSheet.appendRow(['START_TIME', DEFAULT_SETTINGS.START_TIME, 'เวลาเริ่มงาน']);
    settingsSheet.appendRow(['ACADEMIC_YEAR', DEFAULT_SETTINGS.ACADEMIC_YEAR, 'ปีการศึกษาปัจจุบัน']);
    settingsSheet.appendRow(['CURRENT_TERM', DEFAULT_SETTINGS.CURRENT_TERM, 'ภาคเรียนปัจจุบัน']);
    settingsSheet.appendRow(['TERM_START_DATE', DEFAULT_SETTINGS.TERM_START_DATE, 'วันเริ่มภาคเรียน']);
    settingsSheet.appendRow(['TERM_END_DATE', DEFAULT_SETTINGS.TERM_END_DATE, 'วันสิ้นสุดภาคเรียน']);
    settingsSheet.appendRow(['LEAVE_DOC_PREFIX', DEFAULT_SETTINGS.LEAVE_DOC_PREFIX, 'คำนำหน้าเลขที่หนังสือใบลา']);
    settingsSheet.appendRow(['LEAVE_DOC_START_NO', DEFAULT_SETTINGS.LEAVE_DOC_START_NO, 'เลขเริ่มต้นใบลา']);
  }
  ensureSheetHeaders(settingsSheet, SETTINGS_HEADERS);

  let positionSheet = ss.getSheetByName(POSITION_SETTINGS_SHEET);
  if (!positionSheet) {
    positionSheet = ss.insertSheet(POSITION_SETTINGS_SHEET);
    positionSheet.appendRow(POSITION_HEADERS);
    positionSheet.appendRow(['ครู', '16:30']);
    positionSheet.appendRow(['ภารโรง', '18:00']);
    positionSheet.appendRow(['เจ้าหน้าที่', '16:30']);
    positionSheet.appendRow(['ผู้บริหาร', '16:30']);
  }
  ensureSheetHeaders(positionSheet, POSITION_HEADERS);

  positionSheet.getRange('B:B').setNumberFormat('@');

  let pdfLogSheet = ss.getSheetByName(PDF_LOGS_SHEET);
  if (!pdfLogSheet) {
    pdfLogSheet = ss.insertSheet(PDF_LOGS_SHEET);
    pdfLogSheet.appendRow(PDF_LOG_HEADERS);
  }
  ensureSheetHeaders(pdfLogSheet, PDF_LOG_HEADERS);
  let permissionSheet = ss.getSheetByName(PERMISSION_REQUESTS_SHEET);
  if (!permissionSheet) {
    permissionSheet = ss.insertSheet(PERMISSION_REQUESTS_SHEET);
    permissionSheet.appendRow(PERMISSION_REQUEST_HEADERS);
  }
  ensureSheetHeaders(permissionSheet, PERMISSION_REQUEST_HEADERS);
  permissionSheet.getRange('D:D').setNumberFormat('@');
  permissionSheet.getRange('G:H').setNumberFormat('@');
  permissionSheet.getRange('O:P').setNumberFormat('@');
  permissionSheet.getRange('R:S').setNumberFormat('@');
  permissionSheet.getRange('U:V').setNumberFormat('@');
}

function ensureSheetHeaders(sheet, headers) {
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    return;
  }

  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];

  headers.forEach((header, index) => {
    if (!String(current[index] || '').trim()) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
}

function checkSystemSheetStructure() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = [
    {
      sheetName: CHECKIN_SHEET,
      headers: CHECKIN_HEADERS
    },
    {
      sheetName: USERS_SHEET,
      headers: USERS_HEADERS
    },
    {
      sheetName: SETTINGS_SHEET,
      headers: SETTINGS_HEADERS
    },
    {
      sheetName: POSITION_SETTINGS_SHEET,
      headers: POSITION_HEADERS
    },
    {
      sheetName: PDF_LOGS_SHEET,
      headers: PDF_LOG_HEADERS
    }
  ];

  const result = checks.map((item) => {
    const sheet = ss.getSheetByName(item.sheetName);

    if (!sheet) {
      return {
        sheet: item.sheetName,
        ok: false,
        message: 'ไม่พบชีต'
      };
    }

    const actual = sheet
      .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), item.headers.length))
      .getValues()[0]
      .map((value) => String(value || '').trim());

    const issues = [];

    item.headers.forEach((expected, index) => {
      const got = actual[index] || '';

      if (got !== expected) {
        issues.push({
          column: index + 1,
          expected: expected,
          actual: got
        });
      }
    });

    return {
      sheet: item.sheetName,
      ok: issues.length === 0,
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn(),
      issues: issues
    };
  });

  return {
    success: true,
    ok: result.every((item) => item.ok),
    result: result
  };
}

function repairSystemSheetHeaders() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};

  map[CHECKIN_SHEET] = CHECKIN_HEADERS;
  map[USERS_SHEET] = USERS_HEADERS;
  map[SETTINGS_SHEET] = SETTINGS_HEADERS;
  map[POSITION_SETTINGS_SHEET] = POSITION_HEADERS;
  map[PDF_LOGS_SHEET] = PDF_LOG_HEADERS;
  map[PERMISSION_REQUESTS_SHEET] = PERMISSION_REQUEST_HEADERS;

  Object.keys(map).forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      sheet.getRange(1, 1, 1, map[sheetName].length).setValues([map[sheetName]]);
    }
  });

  return checkSystemSheetStructure();
}

function clearTestData() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [
    {
      sheetName: CHECKIN_SHEET,
      headers: CHECKIN_HEADERS
    },
    {
      sheetName: PDF_LOGS_SHEET,
      headers: PDF_LOG_HEADERS
    }
  ];
  const result = [];

  targets.forEach((item) => {
    let sheet = ss.getSheetByName(item.sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(item.sheetName);
    }

    ensureSheetHeaders(sheet, item.headers);

    const lastRow = sheet.getLastRow();
    const lastColumn = Math.max(sheet.getLastColumn(), item.headers.length);
    const deletedRows = Math.max(lastRow - 1, 0);

    if (deletedRows > 0) {
      sheet.getRange(2, 1, deletedRows, lastColumn).clearContent();
    }

    if (item.sheetName === CHECKIN_SHEET) {
      sheet.getRange('C:C').setNumberFormat('@');
    }

    result.push({
      sheet: item.sheetName,
      clearedRows: deletedRows
    });
  });

  return {
    success: true,
    message: 'ล้างข้อมูลทดสอบเรียบร้อยแล้ว โดยคง Users, Settings และ PositionSettings ไว้',
    result: result,
    updated: result.reduce((sum, item) => sum + item.clearedRows, 0)
  };
}
function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function doPost(e) {

  try {

    const data = JSON.parse(e.postData.contents);

    const groupId =
      data.events[0].source.groupId;

    // บันทึกลงชีตแทน Logger
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheet =
      ss.getSheetByName('LineLogs');

    if (!sheet) {

      sheet = ss.insertSheet('LineLogs');

      sheet.appendRow([
        'Time',
        'Group ID',
        'Raw JSON'
      ]);
    }

    sheet.appendRow([
      new Date(),
      groupId,
      JSON.stringify(data)
    ]);

    return ContentService
      .createTextOutput('OK');

  } catch(err) {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    let sheet =
      ss.getSheetByName('LineLogs');

    if (!sheet) {

      sheet = ss.insertSheet('LineLogs');
    }

    sheet.appendRow([
      new Date(),
      'ERROR',
      err.toString()
    ]);

    return ContentService
      .createTextOutput('ERROR');
  }
}

function registerUser(data) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);

  const name = String(data.name || '').trim();
  const phone = normalizePhone(data.phone);
  const position = String(data.position || '').trim();
  const profilePhoto = data.profilePhoto || null;
  const signature = data.signature || null;
  const pin = String(data.pin || '').trim();

  if (!name || !phone || !position || !isValidPin(pin)) {
    return {
      success: false,
      message: 'กรุณากรอกชื่อ เบอร์โทร ตำแหน่ง และ PIN 4 หลัก'
    };
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (isSamePhone(values[i][1], phone)) {
      return {
        success: false,
        message: 'เบอร์นี้ถูกใช้งานแล้ว'
      };
    }
  }

  const profileFile = saveUserProfileFile(profilePhoto, phone, name, 'รูปประจำตัว');
  const signatureFile = saveUserProfileFile(signature, phone, name, 'ลายเซ็น');
  const pinSalt = createPinSalt();
  const pinHash = hashPin(pin, pinSalt);

  sheet.appendRow([
    name,
    formatPhoneForSheet(phone),
    position,
    'user',
    'active',
    profileFile ? profileFile.url : '',
    signatureFile ? signatureFile.url : '',
    profileFile ? profileFile.id : '',
    signatureFile ? signatureFile.id : '',
    pinHash,
    pinSalt,
    ''
  ]);

  return {
    success: true,
    message: 'สมัครสมาชิกสำเร็จ'
  };
}

function loginUser(phone, pin) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);

  const loginPhone = normalizePhone(phone);
  const loginPin = String(pin || '').trim();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (isSamePhone(values[i][1], loginPhone)) {
      const accountStatus = String(values[i][4] || '').trim();
      const pinHash = String(values[i][9] || '').trim();
      const pinSalt = String(values[i][10] || '').trim();

      if (accountStatus && accountStatus !== 'active') {
        return {
          success: false,
          message: 'บัญชีนี้ยังไม่พร้อมใช้งาน กรุณาติดต่อผู้ดูแลระบบ'
        };
      }

      if (pinHash && pinSalt) {
        if (!isValidPin(loginPin) || hashPin(loginPin, pinSalt) !== pinHash) {
          return {
            success: false,
            message: 'เบอร์โทรหรือ PIN ไม่ถูกต้อง'
          };
        }
      }

      sheet.getRange(i + 1, 12).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'));

      const user = {
        name: String(values[i][0] || ''),
        phone: normalizePhone(loginPhone),
        position: String(values[i][2] || ''),
        role: String(values[i][3] || ''),
        accountStatus: accountStatus,
        profilePhotoUrl: String(values[i][7] || '') ? getDriveThumbnailUrl(String(values[i][7] || '')) : String(values[i][5] || ''),
        signatureUrl: String(values[i][6] || ''),
        profilePhotoFileId: String(values[i][7] || ''),
        signatureFileId: String(values[i][8] || ''),
        requirePinSetup: !(pinHash && pinSalt)
      };

      return {
        success: true,
        name: user.name,
        phone: user.phone,
        position: user.position,
        role: user.role,
        accountStatus: user.accountStatus,
        profilePhotoUrl: user.profilePhotoUrl,
        signatureUrl: user.signatureUrl,
        profilePhotoFileId: user.profilePhotoFileId,
        signatureFileId: user.signatureFileId,
        requirePinSetup: user.requirePinSetup,
        todayCheckin: checkTodayCheckinForUser(user)
      };
    }
  }

  return {
    success: false,
    message: 'ไม่พบข้อมูลผู้ใช้'
  };
}

function isValidPin(pin) {
  return /^\d{4}$/.test(String(pin || '').trim());
}

function createPinSalt() {
  return Utilities.getUuid() + '-' + Utilities.getUuid();
}

function hashPin(pin, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt || '') + ':' + String(pin || ''),
    Utilities.Charset.UTF_8
  );

  return raw
    .map((byte) => {
      const value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    })
    .join('');
}

function setUserPin(data) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  const phone = normalizePhone(data && data.phone);
  const newPin = String(data && data.newPin || '').trim();
  const currentPin = String(data && data.currentPin || '').trim();

  if (!sheet || !phone || sheet.getLastRow() < 2) {
    return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
  }

  if (!isValidPin(newPin)) {
    return { success: false, message: 'PIN ต้องเป็นตัวเลข 4 หลัก' };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), USERS_HEADERS.length)).getValues();

  for (let i = 0; i < values.length; i++) {
    if (!isSamePhone(values[i][1], phone)) {
      continue;
    }

    const rowNumber = i + 2;
    const oldHash = String(values[i][9] || '').trim();
    const oldSalt = String(values[i][10] || '').trim();

    if (oldHash && oldSalt) {
      if (!isValidPin(currentPin) || hashPin(currentPin, oldSalt) !== oldHash) {
        return { success: false, message: 'PIN เดิมไม่ถูกต้อง' };
      }
    }

    const salt = createPinSalt();
    const hash = hashPin(newPin, salt);

    sheet.getRange(rowNumber, 10).setValue(hash);
    sheet.getRange(rowNumber, 11).setValue(salt);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: oldHash && oldSalt ? 'เปลี่ยน PIN เรียบร้อยแล้ว' : 'ตั้งค่า PIN เรียบร้อยแล้ว'
    };
  }

  return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
}

function resetUserPinByAdmin(data) {
  initializeSheets();

  const actorPhone = normalizePhone(data && data.actorPhone);
  const targetPhone = normalizePhone(data && data.targetPhone);
  const newPin = String(data && data.newPin || '').trim();

  if (!isValidPin(newPin)) {
    return { success: false, message: 'PIN ใหม่ต้องเป็นตัวเลข 4 หลัก' };
  }

  const actor = getUserProfileWithRoleByPhone(actorPhone);

  if (!actor || (actor.role !== 'admin' && actor.role !== 'director')) {
    return { success: false, message: 'เฉพาะผู้ดูแลระบบหรือผู้บริหารเท่านั้นที่รีเซ็ต PIN ได้' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);

  if (!sheet || !targetPhone || sheet.getLastRow() < 2) {
    return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), USERS_HEADERS.length)).getValues();

  for (let i = 0; i < values.length; i++) {
    if (!isSamePhone(values[i][1], targetPhone)) {
      continue;
    }

    const salt = createPinSalt();
    const hash = hashPin(newPin, salt);

    sheet.getRange(i + 2, 10).setValue(hash);
    sheet.getRange(i + 2, 11).setValue(salt);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'รีเซ็ต PIN เรียบร้อยแล้ว'
    };
  }

  return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
}

function getUserProfileWithRoleByPhone(phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  const targetPhone = normalizePhone(phone);

  if (!sheet || !targetPhone || sheet.getLastRow() < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), USERS_HEADERS.length)).getValues();

  for (let i = 0; i < values.length; i++) {
    if (isSamePhone(values[i][1], targetPhone)) {
      return {
        name: String(values[i][0] || ''),
        phone: normalizePhone(values[i][1]),
        position: String(values[i][2] || ''),
        role: String(values[i][3] || '')
      };
    }
  }

  return null;
}

function saveUserProfileFile(attachment, phone, name, label) {
  if (!attachment || !attachment.data) {
    return null;
  }

  const base64 = String(attachment.data || '').replace(/^data:[^,]+,/, '');
  const originalName = sanitizeFileName(attachment.name || label || 'profile');
  const contentType = String(attachment.type || 'application/octet-stream');

  if (!contentType.includes('image/')) {
    throw new Error(label + ' ต้องเป็นไฟล์รูปภาพเท่านั้น');
  }

  const bytes = Utilities.base64Decode(base64);

  if (bytes.length > 2 * 1024 * 1024) {
    throw new Error(label + ' มีขนาดใหญ่เกิน 2 MB');
  }

  const settings = getSettings();
  const rootFolder = getReportFolder(settings);
  const profileFolder = getOrCreateChildFolder(rootFolder, 'ข้อมูลโปรไฟล์บุคลากร');
  const personFolder = getOrCreateChildFolder(profileFolder, sanitizeFileName(name || phone || 'user'));
  const extension = getFileExtension(originalName, contentType) || '.png';
  const filename = sanitizeFileName(label) + '_' + sanitizeFileName(name || phone) + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMddHHmmss') + extension;
  const blob = Utilities.newBlob(bytes, contentType, filename);
  const file = personFolder.createFile(blob);

  return {
    id: file.getId(),
    url: getDriveThumbnailUrl(file.getId()),
    name: filename
  };
}

function getDriveThumbnailUrl(fileId) {
  return fileId
    ? 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w240'
    : '';
}

function getUserProfileByPhone(phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  const targetPhone = normalizePhone(phone);

  if (!sheet || !targetPhone || sheet.getLastRow() < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), USERS_HEADERS.length)).getValues();

  for (let i = 0; i < values.length; i++) {
    if (isSamePhone(values[i][1], targetPhone)) {
      return {
        name: String(values[i][0] || ''),
        profilePhotoUrl: String(values[i][7] || '') ? getDriveThumbnailUrl(String(values[i][7] || '')) : String(values[i][5] || ''),
        signatureUrl: String(values[i][6] || ''),
        profilePhotoFileId: String(values[i][7] || ''),
        signatureFileId: String(values[i][8] || '')
      };
    }
  }

  return null;
}

function getApproverSignatureProfile(approverName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  const targetName = normalizeName(approverName);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), USERS_HEADERS.length)).getValues();
  let directorProfile = null;

  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0] || '');
    const role = String(values[i][3] || '').trim();
    const accountStatus = String(values[i][4] || '').trim();
    const signatureFileId = String(values[i][8] || '').trim();

    if (!signatureFileId || (accountStatus && accountStatus !== 'active')) {
      continue;
    }

    const profile = {
      name: name,
      signatureUrl: String(values[i][6] || ''),
      signatureFileId: signatureFileId,
      role: role
    };

    if (targetName && normalizeName(name) === targetName) {
      return profile;
    }

    if (!directorProfile && role === 'director') {
      directorProfile = profile;
    }
  }

  return directorProfile;
}

function updateUserProfile(data) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  const phone = normalizePhone(data && data.phone);

  if (!sheet || !phone || sheet.getLastRow() < 2) {
    return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), USERS_HEADERS.length)).getValues();
  let rowNumber = -1;
  let rowData = null;

  for (let i = 0; i < values.length; i++) {
    if (isSamePhone(values[i][1], phone)) {
      rowNumber = i + 2;
      rowData = values[i];
      break;
    }
  }

  if (rowNumber < 2 || !rowData) {
    return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
  }

  const name = String(rowData[0] || '').trim();
  const removeProfilePhoto = Boolean(data && data.removeProfilePhoto);
  const removeSignature = Boolean(data && data.removeSignature);
  const profileFile = saveUserProfileFile(data.profilePhoto || null, phone, name, 'รูปประจำตัว');
  const signatureFile = saveUserProfileFile(data.signature || null, phone, name, 'ลายเซ็น');

  if (removeProfilePhoto) {
    trashDriveFileIfPossible(rowData[7]);
    sheet.getRange(rowNumber, 6).clearContent();
    sheet.getRange(rowNumber, 8).clearContent();
  }

  if (removeSignature) {
    trashDriveFileIfPossible(rowData[8]);
    sheet.getRange(rowNumber, 7).clearContent();
    sheet.getRange(rowNumber, 9).clearContent();
  }

  if (profileFile) {
    sheet.getRange(rowNumber, 6).setValue(profileFile.url);
    sheet.getRange(rowNumber, 8).setValue(profileFile.id);
  }

  if (signatureFile) {
    sheet.getRange(rowNumber, 7).setValue(signatureFile.url);
    sheet.getRange(rowNumber, 9).setValue(signatureFile.id);
  }

  SpreadsheetApp.flush();

  const profile = getUserProfileByPhone(phone) || {};

  return {
    success: true,
    message: 'อัปเดตโปรไฟล์เรียบร้อยแล้ว',
    profilePhotoUrl: profile.profilePhotoUrl || '',
    signatureUrl: profile.signatureUrl || '',
    profilePhotoFileId: profile.profilePhotoFileId || '',
    signatureFileId: profile.signatureFileId || ''
  };
}

function trashDriveFileIfPossible(fileId) {
  const id = String(fileId || '').trim();

  if (!id) {
    return;
  }

  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (err) {
    Logger.log('Cannot trash profile file: ' + (err && err.message ? err.message : err));
  }
}

function saveCheckin(data) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CHECKIN_SHEET);

  const settings = getSettings();

  const schoolLat = Number(settings.SCHOOL_LAT);
  const schoolLng = Number(settings.SCHOOL_LNG);
  const allowRadius = Number(settings.ALLOW_RADIUS);
  const startTime = settings.START_TIME;

  const phone = normalizePhone(data.phone);
  const name = String(data.name || '').trim();
  const position = String(data.position || '').trim();

  const checked = checkTodayCheckinForUser({
    phone: phone,
    name: name
  });
  if (checked.checkedIn) {
    return {
      success: false,
      message: 'วันนี้คุณลงเวลาแล้ว',
      status: checked.timeStatus,
      timeStatus: checked.timeStatus,
      approvalStatus: checked.approvalStatus,
      time: checked.time
    };
  }

  const distance = calculateDistance(
    schoolLat,
    schoolLng,
    Number(data.lat),
    Number(data.lng)
  );

  if (distance > allowRadius) {
    return {
      success: false,
      message: 'คุณอยู่นอกพื้นที่โรงเรียน'
    };
  }

  const now = new Date();
  const checkinTime = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');
  const currentMinutes = timeToMinutes(checkinTime);
  const startMinutes = timeToMinutes(startTime);

  const isLate = currentMinutes > startMinutes;

  if (isLate && !data.reason && !data.submitWithoutReason) {
    return {
      success: false,
      requireLateReason: true,
      message: 'คุณมาสาย กรุณาระบุเหตุผล'
    };
  }

  let timeStatus = '🟢 ปกติ';
  let approvalStatus = 'ไม่ต้องอนุมัติ';
  let lateReason = '-';
  let timeStatusCode = 'normal';
  let approvalStatusCode = 'none';

  if (isLate) {
    timeStatus = '🔴 มาสาย';
    timeStatusCode = 'late';

    if (data.reason) {
      approvalStatus = '🟡 รออนุมัติ';
      approvalStatusCode = 'pending';
      lateReason = String(data.reason).trim();
    } else {
      approvalStatus = '🔴 มาสายโดยไม่ได้ขอ';
      approvalStatusCode = 'unrequested';
      lateReason = 'ไม่ได้ระบุเหตุผล';
    }
  }

  const autoCheckoutTime = getAutoCheckoutTime(position);

  sheet.appendRow([
    now,
    name,
    formatPhoneForSheet(phone),
    position,
    data.lat,
    data.lng,
    Math.round(distance),
    checkinTime,
    autoCheckoutTime,
    timeStatus,
    approvalStatus,
    lateReason,
    '',
    '',
    '',
    timeStatusCode,
    approvalStatusCode
  ]);

  notifyLineCheckin({
    checkinDate: now,
    name: name,
    position: position,
    checkinTime: checkinTime,
    autoCheckoutTime: autoCheckoutTime,
    timeStatus: timeStatus,
    approvalStatus: approvalStatus,
    lateReason: lateReason,
    distance: Math.round(distance)
  });

  return {
  success: true,
  message: 'บันทึกเวลาเรียบร้อย',
  timeStatus: timeStatus,
  approvalStatus: approvalStatus,
  timeStatusCode: timeStatusCode,
  approvalStatusCode: approvalStatusCode,
  checkinTime: checkinTime,
  autoCheckoutTime: autoCheckoutTime,
  distance: Math.round(distance)
}
}

function calculateDistance(lat1, lon1, lat2, lon2) {

  const R = 6371000;

  const dLat = (lat2-lat1) * Math.PI / 180;
  const dLon = (lon2-lon1) * Math.PI / 180;

  const a =
    Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1 * Math.PI / 180) *
    Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon/2) * Math.sin(dLon/2);

  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));

  return R * c;
}

function getDashboardData() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checkSheet = ss.getSheetByName(CHECKIN_SHEET);
  const userSheet = ss.getSheetByName(USERS_SHEET);
  const permissionSheet = getPermissionRequestsSheet();
  const today = getBangkokDateKey(new Date());
  const activeUsers = [];

  if (userSheet && userSheet.getLastRow() >= 2) {
    const userValues = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 5).getValues();

    userValues.forEach((row) => {
      const name = String(row[0] || '').trim();
      const phone = normalizePhone(row[1]);
      const status = String(row[4] || 'active').trim().toLowerCase();

      if (name && status !== 'inactive' && status !== 'disabled' && status !== 'ปิดใช้งาน') {
        activeUsers.push({
          name: name,
          phone: phone
        });
      }
    });
  }

  let normal = 0;
  let late = 0;
  let pending = 0;
  let approved = 0;
  let rejected = 0;
  let leave = 0;
  let sickLeave = 0;
  let personalLeave = 0;
  let officialDuty = 0;
  const leaveNames = [];
  const sickLeaveNames = [];
  const personalLeaveNames = [];
  const officialDutyNames = [];
  const checkedMap = {};
  const noCheckMap = {};

  if (permissionSheet && permissionSheet.getLastRow() >= 2) {
    const permissionValues = permissionSheet.getRange(2, 1, permissionSheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();

    permissionValues.forEach((row) => {
      const requestType = String(row[5] || '').trim();
      const requestDate = getBangkokDateKey(row[6]);
      const approvalStatus = normalizeStatusText(row[9]);

      if (requestDate !== today || !approvalStatus.includes('อนุมัติแล้ว')) {
        return;
      }

      if (!isNoCheckPermissionType(requestType)) {
        return;
      }

      const name = String(row[2] || '').trim();
      const phone = normalizePhone(row[3]);
      const key = phone || normalizeName(name);

      if (key) {
        noCheckMap[key] = true;
      }

      if (isLeavePermissionType(requestType)) {
        leave++;
        if (name) {
          leaveNames.push(name + ' (' + requestType + ')');
        }
      }

      if (requestType === 'ลาป่วย') {
        sickLeave++;
        if (name) {
          sickLeaveNames.push(name);
        }
      }

      if (requestType === 'ลากิจ' || requestType === 'ลา') {
        personalLeave++;
        if (name) {
          personalLeaveNames.push(name);
        }
      }

      if (requestType === 'ไปราชการ') {
        officialDuty++;
        if (name) {
          officialDutyNames.push(name);
        }
      }
    });
  }

  if (checkSheet && checkSheet.getLastRow() >= 2) {
    const values = checkSheet.getRange(2, 1, checkSheet.getLastRow() - 1, Math.max(checkSheet.getLastColumn(), 17)).getValues();

    values.forEach((row) => {
      const rowDate = getBangkokDateKey(row[0]);

      if (rowDate !== today) {
        return;
      }

      const name = String(row[1] || '').trim();
      const phone = normalizePhone(row[2]);
      const key = phone || normalizeName(name);
      const timeStatus = String(row[9] || '').trim();
      const approvalStatus = String(row[10] || '').trim();

      if (key) {
        checkedMap[key] = true;
      }

      if (timeStatus.includes('ปกติ')) {
        normal++;
      }

      if (timeStatus.includes('มาสาย')) {
        late++;
      }

      if (approvalStatus.includes('รออนุมัติ')) {
        pending++;
      }

      if (approvalStatus.includes('อนุมัติแล้ว')) {
        approved++;
      }

      if (approvalStatus.includes('ไม่อนุมัติ')) {
        rejected++;
      }
    });
  }

  const absentNames = activeUsers
    .filter((user) => {
      const key = user.phone || normalizeName(user.name);
      return key && !checkedMap[key] && !noCheckMap[key];
    })
    .map((user) => user.name);

  const totalActive = activeUsers.length;
  const noCheckTotal = Object.keys(noCheckMap).length;
  const checkedIn = Math.max(0, totalActive - absentNames.length - noCheckTotal);

  return {
    date: today,
    totalActive: totalActive,
    checkedIn: checkedIn,
    absent: absentNames.length,
    absentNames: absentNames,
    noCheckTotal: noCheckTotal,
    leave: leave,
    sickLeave: sickLeave,
    personalLeave: personalLeave,
    officialDuty: officialDuty,
    leaveNames: leaveNames,
    sickLeaveNames: sickLeaveNames,
    personalLeaveNames: personalLeaveNames,
    officialDutyNames: officialDutyNames,
    total: checkedIn,
    normal: normal,
    late: late,
    pending: pending,
    approved: approved,
    rejected: rejected
  };
}
function getAttendanceReport(filters) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CHECKIN_SHEET);
  const options = normalizeReportFilters(filters);
  const rows = [];
  const values = sheet && sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), 17)).getValues()
    : [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const dateKey = getBangkokDateKey(row[0]);

    if (options.startDate && dateKey < options.startDate) {
      continue;
    }

    if (options.endDate && dateKey > options.endDate) {
      continue;
    }

    const position = String(row[3] || '');
    const timeStatus = String(row[9] || '');
    const approvalStatus = String(row[10] || '');

    if (options.position && position !== options.position) {
      continue;
    }

    if (options.status && !matchesReportStatus(options.status, timeStatus, approvalStatus)) {
      continue;
    }

    rows.push({
      rowNumber: i + 2,
      date: dateKey,
      name: String(row[1] || ''),
      phone: normalizePhone(row[2]),
      position: position,
      checkinTime: formatTimeValue(row[7], ''),
      autoCheckoutTime: formatTimeValue(row[8], ''),
      distance: String(row[6] || ''),
      timeStatus: timeStatus,
      approvalStatus: approvalStatus,
      timeStatusCode: String(row[15] || getTimeStatusCode(timeStatus)),
      approvalStatusCode: String(row[16] || getApprovalStatusCode(approvalStatus)),
      reason: String(row[11] || ''),
      approver: String(row[12] || ''),
      approvalTime: formatDateTimeValue(row[13]),
      approvalNote: String(row[14] || '')
    });
  }


  appendApprovedNoCheckPermissionRows(rows, options);
  rows.sort((a, b) => {
    if (a.date === b.date) {
      return String(a.checkinTime).localeCompare(String(b.checkinTime));
    }

    return String(a.date).localeCompare(String(b.date));
  });

  return {
    success: true,
    rows: rows,
    summary: buildReportSummary(rows)
  };
}

function appendApprovedNoCheckPermissionRows(rows, options) {
  const sheet = getPermissionRequestsSheet();

  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();

  values.forEach((row, index) => {
    const requestType = String(row[5] || '').trim();

    if (!isNoCheckPermissionType(requestType)) {
      return;
    }

    const dateKey = getBangkokDateKey(row[6]);

    if (options.startDate && dateKey < options.startDate) {
      return;
    }

    if (options.endDate && dateKey > options.endDate) {
      return;
    }

    const position = String(row[4] || '').trim();
    const approvalStatus = String(row[9] || '').trim();
    const timeStatus = requestType;

    if (!normalizeStatusText(approvalStatus).includes('อนุมัติแล้ว')) {
      return;
    }

    if (options.position && position !== options.position) {
      return;
    }

    if (options.status && !matchesReportStatus(options.status, timeStatus, approvalStatus)) {
      return;
    }

    rows.push({
      rowNumber: 'P' + (index + 2),
      date: dateKey,
      name: String(row[2] || ''),
      phone: normalizePhone(row[3]),
      position: position,
      checkinTime: '-',
      autoCheckoutTime: '-',
      distance: '',
      timeStatus: timeStatus,
      approvalStatus: approvalStatus,
      timeStatusCode: isLeavePermissionType(requestType) ? 'leave' : 'official_duty',
      approvalStatusCode: String(row[13] || getPermissionApprovalStatusCode(approvalStatus)),
      reason: String(row[8] || ''),
      approver: String(row[10] || ''),
      approvalTime: String(row[11] || ''),
      approvalNote: String(row[12] || '')
    });
  });
}
function exportAttendanceReportCsv(filters) {
  const report = getAttendanceReport(filters);
  const headers = [
    'วันที่',
    'ชื่อ',
    'เบอร์โทร',
    'ตำแหน่ง',
    'เวลาเข้า',
    'เวลาออกอัตโนมัติ',
    'ระยะเมตร',
    'สถานะเวลา',
    'สถานะอนุมัติ',
    'TimeStatusCode',
    'ApprovalStatusCode',
    'เหตุผล',
    'ผู้อนุมัติ',
    'เวลาอนุมัติ',
    'หมายเหตุอนุมัติ'
  ];

  const lines = [headers];

  report.rows.forEach((row) => {
    lines.push([
      row.date,
      row.name,
      row.phone,
      row.position,
      row.checkinTime,
      row.autoCheckoutTime,
      row.distance,
      row.timeStatus,
      displayApprovalStatusForReport(row.approvalStatus),
      row.timeStatusCode,
      row.approvalStatusCode,
      row.reason,
      row.approver,
      row.approvalTime,
      row.approvalNote
    ]);
  });

  const csv = lines
    .map((line) => line.map(escapeCsvValue).join(','))
    .join('\r\n');

  return {
    success: true,
    filename: 'attendance_report_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '.csv',
    csv: '\ufeff' + csv
  };
}

function createDailyPdfReport(filters) {
  initializeSheets();

  const options = normalizeReportFilters(filters);
  const reportDate = options.date || options.startDate;

  try {
    const result = createDailyPdfReportCore(reportDate);

    if (result.success) {
      logPdfReport(reportDate, 'สำเร็จ', result.fileName || '', result.url || '', result.message || '');
    } else {
      logPdfReport(reportDate, 'ไม่สำเร็จ', result.fileName || '', result.url || '', result.message || 'สร้าง PDF ไม่สำเร็จ');
    }

    return result;
  } catch (err) {
    logPdfReport(reportDate, 'ผิดพลาด', '', '', err.message || String(err));
    throw err;
  }
}

function createDailyPdfReportCore(reportDate) {
  const report = getAttendanceReport({
    date: reportDate
  });
  const settings = getSettings();
  const templateId = String(settings.REPORT_TEMPLATE_DOC_ID || REPORT_TEMPLATE_DOC_ID).trim();

  if (!templateId) {
    return {
      success: false,
      message: 'ยังไม่ได้ตั้งค่า Google Docs Template รายงาน'
    };
  }

  const rootFolder = getReportFolder(settings);
  const folder = getMonthlyReportFolder(rootFolder, reportDate);
  const filename = 'รายงานลงเวลาปฏิบัติงาน_' + reportDate;
  const templateFile = DriveApp.getFileById(templateId);
  const workingFile = templateFile.makeCopy(filename + '_draft', folder);
  const doc = DocumentApp.openById(workingFile.getId());

  fillDailyReportDocument(doc, reportDate, report.rows, settings);
  doc.saveAndClose();

  const pdfBlob = workingFile
    .getAs(MimeType.PDF)
    .setName(filename + '.pdf');

  removeExistingFilesByName(folder, filename + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);

  workingFile.setTrashed(true);

  return {
    success: true,
    message: 'สร้าง PDF รายวันเรียบร้อย',
    url: pdfFile.getUrl(),
    fileName: pdfFile.getName(),
    reportDate: reportDate,
    folderName: folder.getName()
  };
}

function fillDailyReportDocument(doc, reportDate, rows, settings) {
  const body = doc.getBody();
  const dateText = formatThaiDate(reportDate);
  const summary = buildDailyPdfSummary(reportDate, rows);
  const reportRows = rows.filter((row) => !isNoCheckPermissionType(row.timeStatus));

  replaceReportPlaceholders(body, {
    '{{วันที่}}': dateText,
    '{{ทั้งหมด}}': summary.totalActive,
    '{{มา}}': summary.present,
    '{{ที่มา}}': summary.present,
    '{{ลาป่วย}}': summary.sickLeave,
    '{{ลากิจ}}': summary.personalLeave,
    '{{ไปราชการ}}': summary.officialDuty,
    '{{สาย}}': summary.late,
    '{{ไม่มา}}': summary.absent,
    '{{หมายเหตุ}}': summary.note,
    '{{ลายเซ็น}}': String(settings.DIRECTOR_NAME || '')
  });

  fillDailyReportTable(body, reportRows.slice(0, 12));
}

function getReportFolder(settings) {
  const folderId = String(settings.REPORT_FOLDER_ID || '').trim();

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (err) {
      // ถ้า folder id ใช้ไม่ได้ ให้สร้าง/ใช้โฟลเดอร์ค่าเริ่มต้นแทน
    }
  }

  const folders = DriveApp.getFoldersByName(REPORT_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(REPORT_FOLDER_NAME);
}

function getMonthlyReportFolder(rootFolder, dateKey) {
  const academicYear = getAcademicYearThai(dateKey);
  const monthFolderName = getMonthFolderName(dateKey);
  const yearFolder = getOrCreateChildFolder(rootFolder, 'ปีการศึกษา ' + academicYear);

  return getOrCreateChildFolder(yearFolder, monthFolderName);
}

function getOrCreateChildFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(folderName);
}

function getAcademicYearThai(dateKey) {
  const parts = String(dateKey || '').split('-');
  const year = Number(parts[0] || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy'));
  const month = Number(parts[1] || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'MM'));
  const academicYear = month >= 5 ? year : year - 1;

  return academicYear + 543;
}

function getMonthFolderName(dateKey) {
  const parts = String(dateKey || '').split('-');
  const month = Number(parts[1] || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'MM'));
  const months = [
    '',
    'มกราคม',
    'กุมภาพันธ์',
    'มีนาคม',
    'เมษายน',
    'พฤษภาคม',
    'มิถุนายน',
    'กรกฎาคม',
    'สิงหาคม',
    'กันยายน',
    'ตุลาคม',
    'พฤศจิกายน',
    'ธันวาคม'
  ];

  return String(month).padStart(2, '0') + ' ' + (months[month] || '');
}

function removeExistingFilesByName(folder, fileName) {
  const files = folder.getFilesByName(fileName);

  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function createTodayPdfReportAuto() {
  const reportDate = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  return createDailyPdfReport({
    date: reportDate
  });
}

function createMonthlyPdfReports(payload) {
  initializeSheets();

  const monthKey = String(payload && payload.month ? payload.month : '').trim();

  if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) {
    return {
      success: false,
      message: 'กรุณาเลือกเดือนที่ต้องการสร้าง PDF'
    };
  }

  const parts = monthKey.split('-');
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const lastDay = new Date(year, month, 0).getDate();
  const results = [];

  for (let day = 1; day <= lastDay; day++) {
    const reportDate =
      year + '-' +
      String(month).padStart(2, '0') + '-' +
      String(day).padStart(2, '0');

    try {
      const result = createDailyPdfReport({
        date: reportDate
      });

      results.push({
        date: reportDate,
        success: !!result.success,
        fileName: result.fileName || '',
        url: result.url || '',
        message: result.message || ''
      });
    } catch (err) {
      results.push({
        date: reportDate,
        success: false,
        fileName: '',
        url: '',
        message: err.message || String(err)
      });
    }
  }

  const successCount = results.filter((item) => item.success).length;

  return {
    success: true,
    message: 'สร้าง PDF รายเดือนเรียบร้อย ' + successCount + '/' + results.length + ' วัน',
    month: monthKey,
    total: results.length,
    successCount: successCount,
    failedCount: results.length - successCount,
    results: results
  };
}

function createMonthlyCombinedPdfReport(payload) {
  initializeSheets();

  const monthKey = String(payload && payload.month ? payload.month : '').trim();

  if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) {
    return {
      success: false,
      message: 'กรุณาเลือกเดือนที่ต้องการสร้าง PDF รวมเดือน'
    };
  }

  const settings = getSettings();
  const templateId = String(settings.REPORT_TEMPLATE_DOC_ID || REPORT_TEMPLATE_DOC_ID).trim();

  if (!templateId) {
    return {
      success: false,
      message: 'ยังไม่ได้ตั้งค่า Google Docs Template รายงาน'
    };
  }

  const workDates = getReportDatesWithCheckinsInMonth(monthKey);

  if (!workDates.length) {
    return {
      success: false,
      message: 'ไม่พบข้อมูลลงเวลาในเดือนที่เลือก'
    };
  }

  const rootFolder = getReportFolder(settings);
  const folder = getMonthlyReportFolder(rootFolder, workDates[0]);
  const filename = 'รายงานลงเวลาปฏิบัติงาน_' + monthKey + '_รวมเดือน';
  const templateFile = DriveApp.getFileById(templateId);
  const tempFiles = [];

  try {
    const firstFile = templateFile.makeCopy(filename + '_draft', folder);
    tempFiles.push(firstFile);

    const combinedDoc = DocumentApp.openById(firstFile.getId());
    const combinedBody = combinedDoc.getBody();
    fillMonthlyCombinedPage(combinedDoc, workDates[0], settings);

    for (let i = 1; i < workDates.length; i++) {
      const tempFile = templateFile.makeCopy(filename + '_page_' + i, folder);
      tempFiles.push(tempFile);

      const tempDoc = DocumentApp.openById(tempFile.getId());
      fillMonthlyCombinedPage(tempDoc, workDates[i], settings);
      tempDoc.saveAndClose();

      const sourceDoc = DocumentApp.openById(tempFile.getId());

      combinedBody.appendPageBreak();
      appendBodyContents(combinedBody, sourceDoc.getBody());
      sourceDoc.saveAndClose();
    }

    combinedDoc.saveAndClose();

    const pdfBlob = firstFile
      .getAs(MimeType.PDF)
      .setName(filename + '.pdf');

    removeExistingFilesByName(folder, filename + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);

    tempFiles.forEach((file) => file.setTrashed(true));

    logPdfReport(monthKey, 'สำเร็จ', pdfFile.getName(), pdfFile.getUrl(), 'สร้าง PDF รวมเดือนจากวันที่มีการลงเวลาจริง ' + workDates.length + ' หน้า');

    return {
      success: true,
      message: 'สร้าง PDF รวมเดือนเรียบร้อย',
      fileName: pdfFile.getName(),
      url: pdfFile.getUrl(),
      month: monthKey,
      pageCount: workDates.length
    };
  } catch (err) {
    tempFiles.forEach((file) => {
      try {
        file.setTrashed(true);
      } catch (trashErr) {}
    });

    logPdfReport(monthKey, 'ผิดพลาด', '', '', err.message || String(err));
    throw err;
  }
}

function fillMonthlyCombinedPage(doc, reportDate, settings) {
  const report = getAttendanceReport({
    date: reportDate
  });

  fillDailyReportDocument(doc, reportDate, report.rows, settings);
}

function appendBodyContents(targetBody, sourceBody) {
  const count = sourceBody.getNumChildren();

  for (let i = 0; i < count; i++) {
    const child = sourceBody.getChild(i).copy();
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetBody.appendParagraph(child.asParagraph());
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(child.asTable());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetBody.appendListItem(child.asListItem());
    } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      targetBody.appendHorizontalRule();
    } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
      targetBody.appendPageBreak();
    } else {
      targetBody.appendParagraph(child.getText ? child.getText() : '');
    }
  }
}

function getReportDatesWithCheckinsInMonth(monthKey) {
  const parts = String(monthKey || '').split('-');
  const year = Number(parts[0]);
  const month = Number(parts[1]);

  if (!year || !month) {
    return [];
  }

  const startDate = year + '-' + String(month).padStart(2, '0') + '-01';
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = year + '-' + String(month).padStart(2, '0') + '-' + String(lastDay).padStart(2, '0');
  const report = getAttendanceReport({
    startDate: startDate,
    endDate: endDate
  });
  const dateMap = {};

  report.rows.forEach((row) => {
    if (row.date) {
      dateMap[row.date] = true;
    }
  });

  return Object.keys(dateMap).sort();
}
function installDailyPdfTrigger() {
  removeDailyPdfTriggers();

  ScriptApp.newTrigger('createTodayPdfReportAuto')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .nearMinute(15)
    .inTimezone('Asia/Bangkok')
    .create();

  return {
    success: true,
    message: 'ตั้งค่า Trigger สร้าง PDF อัตโนมัติทุกวันเวลาประมาณ 18:15 แล้ว'
  };
}

function removeDailyPdfTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'createTodayPdfReportAuto') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function installMorningLineTrigger() {
  removeMorningLineTriggers();

  ScriptApp.newTrigger('sendMorningCheckinAlert')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0)
    .inTimezone('Asia/Bangkok')
    .create();

  return {
    success: true,
    message: 'ตั้งค่าแจ้งเตือน LINE ช่วงเช้าทุกวันเวลาประมาณ 08:00 แล้ว'
  };
}

function removeMorningLineTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'sendMorningCheckinAlert') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function installTomorrowPermissionLineTrigger() {
  removeTomorrowPermissionLineTriggers();

  ScriptApp.newTrigger('sendTomorrowPermissionAlert')
    .timeBased()
    .everyDays(1)
    .atHour(16)
    .nearMinute(30)
    .inTimezone('Asia/Bangkok')
    .create();

  return {
    success: true,
    message: '\u0e15\u0e31\u0e49\u0e07\u0e04\u0e48\u0e32\u0e41\u0e08\u0e49\u0e07\u0e40\u0e15\u0e37\u0e2d\u0e19 LINE \u0e04\u0e19\u0e25\u0e32/\u0e44\u0e1b\u0e23\u0e32\u0e0a\u0e01\u0e32\u0e23\u0e25\u0e48\u0e27\u0e07\u0e2b\u0e19\u0e49\u0e32\u0e17\u0e38\u0e01\u0e27\u0e31\u0e19\u0e40\u0e27\u0e25\u0e32\u0e1b\u0e23\u0e30\u0e21\u0e32\u0e13 16:30 \u0e41\u0e25\u0e49\u0e27'
  };
}

function removeTomorrowPermissionLineTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'sendTomorrowPermissionAlert') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function logPdfReport(reportDate, status, fileName, fileUrl, message) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PDF_LOGS_SHEET);

  sheet.appendRow([
    reportDate,
    new Date(),
    status,
    fileName,
    fileUrl,
    message
  ]);
}

function getPdfLogs(limit) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PDF_LOGS_SHEET);
  const maxRows = Number(limit || 10);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      success: true,
      rows: []
    };
  }

  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - maxRows + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, PDF_LOG_HEADERS.length).getValues();

  const rows = values.reverse().map((row) => ({
    reportDate: formatDateValue(row[0]) || String(row[0] || ''),
    runTime: formatDateTimeValue(row[1]),
    status: String(row[2] || ''),
    fileName: String(row[3] || ''),
    url: String(row[4] || ''),
    message: String(row[5] || '')
  }));

  return {
    success: true,
    rows: rows
  };
}

function replaceReportPlaceholders(body, values) {
  Object.keys(values).forEach((key) => {
    body.replaceText(escapeForReplaceText(key), String(values[key] || ''));
  });
}

function replacePlaceholderWithImage(body, placeholder, fileId, fallbackText, width) {
  const pattern = escapeForReplaceText(placeholder);

  if (!fileId) {
    body.replaceText(pattern, String(fallbackText || ''));
    return;
  }

  let found = body.findText(pattern);

  if (!found) {
    return;
  }

  while (found) {
    const textElement = found.getElement().asText();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();
    const parent = textElement.getParent();

    textElement.deleteText(start, end);

    try {
      const image = parent.insertInlineImage(parent.getChildIndex(textElement) + 1, DriveApp.getFileById(fileId).getBlob());

      if (width) {
        const originalWidth = image.getWidth();
        const originalHeight = image.getHeight();
        image.setWidth(width);

        if (originalWidth) {
          image.setHeight(Math.round(originalHeight * width / originalWidth));
        }
      }
    } catch (err) {
      textElement.insertText(start, String(fallbackText || ''));
    }

    found = body.findText(pattern, found);
  }
}

function fillDailyReportTable(body, rows) {
  const tables = body.getTables();
  let targetTable = null;

  for (let i = 0; i < tables.length; i++) {
    if (tables[i].getNumRows() >= 13 && tables[i].getRow(0).getNumCells() >= 8) {
      targetTable = tables[i];
      break;
    }
  }

  if (!targetTable) {
    body.replaceText(escapeForReplaceText('{{ATTENDANCE_TABLE}}'), 'ไม่พบตารางรายงาน');
    return;
  }

  for (let i = 0; i < 12; i++) {
    const tableRow = targetTable.getRow(i + 1);
    const item = rows[i];

    setTableCellText(tableRow, 0, toThaiNumber(i + 1));
    setTableCellText(tableRow, 1, item ? item.name : '');
    setTableCellText(tableRow, 2, item ? item.position : '');
    setTableCellText(tableRow, 3, item ? item.checkinTime : '');
    setTableCellText(tableRow, 4, item ? cleanStatusForReport(item.timeStatus) : '');
    setTableCellText(tableRow, 5, item ? item.autoCheckoutTime : '');
    setTableCellText(tableRow, 6, '');
    setTableCellText(tableRow, 7, item ? buildReportReason(item) : '');
  }
}

function setTableCellText(tableRow, cellIndex, value) {
  if (tableRow.getNumCells() <= cellIndex) {
    return;
  }

  tableRow.getCell(cellIndex).setText(String(value || ''));
}

function buildDailyPdfSummary(reportDate, rows) {
  const totalActive = countActiveUsers();
  const presentPhones = {};
  let late = 0;
  let sickLeave = 0;
  let personalLeave = 0;
  let officialDuty = 0;
  const noteItems = [];

  rows.forEach((row) => {
    const timeStatus = cleanStatusForReport(row.timeStatus);
    const key = normalizePhone(row.phone) || normalizeName(row.name);

    if (!isNoCheckPermissionType(timeStatus) && key) {
      presentPhones[key] = true;
    }

    if (normalizeStatusText(timeStatus).includes('มาสาย')) {
      late++;
    }

    if (timeStatus === 'ลาป่วย') {
      sickLeave++;
      noteItems.push(buildPdfNoteItem(row, 'ลาป่วย'));
    }

    if (timeStatus === 'ลากิจ') {
      personalLeave++;
      noteItems.push(buildPdfNoteItem(row, 'ลากิจ'));
    }

    if (timeStatus === 'ลา') {
      personalLeave++;
      noteItems.push(buildPdfNoteItem(row, 'ลา'));
    }

    if (timeStatus === 'ไปราชการ') {
      officialDuty++;
      noteItems.push(buildPdfNoteItem(row, 'ไปราชการ'));
    }
  });

  const present = Object.keys(presentPhones).length;
  const noCheck = sickLeave + personalLeave + officialDuty;

  return {
    totalActive: totalActive,
    present: present,
    sickLeave: sickLeave,
    personalLeave: personalLeave,
    officialDuty: officialDuty,
    late: late,
    absent: Math.max(0, totalActive - present - noCheck),
    note: noteItems.length ? noteItems.join(', ') : '-'
  };
}

function buildPdfNoteItem(row, label) {
  const name = String(row && row.name ? row.name : '').trim();
  const reason = String(row && row.reason ? row.reason : '').trim();
  const finalLabel = inferPdfLeaveLabel(label, reason);
  const detail = reason && reason !== '-' ? ' (' + reason + ')' : '';

  return [name, finalLabel].filter(Boolean).join(' ') + detail;
}

function inferPdfLeaveLabel(label, reason) {
  const text = String(label || '').trim();
  const reasonText = normalizeStatusText(reason);

  if (text !== 'ลา') {
    return text;
  }

  if (
    reasonText.includes('ป่วย') ||
    reasonText.includes('ปวด') ||
    reasonText.includes('ไข้') ||
    reasonText.includes('หมอ') ||
    reasonText.includes('แพทย์') ||
    reasonText.includes('โรงพยาบาล')
  ) {
    return 'ลาป่วย';
  }

  return 'ลากิจ';
}

function countActiveUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  let count = 0;

  values.forEach((row) => {
    const status = String(row[4] || 'active').trim().toLowerCase();

    if (status !== 'inactive' && status !== 'disabled' && status !== 'ปิดใช้งาน') {
      count++;
    }
  });

  return count;
}

function cleanStatusForReport(status) {
  return String(status || '')
    .replace(/[🟢🔴🟡✅]/g, '')
    .trim();
}

function displayApprovalStatusForReport(status) {
  const text = cleanStatusForReport(status);

  if (!text || text.includes('ไม่ต้องอนุมัติ')) {
    return '-';
  }

  return text;
}

function getTimeStatusCode(status) {
  const text = normalizeStatusText(status);

  if (text.includes('มาสาย')) {
    return 'late';
  }

  if (text.includes('ปกติ')) {
    return 'normal';
  }

  return '';
}

function getApprovalStatusCode(status) {
  const text = normalizeStatusText(status);

  if (text.includes('มาสายโดยไม่ได้ขอ')) {
    return 'unrequested';
  }

  if (text.includes('รออนุมัติ')) {
    return 'pending';
  }

  if (text.includes('อนุมัติแล้ว')) {
    return 'approved';
  }

  if (text.includes('ไม่อนุมัติ')) {
    return 'rejected';
  }

  if (text.includes('ไม่ต้องอนุมัติ')) {
    return 'none';
  }

  return '';
}

function backfillStatusCodes() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CHECKIN_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      success: true,
      message: 'ไม่มีข้อมูลให้ปรับปรุง',
      updated: 0
    };
  }

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, Math.max(sheet.getLastColumn(), 17)).getValues();
  const timeCodes = [];
  const approvalCodes = [];
  let updated = 0;

  values.forEach((row) => {
    const timeCode = String(row[15] || getTimeStatusCode(row[9]));
    const approvalCode = String(row[16] || getApprovalStatusCode(row[10]));

    if (!row[15] || !row[16]) {
      updated++;
    }

    timeCodes.push([timeCode]);
    approvalCodes.push([approvalCode]);
  });

  sheet.getRange(2, 16, rowCount, 1).setValues(timeCodes);
  sheet.getRange(2, 17, rowCount, 1).setValues(approvalCodes);

  return {
    success: true,
    message: 'ปรับปรุง Status Code ข้อมูลเก่าเรียบร้อย',
    updated: updated
  };
}

function buildReportReason(item) {
  const reason = String(item.reason || '').trim();

  if (!reason || reason === '-') {
    return '';
  }

  return reason;
}

function formatThaiDate(dateKey) {
  const parts = String(dateKey || '').split('-');

  if (parts.length !== 3) {
    return String(dateKey || '');
  }

  const months = [
    '',
    'มกราคม',
    'กุมภาพันธ์',
    'มีนาคม',
    'เมษายน',
    'พฤษภาคม',
    'มิถุนายน',
    'กรกฎาคม',
    'สิงหาคม',
    'กันยายน',
    'ตุลาคม',
    'พฤศจิกายน',
    'ธันวาคม'
  ];
  const day = Number(parts[2]);
  const month = months[Number(parts[1])] || '';
  const year = Number(parts[0]) + 543;

  return toThaiNumber(day) + ' ' + month + ' ' + toThaiNumber(year);
}

function toThaiNumber(value) {
  const thaiDigits = ['๐', '๑', '๒', '๓', '๔', '๕', '๖', '๗', '๘', '๙'];

  return String(value).replace(/\d/g, (digit) => thaiDigits[Number(digit)]);
}

function formatLeaveDocumentNoForPdf(documentNo) {
  const text = String(documentNo || '').trim();
  const match = text.match(/^(.+?)\s+(\d+\s*\/\s*\d+)$/);

  if (!match) {
    return toThaiNumber(text);
  }

  return toThaiNumber(match[2].replace(/\s+/g, ''));
}

function formatFullLeaveDocumentNoForPdf(documentNo) {
  return toThaiNumber(String(documentNo || '').replace(/\s+/g, ' ').trim());
}

function escapeForReplaceText(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeReportFilters(filters) {
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const source = filters || {};
  const singleDate = String(source.date || '').trim();

  return {
    startDate: singleDate || String(source.startDate || today).trim(),
    endDate: singleDate || String(source.endDate || today).trim(),
    date: singleDate || String(source.startDate || today).trim(),
    position: String(source.position || '').trim(),
    status: String(source.status || '').trim()
  };
}

function matchesReportStatus(filterStatus, timeStatus, approvalStatus) {
  const timeText = normalizeStatusText(timeStatus);
  const approvalText = normalizeStatusText(approvalStatus);

  if (filterStatus === 'normal') {
    return timeText.includes('ปกติ');
  }

  if (filterStatus === 'late') {
    return timeText.includes('มาสาย');
  }

  if (filterStatus === 'pending') {
    return approvalText.includes('รออนุมัติ');
  }

  if (filterStatus === 'approved') {
    return approvalText.includes('อนุมัติแล้ว');
  }

  if (filterStatus === 'rejected') {
    return approvalText.includes('ไม่อนุมัติ');
  }

  if (filterStatus === 'unrequested') {
    return approvalText.includes('มาสายโดยไม่ได้ขอ');
  }

  if (filterStatus === 'leave') {
    return timeText.includes('ลา');
  }

  if (filterStatus === 'official_duty') {
    return timeText.includes('ไปราชการ');
  }

  return true;
}

function buildReportSummary(rows) {
  const summary = {
    total: rows.length,
    normal: 0,
    late: 0,
    pending: 0,
    approved: 0,
    rejected: 0,
    unrequested: 0,
    leave: 0,
    officialDuty: 0
  };

  rows.forEach((row) => {
    const timeStatus = normalizeStatusText(row.timeStatus);
    const approvalStatus = normalizeStatusText(row.approvalStatus);

    if (timeStatus.includes('ปกติ')) {
      summary.normal++;
    }

    if (timeStatus.includes('มาสาย')) {
      summary.late++;
    }

    if (approvalStatus.includes('รออนุมัติ')) {
      summary.pending++;
    }

    if (approvalStatus.includes('อนุมัติแล้ว')) {
      summary.approved++;
    }

    if (approvalStatus.includes('ไม่อนุมัติ')) {
      summary.rejected++;
    }

    if (approvalStatus.includes('มาสายโดยไม่ได้ขอ')) {
      summary.unrequested++;
    }

    if (timeStatus.includes('ลา')) {
      summary.leave++;
    }

    if (timeStatus.includes('ไปราชการ')) {
      summary.officialDuty++;
    }
  });

  return summary;
}

function formatDateTimeValue(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  }

  return String(value);
}

function escapeCsvValue(value) {
  const text = String(value || '');
  return '"' + text.replace(/"/g, '""') + '"';
}

function openDashboard() {

  return HtmlService
    .createHtmlOutputFromFile(DASHBOARD_FILE)
    .setTitle('Dashboard');
}

function canOpenDashboard(phone) {
  const user = getUserProfileWithRoleByPhone(phone);

  return Boolean(user && (user.role === 'admin' || user.role === 'director'));
}

function canOpenSettings(phone) {
  const user = getUserProfileWithRoleByPhone(phone);

  return Boolean(user && user.role === 'director');
}

function getDashboardUrl(phone) {
  if (!canOpenDashboard(phone)) {
    throw new Error('คุณไม่มีสิทธิ์เข้า Dashboard');
  }

  return ScriptApp.getService().getUrl() +
    '?token=' + encodeURIComponent(QR_TOKEN) +
    '&page=dashboard' +
    '&user=' + encodeURIComponent(normalizePhone(phone));
}

function exportExcel(phone) {
  if (!canOpenDashboard(phone)) {
    throw new Error('คุณไม่มีสิทธิ์ Export Excel');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const url =
    'https://docs.google.com/spreadsheets/d/' +
    ss.getId() +
    '/export?format=xlsx';

  return url;
}

function getPermissionRequestsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PERMISSION_REQUESTS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(PERMISSION_REQUESTS_SHEET);
    sheet.appendRow(PERMISSION_REQUEST_HEADERS);
  }

  ensureSheetHeaders(sheet, PERMISSION_REQUEST_HEADERS);
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('G:H').setNumberFormat('@');
  sheet.getRange('O:P').setNumberFormat('@');
  sheet.getRange('R:S').setNumberFormat('@');
  sheet.getRange('U:V').setNumberFormat('@');

  return sheet;
}
function submitPermissionRequest(data) {
  initializeSheets();

  const name = String(data && data.name ? data.name : '').trim();
  const phone = normalizePhone(data && data.phone ? data.phone : '');
  const position = String(data && data.position ? data.position : '').trim();
  const requestType = normalizePermissionRequestType(data && data.requestType ? data.requestType : 'มาสาย');
  const requestDate = String(data && data.requestDate ? data.requestDate : '').trim();
  const endDate = String(data && data.endDate ? data.endDate : requestDate).trim();
  const expectedTime = normalizeTimeInput(data && data.expectedTime ? data.expectedTime : '');
  const reason = String(data && data.reason ? data.reason : '').trim();
  const days = calculateInclusiveDays(requestDate, endDate);
  const attachment = data && data.attachment ? data.attachment : null;

  if (!name || !phone) {
    return { success: false, message: 'ไม่พบข้อมูลผู้ใช้ กรุณาเข้าสู่ระบบใหม่' };
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(requestDate)) {
    return { success: false, message: 'กรุณาเลือกวันที่ที่ต้องการขออนุญาต' };
  }

  if (isLeavePermissionType(requestType) && !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
    return { success: false, message: 'กรุณาเลือกวันที่สิ้นสุดการลา' };
  }

  if (isLeavePermissionType(requestType) && days < 1) {
    return { success: false, message: 'วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มลา' };
  }

  if (!requestType) {
    return { success: false, message: 'กรุณาเลือกประเภทคำขอ' };
  }

  if (requestType === 'มาสาย' && !expectedTime) {
    return { success: false, message: 'กรุณาระบุเวลาที่คาดว่าจะมาถึง' };
  }

  if (!reason) {
    return { success: false, message: 'กรุณาระบุเหตุผล' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPermissionRequestsSheet();
    const duplicate = findActivePermissionRequest(sheet, phone, name, requestDate);

    if (duplicate) {
      return {
        success: false,
        message: 'วันนี้คุณมีคำขออนุญาตอยู่แล้ว (' + duplicate.requestType + ' / ' + duplicate.approvalStatus + ') กรุณารอผู้บริหารพิจารณา หรือแจ้งผู้บริหารให้แก้ไข'
      };
    }

    const now = new Date();
    const requestId = getPermissionRequestPrefix(requestType) + '-' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
    const submittedAt = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const approvalStatus = '🟡 รออนุมัติ';
    const evidence = isLeavePermissionType(requestType)
      ? savePermissionEvidenceFile(attachment, requestId, requestDate, name, requestType)
      : null;
    const leaveSequence = isLeavePermissionType(requestType)
      ? countLeaveRequestsForPerson(sheet, phone, name, requestDate) + 1
      : '';
    const leaveDocument = isLeavePermissionType(requestType)
      ? buildNextLeaveDocumentNumber(sheet, requestDate)
      : { documentNo: '', academicYear: '' };
    const termStats = isLeavePermissionType(requestType)
      ? buildLeaveTermStats(sheet, phone, name, requestDate, days)
      : { term: '', count: '', days: '', summary: '' };

    sheet.appendRow([
      requestId,
      submittedAt,
      name,
      phone,
      position,
      requestType,
      requestDate,
      expectedTime,
      reason,
      approvalStatus,
      '',
      '',
      '',
      'pending',
      isLeavePermissionType(requestType) ? endDate : '',
      isLeavePermissionType(requestType) ? days : '',
      '',
      evidence ? evidence.url : '',
      evidence ? evidence.name : '',
      leaveSequence,
      leaveDocument.documentNo,
      leaveDocument.academicYear,
      termStats.term,
      termStats.count,
      termStats.days
    ]);

    notifyLinePermissionRequest({
      requestId: requestId,
      documentNo: leaveDocument.documentNo,
      submittedAt: submittedAt,
      name: name,
      position: position,
      requestType: requestType,
      requestDate: requestDate,
      endDate: endDate,
      days: days,
      expectedTime: expectedTime,
      reason: reason,
      approvalStatus: approvalStatus,
      evidenceUrl: evidence ? evidence.url : '',
      leaveSequence: leaveSequence,
      documentNo: leaveDocument.documentNo,
      leaveTermSummary: termStats.summary
    });

    return {
      success: true,
      message: 'ส่งคำขออนุญาต' + requestType + 'เรียบร้อยแล้ว',
      requestId: requestId,
      documentNo: leaveDocument.documentNo,
      approvalStatus: approvalStatus,
      leaveSequence: leaveSequence,
      leaveTermSummary: termStats.summary,
      leaveTerm: termStats.term,
      leaveTermCount: termStats.count,
      leaveTermDays: termStats.days
    };
  } finally {
    lock.releaseLock();
  }
}

function findActivePermissionRequest(sheet, phone, name, requestDate) {
  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  const targetPhone = normalizePhone(phone);
  const targetName = normalizeName(name);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowPhone = normalizePhone(row[3]);
    const rowName = normalizeName(row[2]);
    const rowDate = getBangkokDateKey(row[6]);
    const approvalStatus = normalizeStatusText(row[9]);
    const isSamePerson = targetPhone
      ? rowPhone === targetPhone
      : rowName === targetName;

    if (!isSamePerson || rowDate !== requestDate) {
      continue;
    }

    if (approvalStatus.includes('รออนุมัติ') || approvalStatus.includes('อนุมัติแล้ว')) {
      return {
        rowNumber: i + 2,
        requestType: String(row[5] || '-'),
        approvalStatus: cleanStatusForReport(row[9]) || '-'
      };
    }
  }

  return null;
}

function countLeaveRequestsForPerson(sheet, phone, name, requestDate) {
  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const targetPhone = normalizePhone(phone);
  const targetName = normalizeName(name);
  const targetAcademicYear = getAcademicYearThai(requestDate);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  let count = 0;

  values.forEach((row) => {
    const requestType = String(row[5] || '').trim();

    if (!isLeavePermissionType(requestType)) {
      return;
    }

    const rowPhone = normalizePhone(row[3]);
    const rowName = normalizeName(row[2]);
    const isSamePerson = targetPhone
      ? rowPhone === targetPhone
      : rowName === targetName;

    if (!isSamePerson) {
      return;
    }

    if (getAcademicYearThai(getBangkokDateKey(row[6])) !== targetAcademicYear) {
      return;
    }

    if (normalizeStatusText(row[9]).includes('ไม่อนุมัติ')) {
      return;
    }

    count++;
  });

  return count;
}

function getCurrentTermSettings(referenceDate) {
  const settings = getSettings();
  const term = String(settings.CURRENT_TERM || DEFAULT_SETTINGS.CURRENT_TERM || '1').trim() || '1';
  const configuredStartDate = String(settings.TERM_START_DATE || '').trim();
  const configuredEndDate = String(settings.TERM_END_DATE || '').trim();
  const defaultRange = getDefaultTermDateRange(referenceDate, term, settings);
  const startDate = /^\d{4}-\d{2}-\d{2}$/.test(configuredStartDate)
    ? configuredStartDate
    : defaultRange.startDate;
  const endDate = /^\d{4}-\d{2}-\d{2}$/.test(configuredEndDate)
    ? configuredEndDate
    : defaultRange.endDate;

  return {
    term: term,
    startDate: startDate,
    endDate: endDate
  };
}

function buildLeaveTermStats(sheet, phone, name, requestDate, currentDays) {
  const termSettings = getCurrentTermSettings(requestDate);
  const stats = countLeaveTermStatsForPerson(sheet, phone, name, termSettings, requestDate);
  const totalCount = stats.count + 1;
  const totalDays = stats.days + (Number(currentDays || 0) || 0);

  return {
    term: termSettings.term,
    count: totalCount,
    days: totalDays,
    summary: 'ภาคเรียนที่ ' + termSettings.term + ': ลาแล้ว ' + totalCount + ' ครั้ง / ' + totalDays + ' วัน'
  };
}

function getDefaultTermDateRange(referenceDate, term, settings) {
  const dateKey = getBangkokDateKey(referenceDate) || getBangkokDateKey(new Date());
  const fallbackAcademicYear = getAcademicYearThai(dateKey);
  const academicYearThai = Number(
    (settings && settings.ACADEMIC_YEAR) ||
    fallbackAcademicYear ||
    DEFAULT_SETTINGS.ACADEMIC_YEAR
  );
  const academicYearAd = academicYearThai > 2400
    ? academicYearThai - 543
    : academicYearThai;
  const normalizedTerm = String(term || '1').trim();

  if (normalizedTerm === '2') {
    return {
      startDate: academicYearAd + '-11-01',
      endDate: (academicYearAd + 1) + '-04-30'
    };
  }

  return {
    startDate: academicYearAd + '-05-01',
    endDate: academicYearAd + '-10-31'
  };
}

function countLeaveTermStatsForPerson(sheet, phone, name, termSettings, fallbackDate) {
  if (!sheet || sheet.getLastRow() < 2) {
    return { count: 0, days: 0 };
  }

  const targetPhone = normalizePhone(phone);
  const targetName = normalizeName(name);
  const startDate = termSettings.startDate || getBangkokDateKey(fallbackDate);
  const endDate = termSettings.endDate || getBangkokDateKey(fallbackDate);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  let count = 0;
  let days = 0;

  values.forEach((row) => {
    const requestType = String(row[5] || '').trim();

    if (!isLeavePermissionType(requestType)) {
      return;
    }

    const approvalCode = String(row[13] || getPermissionApprovalStatusCode(row[9]));

    if (approvalCode === 'rejected') {
      return;
    }

    const rowDate = getBangkokDateKey(row[6]);

    if (rowDate < startDate || rowDate > endDate) {
      return;
    }

    const rowPhone = normalizePhone(row[3]);
    const rowName = normalizeName(row[2]);
    const isSamePerson = targetPhone
      ? rowPhone === targetPhone
      : rowName === targetName;

    if (!isSamePerson) {
      return;
    }

    count++;
    days += Number(row[15] || calculateInclusiveDays(rowDate, getBangkokDateKey(row[14] || row[6])) || 0) || 0;
  });

  return {
    count: count,
    days: days
  };
}

function savePermissionEvidenceFile(attachment, requestId, requestDate, name, requestType) {
  if (!attachment || !attachment.data) {
    return null;
  }

  const base64 = String(attachment.data || '').replace(/^data:[^,]+,/, '');
  const originalName = sanitizeFileName(attachment.name || 'evidence');
  const contentType = String(attachment.type || 'application/octet-stream');
  const bytes = Utilities.base64Decode(base64);

  if (bytes.length > 5 * 1024 * 1024) {
    throw new Error('ไฟล์หลักฐานมีขนาดใหญ่เกิน 5 MB');
  }

  const settings = getSettings();
  const rootFolder = getReportFolder(settings);
  const monthFolder = getMonthlyReportFolder(rootFolder, requestDate);
  const evidenceFolder = getOrCreateChildFolder(monthFolder, 'หลักฐานการลา');
  const extension = getFileExtension(originalName, contentType);
  const filename = 'หลักฐาน_' +
    requestDate + '_' +
    sanitizeFileName(name) + '_' +
    sanitizeFileName(requestType) + '_' +
    sanitizeFileName(requestId) +
    extension;
  const blob = Utilities.newBlob(bytes, contentType, filename);
  const file = evidenceFolder.createFile(blob);

  return {
    name: filename,
    url: file.getUrl()
  };
}

function getFileExtension(filename, contentType) {
  const cleanName = String(filename || '');
  const match = cleanName.match(/(\.[a-zA-Z0-9]{1,8})$/);

  if (match) {
    return match[1].toLowerCase();
  }

  if (contentType === 'application/pdf') return '.pdf';
  if (contentType === 'image/png') return '.png';
  if (contentType === 'image/jpeg') return '.jpg';

  return '';
}

function submitLatePermissionRequest(data) {
  data = data || {};
  data.requestType = 'มาสาย';
  return submitPermissionRequest(data);
}

function getPendingPermissionRequests() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getPermissionRequestsSheet();

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  const result = [];

  values.forEach((row, index) => {
    const approvalStatus = normalizeStatusText(row[9]);
    const requestType = String(row[5] || '').trim();

    if (isPermissionRequestType(requestType) && approvalStatus.includes('รออนุมัติ')) {
      result.push({
        rowNumber: index + 2,
        requestId: String(row[0] || ''),
        submittedAt: String(row[1] || ''),
        name: String(row[2] || ''),
        phone: normalizePhone(row[3]),
        position: String(row[4] || ''),
        requestType: requestType,
        requestDate: String(row[6] || ''),
        endDate: String(row[14] || row[6] || ''),
        days: String(row[15] || ''),
        leavePdfUrl: String(row[16] || ''),
        evidenceUrl: String(row[17] || ''),
        evidenceName: String(row[18] || ''),
        leaveSequence: String(row[19] || ''),
        expectedTime: formatLineTime(row[7]),
        reason: String(row[8] || '-'),
        approvalStatus: String(row[9] || '🟡 รออนุมัติ')
      });
    }
  });

  result.sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt)));
  return result;
}

function getLeaveRequests(filters) {
  initializeSheets();

  const sheet = getPermissionRequestsSheet();
  const options = normalizeLeaveFilters(filters);
  const rows = [];

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      success: true,
      rows: [],
      summary: buildLeaveSummary([])
    };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();

  values.forEach((row) => {
    const requestType = String(row[5] || '').trim();

    if (!isLeavePermissionType(requestType)) {
      return;
    }

    const startDate = getBangkokDateKey(row[6]);
    const endDate = getBangkokDateKey(row[14] || row[6]);
    const approvalStatus = String(row[9] || '').trim();
    const statusCode = String(row[13] || getPermissionApprovalStatusCode(approvalStatus));
    const name = String(row[2] || '').trim();

    if (options.startDate && endDate < options.startDate) {
      return;
    }

    if (options.endDate && startDate > options.endDate) {
      return;
    }

    if (options.type && requestType !== options.type) {
      return;
    }

    if (options.status && statusCode !== options.status) {
      return;
    }

    if (options.keyword && !normalizeName(name).includes(normalizeName(options.keyword))) {
      return;
    }

    const displayTermStats = buildLeaveTermStatsForExistingRow(values, row);

    rows.push({
      requestId: String(row[0] || ''),
      submittedAt: String(row[1] || ''),
      name: name,
      phone: normalizePhone(row[3]),
      position: String(row[4] || ''),
      requestType: requestType,
      startDate: startDate,
      endDate: endDate,
      days: String(row[15] || calculateInclusiveDays(startDate, endDate) || ''),
      reason: String(row[8] || '-'),
      approvalStatus: approvalStatus,
      approvalStatusCode: statusCode,
      approver: String(row[10] || ''),
      approvalTime: String(row[11] || ''),
      approvalNote: String(row[12] || ''),
      leavePdfUrl: String(row[16] || ''),
      evidenceUrl: String(row[17] || ''),
      evidenceName: String(row[18] || ''),
      leaveSequence: String(row[19] || ''),
      documentNo: String(row[20] || ''),
      academicYear: String(row[21] || ''),
      leaveTerm: String(displayTermStats.term || row[22] || ''),
      leaveTermCount: String(displayTermStats.count || row[23] || ''),
      leaveTermDays: String(displayTermStats.days || row[24] || '')
    });
  });

  rows.sort((a, b) =>
    getLeaveStatusSortRank(a.approvalStatusCode) - getLeaveStatusSortRank(b.approvalStatusCode) ||
    getLeaveDocumentSortValue(b.documentNo) - getLeaveDocumentSortValue(a.documentNo) ||
    String(b.startDate).localeCompare(String(a.startDate)) ||
    String(a.name).localeCompare(String(b.name))
  );

  return {
    success: true,
    rows: rows,
    summary: buildLeaveSummary(rows)
  };
}

function getMyHistory(phone) {
  initializeSheets();

  const user = getUserProfileWithRoleByPhone(phone);

  if (!user) {
    return {
      success: false,
      message: 'ไม่พบข้อมูลผู้ใช้'
    };
  }

  if (user.role === 'director') {
    return {
      success: false,
      message: 'ผู้บริหารดูข้อมูลภาพรวมได้จาก Dashboard'
    };
  }

  const checkins = getMyCheckinHistoryRows(user);
  const permissions = getMyPermissionHistoryRows(user);

  return {
    success: true,
    user: {
      name: user.name,
      phone: user.phone,
      position: user.position || '',
      role: user.role || ''
    },
    checkins: checkins,
    permissions: permissions
  };
}

function getMyCheckinHistoryRows(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CHECKIN_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CHECKIN_HEADERS.length).getValues();
  const targetPhone = normalizePhone(user.phone);
  const targetName = normalizeName(user.name);
  const rows = [];

  values.forEach((row) => {
    const rowPhone = normalizePhone(row[2]);
    const rowName = normalizeName(row[1]);
    const isSamePerson = targetPhone
      ? isSamePhone(rowPhone, targetPhone)
      : rowName === targetName;

    if (!isSamePerson) {
      return;
    }

    rows.push({
      date: getBangkokDateKey(row[0]),
      checkinTime: formatTimeValue(row[7], ''),
      checkoutTime: formatTimeValue(row[8], ''),
      timeStatus: cleanStatusForReport(row[9]) || '-',
      approvalStatus: displayApprovalStatusForReport(row[10]),
      reason: String(row[11] || '-')
    });
  });

  rows.sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));

  return rows.slice(0, 120);
}

function getMyPermissionHistoryRows(user) {
  const sheet = getPermissionRequestsSheet();

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  const targetPhone = normalizePhone(user.phone);
  const targetName = normalizeName(user.name);
  const rows = [];

  values.forEach((row) => {
    const rowPhone = normalizePhone(row[3]);
    const rowName = normalizeName(row[2]);
    const isSamePerson = targetPhone
      ? isSamePhone(rowPhone, targetPhone)
      : rowName === targetName;

    if (!isSamePerson) {
      return;
    }

    rows.push({
      requestId: String(row[0] || ''),
      submittedAt: String(row[1] || ''),
      requestType: cleanStatusForReport(row[5]) || '-',
      startDate: getBangkokDateKey(row[6]),
      endDate: getBangkokDateKey(row[14] || row[6]),
      expectedTime: formatTimeValue(row[7], ''),
      reason: String(row[8] || '-'),
      approvalStatus: cleanStatusForReport(row[9]) || '-',
      approvalStatusCode: String(row[13] || getPermissionApprovalStatusCode(row[9])),
      days: String(row[15] || calculateInclusiveDays(getBangkokDateKey(row[6]), getBangkokDateKey(row[14] || row[6])) || ''),
      leavePdfUrl: String(row[16] || ''),
      evidenceUrl: String(row[17] || ''),
      evidenceName: String(row[18] || ''),
      leaveSequence: String(row[19] || ''),
      documentNo: String(row[20] || ''),
      approver: String(row[10] || '')
    });
  });

  rows.sort((a, b) =>
    String(b.startDate || '').localeCompare(String(a.startDate || '')) ||
    String(b.submittedAt || '').localeCompare(String(a.submittedAt || ''))
  );

  return rows.slice(0, 120);
}

function getLeaveStatusSortRank(statusCode) {
  if (statusCode === 'pending') return 0;
  if (statusCode === 'approved') return 1;
  if (statusCode === 'rejected') return 2;
  return 3;
}

function buildLeaveTermStatsForExistingRow(values, currentRow) {
  const requestType = String(currentRow[5] || '').trim();

  if (!isLeavePermissionType(requestType)) {
    return { term: '', count: '', days: '' };
  }

  const currentDate = getBangkokDateKey(currentRow[6]);
  const termSettings = getCurrentTermSettings(currentDate);
  const targetPhone = normalizePhone(currentRow[3]);
  const targetName = normalizeName(currentRow[2]);
  const currentDocSort = getLeaveDocumentSortValue(currentRow[20]);
  const currentSubmittedAt = String(currentRow[1] || '');
  let count = 0;
  let days = 0;

  values.forEach((row) => {
    if (!isLeavePermissionType(String(row[5] || '').trim())) {
      return;
    }

    const approvalCode = String(row[13] || getPermissionApprovalStatusCode(row[9]));

    if (approvalCode === 'rejected') {
      return;
    }

    const rowPhone = normalizePhone(row[3]);
    const rowName = normalizeName(row[2]);
    const isSamePerson = targetPhone
      ? rowPhone === targetPhone
      : rowName === targetName;

    if (!isSamePerson) {
      return;
    }

    const rowDate = getBangkokDateKey(row[6]);

    if (rowDate < termSettings.startDate || rowDate > termSettings.endDate) {
      return;
    }

    const rowDocSort = getLeaveDocumentSortValue(row[20]);

    if (currentDocSort > 0 && rowDocSort > 0) {
      if (rowDocSort > currentDocSort) {
        return;
      }
    } else if (String(row[1] || '') > currentSubmittedAt) {
      return;
    }

    count++;
    days += Number(row[15] || calculateInclusiveDays(rowDate, getBangkokDateKey(row[14] || row[6])) || 0) || 0;
  });

  return {
    term: termSettings.term,
    count: count,
    days: days
  };
}

function repairLeaveTermStats() {
  initializeSheets();

  const sheet = getPermissionRequestsSheet();

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      success: true,
      updated: 0,
      message: 'ไม่มีข้อมูลการลาให้ปรับปรุง'
    };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  let updated = 0;

  values.forEach((row, index) => {
    if (!isLeavePermissionType(String(row[5] || '').trim())) {
      return;
    }

    const stats = buildLeaveTermStatsForExistingRow(values, row);
    const rowNumber = index + 2;

    sheet.getRange(rowNumber, 23).setValue(stats.term || '');
    sheet.getRange(rowNumber, 24).setValue(stats.count || '');
    sheet.getRange(rowNumber, 25).setValue(stats.days || '');
    updated++;
  });

  SpreadsheetApp.flush();

  return {
    success: true,
    updated: updated,
    message: 'ปรับปรุงสถิติการลาสะสมเรียบร้อย ' + updated + ' รายการ'
  };
}

function getLeaveDocumentSortValue(documentNo) {
  const match = String(documentNo || '').match(/(\d+)\s*\/\s*(\d+)/);

  if (!match) {
    return 0;
  }

  return (Number(match[2]) || 0) * 100000 + (Number(match[1]) || 0);
}

function normalizeLeaveFilters(filters) {
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const source = filters || {};
  const month = String(source.month || '').trim();
  let startDate = String(source.startDate || '').trim();
  let endDate = String(source.endDate || '').trim();

  if (month && /^\d{4}-\d{2}$/.test(month)) {
    const parts = month.split('-');
    const year = Number(parts[0]);
    const monthNumber = Number(parts[1]);
    const lastDay = new Date(year, monthNumber, 0).getDate();
    startDate = month + '-01';
    endDate = month + '-' + String(lastDay).padStart(2, '0');
  }

  if (!startDate && !endDate) {
    startDate = today.slice(0, 7) + '-01';
    endDate = today;
  }

  return {
    startDate: startDate || endDate,
    endDate: endDate || startDate,
    month: month,
    keyword: String(source.keyword || '').trim(),
    type: normalizePermissionRequestType(source.type || ''),
    status: String(source.status || '').trim()
  };
}

function buildLeaveSummary(rows) {
  const summary = {
    total: rows.length,
    sickLeave: 0,
    personalLeave: 0,
    pending: 0,
    approved: 0,
    rejected: 0,
    evidence: 0,
    pdf: 0
  };

  rows.forEach((row) => {
    if (row.requestType === 'ลาป่วย') summary.sickLeave++;
    if (row.requestType === 'ลากิจ' || row.requestType === 'ลา') summary.personalLeave++;
    if (row.approvalStatusCode === 'pending') summary.pending++;
    if (row.approvalStatusCode === 'approved') summary.approved++;
    if (row.approvalStatusCode === 'rejected') summary.rejected++;
    if (row.evidenceUrl) summary.evidence++;
    if (row.leavePdfUrl) summary.pdf++;
  });

  return summary;
}

function updatePermissionRequestApproval(requestId, newStatus, approverName, note) {
  initializeSheets();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPermissionRequestsSheet();

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, message: 'ไม่พบรายการคำขอ' };
    }

    const id = String(requestId || '').trim();
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
    let targetRow = -1;

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === id) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow < 2) {
      return { success: false, message: 'ไม่พบรายการคำขอที่เลือก' };
    }

    const currentStatus = normalizeStatusText(sheet.getRange(targetRow, 10).getValue());

    if (!currentStatus.includes('รออนุมัติ')) {
      return { success: false, message: 'รายการนี้ถูกจัดการแล้ว' };
    }

    const now = new Date();
    const approvalTime = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const statusCode = getPermissionApprovalStatusCode(newStatus);

    sheet.getRange(targetRow, 10).setValue(newStatus);
    sheet.getRange(targetRow, 11).setValue(String(approverName || 'ผู้บริหาร'));
    sheet.getRange(targetRow, 12).setValue(approvalTime);
    sheet.getRange(targetRow, 13).setValue(String(note || '-'));
    sheet.getRange(targetRow, 14).setValue(statusCode);

    SpreadsheetApp.flush();

    let rowData = sheet.getRange(targetRow, 1, 1, PERMISSION_REQUEST_HEADERS.length).getValues()[0];
    let leavePdf = null;

    if (statusCode === 'approved' && isLeavePermissionType(rowData[5])) {
      try {
        leavePdf = createLeavePdfForRow(rowData);

        if (leavePdf && leavePdf.success) {
          sheet.getRange(targetRow, 17).setValue(leavePdf.url || '');
        }
      } catch (pdfErr) {
        leavePdf = {
          success: false,
          message: pdfErr.message || String(pdfErr)
        };
      }
    }

    notifyLinePermissionApproval({
      requestId: rowData[0],
      name: rowData[2],
      position: rowData[4],
      requestType: rowData[5],
      requestDate: rowData[6],
      endDate: rowData[14] || rowData[6],
      days: rowData[15] || '',
      expectedTime: rowData[7],
      reason: rowData[8],
      approvalStatus: newStatus,
      approver: approverName || 'ผู้บริหาร',
      approvalNote: note || '-',
      approvalTime: approvalTime,
      documentNo: rowData[20] || '',
      leaveTerm: rowData[22] || '',
      leaveTermCount: rowData[23] || '',
      leaveTermDays: rowData[24] || '',
      leavePdfUrl: leavePdf && leavePdf.success ? leavePdf.url : ''
    });

    return {
      success: true,
      message: 'อัปเดตคำขอเรียบร้อยแล้ว',
      requestId: id,
      approvalStatus: newStatus,
      approvalTime: approvalTime,
      leavePdfUrl: leavePdf && leavePdf.success ? leavePdf.url : '',
      leavePdfMessage: leavePdf && !leavePdf.success ? leavePdf.message : ''
    };
  } finally {
    lock.releaseLock();
  }
}

function approvePermissionRequest(requestId, approverName, note) {
  return updatePermissionRequestApproval(requestId, '🟢 อนุมัติแล้ว', approverName, note);
}

function rejectPermissionRequest(requestId, approverName, note) {
  return updatePermissionRequestApproval(requestId, '🔴 ไม่อนุมัติ', approverName, note);
}

function createLeavePdf(requestId) {
  initializeSheets();

  const sheet = getPermissionRequestsSheet();
  const id = String(requestId || '').trim();

  if (!id || !sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'ไม่พบคำขอลา' };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      let rowData = values[i];
      const statusCode = String(rowData[13] || getPermissionApprovalStatusCode(rowData[9]));

      const result = createLeavePdfForRow(rowData);

      if (result.success) {
        sheet.getRange(i + 2, 17).setValue(result.url || '');
      }

      return result;
    }
  }

  return { success: false, message: 'ไม่พบคำขอลา' };
}

function assignLeaveDocumentNumberIfNeeded(sheet, targetRow, rowData) {
  const currentDocumentNo = String(rowData[20] || '').trim();
  const currentAcademicYear = String(rowData[21] || '').trim();

  if (currentDocumentNo && currentAcademicYear) {
    return rowData;
  }

  const settings = getSettings();
  const academicYear = String(settings.ACADEMIC_YEAR || getAcademicYearThai(getBangkokDateKey(rowData[6])) || DEFAULT_SETTINGS.ACADEMIC_YEAR).trim();
  const prefix = String(settings.LEAVE_DOC_PREFIX || DEFAULT_SETTINGS.LEAVE_DOC_PREFIX).trim() || 'ผอ';
  const parsedStartNo = Number(settings.LEAVE_DOC_START_NO || DEFAULT_SETTINGS.LEAVE_DOC_START_NO || 1);
  const startNo = Number.isFinite(parsedStartNo) && parsedStartNo > 0 ? Math.floor(parsedStartNo) : 1;
  const nextNo = getNextLeaveDocumentNumber(sheet, academicYear, prefix, startNo);
  const documentNo = prefix + ' ' + nextNo + '/' + academicYear;

  sheet.getRange(targetRow, 21).setValue(documentNo);
  sheet.getRange(targetRow, 22).setValue(academicYear);
  SpreadsheetApp.flush();

  rowData[20] = documentNo;
  rowData[21] = academicYear;

  return rowData;
}

function buildNextLeaveDocumentNumber(sheet, dateKey) {
  const settings = getSettings();
  const academicYear = String(settings.ACADEMIC_YEAR || getAcademicYearThai(getBangkokDateKey(dateKey)) || DEFAULT_SETTINGS.ACADEMIC_YEAR).trim();
  const prefix = String(settings.LEAVE_DOC_PREFIX || DEFAULT_SETTINGS.LEAVE_DOC_PREFIX).trim() || 'ผอ';
  const parsedStartNo = Number(settings.LEAVE_DOC_START_NO || DEFAULT_SETTINGS.LEAVE_DOC_START_NO || 1);
  const startNo = Number.isFinite(parsedStartNo) && parsedStartNo > 0 ? Math.floor(parsedStartNo) : 1;
  const nextNo = getNextLeaveDocumentNumber(sheet, academicYear, prefix, startNo);

  return {
    documentNo: prefix + ' ' + nextNo + '/' + academicYear,
    academicYear: academicYear
  };
}

function getNextLeaveDocumentNumber(sheet, academicYear, prefix, startNo) {
  if (!sheet || sheet.getLastRow() < 2) {
    return startNo;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  let maxNo = 0;

  values.forEach((row) => {
    const rowAcademicYear = String(row[21] || '').trim();
    const documentNo = String(row[20] || '').trim();

    if (rowAcademicYear !== academicYear || !documentNo) {
      return;
    }

    const parsed = parseLeaveDocumentNumber(documentNo, prefix, academicYear);

    if (parsed > maxNo) {
      maxNo = parsed;
    }
  });

  return maxNo >= startNo ? maxNo + 1 : startNo;
}

function parseLeaveDocumentNumber(value, prefix, academicYear) {
  const escapedPrefix = String(prefix || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const escapedYear = String(academicYear || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp('^' + escapedPrefix + '\\s*(\\d+)\\s*/\\s*' + escapedYear + '$');
  const match = String(value || '').trim().match(pattern);

  if (!match) {
    return 0;
  }

  return Number(match[1] || 0) || 0;
}

function createLeavePdfForRow(rowData) {
  const requestType = String(rowData[5] || '').trim();

  if (!isLeavePermissionType(requestType)) {
    return { success: false, message: 'รายการนี้ไม่ใช่คำขอลา' };
  }

  const settings = getSettings();
  const templateId = String(settings.LEAVE_TEMPLATE_DOC_ID || LEAVE_TEMPLATE_DOC_ID).trim();

  if (!templateId) {
    return { success: false, message: 'ยังไม่ได้ตั้งค่า Template ใบลา' };
  }

  const startDate = getBangkokDateKey(rowData[6]);
  const endDate = getBangkokDateKey(rowData[14] || rowData[6]);
  const days = Number(rowData[15] || calculateInclusiveDays(startDate, endDate) || 1);
  const rootFolder = getReportFolder(settings);
  const folder = getMonthlyReportFolder(rootFolder, startDate);
  const safeName = sanitizeFileName(String(rowData[2] || 'ผู้ขอลา'));
  const filename = 'ใบลา_' + startDate + '_' + safeName + '_' + requestType;
  const templateFile = DriveApp.getFileById(templateId);
  const workingFile = templateFile.makeCopy(filename + '_draft', folder);
  const doc = DocumentApp.openById(workingFile.getId());
  const body = doc.getBody();
  const approvalNote = String(rowData[12] || '').trim();
  const userProfile = getUserProfileByPhone(rowData[3]);
  const approverProfile = getApproverSignatureProfile(rowData[10]);
  const leaveDocumentNo = String(rowData[20] || rowData[0] || '');

  replaceReportPlaceholders(body, {
    '{{leaveId}}': formatLeaveDocumentNoForPdf(leaveDocumentNo),
    '{{leaveDocumentNo}}': formatFullLeaveDocumentNoForPdf(leaveDocumentNo),
    '{{today}}': formatThaiDate(getBangkokDateKey(new Date())),
    '{{subject}}': 'ขออนุญาต' + requestType,
    '{{leaveType}}': requestType,
    '{{ใบรับรอง}}': rowData[17] ? 'มีหลักฐานแนบ' : '-',
    '{{fullName}}': String(rowData[2] || ''),
    '{{position}}': String(rowData[4] || ''),
    '{{reason}}': String(rowData[8] || '-'),
    '{{startDate}}': formatThaiDate(startDate),
    '{{endDate}}': formatThaiDate(endDate),
    '{{days}}': toThaiNumber(days),
    '{{note}}': approvalNote && approvalNote !== '-' ? approvalNote : 'อนุญาต',
  });

  replacePlaceholderWithImage(body, '{{signature}}', userProfile && userProfile.signatureFileId, '________________________', 180);
  replacePlaceholderWithImage(body, '{{sigDirector}}', approverProfile && approverProfile.signatureFileId, '________________________', 180);

  doc.saveAndClose();

  const pdfBlob = workingFile
    .getAs(MimeType.PDF)
    .setName(filename + '.pdf');

  removeExistingFilesByName(folder, filename + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  workingFile.setTrashed(true);

  return {
    success: true,
    message: 'สร้าง PDF ใบลาเรียบร้อย',
    url: pdfFile.getUrl(),
    fileName: pdfFile.getName()
  };
}

function getPermissionApprovalStatusCode(status) {
  const text = normalizeStatusText(status);

  if (text.includes('อนุมัติแล้ว')) return 'approved';
  if (text.includes('ไม่อนุมัติ')) return 'rejected';
  return 'pending';
}

function normalizeTimeInput(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{1,2})[:.](\d{2})$/);

  if (!match) return '';

  const hour = Number(match[1]);
  const minute = Number(match[2]);

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return '';

  return String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
}

function calculateInclusiveDays(startDate, endDate) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(startDate || '')) ||
      !/^\d{4}-\d{2}-\d{2}$/.test(String(endDate || ''))) {
    return 0;
  }

  const start = new Date(String(startDate) + 'T00:00:00+07:00');
  const end = new Date(String(endDate) + 'T00:00:00+07:00');
  const diff = Math.round((end.getTime() - start.getTime()) / 86400000) + 1;

  return diff > 0 ? diff : 0;
}

function sanitizeFileName(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%\{\}]/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 80) || 'file';
}

function notifyLineLatePermissionRequest(data) {
  data = data || {};
  data.requestType = 'มาสาย';
  return notifyLinePermissionRequest(data);
}

function notifyLinePermissionRequest(data) {
  data = data || {};
  const requestType = normalizePermissionRequestType(data.requestType || 'มาสาย') || 'มาสาย';
  let title = 'คำขอ';
  const lines = [];

  if (requestType === 'มาสาย') {
    title = 'ขอมาสาย';
    lines.push('เวลา: ' + formatLineTime(data.expectedTime));
    lines.push('เหตุผล: ' + String(data.reason || '-'));
    lines.push('สถานะ: ' + cleanStatusForReport(data.approvalStatus || 'รออนุมัติ'));
  } else if (isLeavePermissionType(requestType)) {
    title = 'ขอ' + requestType;
    if (data.documentNo) lines.push('เลขที่: ' + String(data.documentNo));
    lines.push('วันที่: ' + formatLineThaiDate(data.requestDate));
    lines.push('จำนวน: ' + String(data.days || '-') + ' วัน');
    lines.push('เหตุผล: ' + String(data.reason || '-'));
    lines.push(data.evidenceUrl ? 'หลักฐาน: แนบแล้ว' : 'หลักฐาน: -');
    lines.push('สถานะ: ' + cleanStatusForReport(data.approvalStatus || 'รออนุมัติ'));
  } else if (requestType === 'ไปราชการ') {
    title = 'ขอไปราชการ';
    lines.push('วันที่: ' + formatLineThaiDate(data.requestDate));
    lines.push('เหตุผล: ' + String(data.reason || '-'));
    lines.push('สถานะ: ' + cleanStatusForReport(data.approvalStatus || 'รออนุมัติ'));
  } else {
    lines.push('วันที่: ' + formatLineThaiDate(data.requestDate));
    lines.push('เหตุผล: ' + String(data.reason || '-'));
    lines.push('สถานะ: ' + cleanStatusForReport(data.approvalStatus || 'รออนุมัติ'));
  }

  const message = buildLineCard(title, lines, {
    headerName: data.name || '-',
    primaryLabel: data.evidenceUrl ? 'ดูหลักฐาน' : '',
    primaryUrl: data.evidenceUrl || ''
  });

  return sendLineMessage(message);
}

function notifyLinePermissionApproval(data) {
  data = data || {};
  const requestType = normalizePermissionRequestType(data.requestType || 'มาสาย') || 'มาสาย';
  const approved = getPermissionApprovalStatusCode(data.approvalStatus) === 'approved';
  const title = approved ? 'อนุมัติแล้ว' : 'ไม่อนุมัติ';
  const lines = [];

  if (isLeavePermissionType(requestType)) {
    if (data.documentNo) lines.push('เลขที่: ' + String(data.documentNo));
    lines.push('ประเภท: ' + requestType);
    lines.push('จำนวน: ' + String(data.days || '-') + ' วัน');
    lines.push('ผู้อนุมัติ: ' + String(data.approver || '-'));
    lines.push(data.leavePdfUrl ? 'PDF: พร้อมแล้ว' : 'PDF: -');
  } else if (requestType === 'มาสาย') {
    lines.push('เวลา: ' + formatLineTime(data.expectedTime));
    lines.push('ผล: ' + cleanStatusForReport(data.approvalStatus || ''));
    lines.push('ผู้อนุมัติ: ' + String(data.approver || '-'));
  } else {
    lines.push('ประเภท: ' + requestType);
    lines.push('ผล: ' + cleanStatusForReport(data.approvalStatus || ''));
    lines.push('ผู้อนุมัติ: ' + String(data.approver || '-'));
  }

  if (data.approvalNote && data.approvalNote !== '-') {
    lines.push('หมายเหตุ: ' + String(data.approvalNote));
  }

  const message = buildLineCard(title, lines, {
    headerName: data.name || '-',
    primaryLabel: data.leavePdfUrl ? 'ดูใบลา PDF' : '',
    primaryUrl: data.leavePdfUrl || ''
  });

  return sendLineMessage(message);
}

function normalizePermissionRequestType(value) {
  const text = String(value || '').trim();

  if (
    text === 'มาสาย' ||
    text === 'ลา' ||
    text === 'ลาป่วย' ||
    text === 'ลากิจ' ||
    text === 'ไปราชการ'
  ) {
    return text;
  }

  return '';
}

function isPermissionRequestType(value) {
  return Boolean(normalizePermissionRequestType(value));
}

function isLeavePermissionType(value) {
  const text = normalizePermissionRequestType(value);
  return text === 'ลา' || text === 'ลาป่วย' || text === 'ลากิจ';
}

function isNoCheckPermissionType(value) {
  const text = normalizePermissionRequestType(value);
  return isLeavePermissionType(text) || text === 'ไปราชการ';
}

function getPermissionRequestPrefix(requestType) {
  if (isLeavePermissionType(requestType)) return 'LV';
  if (requestType === 'ไปราชการ') return 'OD';
  return 'LR';
}
function sendLineMessage(message) {
  if (!CHANNEL_ACCESS_TOKEN || !GROUP_ID) {
    logLineError('LINE_CONFIG', 'ยังไม่ได้ตั้งค่า CHANNEL_ACCESS_TOKEN หรือ GROUP_ID');
    return {
      success: false,
      message: 'ยังไม่ได้ตั้งค่า LINE'
    };
  }

  const url = 'https://api.line.me/v2/bot/message/push';
  const lineMessage = buildLineMessagePayload(message);

  const payload = {

    to: GROUP_ID,

    messages: [lineMessage]
  };

  const options = {

    method: 'post',

    contentType: 'application/json',

    headers: {
      Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },

    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);

    return {
      success: true,
      statusCode: response.getResponseCode(),
      body: response.getContentText()
    };

  } catch (err) {
    logLineError('SEND_LINE', err.toString());

    if (message && typeof message === 'object' && message.fallbackText) {
      try {
        const fallbackPayload = {
          to: GROUP_ID,
          messages: [
            {
              type: 'text',
              text: message.fallbackText
            }
          ]
        };
        const fallbackOptions = {
          method: 'post',
          contentType: 'application/json',
          headers: {
            Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN
          },
          payload: JSON.stringify(fallbackPayload)
        };
        const fallbackResponse = UrlFetchApp.fetch(url, fallbackOptions);

        return {
          success: true,
          statusCode: fallbackResponse.getResponseCode(),
          body: fallbackResponse.getContentText(),
          fallback: true
        };
      } catch (fallbackErr) {
        logLineError('SEND_LINE_FALLBACK', fallbackErr.toString());
      }
    }

    return {
      success: false,
      message: err.toString()
    };
  }
}

function buildLineMessagePayload(message) {
  if (message && typeof message === 'object' && message.type === 'flex') {
    return {
      type: 'flex',
      altText: String(message.altText || 'แจ้งเตือนระบบลงเวลา').slice(0, 400),
      contents: message.contents
    };
  }

  return {
    type: 'text',
    text: String(message || '-')
  };
}

function notifyLineCheckin(data) {
  data = data || {};
  const timeStatus = cleanStatusForReport(data.timeStatus);
  const isLate = normalizeStatusText(timeStatus).includes('มาสาย');
  const title = isLate ? 'มาสาย' : 'ลงเวลาแล้ว';
  const lines = [
    'เวลาเข้า: ' + formatLineTime(data.checkinTime),
    'สถานะ: ' + timeStatus
  ];

  if (isLate && data.lateReason && data.lateReason !== '-') {
    lines.push('เหตุผล: ' + data.lateReason);
  }

  if (isLate) {
    lines.push('อนุมัติ: ' + displayApprovalStatusForReport(data.approvalStatus));
  }

  const message = buildLineCard(title, lines, {
    headerName: data.name || '-'
  });

  return sendLineMessage(message);
}

function notifyLineApproval(data) {
  data = data || {};
  const approved = getApprovalStatusCode(data.approvalStatus) === 'approved';
  const title = approved ? 'อนุมัติมาสาย' : 'ไม่อนุมัติ';
  const lines = [
    'เวลาเข้า: ' + formatLineTime(data.checkinTime),
    'เหตุผล: ' + String(data.reason || '-'),
    'ผู้อนุมัติ: ' + String(data.approver || '-')
  ];

  if (data.approvalNote && data.approvalNote !== '-') {
    lines.push('หมายเหตุ: ' + String(data.approvalNote));
  }

  const message = buildLineCard(title, lines, {
    headerName: data.name || '-'
  });

  return sendLineMessage(message);
}

function buildLineCard(title, lines, options) {
  options = options || {};

  const cleanLines = (lines || [])
    .map((line) => String(line || '').trim())
    .filter(Boolean);
  const cleanTitle = String(title || '\u0e41\u0e08\u0e49\u0e07\u0e40\u0e15\u0e37\u0e2d\u0e19').trim();
  const fallbackText = cleanTitle + '\n' +
    '━━━━━━━━━━━━\n' +
    cleanLines.join('\n');
  const settings = getSettings();
  const schoolName = String(settings.SCHOOL_NAME || '\u0e23\u0e30\u0e1a\u0e1a\u0e25\u0e07\u0e40\u0e27\u0e25\u0e32\u0e1b\u0e0f\u0e34\u0e1a\u0e31\u0e15\u0e34\u0e07\u0e32\u0e19').trim();
  const headerName = String(options.headerName || schoolName).trim();
  const sentAt = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') + ' \u0e19.';
  const bodyContents = cleanLines.map((line, index) => buildFlexLineRow(line, index));
  bodyContents.push({
    type: 'text',
    text: sentAt,
    color: '#94a3b8',
    size: 'xxs',
    align: 'end',
    margin: 'xs'
  });
  const footerContents = buildFlexFooterButtons(options || {});

  return {
    type: 'flex',
    altText: cleanTitle,
    fallbackText: fallbackText,
    contents: {
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#7f1d1d',
        paddingAll: '10px',
        spacing: 'xs',
        contents: [
          {
            type: 'text',
            text: headerName,
            color: '#facc15',
            weight: 'bold',
            size: 'sm',
            wrap: true,
            maxLines: 2
          },
          {
            type: 'text',
            text: '● ' + cleanTitle,
            color: '#ffffff',
            weight: 'bold',
            size: 'xs',
            maxLines: 1
          }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '12px',
        spacing: 'sm',
        contents: bodyContents.length ? bodyContents : [
          {
            type: 'text',
            text: '-',
            color: '#64748b',
            size: 'xs'
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        paddingAll: footerContents.length ? '10px' : '0px',
        spacing: 'sm',
        contents: footerContents
      },
      styles: {
        footer: {
          separator: footerContents.length > 0
        }
      }
    }
  };
}

function buildFlexLineRow(line, index) {
  const text = String(line || '').trim();
  const colonIndex = text.indexOf(':');
  const label = colonIndex > 0 ? text.slice(0, colonIndex).trim() : '';
  const shouldSplit = colonIndex > 0 &&
    colonIndex <= 14 &&
    text.indexOf(' | ') < 0 &&
    !/\d$/.test(label);

  if (shouldSplit) {
    return {
      type: 'box',
      layout: 'horizontal',
      spacing: 'sm',
      contents: [
        {
          type: 'text',
          text: label,
          color: '#94a3b8',
          size: 'xs',
          flex: 2,
          wrap: true
        },
        {
          type: 'text',
          text: text.slice(colonIndex + 1).trim() || '-',
          color: '#0f172a',
          size: 'xs',
          flex: 5,
          weight: 'bold',
          wrap: true
        }
      ]
    };
  }

  return {
    type: 'text',
    text: text,
    color: '#0f172a',
    size: index === 0 ? 'sm' : 'xs',
    weight: index === 0 ? 'bold' : 'regular',
    wrap: true
  };
}

function buildFlexFooterButtons(options) {
  const buttons = [];

  addFlexUriButton(buttons, options.primaryLabel, options.primaryUrl, '#f59e0b');
  addFlexUriButton(buttons, options.secondaryLabel, options.secondaryUrl, '#7f1d1d');
  addFlexUriButton(buttons, options.evidenceLabel, options.evidenceUrl, '#334155');

  return buttons;
}

function addFlexUriButton(buttons, label, url, color) {
  if (!label || !url) {
    return;
  }

  buttons.push({
    type: 'button',
    style: 'primary',
    height: 'sm',
    color: color || '#f59e0b',
    action: {
      type: 'uri',
      label: String(label).slice(0, 20),
      uri: String(url)
    }
  });
}
function formatLineTime(value) {
  const text = formatTimeValue(value, '');

  if (!text) {
    return '';
  }

  const match = String(text).match(/(\d{1,2}):(\d{2})/);

  if (!match) {
    return String(text);
  }

  return match[1].padStart(2, '0') + ':' + match[2] + ' น.';
}

function formatLineThaiDate(value) {
  const date =
    Object.prototype.toString.call(value) === '[object Date]'
      ? value
      : new Date(value);

  if (isNaN(date.getTime())) {
    return '';
  }

  const days = [
    'วันอาทิตย์',
    'วันจันทร์',
    'วันอังคาร',
    'วันพุธ',
    'วันพฤหัสบดี',
    'วันศุกร์',
    'วันเสาร์'
  ];
  const months = [
    'ม.ค.',
    'ก.พ.',
    'มี.ค.',
    'เม.ย.',
    'พ.ค.',
    'มิ.ย.',
    'ก.ค.',
    'ส.ค.',
    'ก.ย.',
    'ต.ค.',
    'พ.ย.',
    'ธ.ค.'
  ];
  const dayName = days[Number(Utilities.formatDate(date, 'Asia/Bangkok', 'u')) % 7];
  const day = Utilities.formatDate(date, 'Asia/Bangkok', 'd');
  const monthIndex = Number(Utilities.formatDate(date, 'Asia/Bangkok', 'M')) - 1;
  const year = String(Number(Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy')) + 543).slice(-2);

  return dayName + 'ที่ ' + day + ' ' + months[monthIndex] + ' ' + year;
}

function logLineError(type, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('LineLogs');

    if (!sheet) {
      sheet = ss.insertSheet('LineLogs');
      sheet.appendRow(['Time', 'Type', 'Detail']);
    }

    sheet.appendRow([
      new Date(),
      type,
      detail
    ]);

  } catch (err) {
    // ไม่ให้ logging ทำให้ workflow หลักล้ม
  }
}

function sendDailySummary() {

  const data = getDashboardData();

  const message = buildLineCard('📊 สรุปวันนี้', [
    formatLineThaiDate(data.date),
    'ทั้งหมด ' + data.totalActive + ' | ลงเวลา ' + data.checkedIn,
    'ปกติ ' + data.normal + ' | มาสาย ' + data.late,
    'ลา ' + data.leave + ' | ไปราชการ ' + data.officialDuty,
    'ยังไม่ลงเวลา ' + data.absent
  ]);

  sendLineMessage(message);
}

function sendMorningCheckinAlert() {
  initializeSheets();

  const data = getDashboardData();
  const settings = getSettings();
  const startTime = formatLineTime(settings.START_TIME || DEFAULT_SETTINGS.START_TIME);
  const absentText = data.absentNames && data.absentNames.length
    ? data.absentNames.slice(0, 12).join(', ')
    : 'ไม่มี';
  const noCheckTextParts = [];

  if (data.sickLeave) noCheckTextParts.push('ลาป่วย ' + data.sickLeave);
  if (data.personalLeave) noCheckTextParts.push('ลากิจ ' + data.personalLeave);
  if (data.officialDuty) noCheckTextParts.push('ไปราชการ ' + data.officialDuty);

  const message = buildLineCard('⏰ สรุปหลังเวลาเข้า', [
    formatLineThaiDate(data.date) + ' | เวลาเข้า ' + startTime,
    'ทั้งหมด ' + data.totalActive + ' | ลงเวลา ' + data.checkedIn,
    'ปกติ ' + data.normal + ' | มาสาย ' + data.late,
    noCheckTextParts.length ? noCheckTextParts.join(' | ') : 'ลา/ไปราชการ 0',
    'ยังไม่ลงเวลา ' + data.absent,
    'รายชื่อ: ' + absentText
  ]);

  return sendLineMessage(message);
}

function sendTomorrowPermissionAlert() {
  initializeSheets();

  const tomorrow = getRelativeBangkokDateKey(1);
  const items = getApprovedNoCheckPermissionsForDate(tomorrow);

  if (!items.length) {
    return {
      success: true,
      skipped: true,
      message: '\u0e1e\u0e23\u0e38\u0e48\u0e07\u0e19\u0e35\u0e49\u0e44\u0e21\u0e48\u0e21\u0e35\u0e1c\u0e39\u0e49\u0e25\u0e32\u0e2b\u0e23\u0e37\u0e2d\u0e44\u0e1b\u0e23\u0e32\u0e0a\u0e01\u0e32\u0e23'
    };
  }

  const lines = [
    '\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48: ' + formatLineThaiDate(tomorrow),
    '\u0e08\u0e33\u0e19\u0e27\u0e19: ' + items.length + ' \u0e04\u0e19'
  ];

  items.slice(0, 12).forEach((item, index) => {
    const reason = item.reason && item.reason !== '-'
      ? ' - ' + item.reason
      : '';
    lines.push((index + 1) + '. ' + item.name + ' | ' + item.requestType + reason);
  });

  if (items.length > 12) {
    lines.push('\u0e41\u0e25\u0e30\u0e2d\u0e35\u0e01 ' + (items.length - 12) + ' \u0e04\u0e19');
  }

  const message = buildLineCard('\u0e1e\u0e23\u0e38\u0e48\u0e07\u0e19\u0e35\u0e49\u0e21\u0e35\u0e1c\u0e39\u0e49\u0e25\u0e32/\u0e44\u0e1b\u0e23\u0e32\u0e0a\u0e01\u0e32\u0e23', lines, {
    headerName: '\u0e41\u0e08\u0e49\u0e07\u0e40\u0e15\u0e37\u0e2d\u0e19\u0e25\u0e48\u0e27\u0e07\u0e2b\u0e19\u0e49\u0e32'
  });

  return sendLineMessage(message);
}

function getApprovedNoCheckPermissionsForDate(dateKey) {
  const sheet = getPermissionRequestsSheet();

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const targetDate = getBangkokDateKey(dateKey);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMISSION_REQUEST_HEADERS.length).getValues();
  const items = [];

  values.forEach((row) => {
    const requestType = normalizePermissionRequestType(row[5]);

    if (!isNoCheckPermissionType(requestType)) {
      return;
    }

    const statusCode = String(row[13] || getPermissionApprovalStatusCode(row[9]));

    if (statusCode !== 'approved') {
      return;
    }

    const startDate = getBangkokDateKey(row[6]);
    const endDate = getBangkokDateKey(row[14] || row[6]);

    if (!startDate || targetDate < startDate || targetDate > endDate) {
      return;
    }

    items.push({
      name: String(row[2] || '-'),
      requestType: cleanStatusForReport(requestType) || '-',
      startDate: startDate,
      endDate: endDate,
      days: String(row[15] || calculateInclusiveDays(startDate, endDate) || ''),
      reason: String(row[8] || '-'),
      documentNo: String(row[20] || '')
    });
  });

  items.sort((a, b) =>
    String(a.requestType || '').localeCompare(String(b.requestType || '')) ||
    String(a.name || '').localeCompare(String(b.name || ''))
  );

  return items;
}

function getRelativeBangkokDateKey(offsetDays) {
  const now = new Date();
  now.setDate(now.getDate() + Number(offsetDays || 0));

  return Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function checkTodayCheckin(phone) {
  return checkTodayCheckinForUser({
    phone: phone
  });
}

function checkTodayCheckinForUser(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CHECKIN_SHEET);

  if (!sheet) {
    return {
      checkedIn: false
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      checkedIn: false
    };
  }

  const maxRowsToCheck = 2000;
  const startRow = Math.max(2, lastRow - maxRowsToCheck + 1);
  const rowCount = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, rowCount, 12).getValues();

  const today = getBangkokDateKey(new Date());

  const targetPhone = normalizePhone(user && user.phone);
  const targetName = normalizeName(user && user.name);

  for (let i = data.length - 1; i >= 0; i--) {
    const rowDate = getBangkokDateKey(data[i][0]);

    const rowPhone = normalizePhone(data[i][2]);
    const rowName = normalizeName(data[i][1]);
    const samePhone = isSamePhone(rowPhone, targetPhone);
    const sameName = targetName && rowName === targetName;

    if (rowDate === today && (samePhone || sameName)) {
      return {
        checkedIn: true,
        time: formatTimeValue(data[i][7], ''),
        autoCheckoutTime: formatTimeValue(data[i][8], ''),
        timeStatus: String(data[i][9] || ''),
        approvalStatus: String(data[i][10] || ''),
        reason: String(data[i][11] || '')
      };
    }
  }

  return {
    checkedIn: false
  };
}

function testLine() {

  return sendLineMessage(buildLineCard('LINE Flex Test', [
    'System: Checkin',
    'Status: Ready',
    'Type: Flex Message'
  ]));
}
function normalizePhone(phone) {
  return String(phone || '')
    .trim()
    .replace(/[^\d]/g, '');
}

function formatPhoneForSheet(phone) {
  const normalized = normalizePhone(phone);

  if (!normalized) {
    return '';
  }

  return "'" + normalized;
}

function normalizeName(name) {
  return String(name || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function isSamePhone(a, b) {
  const phoneA = normalizePhone(a);
  const phoneB = normalizePhone(b);

  if (!phoneA || !phoneB) {
    return false;
  }

  if (phoneA === phoneB) {
    return true;
  }

  const tailA = phoneA.slice(-9);
  const tailB = phoneB.slice(-9);

  return tailA.length >= 9 && tailA === tailB;
}

function getBangkokDateKey(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM-dd');
  }

  const text = String(value).trim();

  if (!text) {
    return '';
  }

  const parsed = new Date(text);

  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'Asia/Bangkok', 'yyyy-MM-dd');
  }

  const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);

  if (match) {
    return match[1] + '-' +
      String(match[2]).padStart(2, '0') + '-' +
      String(match[3]).padStart(2, '0');
  }

  return '';
}

function getSettings() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET);
  const values = sheet.getDataRange().getValues();

  const settings = Object.assign({}, DEFAULT_SETTINGS);

  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    const value = values[i][1];

    if (key) {
      if (String(key).trim() === 'START_TIME') {
        settings[key] = formatTimeValue(value, DEFAULT_SETTINGS.START_TIME);
      } else if (String(key).trim() === 'REPORT_FOLDER_ID') {
        settings[key] = String(value || DEFAULT_SETTINGS.REPORT_FOLDER_ID);
      } else {
        settings[key] = String(value);
      }
    }
  }

  return settings;
}

function getSystemSettingsForUi() {
  initializeSheets();

  const settings = getSettings();

  return {
    success: true,
    settings: {
      SCHOOL_NAME: String(settings.SCHOOL_NAME || ''),
      SCHOOL_LOGO_URL: String(settings.SCHOOL_LOGO_URL || ''),
      SCHOOL_LOGO_DATA: String(settings.SCHOOL_LOGO_DATA || ''),
      SCHOOL_LOGO_SIZE_PERCENT: String(settings.SCHOOL_LOGO_SIZE_PERCENT || DEFAULT_SETTINGS.SCHOOL_LOGO_SIZE_PERCENT),
      APP_VERSION: String(settings.APP_VERSION || DEFAULT_SETTINGS.APP_VERSION),
      REPORT_TEMPLATE_DOC_ID: String(settings.REPORT_TEMPLATE_DOC_ID || DEFAULT_SETTINGS.REPORT_TEMPLATE_DOC_ID),
      LEAVE_TEMPLATE_DOC_ID: String(settings.LEAVE_TEMPLATE_DOC_ID || DEFAULT_SETTINGS.LEAVE_TEMPLATE_DOC_ID),
      REPORT_FOLDER_ID: String(settings.REPORT_FOLDER_ID || ''),
      SCHOOL_LAT: String(settings.SCHOOL_LAT || ''),
      SCHOOL_LNG: String(settings.SCHOOL_LNG || ''),
      ALLOW_RADIUS: String(settings.ALLOW_RADIUS || ''),
      START_TIME: String(settings.START_TIME || ''),
      ACADEMIC_YEAR: String(settings.ACADEMIC_YEAR || DEFAULT_SETTINGS.ACADEMIC_YEAR),
      CURRENT_TERM: String(settings.CURRENT_TERM || DEFAULT_SETTINGS.CURRENT_TERM),
      TERM_START_DATE: String(settings.TERM_START_DATE || ''),
      TERM_END_DATE: String(settings.TERM_END_DATE || ''),
      LEAVE_DOC_PREFIX: String(settings.LEAVE_DOC_PREFIX || DEFAULT_SETTINGS.LEAVE_DOC_PREFIX),
      LEAVE_DOC_START_NO: String(settings.LEAVE_DOC_START_NO || DEFAULT_SETTINGS.LEAVE_DOC_START_NO)
    },
    positions: getPositionSettingsForUi()
  };
}

function saveSystemSettings(payload) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  const settings = payload && payload.settings ? payload.settings : {};

  upsertSettingIfPresent(settingsSheet, settings, 'SCHOOL_NAME', 'ชื่อโรงเรียน');
  upsertSettingIfPresent(settingsSheet, settings, 'SCHOOL_LOGO_URL', 'URL โลโก้โรงเรียน');
  upsertSettingIfPresent(settingsSheet, settings, 'SCHOOL_LOGO_SIZE_PERCENT', 'ขนาดโลโก้ หน่วยเปอร์เซ็นต์');
  upsertSettingIfPresent(settingsSheet, settings, 'APP_VERSION', 'Version ของแอพ');
  upsertSettingIfPresent(settingsSheet, settings, 'REPORT_TEMPLATE_DOC_ID', 'รหัส Google Docs Template รายงาน');
  upsertSettingIfPresent(settingsSheet, settings, 'LEAVE_TEMPLATE_DOC_ID', 'รหัส Google Docs Template ใบลา');
  upsertSettingIfPresent(settingsSheet, settings, 'REPORT_FOLDER_ID', 'รหัสโฟลเดอร์เก็บ PDF รายงาน');

  if (settings.SCHOOL_LOGO_DATA !== undefined) {
    if (String(settings.SCHOOL_LOGO_DATA || '').length > 49000) {
      return {
        success: false,
        message: 'ไฟล์โลโก้มีขนาดใหญ่เกินไป กรุณาเลือกรูปขนาดเล็กลง'
      };
    }

    upsertSetting(settingsSheet, 'SCHOOL_LOGO_DATA', settings.SCHOOL_LOGO_DATA, 'ข้อมูลรูปโลโก้โรงเรียน');
  }

  upsertSettingIfPresent(settingsSheet, settings, 'SCHOOL_LAT', 'ละติจูดโรงเรียน');
  upsertSettingIfPresent(settingsSheet, settings, 'SCHOOL_LNG', 'ลองจิจูดโรงเรียน');
  upsertSettingIfPresent(settingsSheet, settings, 'ALLOW_RADIUS', 'รัศมีที่อนุญาต หน่วยเมตร');
  upsertSettingIfPresent(settingsSheet, settings, 'ACADEMIC_YEAR', 'ปีการศึกษาปัจจุบัน');
  upsertSettingIfPresent(settingsSheet, settings, 'CURRENT_TERM', 'ภาคเรียนปัจจุบัน');
  upsertSettingIfPresent(settingsSheet, settings, 'TERM_START_DATE', 'วันเริ่มภาคเรียน');
  upsertSettingIfPresent(settingsSheet, settings, 'TERM_END_DATE', 'วันสิ้นสุดภาคเรียน');
  upsertSettingIfPresent(settingsSheet, settings, 'LEAVE_DOC_PREFIX', 'คำนำหน้าเลขที่หนังสือใบลา');
  upsertSettingIfPresent(settingsSheet, settings, 'LEAVE_DOC_START_NO', 'เลขเริ่มต้นใบลา');
  if (Object.prototype.hasOwnProperty.call(settings, 'START_TIME')) {
    upsertSetting(settingsSheet, 'START_TIME', formatTimeValue(settings.START_TIME, DEFAULT_SETTINGS.START_TIME), 'เวลาเริ่มงาน');
  }

  if (payload && payload.positions) {
    savePositionSettings(payload.positions);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'บันทึกการตั้งค่าเรียบร้อย'
  };
}

function upsertSettingIfPresent(sheet, settings, key, description) {
  if (Object.prototype.hasOwnProperty.call(settings, key)) {
    upsertSetting(sheet, key, settings[key], description);
  }
}

function upsertSetting(sheet, key, value, description) {
  const values = sheet.getDataRange().getValues();
  const cleanValue = String(value || '').trim();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(cleanValue);
      sheet.getRange(i + 1, 3).setValue(description || values[i][2] || '');
      return;
    }
  }

  sheet.appendRow([key, cleanValue, description || '']);
}

function getPublicAppSettings() {
  const settings = getSettings();

  return {
    success: true,
    schoolName: String(settings.SCHOOL_NAME || DEFAULT_SETTINGS.SCHOOL_NAME),
    schoolLogoUrl: String(settings.SCHOOL_LOGO_URL || ''),
    schoolLogoData: String(settings.SCHOOL_LOGO_DATA || ''),
    schoolLogoSizePercent: String(settings.SCHOOL_LOGO_SIZE_PERCENT || DEFAULT_SETTINGS.SCHOOL_LOGO_SIZE_PERCENT),
    appVersion: String(settings.APP_VERSION || DEFAULT_SETTINGS.APP_VERSION)
  };
}

function getPositionSettingsForUi() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSITION_SETTINGS_SHEET);
  const values = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const position = String(values[i][0] || '').trim();

    if (position) {
      result.push({
        position: position,
        checkoutTime: formatTimeValue(values[i][1], '16:30')
      });
    }
  }

  return result;
}

function savePositionSettings(positions) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSITION_SETTINGS_SHEET);
  sheet.getRange('B:B').setNumberFormat('@');

  const current = sheet.getDataRange().getValues();

  const map = {};

  for (let i = 0; i < positions.length; i++) {
    const position = String(positions[i].position || '').trim();
    const checkoutTime = formatTimeValue(positions[i].checkoutTime, '');

    if (position && checkoutTime) {
      map[position] = checkoutTime;
    }
  }

  for (let i = 1; i < current.length; i++) {
    const position = String(current[i][0] || '').trim();

    if (position && map[position]) {
      sheet.getRange(i + 1, 2).setValue(map[position]);
      delete map[position];
    }
  }

  Object.keys(map).forEach((position) => {
    sheet.appendRow([position, map[position]]);
  });
}

function getAutoCheckoutTime(position) {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(POSITION_SETTINGS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(POSITION_SETTINGS_SHEET);
    sheet.appendRow(['ตำแหน่ง', 'เวลาออกอัตโนมัติ']);
    sheet.appendRow(['ครู', '16:30']);
    sheet.appendRow(['ภารโรง', '18:00']);
    sheet.appendRow(['เจ้าหน้าที่', '16:30']);
    sheet.appendRow(['ผู้บริหาร', '16:30']);
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(position).trim()) {
      return formatTimeValue(values[i][1], '16:30');
    }
  }

  return '16:30';
}

function timeToMinutes(timeText) {
  const parts = formatTimeValue(timeText, '00:00').split(':');
  const hour = Number(parts[0] || 0);
  const minute = Number(parts[1] || 0);

  if (isNaN(hour) || isNaN(minute)) {
    return 0;
  }

  return hour * 60 + minute;
}

function formatTimeValue(value, fallback) {
  if (!value) {
    return fallback || '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      'Asia/Bangkok',
      'HH:mm'
    );
  }

  if (typeof value === 'number') {
    if (value > 0 && value < 1) {
      const totalMinutes = Math.round(value * 24 * 60);
      const hourFromSheetTime = Math.floor(totalMinutes / 60) % 24;
      const minuteFromSheetTime = totalMinutes % 60;

      return padTimePart(hourFromSheetTime) + ':' + padTimePart(minuteFromSheetTime);
    }

    return formatClockText(String(value), fallback);
  }

  return formatClockText(String(value).trim(), fallback);
}

function formatClockText(value, fallback) {
  const text = String(value || '').trim();

  if (!text) {
    return fallback || '';
  }

  const match = text.match(/^(\d{1,2})(?::|\.)(\d{1,2})(?::\d{1,2})?$/);

  if (match) {
    const hour = Number(match[1]);
    let minuteText = match[2];

    if (minuteText.length === 1) {
      minuteText = minuteText + '0';
    }

    const minute = Number(minuteText);

    if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
      return padTimePart(hour) + ':' + padTimePart(minute);
    }
  }

  const dateTimeMatch = text.match(/\b(\d{1,2}):(\d{2})(?::\d{2})?\b/);

  if (dateTimeMatch) {
    const hour = Number(dateTimeMatch[1]);
    const minute = Number(dateTimeMatch[2]);

    if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
      return padTimePart(hour) + ':' + padTimePart(minute);
    }
  }

  const compactMatch = text.match(/^(\d{1,2})$/);

  if (compactMatch) {
    const hour = Number(compactMatch[1]);

    if (hour >= 0 && hour <= 23) {
      return padTimePart(hour) + ':00';
    }
  }

  return text;
}

function padTimePart(value) {
  return String(Number(value || 0)).padStart(2, '0');
}

function getPendingLateRequests() {
  initializeSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CHECKIN_SHEET);

  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const approvalStatus = normalizeStatusText(row[10]);

    if (approvalStatus.includes('รออนุมัติ')) {
      result.push({
        rowNumber: i + 1,
        date: formatDateValue(row[0]),
        name: String(row[1] || ''),
        phone: normalizePhone(row[2]),
        position: String(row[3] || ''),
        checkinTime: formatTimeValue(row[7], ''),
        timeStatus: String(row[9] || ''),
        approvalStatus: String(row[10] || ''),
        reason: String(row[11] || '-')
      });
    }
  }

  return result;
}

function updateLateApproval(rowNumber, newStatus, approverName, note) {
  initializeSheets();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CHECKIN_SHEET);

    if (!sheet) {
      return {
        success: false,
        message: 'ไม่พบชีต CheckinData'
      };
    }

    const row = Number(rowNumber);

    if (!row || row < 2 || row > sheet.getLastRow()) {
      return {
        success: false,
        message: 'ไม่พบรายการที่ต้องการอนุมัติ'
      };
    }

    const currentStatus = normalizeStatusText(
      sheet.getRange(row, 11).getValue()
    );

    if (!currentStatus.includes('รออนุมัติ')) {
      return {
        success: false,
        message: 'รายการนี้ไม่ได้อยู่ในสถานะรออนุมัติแล้ว'
      };
    }

    const now = new Date();
    const approvalTime = Utilities.formatDate(
      now,
      'Asia/Bangkok',
      'yyyy-MM-dd HH:mm:ss'
    );

    sheet.getRange(row, 11).setValue(newStatus);
    sheet.getRange(row, 13).setValue(String(approverName || 'ผู้บริหาร'));
    sheet.getRange(row, 14).setValue(approvalTime);
    sheet.getRange(row, 15).setValue(String(note || '-'));
    sheet.getRange(row, 17).setValue(getApprovalStatusCode(newStatus));

    SpreadsheetApp.flush();

    const rowData = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 17)).getValues()[0];

    notifyLineApproval({
      name: rowData[1],
      position: rowData[3],
      checkinTime: rowData[7],
      timeStatus: rowData[9],
      approvalStatus: newStatus,
      reason: rowData[11],
      approver: approverName || 'ผู้บริหาร',
      approvalNote: note || '-',
      approvalTime: approvalTime
    });

    return {
      success: true,
      message: 'อัปเดตสถานะเรียบร้อย',
      rowNumber: row,
      approvalStatus: newStatus,
      approvalTime: approvalTime
    };

  } finally {
    lock.releaseLock();
  }
}

function approveLateRequest(rowNumber, approverName, note) {
  return updateLateApproval(
    rowNumber,
    '🟢 อนุมัติแล้ว',
    approverName,
    note
  );
}

function rejectLateRequest(rowNumber, approverName, note) {
  return updateLateApproval(
    rowNumber,
    '🔴 ไม่อนุมัติ',
    approverName,
    note
  );
}

function formatDateValue(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(
      value,
      'Asia/Bangkok',
      'yyyy-MM-dd'
    );
  }

  return String(value);
}

function normalizeStatusText(value) {
  return String(value || '')
    .replace(/[🟢🔴🟡✅❌]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
