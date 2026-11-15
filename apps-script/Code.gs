// ============================================================
// ระบบนิเทศออนไลน์ สำนักงานเขตพื้นที่การศึกษามัธยมศึกษานราธิวาส
// Google Apps Script Backend
// ============================================================

const SPREADSHEET_ID = '1ojkushQJ6-ueIvmTwhxqOQbsYAXMzVNMKS-UcyKfAQ0';
const DRIVE_FOLDER_ID = '1FDj-gDOWuvq8A_kugvPrBhPZ8X4pHVUo';
const SETTINGS_SHEET_NAME = 'Settings';
const SETTINGS_CACHE_KEY = 'nited_system_settings_v1';
const DEFAULT_SYSTEM_SETTINGS = [
  ['SYSTEM_NAME', 'ระบบนิเทศออนไลน์', 'ชื่อระบบ'],
  ['OFFICE_NAME', 'สำนักงานเขตพื้นที่การศึกษามัธยมศึกษานราธิวาส', 'ชื่อหน่วยงาน'],
  ['TIME_ZONE', 'Asia/Bangkok', 'เขตเวลา'],
  ['DRIVE_FOLDER_ID', DRIVE_FOLDER_ID, 'รหัสโฟลเดอร์หลักสำหรับเก็บไฟล์'],
  ['BOOKING_PENDING', 'รอดำเนินการ', 'สถานะกำหนดนิเทศ: รอดำเนินการ'],
  ['BOOKING_CONFIRMED', 'ยืนยันแล้ว', 'สถานะกำหนดนิเทศ: ยืนยันแล้ว'],
  ['BOOKING_REJECTED', 'ปฏิเสธ', 'สถานะกำหนดนิเทศ: ปฏิเสธ'],
  ['BOOKING_COMPLETED', 'นิเทศสำเร็จ', 'สถานะกำหนดนิเทศ: นิเทศสำเร็จ'],
  ['FILE_PENDING', 'รอตรวจสอบ', 'สถานะหลักฐาน: รอตรวจสอบ'],
  ['FILE_APPROVED', 'ผ่าน', 'สถานะหลักฐาน: ผ่าน'],
  ['FILE_REVISE', 'ปรับปรุง', 'สถานะหลักฐาน: ปรับปรุง'],
  ['QUALITY_EXCELLENT', 'ดีมาก', 'ระดับคุณภาพ: ดีมาก'],
  ['QUALITY_GOOD', 'ดี', 'ระดับคุณภาพ: ดี'],
  ['QUALITY_FAIR', 'พอใช้', 'ระดับคุณภาพ: พอใช้'],
  ['QUALITY_POOR', 'ปรับปรุง', 'ระดับคุณภาพ: ปรับปรุง']
];
// ตั้งค่า NITED_ADMIN_PASSWORD ใน Script Properties ก่อนเปิดใช้งานจริง
const ADMIN_PASSWORD_PROPERTY = 'NITED_ADMIN_PASSWORD';

const SHEETS = {
  BOOKING: 'Booking',
  FILES: 'Files',
  EVALUATION: 'Inspections',
  SUPERVISION: 'Supervision',
  USERS: 'Users',
  SCHOOLS: 'Schools',
  LOGFILES: 'logfiles'
};

const STATUS = {
  PENDING: 'รอดำเนินการ',
  CONFIRMED: 'ยืนยันแล้ว',
  REJECTED: 'ปฏิเสธ',
  COMPLETED: 'นิเทศสำเร็จ'
};

const FILE_STATUS = {
  PENDING: 'รอตรวจสอบ',
  APPROVED: 'ผ่าน',
  REVISE: 'ปรับปรุง'
};

const QUALITY_LEVELS = {
  EXCELLENT: 'ดีมาก',
  GOOD: 'ดี',
  FAIR: 'พอใช้',
  POOR: 'ปรับปรุง'
};

// ============================================================
// Web App Entry Points
// ============================================================

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'publicSchools') {
    const callback = String(e.parameter.callback || '');
    if (!/^[A-Za-z_$][\w$]{0,80}$/.test(callback)) {
      return ContentService.createTextOutput('/* Invalid callback */')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    const result = getPublicSchools_();
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  const page = e.parameter.page || 'index';
  const template = HtmlService.createTemplateFromFile(page);
  const html = template.evaluate()
    .setTitle(getSystemSetting_('SYSTEM_NAME', 'ระบบนิเทศออนไลน์') + ' สพม.นราธิวาส')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// Public directory exposes school location, but excludes personal contact fields.
function getPublicSchools_() {
  try {
    const rows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
    const columns = getSchoolColumns_(rows[0] || []);
    const schools = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const schoolId = schoolValue_(row, columns, 'schoolId');
      if (!schoolId || !isSchoolActive_(schoolValue_(row, columns, 'active'))) continue;
      schools.push({
        schoolId: schoolId,
        name: schoolValue_(row, columns, 'name'),
        code: schoolValue_(row, columns, 'code'),
        district: schoolValue_(row, columns, 'district'),
        director: schoolValue_(row, columns, 'director'),
        address: schoolValue_(row, columns, 'address'),
        website: schoolValue_(row, columns, 'website'),
        latitude: schoolValue_(row, columns, 'latitude'),
        longitude: schoolValue_(row, columns, 'longitude'),
        coordinates: schoolValue_(row, columns, 'coordinates')
      });
    }
    schools.sort((a, b) => String(a.name).localeCompare(String(b.name), 'th'));
    return { success: true, schools };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function doPost(e) {
  try {
    getSystemSettings_();
    const data = typeof e === 'string' ? JSON.parse(e) : JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'login') return jsonResponse(login(data));
    if (action === 'forgotPassword') return jsonResponse(requestPasswordReset_(data));
    if (action === 'logout') return jsonResponse(logout_(data));
    const auth = authenticateRequest_(data);
    if (!auth.success) return jsonResponse(auth);
    if (auth.user.mustResetPassword && !['changePassword','logout'].includes(action)) {
      return jsonResponse({success:false,message:'กรุณากำหนดรหัสผ่านใหม่ก่อนใช้งานระบบ',passwordChangeRequired:true});
    }
    data._user = auth.user;
    if (auth.user.role === 'school') data.schoolId = auth.user.schoolId;
    auditLog_(auth.user, action, 'request', 'รับคำขอ API');
    if (!['admin','supervisor','school'].includes(String(auth.user.role || '').toLowerCase())) return jsonResponse({success:false,message:'บัญชีนี้ไม่มีสิทธิ์ใช้ระบบ'});
    if (action === 'changePassword') {
      const result = changePassword_(data, auth.user);
      if (result.success) {
        const updatedUser = Object.assign({}, auth.user, {mustResetPassword:false});
        CacheService.getScriptCache().put('nited_session_' + String(data.token || ''), JSON.stringify(updatedUser), 21600);
        result.user = updatedUser;
      }
      return jsonResponse(result);
    }
    const supervisorActions = ['createEvaluation','updateEvaluation'];
    const adminActions = ['createSchool','createUser','updateFileStatus','deleteEvaluation','setSchoolPassword','getTeachers','getTeacherReport','getDepartmentReport'];
    if (supervisorActions.includes(action) && !isPrivileged_(auth.user)) return jsonResponse({success:false,message:'ไม่มีสิทธิ์ดำเนินการนี้'});
    if (adminActions.includes(action) && auth.user.role !== 'admin') return jsonResponse({success:false,message:'เมนูผู้ดูแลระบบสำหรับบัญชีผู้ดูแลระบบเท่านั้น'});
    if (action === 'updateSchool' && !['admin','school'].includes(auth.user.role)) return jsonResponse({success:false,message:'ไม่มีสิทธิ์ดำเนินการนี้'});
    if (['deleteBooking','deleteFile'].includes(action) && !['admin','school'].includes(auth.user.role)) return jsonResponse({success:false,message:'ไม่มีสิทธิ์ดำเนินการนี้'});
    if (action === 'updateBooking' && data.status !== undefined && !isPrivileged_(auth.user)) return jsonResponse({success:false,message:'เฉพาะผู้ดูแลระบบหรือผู้นิเทศเท่านั้นที่เปลี่ยนสถานะกำหนดนิเทศได้'});
    if (auth.user.role === 'school' && !['createBooking','updateBooking','deleteBooking','getBookings','getBookingById','uploadFile','getFiles','deleteFile','getEvaluations','getDashboardStats','getCalendarData','getReport','getTeacherReport','getDepartmentReport','getSchools','updateSchool'].includes(action)) return jsonResponse({success:false,message:'ไม่มีสิทธิ์ดำเนินการนี้'});

    switch (action) {
      case 'createBooking': return jsonResponse(createBooking(data));
      case 'updateBooking': return jsonResponse(updateBooking(data));
      case 'deleteBooking': return jsonResponse(deleteBooking(data));
      case 'getBookings': return jsonResponse(getBookings(data));
      case 'getBookingById': return jsonResponse(getBookingById(data));
      case 'uploadFile': return jsonResponse(uploadFile(data));
      case 'updateFileStatus': return jsonResponse(updateFileStatus(data));
      case 'getFiles': return jsonResponse(getFiles(data));
      case 'deleteFile': return jsonResponse(deleteFile(data));
      case 'createEvaluation': return jsonResponse(createEvaluation(data));
      case 'updateEvaluation': return jsonResponse(updateEvaluation(data));
      case 'getEvaluations': return jsonResponse(getEvaluations(data));
      case 'deleteEvaluation': return jsonResponse(deleteEvaluation(data));
      case 'getDashboardStats': return jsonResponse(getDashboardStats(data));
      case 'getCalendarData': return jsonResponse(getCalendarData(data));
      case 'getReport': return jsonResponse(getReport(data));
      case 'getTeacherReport': return jsonResponse(getTeacherReport(data));
      case 'getDepartmentReport': return jsonResponse(getDepartmentReport(data));
      case 'getSchools': return jsonResponse(getSchools(data));
      case 'getTeachers': return jsonResponse(getTeachers());
      case 'createSchool': return jsonResponse(createSchool(data));
      case 'updateSchool': return jsonResponse(updateSchool(data));
      case 'setSchoolPassword': return jsonResponse(setSchoolPassword_(data));
      default: return jsonResponse({ success: false, message: 'ไม่รู้จักคำสั่ง' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function authenticateRequest_(data) {
  const token = String(data.token || '');
  if (!token) return {success:false,message:'กรุณาเข้าสู่ระบบใหม่',authRequired:true};
  const raw = CacheService.getScriptCache().get('nited_session_' + token);
  if (!raw) return {success:false,message:'หมดเวลาเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่',authRequired:true};
  try { return {success:true,user:JSON.parse(raw)}; } catch (_) { return {success:false,message:'session ไม่ถูกต้อง',authRequired:true}; }
}

function isPrivileged_(user) { return ['admin','supervisor'].includes(String(user.role || '').toLowerCase()); }

function isBookingCompleted_(status) {
  return status === STATUS.COMPLETED || status === 'นิเทศแล้ว' || status === 'นิเทศสำเร็จ';
}

function auditLog_(user, action, result, details) {
  try {
    const sheet = getSheet(SHEETS.LOGFILES);
    sheet.appendRow([formatDateTime(new Date()), safeLogCell_(user.username), safeLogCell_(user.role), safeLogCell_(user.schoolCode), safeLogCell_(action), safeLogCell_(result), safeLogCell_(details, 500)]);
  } catch (err) { console.error('audit log failed: ' + err); }
}

function safeLogCell_(value, maxLength) {
  let text = String(value == null ? '' : value).slice(0, maxLength || 200);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// Utility Functions
// ============================================================

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  getSystemSettings_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheet(sheet, name);
  }
  if (name === SETTINGS_SHEET_NAME) return sheet;
  ensureSheetHeaders(sheet, name);
  if (name === SHEETS.BOOKING) ensureBookingSchema_(sheet);
  return sheet;
}

// Central configuration lives in the Settings sheet as Key / Value / Description.
// The spreadsheet ID remains a bootstrap constant because it is required to open
// the spreadsheet that contains Settings in the first place.
function getSystemSettings_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SETTINGS_CACHE_KEY);
  if (cached) {
    try {
      const values = JSON.parse(cached);
      applySystemSettings_(values);
      return values;
    } catch (_) {}
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['Setting Key', 'Setting Value', 'Description']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1a5276').setFontColor('white');
  }
  const existing = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
  const values = {};
  existing.forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) values[key] = String(row[1] == null ? '' : row[1]).trim();
  });
  if (values.BOOKING_COMPLETED === 'นิเทศแล้ว') {
    values.BOOKING_COMPLETED = 'นิเทศสำเร็จ';
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || '').trim() === 'BOOKING_COMPLETED') {
        sheet.getRange(i + 2, 2).setValue(values.BOOKING_COMPLETED);
        break;
      }
    }
  }
  const missing = DEFAULT_SYSTEM_SETTINGS.filter(row => !Object.prototype.hasOwnProperty.call(values, row[0]));
  if (missing.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, missing.length, 3).setValues(missing);
    missing.forEach(row => { values[row[0]] = row[1]; });
  }
  applySystemSettings_(values);
  cache.put(SETTINGS_CACHE_KEY, JSON.stringify(values), 60);
  return values;
}

function applySystemSettings_(values) {
  STATUS.PENDING = values.BOOKING_PENDING || 'รอดำเนินการ';
  STATUS.CONFIRMED = values.BOOKING_CONFIRMED || 'ยืนยันแล้ว';
  STATUS.REJECTED = values.BOOKING_REJECTED || 'ปฏิเสธ';
  STATUS.COMPLETED = values.BOOKING_COMPLETED || 'นิเทศสำเร็จ';
  FILE_STATUS.PENDING = values.FILE_PENDING || 'รอตรวจสอบ';
  FILE_STATUS.APPROVED = values.FILE_APPROVED || 'ผ่าน';
  FILE_STATUS.REVISE = values.FILE_REVISE || 'ปรับปรุง';
  QUALITY_LEVELS.EXCELLENT = values.QUALITY_EXCELLENT || 'ดีมาก';
  QUALITY_LEVELS.GOOD = values.QUALITY_GOOD || 'ดี';
  QUALITY_LEVELS.FAIR = values.QUALITY_FAIR || 'พอใช้';
  QUALITY_LEVELS.POOR = values.QUALITY_POOR || 'ปรับปรุง';
}

function getSystemSetting_(key, fallback) {
  const values = getSystemSettings_();
  const value = values[key];
  return value === undefined || value === '' ? fallback : value;
}

function ensureSheetHeaders(sheet, name) {
  const headers = {
    [SHEETS.BOOKING]: ['Timestamp', 'Date', 'Time', 'Booking ID', 'School ID', 'School Name', 'Supervisor Name', 'Teacher Name', 'Department', 'Period', 'Subject Name', 'Subject Code', 'Class Level', 'Room', 'Status', 'Notes', 'Username'],
    [SHEETS.FILES]: ['Timestamp', 'Teacher Name', 'Booking ID', 'File Type', 'File Name', 'File URL', 'Drive File ID', 'Status', 'Admin Note', 'Reviewed By', 'Reviewed Date', 'School ID'],
    [SHEETS.EVALUATION]: ['Timestamp', 'Supervision Type', 'Supervisor Name', 'School Name', 'School Code', 'Teacher Name', 'Subject Group', 'Subject Name', 'Grade Level', 'Period', 'Teaching Date', 'Topic', 'Techniques', 'Scores', 'Total Score', 'Strengths', 'Improvements', 'Suggestions', 'Evaluated By', 'Booking ID'],
    [SHEETS.SUPERVISION]: ['Timestamp', 'Supervision Type', 'Supervisor Name', 'School Name', 'School Code', 'Teacher Name', 'Subject Group', 'Subject Name', 'Grade Level', 'Period', 'Teaching Date', 'Topic', 'Techniques', 'Scores', 'Total Score', 'Strengths', 'Improvements', 'Suggestions', 'Evaluated By', 'Booking ID'],
    [SHEETS.USERS]: ['Username', 'Password', 'Full Name', 'Role', 'Department', 'Active', 'School ID', 'Email', 'Password Reset Required', 'Password Reset Expires At'],
    [SHEETS.SCHOOLS]: ['School ID', 'School Name', 'School Code', 'District', 'Director', 'Phone', 'Email', 'Address', 'Active', 'Created At', 'Website', 'Password Hash', 'Password Reset Required', 'Password Reset Expires At'],
    [SHEETS.LOGFILES]: ['Timestamp', 'Username', 'Role', 'School Code', 'Action', 'Result', 'Details']
  }[name];
  if (!headers) return;

  let lastColumn = sheet.getLastColumn();
  const existing = lastColumn ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
  headers.forEach(header => {
    if (existing.indexOf(header) !== -1) return;
    lastColumn++;
    sheet.getRange(1, lastColumn).setValue(header)
      .setFontWeight('bold').setBackground('#1a5276').setFontColor('white');
    existing.push(header);
  });
}

function bookingColumns_(headers) {
  const normalized = (headers || []).map(value => String(value || '').trim().toLowerCase());
  const aliases = {
    timestamp: ['timestamp'], date: ['date'], time: ['time'], bookingId: ['booking id', 'bookingid'],
    schoolCode: ['school code', 'code', 'รหัสโรงเรียน'], schoolName: ['school name', 'ชื่อโรงเรียน'],
    supervisorName: ['supervisor name', 'ผู้นิเทศ'], teacherName: ['teacher name', 'ครู'], department: ['department', 'กลุ่มสาระ'], period: ['period', 'คาบ'],
    subjectName: ['subject name', 'ชื่อวิชา'], subjectCode: ['subject code', 'รหัสวิชา'],
    classLevel: ['class level', 'ระดับชั้น'], room: ['room', 'ห้อง'], status: ['status', 'สถานะ'],
    notes: ['notes', 'หมายเหตุ'], schoolId: ['school id', 'schoolid'], username: ['username', 'user name']
  };
  const columns = {};
  Object.keys(aliases).forEach(key => { columns[key] = normalized.findIndex(value => aliases[key].includes(value)); });
  return columns;
}

function ensureBookingSchema_(sheet) {
  const expected = ['Timestamp','Date','Time','Booking ID','School ID','School Name','Supervisor Name','Teacher Name','Department','Period','Subject Name','Subject Code','Class Level','Room','Status','Notes','Username'];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const width = sheet.getLastColumn();
    const values = width ? sheet.getRange(1,1,sheet.getLastRow(),width).getValues() : [[]];
    const headers = values[0] || [];
    const alreadyCanonical = expected.length === headers.length && expected.every((header,index) => String(headers[index] || '').trim() === header);
    const legacy = bookingColumns_(headers);
    const schoolRows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
    const schoolColumns = getSchoolColumns_(schoolRows[0] || []);
    const schoolById = {}, schoolByCode = {};
    for (let i=1;i<schoolRows.length;i++) {
      const id=String(schoolValue_(schoolRows[i],schoolColumns,'schoolId') || '');
      const code=String(schoolValue_(schoolRows[i],schoolColumns,'code') || '');
      const info={schoolId:id,code,name:schoolValue_(schoolRows[i],schoolColumns,'name')};
      if(id) schoolById[id]=info;
      if(code) schoolByCode[code]=info;
    }
    const read = (row,key) => legacy[key] >= 0 ? row[legacy[key]] : '';
    const dataRows = values.slice(1).filter(row => row.some(cell => cell !== '' && cell != null));
    const output = dataRows.map(row => {
      const previousId=String(read(row,'schoolId') || '');
      const oldSchoolInfo = schoolById[previousId] || schoolByCode[previousId] || {};
      const previousSchoolValue=String(read(row,'schoolId') || read(row,'schoolCode') || '');
      const schoolInfo=schoolById[previousSchoolValue] || schoolByCode[previousSchoolValue] || oldSchoolInfo;
      const schoolId=schoolInfo.schoolId || previousSchoolValue;
      let notes=String(read(row,'notes') || '');
      const oldSupervisor=String(read(row,'supervisorName') || supervisorFromNotes_(notes));
      notes=bookingNotes_(notes, '');
      return [read(row,'timestamp'),read(row,'date'),read(row,'time'),read(row,'bookingId') || generateId(),schoolId,
        read(row,'schoolName') || schoolInfo.name || '',oldSupervisor,read(row,'teacherName'),read(row,'department'),read(row,'period'),
        read(row,'subjectName'),read(row,'subjectCode'),read(row,'classLevel'),read(row,'room'),read(row,'status'),notes,read(row,'username')];
    });
    if (!alreadyCanonical) {
      sheet.clearContents();
      sheet.getRange(1,1,1,expected.length).setValues([expected]);
      if (output.length) sheet.getRange(2,1,output.length,expected.length).setValues(output);
      sheet.getRange(1,1,1,expected.length).setFontWeight('bold').setBackground('#1a5276').setFontColor('white');
    } else {
      for (let i=1;i<values.length;i++) {
        if (values[i].some(cell => cell !== '' && cell != null) && !values[i][3]) sheet.getRange(i+1,4).setValue(generateId());
        if (values[i][14] === 'นิเทศแล้ว') sheet.getRange(i+1,15).setValue('นิเทศสำเร็จ');
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function initializeSheet(sheet, name) {
  const headers = {
    [SHEETS.BOOKING]: ['Timestamp', 'Date', 'Time', 'Booking ID', 'School ID', 'School Name', 'Supervisor Name', 'Teacher Name', 'Department', 'Period', 'Subject Name', 'Subject Code', 'Class Level', 'Room', 'Status', 'Notes', 'Username'],
    [SHEETS.FILES]: ['Timestamp', 'Teacher Name', 'Booking ID', 'File Type', 'File Name', 'File URL', 'Drive File ID', 'Status', 'Admin Note', 'Reviewed By', 'Reviewed Date', 'School ID'],
    [SHEETS.EVALUATION]: ['Timestamp', 'Supervision Type', 'Supervisor Name', 'School Name', 'School Code', 'Teacher Name', 'Subject Group', 'Subject Name', 'Grade Level', 'Period', 'Teaching Date', 'Topic', 'Techniques', 'Scores', 'Total Score', 'Strengths', 'Improvements', 'Suggestions', 'Evaluated By', 'Booking ID'],
    [SHEETS.SUPERVISION]: ['Timestamp', 'Supervision Type', 'Supervisor Name', 'School Name', 'School Code', 'Teacher Name', 'Subject Group', 'Subject Name', 'Grade Level', 'Period', 'Teaching Date', 'Topic', 'Techniques', 'Scores', 'Total Score', 'Strengths', 'Improvements', 'Suggestions', 'Evaluated By', 'Booking ID'],
    [SHEETS.USERS]: ['Username', 'Password', 'Full Name', 'Role', 'Department', 'Active', 'School ID', 'Email', 'Password Reset Required', 'Password Reset Expires At'],
    [SHEETS.SCHOOLS]: ['School ID', 'School Name', 'School Code', 'District', 'Director', 'Phone', 'Email', 'Address', 'Active', 'Created At', 'Website', 'Password Hash', 'Password Reset Required', 'Password Reset Expires At'],
    [SHEETS.LOGFILES]: ['Timestamp', 'Username', 'Role', 'School Code', 'Action', 'Result', 'Details']
  };

  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers[name].length).setBackground('#1a5276');
    sheet.getRange(1, 1, 1, headers[name].length).setFontColor('white');
  }
}

function generateId() {
  return Utilities.getUuid().split('-')[0];
}

function formatDate(date) {
  const d = new Date(date);
  return Utilities.formatDate(d, getSystemSetting_('TIME_ZONE', 'Asia/Bangkok'), 'yyyy-MM-dd');
}

function formatDateTime(date) {
  const d = new Date(date);
  return Utilities.formatDate(d, getSystemSetting_('TIME_ZONE', 'Asia/Bangkok'), 'yyyy-MM-dd HH:mm:ss');
}

function findRowByValue(sheet, column, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][column - 1] == value) {
      return i + 1;
    }
  }
  return -1;
}

function getDriveFolder() {
  const folder = DriveApp.getFolderById(getSystemSetting_('DRIVE_FOLDER_ID', DRIVE_FOLDER_ID));
  return folder;
}

function getSubFolder(name) {
  const parent = getDriveFolder();
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(name);
}

// ============================================================
// Authentication
// ============================================================

function login(data) {
  try {
    const username = String(data.username || data.schoolCode || '').trim();
    const password = String(data.password || '');
    let user = null;

    const adminPassword = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSWORD_PROPERTY);
    if (adminPassword && username === 'admin' && password === adminPassword) {
      user = {
          username: 'admin',
          fullName: 'ผู้ดูแลระบบ',
          role: 'admin',
          department: 'ทั้งหมด'
      };
    }

    if (!user) {
      const schools = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
      const cols = getSchoolColumns_(schools[0] || []);
      const headers = schools[0] || [];
      const resetRequiredCol = headers.indexOf('Password Reset Required');
      const resetExpiresCol = headers.indexOf('Password Reset Expires At');
      for (let i = 1; i < schools.length; i++) {
        if (String(schoolValue_(schools[i], cols, 'code')).trim() !== username || !isSchoolActive_(schoolValue_(schools[i], cols, 'active'))) continue;
        const storedHash = schoolValue_(schools[i], cols, 'passwordHash');
        if (storedHash && verifySchoolPassword_(password, storedHash)) {
          const mustResetPassword = resetRequiredCol >= 0 && isTruthy_(schools[i][resetRequiredCol]);
          if (mustResetPassword && resetExpiresCol >= 0 && isResetExpired_(schools[i][resetExpiresCol])) {
            return {success:false,message:'รหัสผ่านชั่วคราวหมดอายุแล้ว กรุณาขอรหัสรีเซ็ตใหม่'};
          }
          user = { username, fullName: schoolValue_(schools[i], cols, 'name'), role: 'school', schoolId: schoolValue_(schools[i], cols, 'schoolId'), schoolCode: username, schoolName: schoolValue_(schools[i], cols, 'name'), mustResetPassword };
        }
        break;
      }
    }

    if (!user) {
      const sheet = getSheet(SHEETS.USERS);
      const rowData = sheet.getDataRange().getValues();
      const headers = rowData[0] || [];
      const col = {
        username: headers.indexOf('Username'), password: headers.indexOf('Password'), fullName: headers.indexOf('Full Name'),
        role: headers.indexOf('Role'), department: headers.indexOf('Department'), active: headers.indexOf('Active'),
        schoolId: headers.indexOf('School ID'), resetRequired: headers.indexOf('Password Reset Required'),
        resetExpires: headers.indexOf('Password Reset Expires At')
      };
      for (let i = 1; i < rowData.length; i++) {
        if (String(rowData[i][col.username]) === username && verifyStoredPassword_(password, rowData[i][col.password]) && isUserActive_(rowData[i][col.active])) {
          const mustResetPassword = col.resetRequired >= 0 && isTruthy_(rowData[i][col.resetRequired]);
          if (mustResetPassword && col.resetExpires >= 0 && isResetExpired_(rowData[i][col.resetExpires])) {
            return {success:false,message:'รหัสผ่านชั่วคราวหมดอายุแล้ว กรุณาขอรหัสรีเซ็ตใหม่'};
          }
          user = { username: rowData[i][col.username], fullName: rowData[i][col.fullName], role: String(rowData[i][col.role] || '').toLowerCase(), department: rowData[i][col.department], schoolId: col.schoolId >= 0 ? rowData[i][col.schoolId] || '' : '', mustResetPassword };
          break;
        }
      }
    }
    if (!user) {
      auditLog_({username: username, role: 'unknown', schoolCode: ''}, 'login', 'failed', 'เข้าสู่ระบบไม่สำเร็จ');
      return { success: false, message: 'รหัสผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
    }
    const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('nited_session_' + token, JSON.stringify(user), 21600);
    auditLog_(user, 'login', 'success', 'เข้าสู่ระบบ');
    return { success: true, token, user };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function logout_(data) {
  const token = String(data.token || '');
  const raw = token ? CacheService.getScriptCache().get('nited_session_' + token) : '';
  if (token) CacheService.getScriptCache().remove('nited_session_' + token);
  if (raw) { try { auditLog_(JSON.parse(raw), 'logout', 'success', 'ออกจากระบบ'); } catch (_) {} }
  return {success:true};
}

function hashSchoolPassword_(password, salt) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('NITED_SCHOOL_PASSWORD_SECRET');
  if (!secret) { secret = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('NITED_SCHOOL_PASSWORD_SECRET', secret); }
  const digest = Utilities.computeHmacSha256Signature(salt + ':' + password, secret, Utilities.Charset.UTF_8);
  return 'v1$' + salt + '$' + Utilities.base64Encode(digest);
}

function verifySchoolPassword_(password, stored) {
  const parts = String(stored || '').split('$');
  return parts.length === 3 && parts[0] === 'v1' && hashSchoolPassword_(password, parts[1]) === stored;
}

function verifyStoredPassword_(password, stored) {
  const value = String(stored || '');
  return value.indexOf('v1$') === 0 ? verifySchoolPassword_(password, value) : value === String(password || '');
}

function isTruthy_(value) {
  return value === true || ['true','yes','1','ใช่'].includes(String(value || '').trim().toLowerCase());
}

function isUserActive_(value) {
  return value !== false && !['false','no','0','ปิด','ไม่ใช้งาน'].includes(String(value || '').trim().toLowerCase());
}

function isResetExpired_(value) {
  if (!value) return false;
  const expiresAt = value instanceof Date ? value : new Date(value);
  return !isNaN(expiresAt.getTime()) && expiresAt.getTime() <= Date.now();
}

function requestPasswordReset_(data) {
  const email = String(data.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return {success:false,message:'กรุณากรอกอีเมลให้ถูกต้อง'};
  const generic = {success:true,message:'หากอีเมลนี้ตรงกับบัญชีในระบบ ระบบจะส่งรหัสผ่านชั่วคราวให้ภายในไม่กี่นาที รหัสมีอายุ 30 นาที'};
  const cache = CacheService.getScriptCache();
  const digest = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email));
  const rateKey = 'nited_password_reset_' + digest;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (cache.get(rateKey)) return generic;
    cache.put(rateKey, 'pending', 60);
  } finally {
    lock.releaseLock();
  }

  const changes = [];
  const resetEntries = [];
  const expiresAt = new Date(Date.now() + 30 * 60 * 1000);
  const temporaryCredentials = [];
  try {
    const schoolSheet = getSheet(SHEETS.SCHOOLS);
    const schoolRows = schoolSheet.getDataRange().getValues();
    const schoolHeaders = schoolRows[0] || [];
    const schoolCols = getSchoolColumns_(schoolHeaders);
    const schoolResetCol = schoolHeaders.indexOf('Password Reset Required');
    const schoolExpiresCol = schoolHeaders.indexOf('Password Reset Expires At');
    for (let i = 1; i < schoolRows.length; i++) {
      if (String(schoolValue_(schoolRows[i], schoolCols, 'email') || '').trim().toLowerCase() !== email || !isSchoolActive_(schoolValue_(schoolRows[i], schoolCols, 'active'))) continue;
      const entry = {sheet:schoolSheet,rowNumber:i+1,passwordCol:schoolCols.passwordHash,requiredCol:schoolResetCol,expiresCol:schoolExpiresCol,
        oldPassword:schoolRows[i][schoolCols.passwordHash],oldRequired:schoolRows[i][schoolResetCol],oldExpires:schoolRows[i][schoolExpiresCol],username:schoolValue_(schoolRows[i],schoolCols,'code')};
      resetEntries.push(entry);
    }

    const userSheet = getSheet(SHEETS.USERS);
    const userRows = userSheet.getDataRange().getValues();
    const userHeaders = userRows[0] || [];
    const userEmailCol = userHeaders.indexOf('Email');
    const userPasswordCol = userHeaders.indexOf('Password');
    const userNameCol = userHeaders.indexOf('Username');
    const userActiveCol = userHeaders.indexOf('Active');
    const userRequiredCol = userHeaders.indexOf('Password Reset Required');
    const userExpiresCol = userHeaders.indexOf('Password Reset Expires At');
    if (userEmailCol >= 0) for (let i = 1; i < userRows.length; i++) {
      if (String(userRows[i][userEmailCol] || '').trim().toLowerCase() !== email || !isUserActive_(userRows[i][userActiveCol])) continue;
      resetEntries.push({sheet:userSheet,rowNumber:i+1,passwordCol:userPasswordCol,requiredCol:userRequiredCol,expiresCol:userExpiresCol,
        oldPassword:userRows[i][userPasswordCol],oldRequired:userRows[i][userRequiredCol],oldExpires:userRows[i][userExpiresCol],username:userRows[i][userNameCol]});
    }

    if (!resetEntries.length) return generic;
    resetEntries.forEach(entry => {
      const temporaryPassword = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
      const salt = Utilities.getUuid().replace(/-/g, '');
      entry.sheet.getRange(entry.rowNumber, entry.passwordCol + 1).setValue(hashSchoolPassword_(temporaryPassword, salt));
      entry.sheet.getRange(entry.rowNumber, entry.requiredCol + 1).setValue(true);
      entry.sheet.getRange(entry.rowNumber, entry.expiresCol + 1).setValue(expiresAt);
      changes.push(entry);
      temporaryCredentials.push({username:entry.username,password:temporaryPassword});
    });

    const credentialText = temporaryCredentials.map(item => 'ชื่อผู้ใช้: ' + item.username + '\nรหัสผ่านชั่วคราว: ' + item.password).join('\n\n');
    MailApp.sendEmail(email, 'รหัสผ่านชั่วคราวสำหรับระบบนิเทศออนไลน์ สพม.นราธิวาส',
      'มีคำขอรีเซ็ตรหัสผ่านสำหรับอีเมลนี้\n\n' + credentialText + '\n\nรหัสผ่านชั่วคราวใช้ได้ 30 นาที และระบบจะให้กำหนดรหัสผ่านใหม่หลังเข้าสู่ระบบ\nหากคุณไม่ได้เป็นผู้ขอ สามารถละเว้นอีเมลฉบับนี้ได้');
    return generic;
  } catch (err) {
    changes.forEach(entry => {
      entry.sheet.getRange(entry.rowNumber, entry.passwordCol + 1).setValue(entry.oldPassword || '');
      entry.sheet.getRange(entry.rowNumber, entry.requiredCol + 1).setValue(entry.oldRequired || false);
      entry.sheet.getRange(entry.rowNumber, entry.expiresCol + 1).setValue(entry.oldExpires || '');
    });
    cache.remove(rateKey);
    console.error('password reset email failed: ' + err);
    return {success:false,message:'ส่งอีเมลรีเซ็ตรหัสผ่านไม่สำเร็จ กรุณาติดต่อผู้ดูแลระบบ'};
  }
}

function changePassword_(data, user) {
  const password = String(data.newPassword || '');
  const confirmation = String(data.confirmPassword || '');
  if (password.length < 8) return {success:false,message:'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร'};
  if (password !== confirmation) return {success:false,message:'รหัสผ่านใหม่และการยืนยันไม่ตรงกัน'};
  let sheet, rowNumber, passwordCol, requiredCol, expiresCol;
  if (user.role === 'school') {
    sheet = getSheet(SHEETS.SCHOOLS);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0] || [];
    const cols = getSchoolColumns_(headers);
    passwordCol = cols.passwordHash;
    requiredCol = headers.indexOf('Password Reset Required');
    expiresCol = headers.indexOf('Password Reset Expires At');
    for (let i = 1; i < rows.length; i++) if (String(schoolValue_(rows[i], cols, 'schoolId')) === String(user.schoolId)) { rowNumber = i + 1; break; }
  } else {
    sheet = getSheet(SHEETS.USERS);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0] || [];
    passwordCol = headers.indexOf('Password');
    requiredCol = headers.indexOf('Password Reset Required');
    expiresCol = headers.indexOf('Password Reset Expires At');
    const usernameCol = headers.indexOf('Username');
    for (let i = 1; i < rows.length; i++) if (String(rows[i][usernameCol]) === String(user.username)) { rowNumber = i + 1; break; }
  }
  if (!rowNumber || passwordCol < 0 || requiredCol < 0 || expiresCol < 0) return {success:false,message:'ไม่พบข้อมูลบัญชีสำหรับเปลี่ยนรหัสผ่าน'};
  const salt = Utilities.getUuid().replace(/-/g, '');
  sheet.getRange(rowNumber, passwordCol + 1).setValue(hashSchoolPassword_(password, salt));
  sheet.getRange(rowNumber, requiredCol + 1).setValue(false);
  sheet.getRange(rowNumber, expiresCol + 1).clearContent();
  return {success:true,message:'กำหนดรหัสผ่านใหม่สำเร็จ'};
}

function setSchoolPassword_(data) {
  const password = String(data.password || '');
  if (password.length < 8) return {success:false,message:'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร'};
  const sheet = getSheet(SHEETS.SCHOOLS);
  const rows = sheet.getDataRange().getValues();
  const cols = getSchoolColumns_(rows[0] || []);
  const codeCol = cols.code, hashCol = cols.passwordHash;
  if (hashCol == null) return {success:false,message:'ไม่พบคอลัมน์รหัสผ่านใน Schools'};
  const targetCode = String(data.schoolCode || '').trim();
  for (let i=1;i<rows.length;i++) if (String(rows[i][codeCol]).trim() === targetCode) {
    const salt = Utilities.getUuid().replace(/-/g,'');
    sheet.getRange(i+1, hashCol+1).setValue(hashSchoolPassword_(password, salt));
    const headers = rows[0] || [];
    const resetCol = headers.indexOf('Password Reset Required');
    const expiryCol = headers.indexOf('Password Reset Expires At');
    if (resetCol >= 0) sheet.getRange(i + 1, resetCol + 1).setValue(false);
    if (expiryCol >= 0) sheet.getRange(i + 1, expiryCol + 1).clearContent();
    return {success:true,message:'ตั้งรหัสผ่านโรงเรียนสำเร็จ'};
  }
  return {success:false,message:'ไม่พบรหัสโรงเรียน'};
}

// ============================================================
// Booking Management
// ============================================================

function createBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    if (!isActiveSchool(data.schoolId)) {
      return { success: false, message: 'กรุณาเลือกโรงเรียนผู้รับการนิเทศที่อยู่ในสังกัด' };
    }
    const school = getSchools({}).schools.find(item => String(item.schoolId) === String(data.schoolId));
    if (!school) return {success:false,message:'ไม่พบข้อมูลโรงเรียน'};
    const id = generateId();
    const conflicting = checkBookingConflict(data.date, data.time, data.period, school.schoolId);
    if (conflicting) {
      return { success: false, message: 'วัน-เวลานี้มีการจองแล้ว กรุณาเลือกวัน-เวลาอื่น' };
    }

    const supervisorName = String(data.supervisorName || (data._user && data._user.fullName) || '').trim();
    const row = [formatDateTime(new Date()), data.date, data.time, id, school.schoolId, school.name, supervisorName,
      data.teacherName || '', data.department || '', data.period || '', data.subjectName || '', data.subjectCode || '',
      data.classLevel || '', data.room || '', STATUS.PENDING, bookingNotes_(data.notes, ''),
      (data._user && data._user.username) || ''];

    sheet.appendRow(row);

    return {
      success: true,
      message: 'จองวันนิเทศสำเร็จ รอการยืนยันจากผู้ดูแลระบบ',
      bookingId: id
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function bookingNotes_(notes) {
  const lines = String(notes || '').split(/\r?\n/).filter(line => !/^\s*ผู้นิเทศ\s*:/i.test(line));
  return lines.join('\n').trim();
}

function supervisorFromNotes_(notes) {
  const match = String(notes || '').match(/^\s*ผู้นิเทศ\s*:\s*(.*?)\s*$/im);
  return match ? match[1] : '';
}

function getSchoolIdByCode_(schoolCode) {
  return getSchoolIdMapByCode_()[String(schoolCode || '')] || '';
}

function getSchoolIdMapByCode_() {
  const rows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
  const columns = getSchoolColumns_(rows[0] || []);
  const map = {};
  for (let i=1;i<rows.length;i++) {
    const code = String(schoolValue_(rows[i],columns,'code') || '');
    if (code) map[code] = schoolValue_(rows[i],columns,'schoolId');
  }
  return map;
}

function getSchoolCodeMapById_() {
  const rows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
  const columns = getSchoolColumns_(rows[0] || []);
  const map = {};
  for (let i=1;i<rows.length;i++) {
    const id = String(schoolValue_(rows[i],columns,'schoolId') || '');
    if (id) map[id] = schoolValue_(rows[i],columns,'code');
  }
  return map;
}

function getSchoolCodeById_(schoolId) {
  return getSchoolCodeMapById_()[String(schoolId || '')] || '';
}

function getBookingRecordById_(bookingId) {
  const id = String(bookingId || '').trim();
  if (!id) return null;
  const rows = getSheet(SHEETS.BOOKING).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3] || '').trim() === id) {
      return { rowNumber: i + 1, bookingId: id, schoolId: rows[i][4], schoolName: rows[i][5], supervisorName: rows[i][6], teacherName: rows[i][7], department: rows[i][8], period: rows[i][9], subjectName: rows[i][10], subjectCode: rows[i][11], classLevel: rows[i][12], room: rows[i][13], status: rows[i][14] };
    }
  }
  return null;
}

function checkBookingConflict(date, time, period, schoolId, excludeRowNumber) {
  const sheet = getSheet(SHEETS.BOOKING);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (excludeRowNumber && i + 1 === Number(excludeRowNumber)) continue;
    if (String(data[i][1]) === String(date) && String(data[i][2]) === String(time) && String(data[i][9]) === String(period) && String(data[i][4]) === String(schoolId)) {
      if (data[i][14] !== STATUS.REJECTED) {
        return true;
      }
    }
  }
  return false;
}

function updateBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const rowNum = Number(data.rowNumber);

    if (!Number.isInteger(rowNum) || rowNum < 2 || rowNum > sheet.getLastRow()) {
      return { success: false, message: 'ไม่พบรายการจอง' };
    }

    const existingRow = sheet.getRange(rowNum, 1, 1, 17).getValues()[0];
    const user = data._user || {};
    if (user.role === 'supervisor' && data.status !== undefined && !isBookingCompleted_(data.status)) {
      return {success:false,message:'ผู้นิเทศเปลี่ยนสถานะได้เฉพาะเมื่อตรวจยืนยันว่า นิเทศสำเร็จ'};
    }
    if (isBookingCompleted_(existingRow[14]) && user.role !== 'admin') {
      return {success:false,message:'รายการนี้นิเทศสำเร็จแล้ว จึงไม่สามารถแก้ไขได้'};
    }
    if (user.role === 'supervisor' && String(existingRow[6] || '').trim() !== String(user.fullName || '').trim()) {
      return {success:false,message:'แก้ไขได้เฉพาะกำหนดนิเทศที่ระบุชื่อคุณเป็นผู้นิเทศ'};
    }
    if (data._user && data._user.role === 'school' && String(existingRow[4]) !== String(data._user.schoolId)) return {success:false,message:'ไม่มีสิทธิ์แก้ไขรายการนี้'};
    const isSchoolUser = data._user && data._user.role === 'school';
    if (isSchoolUser) { delete data.status; delete data.supervisorName; }

    const width = sheet.getLastColumn();
    const currentData = sheet.getRange(rowNum, 1, 1, width).getValues()[0];
    const schoolId = currentData[4];
    const nextDate = data.date !== undefined && data.date !== '' ? data.date : currentData[1];
    const nextTime = data.time !== undefined && data.time !== '' ? data.time : currentData[2];
    const nextPeriod = data.period !== undefined && data.period !== '' ? data.period : currentData[9];
    if (checkBookingConflict(nextDate, nextTime, nextPeriod, schoolId, rowNum)) return {success:false,message:'วัน-เวลานี้มีการจองแล้ว กรุณาเลือกวัน-เวลาอื่น'};

    const fieldMap = { date:1, time:2, supervisorName:6, teacherName:7, department:8, period:9, subjectName:10, subjectCode:11, classLevel:12, room:13, status:14 };
    Object.keys(fieldMap).forEach(key => {
      if (data[key] !== undefined) currentData[fieldMap[key]] = data[key];
    });
    const notes = data.notes !== undefined ? data.notes : currentData[15];
    currentData[15] = bookingNotes_(notes, '');
    sheet.getRange(rowNum, 1, 1, width).setValues([currentData]);

    return { success: true, message: isBookingCompleted_(currentData[14]) ? 'บันทึกผลนิเทศสำเร็จ และล็อกรายการแล้ว' : 'อัพเดทการจองสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const rowNumber = Number(data.rowNumber);
    if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) return {success:false,message:'ไม่พบรายการจอง'};
    const row = sheet.getRange(rowNumber, 1, 1, 17).getValues()[0];
    if (isBookingCompleted_(row[14]) && (!data._user || data._user.role !== 'admin')) return {success:false,message:'รายการนิเทศสำเร็จแล้ว ไม่สามารถลบได้'};
    if (data._user && data._user.role === 'school' && String(row[4]) !== String(data._user.schoolId)) return {success:false,message:'ไม่มีสิทธิ์ลบรายการนี้'};
    sheet.deleteRow(rowNumber);
    return { success: true, message: 'ลบการจองสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getBookings(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const schoolCodesById = getSchoolCodeMapById_();
    let bookings = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }
      row['School Code'] = schoolCodesById[String(allData[i][4] || '')] || '';
      row['Supervisor Name'] = allData[i][6];
      row['Notes'] = bookingNotes_(allData[i][15], '');
      row.bookingIndex = i;

      if (data && data.status && allData[i][14] !== data.status) continue;
      if (data && data._user && data._user.role === 'school' && String(allData[i][4]) !== String(data._user.schoolId)) continue;
      if (data && data.teacherName && allData[i][7] !== data.teacherName) continue;
      if (data && data.date && allData[i][1] !== data.date) continue;
      if (data && data.department && allData[i][8] !== data.department) continue;

      bookings.push(row);
    }

    bookings.reverse();

    return { success: true, bookings: bookings };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getBookingById(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    for (let i = 1; i < allData.length; i++) {
      if (i + 1 === data.rowNumber || String(allData[i][3]) === String(data.bookingId)) {
        if (data._user && data._user.role === 'school' && String(allData[i][4]) !== String(data._user.schoolId)) return {success:false,message:'ไม่พบรายการจอง'};
        const row = {};
        row.rowNumber = i + 1;
        for (let j = 0; j < headers.length; j++) {
          row[headers[j]] = allData[i][j];
        }
        row['School Code'] = getSchoolCodeById_(allData[i][4]);
        row['Supervisor Name'] = allData[i][6];
        row['Notes'] = bookingNotes_(allData[i][15], '');
        return { success: true, booking: row };
      }
    }

    return { success: false, message: 'ไม่พบรายการจอง' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// File Management
// ============================================================

function uploadFile(data) {
  try {
    if (data._user && data._user.role === 'school') data.schoolId = data._user.schoolId;
    const booking = data.bookingId ? getBookingRecordById_(data.bookingId) : null;
    if (data.bookingId && !booking) return {success:false,message:'ไม่พบ Booking ID ที่เลือก'};
    if (booking && String(booking.schoolId) !== String(data.schoolId || '')) return {success:false,message:'โรงเรียนของหลักฐานต้องตรงกับโรงเรียนในกำหนดนิเทศ'};
    const sheet = getSheet(SHEETS.FILES);

    const fileFolderMap = {
      'แผนการสอน': 'Plans',
      'สื่อการสอน': 'Media',
      'ภาพกิจกรรม': 'Photos',
      'คลิปวิดีโอ': 'Clips'
    };

    let fileUrl = data.fileUrl || '';
    let driveFileId = '';
    let fileName = data.fileName || '';

    if (data.fileData && data.fileName) {
      const folderName = fileFolderMap[data.fileType] || 'Other';
      const folder = getSubFolder(folderName);

      const decoded = Utilities.base64Decode(data.fileData.split(',')[1] || data.fileData);
      const blob = Utilities.newBlob(decoded, data.mimeType || 'application/octet-stream', data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      fileUrl = file.getUrl();
      driveFileId = file.getId();
      fileName = file.getName();
    }

    const row = [
      formatDateTime(new Date()),
      data.teacherName,
      data.bookingId || '',
      data.fileType,
      fileName,
      fileUrl,
      driveFileId,
      FILE_STATUS.PENDING,
      '',
      '',
      '',
      data.schoolId || ''
    ];

    sheet.appendRow(row);

    return {
      success: true,
      message: 'อัพโหลดไฟล์สำเร็จ',
      fileUrl: fileUrl
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateFileStatus(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการไฟล์' };
    }

    const currentData = sheet.getRange(rowNum, 1, 1, 11).getValues()[0];

    if (data.status) currentData[7] = data.status;
    if (data.adminNote !== undefined) currentData[8] = data.adminNote;
    if (data.reviewedBy) currentData[9] = data.reviewedBy;
    currentData[10] = formatDateTime(new Date());

    sheet.getRange(rowNum, 1, 1, 11).setValues([currentData]);

    return { success: true, message: 'อัพเดทสถานะไฟล์สำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getFiles(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const schoolNames = getSchoolNameMap_();
    let files = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }
      row['School Name'] = schoolNames[String(allData[i][11] || '')] || '';
      if (data && data._user && data._user.role === 'school' && String(allData[i][11]) !== String(data._user.schoolId)) continue;

      if (data && data.status && allData[i][7] !== data.status) continue;
      if (data && data.teacherName && allData[i][1] !== data.teacherName) continue;
      if (data && data.fileType && allData[i][3] !== data.fileType) continue;
      if (data && data.bookingId && allData[i][2] !== data.bookingId) continue;

      files.push(row);
    }

    files.reverse();

    return { success: true, files: files };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteFile(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);
    const row = sheet.getRange(Number(data.rowNumber), 1, 1, 12).getValues()[0];
    if (data._user && data._user.role === 'school' && String(row[11]) !== String(data._user.schoolId)) return {success:false,message:'ไม่มีสิทธิ์ลบไฟล์นี้'};

    if (data.driveFileId) {
      try {
        DriveApp.getFileById(data.driveFileId).setTrashed(true);
      } catch (e) {
        // File may not exist
      }
    }

    sheet.deleteRow(data.rowNumber);
    return { success: true, message: 'ลบไฟล์สำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Evaluation Management
// ============================================================

function createEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    if (!data.bookingId) return {success:false,message:'กรุณาเลือก Booking ID เพื่อเชื่อมโยงผลนิเทศ'};
    const booking = getBookingRecordById_(data.bookingId);
    if (!booking) return {success:false,message:'ไม่พบ Booking ID ที่เลือก'};
    if (String(booking.schoolId) !== String(data.schoolId || '')) return {success:false,message:'โรงเรียนต้องตรงกับ Booking ID ที่เลือก'};
    if (data._user && data._user.role === 'supervisor' && String(booking.supervisorName || '').trim() !== String(data._user.fullName || '').trim()) return {success:false,message:'บันทึกผลได้เฉพาะกำหนดนิเทศที่ระบุชื่อคุณเป็นผู้นิเทศ'};
    if (!isActiveSchool(data.schoolId)) {
      return { success: false, message: 'กรุณาเลือกโรงเรียนผู้รับการนิเทศที่อยู่ในสังกัด' };
    }
    if (!data.supervisionType || !data.teachingDate || !data.topic) return {success:false,message:'กรุณาระบุประเภทการนิเทศ วันที่จัดการเรียนรู้ และหัวข้อที่สอน'};
    if (data.totalScore !== '' && data.totalScore != null && (isNaN(Number(data.totalScore)) || Number(data.totalScore) < 0)) return {success:false,message:'คะแนนรวมต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป'};
    const school = getSchools({}).schools.find(item => String(item.schoolId) === String(data.schoolId));
    const row = [formatDateTime(new Date()), data.supervisionType || '', booking.supervisorName || data.supervisorName || '',
      school.name, school.code, data.teacherName || '', data.subjectGroup || '', data.subjectName || '',
      data.gradeLevel || '', data.period || '', data.teachingDate || '', data.topic || '', data.techniques || '',
      data.scores || '', data.totalScore === '' || data.totalScore == null ? '' : Number(data.totalScore),
      data.strengths || '', data.improvements || '', data.suggestions || '', data.evaluatedBy || '', booking.bookingId];
    appendEvaluationRecord_(sheet, row);
    appendEvaluationRecord_(getSheet(SHEETS.SUPERVISION), row);
    return {
      success: true,
      message: 'บันทึกข้อมูลการนิเทศสำเร็จ'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function appendEvaluationRecord_(sheet, canonicalRow) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const canonicalHeaders = ['Timestamp','Supervision Type','Supervisor Name','School Name','School Code','Teacher Name','Subject Group','Subject Name','Grade Level','Period','Teaching Date','Topic','Techniques','Scores','Total Score','Strengths','Improvements','Suggestions','Evaluated By','Booking ID'];
  const indexByHeader = {};
  canonicalHeaders.forEach((header, index) => { indexByHeader[header] = index; });
  sheet.appendRow(headers.map(header => indexByHeader[header] === undefined ? '' : canonicalRow[indexByHeader[header]]));
}

function updateEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการประเมิน' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const currentData = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    ['Scores','Total Score','Strengths','Improvements','Suggestions'].forEach((header, index) => {
      const key = ['scores','totalScore','strengths','improvements','suggestions'][index];
      const col = headers.indexOf(header);
      if (data[key] !== undefined && col >= 0) currentData[col] = data[key];
    });
    sheet.getRange(rowNum, 1, 1, headers.length).setValues([currentData]);
    const bookingIdCol = headers.indexOf('Booking ID');
    const bookingId = bookingIdCol >= 0 ? currentData[bookingIdCol] : '';
    if (bookingId) {
      const mirror = getSheet(SHEETS.SUPERVISION);
      const mirrorRows = mirror.getDataRange().getValues();
      const mirrorHeaders = mirrorRows[0];
      const mirrorBookingIdCol = mirrorHeaders.indexOf('Booking ID');
      for (let i = 1; i < mirrorRows.length; i++) {
        if (mirrorBookingIdCol >= 0 && String(mirrorRows[i][mirrorBookingIdCol]) === String(bookingId)) {
          ['Scores','Total Score','Strengths','Improvements','Suggestions'].forEach(header => {
            const sourceCol = headers.indexOf(header), targetCol = mirrorHeaders.indexOf(header);
            if (sourceCol >= 0 && targetCol >= 0) mirror.getRange(i + 1, targetCol + 1).setValue(currentData[sourceCol]);
          });
        }
      }
    }

    return { success: true, message: 'อัพเดทผลการประเมินสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getEvaluations(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const schoolNames = getSchoolNameMap_();
    let evaluations = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }
      row['School Name'] = row['School Name'] || '';
      if (data && data._user && data._user.role === 'school' && String(row['School Code']) !== String(data._user.schoolCode)) continue;
      if (data && data.teacherName && row['Teacher Name'] !== data.teacherName) continue;
      if (data && data.schoolId && (!data._user || data._user.role !== 'school') && row['School Code'] !== data.schoolId && row['School Name'] !== data.schoolId) continue;

      evaluations.push(row);
    }

    evaluations.reverse();

    return { success: true, evaluations: evaluations };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = sheet.getRange(Number(data.rowNumber), 1, 1, headers.length).getValues()[0];
    const bookingIdCol = headers.indexOf('Booking ID');
    const bookingId = bookingIdCol >= 0 ? row[bookingIdCol] : '';
    sheet.deleteRow(data.rowNumber);
    if (bookingId) {
      const mirror = getSheet(SHEETS.SUPERVISION);
      const rows = mirror.getDataRange().getValues();
      const idCol = rows[0].indexOf('Booking ID');
      for (let i = rows.length - 1; i >= 1; i--) if (idCol >= 0 && String(rows[i][idCol]) === String(bookingId)) mirror.deleteRow(i + 1);
    }
    return { success: true, message: 'ลบผลการประเมินสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Dashboard & Statistics
// ============================================================

function getDashboardStats(data) {
  try {
    const bookingSheet = getSheet(SHEETS.BOOKING);
    const fileSheet = getSheet(SHEETS.FILES);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let bookingData = bookingSheet.getDataRange().getValues();
    let fileData = fileSheet.getDataRange().getValues();
    let evalData = evalSheet.getDataRange().getValues();
    if (data && data._user && data._user.role === 'school') {
      bookingData = [bookingData[0]].concat(bookingData.slice(1).filter(r => String(r[4]) === String(data._user.schoolId)));
      fileData = [fileData[0]].concat(fileData.slice(1).filter(r => String(r[11]) === String(data._user.schoolId)));
      evalData = [evalData[0]].concat(evalData.slice(1).filter(r => String(r[4]) === String(data._user.schoolCode)));
    }
    const schoolNames = getSchoolNameMap_();
    const schoolCodesById = getSchoolCodeMapById_();

    const today = formatDate(new Date());
    const thisMonth = today.substring(0, 7);

    let stats = {
      totalBookings: bookingData.length - 1,
      pendingBookings: 0,
      confirmedBookings: 0,
      completedBookings: 0,
      rejectedBookings: 0,
      todayBookings: 0,
      thisMonthBookings: 0,

      totalFiles: fileData.length - 1,
      pendingFiles: 0,
      approvedFiles: 0,
      reviseFiles: 0,

      totalEvaluations: evalData.length - 1,
      excellentCount: 0,
      goodCount: 0,
      fairCount: 0,
      poorCount: 0,

      recentBookings: [],
      recentFiles: [],
      recentEvaluations: [],

      departmentStats: {},
      teacherStats: {}
    };

    for (let i = 1; i < bookingData.length; i++) {
      const status = bookingData[i][14];
      const bookingDate = bookingData[i][1];

      if (status === STATUS.PENDING) stats.pendingBookings++;
      if (status === STATUS.CONFIRMED) stats.confirmedBookings++;
      if (status === STATUS.COMPLETED) stats.completedBookings++;
      if (status === STATUS.REJECTED) stats.rejectedBookings++;
      if (bookingDate === today) stats.todayBookings++;
      if (bookingDate && bookingDate.toString().startsWith(thisMonth)) stats.thisMonthBookings++;

      const dept = bookingData[i][8];
      if (dept) {
        if (!stats.departmentStats[dept]) {
          stats.departmentStats[dept] = { total: 0, completed: 0, pending: 0 };
        }
        stats.departmentStats[dept].total++;
        if (status === STATUS.COMPLETED) stats.departmentStats[dept].completed++;
        if (status === STATUS.PENDING) stats.departmentStats[dept].pending++;
      }

      const teacher = bookingData[i][7];
      if (teacher) {
        if (!stats.teacherStats[teacher]) {
          stats.teacherStats[teacher] = { total: 0, completed: 0, files: 0 };
        }
        stats.teacherStats[teacher].total++;
        if (status === STATUS.COMPLETED) stats.teacherStats[teacher].completed++;
      }
    }

    for (let i = 1; i < fileData.length; i++) {
      const status = fileData[i][7];
      if (status === FILE_STATUS.PENDING) stats.pendingFiles++;
      if (status === FILE_STATUS.APPROVED) stats.approvedFiles++;
      if (status === FILE_STATUS.REVISE) stats.reviseFiles++;

      const teacher = fileData[i][1];
      if (teacher && stats.teacherStats[teacher]) {
        stats.teacherStats[teacher].files++;
      }
    }

    for (let i = 1; i < evalData.length; i++) {
      const total = Number(evalData[i][14]) || 0;
      if (total >= 80) stats.excellentCount++;
      else if (total >= 60) stats.goodCount++;
      else if (total >= 40) stats.fairCount++;
      else if (evalData[i][14] !== '') stats.poorCount++;
    }

    stats.recentBookings = bookingData.slice(-10).reverse().map((row, idx) => ({
      rowNumber: bookingData.length - idx,
      date: row[1],
      time: row[2],
      teacherName: row[7],
      department: row[8],
      period: row[9],
      subjectName: row[10],
      status: row[14],
      schoolId: row[4] || '',
      schoolCode: schoolCodesById[String(row[4] || '')] || '',
      schoolName: row[5] || ''
    }));

    stats.recentFiles = fileData.slice(-10).reverse().map((row, idx) => ({
      rowNumber: fileData.length - idx,
      teacherName: row[1],
      fileType: row[3],
      fileName: row[4],
      status: row[7],
      schoolId: row[11] || '',
      schoolName: schoolNames[String(row[11] || '')] || ''
    }));

    stats.recentEvaluations = evalData.slice(-5).reverse().map((row, idx) => ({
      rowNumber: evalData.length - idx,
      teacherName: row[5],
      schoolId: '',
      schoolName: row[3] || '',
      supervisionDate: row[10],
      qualityLevel: scoreToQuality_(row[14]),
      evaluatedBy: row[18]
    }));

    return { success: true, stats: stats };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getCalendarData(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    let events = [];

    const month = data.month || new Date().getMonth();
    const year = data.year || new Date().getFullYear();

    for (let i = 1; i < allData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(allData[i][4]) !== String(data._user.schoolId)) continue;
      const bookingDate = allData[i][1];
      if (!bookingDate) continue;

      const d = new Date(bookingDate);
      if (d.getMonth() == month && d.getFullYear() == year) {
        events.push({
          date: bookingDate,
          title: (allData[i][5] || 'โรงเรียน') + (allData[i][10] ? ' - ' + allData[i][10] : '') + (allData[i][7] ? ' (' + allData[i][7] + ')' : ''),
          time: allData[i][2],
          status: allData[i][14],
          department: allData[i][8]
        });
      }
    }

    return { success: true, events: events };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Reports
// ============================================================

function getReport(data) {
  try {
    const type = data.type || 'all';
    const startDate = data.startDate || '';
    const endDate = data.endDate || '';

    const bookingSheet = getSheet(SHEETS.BOOKING);
    const fileSheet = getSheet(SHEETS.FILES);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let report = {
      bookings: [],
      files: [],
      evaluations: [],
      summary: {}
    };

    const bookingData = bookingSheet.getDataRange().getValues();
    for (let i = 1; i < bookingData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(bookingData[i][4]) !== String(data._user.schoolId)) continue;
      const d = bookingData[i][1];
      if (startDate && d < startDate) continue;
      if (endDate && d > endDate) continue;
      report.bookings.push({
        date: d,
        time: bookingData[i][2],
        teacherName: bookingData[i][7],
        department: bookingData[i][8],
        period: bookingData[i][9],
        subjectName: bookingData[i][10],
        status: bookingData[i][14]
      });
    }

    const fileData = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(fileData[i][11]) !== String(data._user.schoolId)) continue;
      report.files.push({
        timestamp: fileData[i][0],
        teacherName: fileData[i][1],
        fileType: fileData[i][3],
        fileName: fileData[i][4],
        status: fileData[i][7]
      });
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(evalData[i][4]) !== String(data._user.schoolCode)) continue;
      report.evaluations.push({
        teacherName: evalData[i][5],
        schoolName: evalData[i][3],
        supervisionType: evalData[i][1],
        supervisionDate: evalData[i][10],
        strengths: evalData[i][15],
        improvements: evalData[i][16],
        qualityLevel: scoreToQuality_(evalData[i][14])
      });
    }

    report.summary = {
      totalBookings: report.bookings.length,
      completedBookings: report.bookings.filter(b => b.status === STATUS.COMPLETED).length,
      totalFiles: report.files.length,
      approvedFiles: report.files.filter(f => f.status === FILE_STATUS.APPROVED).length,
      totalEvaluations: report.evaluations.length
    };

    return { success: true, report: report };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getTeacherReport(data) {
  try {
    const teacherName = data.teacherName;

    const bookingSheet = getSheet(SHEETS.BOOKING);
    const fileSheet = getSheet(SHEETS.FILES);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let report = {
      teacherName: teacherName,
      bookings: [],
      files: [],
      evaluations: [],
      summary: {}
    };

    const bookingData = bookingSheet.getDataRange().getValues();
    for (let i = 1; i < bookingData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(bookingData[i][4]) !== String(data._user.schoolId)) continue;
      if (bookingData[i][7] === teacherName) {
        report.bookings.push({
          date: bookingData[i][1],
          time: bookingData[i][2],
          subjectName: bookingData[i][10],
          status: bookingData[i][14]
        });
      }
    }

    const fileData = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(fileData[i][11]) !== String(data._user.schoolId)) continue;
      if (fileData[i][1] === teacherName) {
        report.files.push({
          fileType: fileData[i][3],
          fileName: fileData[i][4],
          status: fileData[i][7]
        });
      }
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(evalData[i][4]) !== String(data._user.schoolCode)) continue;
      if (evalData[i][5] === teacherName) {
        report.evaluations.push({
          supervisionDate: evalData[i][10],
          strengths: evalData[i][15],
          improvements: evalData[i][16],
          suggestions: evalData[i][17],
          qualityLevel: scoreToQuality_(evalData[i][14])
        });
      }
    }

    report.summary = {
      totalBookings: report.bookings.length,
      completedBookings: report.bookings.filter(b => b.status === STATUS.COMPLETED).length,
      totalFiles: report.files.length,
      approvedFiles: report.files.filter(f => f.status === FILE_STATUS.APPROVED).length,
      totalEvaluations: report.evaluations.length,
      avgQuality: getAvgQuality(report.evaluations)
    };

    return { success: true, report: report };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getDepartmentReport(data) {
  try {
    const department = data.department;

    const bookingSheet = getSheet(SHEETS.BOOKING);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let report = {
      department: department,
      teachers: {},
      summary: {}
    };

    const bookingData = bookingSheet.getDataRange().getValues();
    for (let i = 1; i < bookingData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(bookingData[i][4]) !== String(data._user.schoolId)) continue;
      if (bookingData[i][8] === department) {
        const teacher = bookingData[i][7];
        if (!report.teachers[teacher]) {
          report.teachers[teacher] = {
            totalBookings: 0,
            completedBookings: 0,
            files: 0,
            evaluations: []
          };
        }
        report.teachers[teacher].totalBookings++;
        if (bookingData[i][14] === STATUS.COMPLETED) {
          report.teachers[teacher].completedBookings++;
        }
      }
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
      if (data && data._user && data._user.role === 'school' && String(evalData[i][4]) !== String(data._user.schoolCode)) continue;
      const teacher = evalData[i][5];
      if (report.teachers[teacher]) {
        report.teachers[teacher].evaluations.push({
          date: evalData[i][10],
          qualityLevel: scoreToQuality_(evalData[i][14])
        });
      }
    }

    const teacherList = Object.keys(report.teachers);
    report.summary = {
      totalTeachers: teacherList.length,
      totalBookings: teacherList.reduce((sum, t) => sum + report.teachers[t].totalBookings, 0),
      completedBookings: teacherList.reduce((sum, t) => sum + report.teachers[t].completedBookings, 0)
    };

    return { success: true, report: report };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAvgQuality(evaluations) {
  if (evaluations.length === 0) return 'ไม่มีข้อมูล';

  const scores = {
    [QUALITY_LEVELS.EXCELLENT]: 4,
    [QUALITY_LEVELS.GOOD]: 3,
    [QUALITY_LEVELS.FAIR]: 2,
    [QUALITY_LEVELS.POOR]: 1
  };

  let total = 0;
  let count = 0;
  evaluations.forEach(e => {
    if (scores[e.qualityLevel]) {
      total += scores[e.qualityLevel];
      count++;
    }
  });

  if (count === 0) return 'ไม่มีข้อมูล';
  const avg = total / count;
  if (avg >= 3.5) return QUALITY_LEVELS.EXCELLENT;
  if (avg >= 2.5) return QUALITY_LEVELS.GOOD;
  if (avg >= 1.5) return QUALITY_LEVELS.FAIR;
  return QUALITY_LEVELS.POOR;
}

function scoreToQuality_(score) {
  const value = Number(score);
  if (score === '' || score == null || isNaN(value)) return '';
  if (value >= 80) return QUALITY_LEVELS.EXCELLENT;
  if (value >= 60) return QUALITY_LEVELS.GOOD;
  if (value >= 40) return QUALITY_LEVELS.FAIR;
  return QUALITY_LEVELS.POOR;
}

// ============================================================
// User Management (Admin)
// ============================================================

function createUser(data) {
  try {
    const sheet = getSheet(SHEETS.USERS);

    const existing = findRowByValue(sheet, 1, data.username);
    if (existing !== -1) {
      return { success: false, message: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
    }

    sheet.appendRow([
      data.username,
      data.password,
      data.fullName,
      data.role || 'teacher',
      data.department || '',
      true,
      data.schoolId || '',
      String(data.email || '').trim().toLowerCase(),
      false,
      ''
    ]);

    return { success: true, message: 'สร้างผู้ใช้สำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getUsers() {
  try {
    const sheet = getSheet(SHEETS.USERS);
    const allData = sheet.getDataRange().getValues();
    let users = [];

    for (let i = 1; i < allData.length; i++) {
      users.push({
        rowNumber: i + 1,
        username: allData[i][0],
        fullName: allData[i][2],
        role: allData[i][3],
        department: allData[i][4],
        active: allData[i][5]
      });
    }

    return { success: true, users: users };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Get Unique Values for Dropdowns
// ============================================================

function getDepartments() {
  const departments = [
    'คณิตศาสตร์', 'วิทยาศาสตร์', 'ภาษาไทย', 'ภาษาอังกฤษ',
    'สังคมศึกษา', 'พลศึกษา', 'ศิลปะ', 'การงานอาชีพ',
    'เทคโนโลยี', 'แนะแนว', '其他'
  ];
  return { success: true, departments: departments };
}

function getTeachers() {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    const teachers = new Set();

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][7]) teachers.add(allData[i][7]);
    }

    return { success: true, teachers: Array.from(teachers).sort() };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getSubjects() {
  const subjects = [
    { name: 'คณิตศาสตร์', code: 'ค21101' },
    { name: 'วิทยาศาสตร์', code: 'ว21101' },
    { name: 'ภาษาไทย', code: 'ท21101' },
    { name: 'ภาษาอังกฤษ', code: 'อ21101' },
    { name: 'สังคมศึกษา', code: 'ส21101' },
    { name: 'พลศึกษา', code: 'พ21101' },
    { name: 'ศิลปะ', code: 'ศ21101' },
    { name: 'การงานอาชีพ', code: 'ก21101' },
    { name: 'เทคโนโลยี', code: 'ท21102' }
  ];
  return { success: true, subjects: subjects };
}

function getTimeSlots() {
  const slots = [
    { time: '08:30-09:00', period: 'คาบที่ 1' },
    { time: '09:00-09:30', period: 'คาบที่ 2' },
    { time: '09:30-10:00', period: 'คาบที่ 3' },
    { time: '10:00-10:30', period: 'พักรับประทานอาหาร' },
    { time: '10:30-11:00', period: 'คาบที่ 4' },
    { time: '11:00-11:30', period: 'คาบที่ 5' },
    { time: '11:30-12:00', period: 'คาบที่ 6' },
    { time: '13:00-13:30', period: 'คาบที่ 7' },
    { time: '13:30-14:00', period: 'คาบที่ 8' },
    { time: '14:00-14:30', period: 'คาบที่ 9' }
  ];
  return { success: true, timeSlots: slots };
}

function getClassLevels() {
  const levels = [
    'ม.1', 'ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'
  ];
  return { success: true, classLevels: levels };
}

// ============================================================
// School directory (one row per school in the service area)
// ============================================================

function getSchoolColumns_(headerRow) {
  const headers = (headerRow || []).map(value => String(value || '').trim().toLowerCase());
  const aliases = {
    schoolId: ['school id', 'schoolid', 'รหัสสถานศึกษา'],
    name: ['school name', 'name', 'ชื่อโรงเรียน'],
    code: ['school code', 'code', 'รหัสโรงเรียน'],
    district: ['district', 'อำเภอ'],
    director: ['director', 'principal', 'school director', 'ชื่อผู้บริหาร', 'ผู้บริหาร'],
    phone: ['phone', 'telephone', 'โทรศัพท์'],
    email: ['email', 'อีเมล'],
    address: ['address', 'school address', 'ที่อยู่'],
    active: ['active', 'status', 'สถานะ', 'ใช้งาน'],
    createdAt: ['created at', 'created', 'วันที่สร้าง'],
    website: ['website', 'website url', 'school website', 'เว็บไซต์'],
    passwordHash: ['password hash', 'passwordhash'],
    latitude: ['latitude', 'lat', 'ละติจูด'],
    longitude: ['longitude', 'long', 'lng', 'ลองจิจูด'],
    coordinates: ['coordinates', 'coordinate', 'map url', 'google maps', 'พิกัด', 'พิกัดสถานศึกษา', 'ที่ตั้งพิกัด']
  };
  const fallback = { schoolId: 0, name: 1, code: 2, district: 3, director: 4, phone: 5, email: 6, address: 7, active: 8, createdAt: 9, website: 10 };
  const columns = {};
  Object.keys(aliases).forEach(key => {
    const found = headers.findIndex(header => aliases[key].includes(header));
    columns[key] = found >= 0 ? found : (fallback[key] === undefined ? -1 : fallback[key]);
  });
  return columns;
}

function schoolValue_(row, columns, key) {
  const index = columns[key];
  return index >= 0 && index < row.length ? row[index] : '';
}

function isSchoolActive_(value) {
  if (value === false || value === 0) return false;
  return !['false', '0', 'inactive', 'ปิดใช้งาน', 'ไม่ใช้งาน'].includes(String(value || '').trim().toLowerCase());
}

function getSchools(data) {
  try {
    const rows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
    const columns = getSchoolColumns_(rows[0] || []);
    const schools = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const schoolId = schoolValue_(row, columns, 'schoolId');
      if (data && data._user && data._user.role === 'school' && String(schoolId) !== String(data._user.schoolId)) continue;
      if (!schoolId || (!data?.includeInactive && !isSchoolActive_(schoolValue_(row, columns, 'active')))) continue;
      schools.push({
        rowNumber: i + 1,
        schoolId: schoolId,
        name: schoolValue_(row, columns, 'name'),
        code: schoolValue_(row, columns, 'code'),
        district: schoolValue_(row, columns, 'district'),
        director: schoolValue_(row, columns, 'director'),
        phone: schoolValue_(row, columns, 'phone'),
        email: schoolValue_(row, columns, 'email'),
        active: isSchoolActive_(schoolValue_(row, columns, 'active')),
        address: schoolValue_(row, columns, 'address'),
        website: schoolValue_(row, columns, 'website'),
        latitude: schoolValue_(row, columns, 'latitude'),
        longitude: schoolValue_(row, columns, 'longitude'),
        coordinates: schoolValue_(row, columns, 'coordinates')
      });
    }
    schools.sort((a, b) => String(a.name).localeCompare(String(b.name), 'th'));
    return { success: true, schools };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function isActiveSchool(schoolId) {
  if (!schoolId) return false;
  const rows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
  const columns = getSchoolColumns_(rows[0] || []);
  return rows.slice(1).some(row => String(schoolValue_(row, columns, 'schoolId')) === String(schoolId) && isSchoolActive_(schoolValue_(row, columns, 'active')));
}

function getSchoolNameMap_() {
  const rows = getSheet(SHEETS.SCHOOLS).getDataRange().getValues();
  const columns = getSchoolColumns_(rows[0] || []);
  const names = {};
  for (let i = 1; i < rows.length; i++) {
    const schoolId = schoolValue_(rows[i], columns, 'schoolId');
    if (schoolId) names[String(schoolId)] = schoolValue_(rows[i], columns, 'name');
  }
  return names;
}

function createSchool(data) {
  try {
    const name = String(data.name || '').trim();
    const code = String(data.code || '').trim();
    if (!name || !code) return { success: false, message: 'กรุณาระบุชื่อโรงเรียนและรหัสโรงเรียน' };
    const sheet = getSheet(SHEETS.SCHOOLS);
    const rows = sheet.getDataRange().getValues();
    const columns = getSchoolColumns_(rows[0] || []);
    if (rows.slice(1).some(row => String(schoolValue_(row, columns, 'code')).trim().toLowerCase() === code.toLowerCase())) {
      return { success: false, message: 'รหัสโรงเรียนนี้มีอยู่แล้ว' };
    }
    const schoolId = generateId();
    const row = Array(sheet.getLastColumn()).fill('');
    const values = { schoolId, name, code, district: data.district || '', director: data.director || '', phone: data.phone || '', email: data.email || '', address: data.address || '', active: true, createdAt: formatDateTime(new Date()), website: data.website || '' };
    Object.keys(values).forEach(key => { if (columns[key] >= 0) row[columns[key]] = values[key]; });
    sheet.appendRow(row);
    return { success: true, message: 'เพิ่มโรงเรียนสำเร็จ', schoolId };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateSchool(data) {
  try {
    const rowNumber = Number(data.rowNumber);
    const sheet = getSheet(SHEETS.SCHOOLS);
    if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      return { success: false, message: 'ไม่พบข้อมูลโรงเรียน' };
    }
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const columns = getSchoolColumns_(headers);
    const row = sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
    const isSchoolUser = data._user && data._user.role === 'school';
    if (isSchoolUser && String(row[columns.schoolId]) !== String(data._user.schoolId)) return {success:false,message:'ไม่มีสิทธิ์แก้ไขข้อมูลโรงเรียนอื่น'};
    const editableFields = isSchoolUser ? ['director','phone','email','address','website'] : ['name','code','district','director','phone','email','address','website'];
    editableFields.forEach(key => {
      if (data[key] !== undefined && columns[key] >= 0) row[columns[key]] = String(data[key]).trim();
    });
    if (!isSchoolUser && data.active !== undefined && columns.active >= 0) row[columns.active] = Boolean(data.active);
    sheet.getRange(rowNumber, 1, 1, lastColumn).setValues([row]);
    return { success: true, message: 'บันทึกข้อมูลโรงเรียนสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
