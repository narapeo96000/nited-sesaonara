// ============================================================
// ระบบนิเทศออนไลน์ สำนักงานเขตพื้นที่การศึกษามัธยมศึกษานราธิวาส
// Google Apps Script Backend
// ============================================================

const SPREADSHEET_ID = '1ojkushQJ6-ueIvmTwhxqOQbsYAXMzVNMKS-UcyKfAQ0';
const DRIVE_FOLDER_ID = '1FDj-gDOWuvq8A_kugvPrBhPZ8X4pHVUo';
// ตั้งค่า NITED_ADMIN_PASSWORD ใน Script Properties ก่อนเปิดใช้งานจริง
const ADMIN_PASSWORD_PROPERTY = 'NITED_ADMIN_PASSWORD';

const SHEETS = {
  BOOKING: 'Booking',
  FILES: 'Files',
  EVALUATION: 'Inspections',
  USERS: 'Users',
  SCHOOLS: 'Schools'
};

const STATUS = {
  PENDING: 'รอดำเนินการ',
  CONFIRMED: 'ยืนยันแล้ว',
  REJECTED: 'ปฏิเสธ',
  COMPLETED: 'นิเทศแล้ว'
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
    .setTitle('ระบบนิเทศออนไลน์ สพม.นราธิวาส')
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
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    switch (action) {
      case 'login': return jsonResponse(login(data));
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
      case 'createSchool': return jsonResponse(createSchool(data));
      case 'updateSchool': return jsonResponse(updateSchool(data));
      default: return jsonResponse({ success: false, message: 'ไม่รู้จักคำสั่ง' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheet(sheet, name);
  }
  ensureSheetHeaders(sheet, name);
  return sheet;
}

function ensureSheetHeaders(sheet, name) {
  const headers = {
    [SHEETS.BOOKING]: ['Timestamp', 'Date', 'Time', 'Teacher Name', 'Department', 'Period', 'Subject Name', 'Subject Code', 'Class Level', 'Room', 'Status', 'Notes', 'School ID'],
    [SHEETS.FILES]: ['Timestamp', 'Teacher Name', 'Booking ID', 'File Type', 'File Name', 'File URL', 'Drive File ID', 'Status', 'Admin Note', 'Reviewed By', 'Reviewed Date', 'School ID'],
    [SHEETS.EVALUATION]: ['Timestamp', 'Supervision Type', 'Supervisor Name', 'School Name', 'School Code', 'Teacher Name', 'Subject Group', 'Subject Name', 'Grade Level', 'Period', 'Teaching Date', 'Topic', 'Techniques', 'Scores', 'Total Score', 'Strengths', 'Improvements', 'Suggestions', 'Evaluated By'],
    [SHEETS.USERS]: ['Username', 'Password', 'Full Name', 'Role', 'Department', 'Active', 'School ID'],
    [SHEETS.SCHOOLS]: ['School ID', 'School Name', 'School Code', 'District', 'Director', 'Phone', 'Email', 'Address', 'Active', 'Created At', 'Website']
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

function initializeSheet(sheet, name) {
  const headers = {
    [SHEETS.BOOKING]: ['Timestamp', 'Date', 'Time', 'Teacher Name', 'Department', 'Period', 'Subject Name', 'Subject Code', 'Class Level', 'Room', 'Status', 'Notes', 'School ID'],
    [SHEETS.FILES]: ['Timestamp', 'Teacher Name', 'Booking ID', 'File Type', 'File Name', 'File URL', 'Drive File ID', 'Status', 'Admin Note', 'Reviewed By', 'Reviewed Date', 'School ID'],
    [SHEETS.EVALUATION]: ['Timestamp', 'Supervision Type', 'Supervisor Name', 'School Name', 'School Code', 'Teacher Name', 'Subject Group', 'Subject Name', 'Grade Level', 'Period', 'Teaching Date', 'Topic', 'Techniques', 'Scores', 'Total Score', 'Strengths', 'Improvements', 'Suggestions', 'Evaluated By'],
    [SHEETS.USERS]: ['Username', 'Password', 'Full Name', 'Role', 'Department', 'Active', 'School ID'],
    [SHEETS.SCHOOLS]: ['School ID', 'School Name', 'School Code', 'District', 'Director', 'Phone', 'Email', 'Address', 'Active', 'Created At', 'Website']
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
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function formatDateTime(date) {
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
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
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
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
    const { username, password } = data;

    const adminPassword = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSWORD_PROPERTY);
    if (adminPassword && username === 'admin' && password === adminPassword) {
      return {
        success: true,
        user: {
          username: 'admin',
          fullName: 'ผู้ดูแลระบบ',
          role: 'admin',
          department: 'ทั้งหมด'
        }
      };
    }

    const sheet = getSheet(SHEETS.USERS);
    const rowData = sheet.getDataRange().getValues();

    for (let i = 1; i < rowData.length; i++) {
      if (rowData[i][0] === username && rowData[i][1] === password && rowData[i][5] !== false) {
        return {
          success: true,
          user: {
            username: rowData[i][0],
            fullName: rowData[i][2],
            role: rowData[i][3],
          department: rowData[i][4],
          schoolId: rowData[i][6] || ''
          }
        };
      }
    }

    return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Booking Management
// ============================================================

function createBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const id = generateId();
    if (!isActiveSchool(data.schoolId)) {
      return { success: false, message: 'กรุณาเลือกโรงเรียนผู้รับการนิเทศที่อยู่ในสังกัด' };
    }
    if (!data.supervisionType || !data.teachingDate || !data.topic) {
      return { success: false, message: 'กรุณาระบุประเภทการนิเทศ วันที่จัดการเรียนรู้ และหัวข้อที่สอน' };
    }
    if (data.totalScore !== '' && data.totalScore != null && (isNaN(Number(data.totalScore)) || Number(data.totalScore) < 0)) {
      return { success: false, message: 'คะแนนรวมต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป' };
    }

    const conflicting = checkBookingConflict(data.date, data.time, data.period, data.schoolId);
    if (conflicting) {
      return { success: false, message: 'วัน-เวลานี้มีการจองแล้ว กรุณาเลือกวัน-เวลาอื่น' };
    }

    const row = [
      formatDateTime(new Date()),
      data.date,
      data.time,
      data.teacherName,
      data.department,
      data.period,
      data.subjectName,
      data.subjectCode,
      data.classLevel,
      data.room,
      STATUS.PENDING,
      data.notes || '',
      data.schoolId || ''
    ];

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

function checkBookingConflict(date, time, period, schoolId) {
  const sheet = getSheet(SHEETS.BOOKING);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === date && data[i][2] === time && data[i][5] === period && data[i][12] === schoolId) {
      if (data[i][10] !== STATUS.REJECTED) {
        return true;
      }
    }
  }
  return false;
}

function updateBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการจอง' };
    }

    const currentData = sheet.getRange(rowNum, 1, 1, 12).getValues()[0];

    if (data.status) currentData[10] = data.status;
    if (data.notes !== undefined) currentData[11] = data.notes;
    if (data.date) currentData[1] = data.date;
    if (data.time) currentData[2] = data.time;
    if (data.period) currentData[5] = data.period;

    sheet.getRange(rowNum, 1, 1, 12).setValues([currentData]);

    return { success: true, message: 'อัพเดทการจองสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    sheet.deleteRow(data.rowNumber);
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
    const schoolNames = getSchoolNameMap_();
    let bookings = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }
      row['School Name'] = schoolNames[String(allData[i][12] || '')] || '';
      row.bookingIndex = i;

      if (data && data.status && allData[i][10] !== data.status) continue;
      if (data && data.teacherName && allData[i][3] !== data.teacherName) continue;
      if (data && data.date && allData[i][1] !== data.date) continue;
      if (data && data.department && allData[i][4] !== data.department) continue;

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
      if (i + 1 === data.rowNumber || allData[i][0] === data.bookingId) {
        const row = {};
        row.rowNumber = i + 1;
        for (let j = 0; j < headers.length; j++) {
          row[headers[j]] = allData[i][j];
        }
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
    if (!isActiveSchool(data.schoolId)) {
      return { success: false, message: 'กรุณาเลือกโรงเรียนผู้รับการนิเทศที่อยู่ในสังกัด' };
    }
    const school = getSchools({}).schools.find(item => String(item.schoolId) === String(data.schoolId));
    const row = [formatDateTime(new Date()), data.supervisionType || '', data.supervisorName || '',
      school.name, school.code, data.teacherName || '', data.subjectGroup || '', data.subjectName || '',
      data.gradeLevel || '', data.period || '', data.teachingDate || '', data.topic || '', data.techniques || '',
      data.scores || '', data.totalScore === '' || data.totalScore == null ? '' : Number(data.totalScore),
      data.strengths || '', data.improvements || '', data.suggestions || '', data.evaluatedBy || ''];
    sheet.appendRow(row);
    return {
      success: true,
      message: 'บันทึกข้อมูลการนิเทศสำเร็จ'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการประเมิน' };
    }

    const currentData = sheet.getRange(rowNum, 1, 1, 19).getValues()[0];
    if (data.scores !== undefined) currentData[13] = data.scores;
    if (data.totalScore !== undefined) currentData[14] = data.totalScore;
    if (data.strengths !== undefined) currentData[15] = data.strengths;
    if (data.improvements !== undefined) currentData[16] = data.improvements;
    if (data.suggestions !== undefined) currentData[17] = data.suggestions;
    sheet.getRange(rowNum, 1, 1, 19).setValues([currentData]);

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
      if (data && data.teacherName && row['Teacher Name'] !== data.teacherName) continue;
      if (data && data.schoolId && row['School Code'] !== data.schoolId && row['School Name'] !== data.schoolId) continue;

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
    sheet.deleteRow(data.rowNumber);
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

    const bookingData = bookingSheet.getDataRange().getValues();
    const fileData = fileSheet.getDataRange().getValues();
    const evalData = evalSheet.getDataRange().getValues();
    const schoolNames = getSchoolNameMap_();

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
      const status = bookingData[i][10];
      const bookingDate = bookingData[i][1];

      if (status === STATUS.PENDING) stats.pendingBookings++;
      if (status === STATUS.CONFIRMED) stats.confirmedBookings++;
      if (status === STATUS.COMPLETED) stats.completedBookings++;
      if (status === STATUS.REJECTED) stats.rejectedBookings++;
      if (bookingDate === today) stats.todayBookings++;
      if (bookingDate && bookingDate.toString().startsWith(thisMonth)) stats.thisMonthBookings++;

      const dept = bookingData[i][4];
      if (dept) {
        if (!stats.departmentStats[dept]) {
          stats.departmentStats[dept] = { total: 0, completed: 0, pending: 0 };
        }
        stats.departmentStats[dept].total++;
        if (status === STATUS.COMPLETED) stats.departmentStats[dept].completed++;
        if (status === STATUS.PENDING) stats.departmentStats[dept].pending++;
      }

      const teacher = bookingData[i][3];
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
      teacherName: row[3],
      department: row[4],
      period: row[5],
      subjectName: row[6],
      status: row[10],
      schoolId: row[12] || '',
      schoolName: schoolNames[String(row[12] || '')] || ''
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
    const schoolNames = getSchoolNameMap_();
    let events = [];

    const month = data.month || new Date().getMonth();
    const year = data.year || new Date().getFullYear();

    for (let i = 1; i < allData.length; i++) {
      const bookingDate = allData[i][1];
      if (!bookingDate) continue;

      const d = new Date(bookingDate);
      if (d.getMonth() == month && d.getFullYear() == year) {
        events.push({
          date: bookingDate,
          title: (schoolNames[String(allData[i][12] || '')] || 'โรงเรียน') + (allData[i][6] ? ' - ' + allData[i][6] : ''),
          time: allData[i][2],
          status: allData[i][10],
          department: allData[i][4]
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
      const d = bookingData[i][1];
      if (startDate && d < startDate) continue;
      if (endDate && d > endDate) continue;
      report.bookings.push({
        date: d,
        time: bookingData[i][2],
        teacherName: bookingData[i][3],
        department: bookingData[i][4],
        period: bookingData[i][5],
        subjectName: bookingData[i][6],
        status: bookingData[i][10]
      });
    }

    const fileData = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileData.length; i++) {
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
      if (bookingData[i][3] === teacherName) {
        report.bookings.push({
          date: bookingData[i][1],
          time: bookingData[i][2],
          subjectName: bookingData[i][6],
          status: bookingData[i][10]
        });
      }
    }

    const fileData = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileData.length; i++) {
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
      if (bookingData[i][4] === department) {
        const teacher = bookingData[i][3];
        if (!report.teachers[teacher]) {
          report.teachers[teacher] = {
            totalBookings: 0,
            completedBookings: 0,
            files: 0,
            evaluations: []
          };
        }
        report.teachers[teacher].totalBookings++;
        if (bookingData[i][10] === STATUS.COMPLETED) {
          report.teachers[teacher].completedBookings++;
        }
      }
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
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
      data.schoolId || ''
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
      if (allData[i][3]) teachers.add(allData[i][3]);
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
    ['name', 'code', 'district', 'director', 'phone', 'email', 'address', 'website'].forEach(key => {
      if (data[key] !== undefined && columns[key] >= 0) row[columns[key]] = String(data[key]).trim();
    });
    if (data.active !== undefined && columns.active >= 0) row[columns.active] = Boolean(data.active);
    sheet.getRange(rowNumber, 1, 1, lastColumn).setValues([row]);
    return { success: true, message: 'บันทึกข้อมูลโรงเรียนสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
