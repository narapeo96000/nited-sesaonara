// ============================================================
// Configuration - ระบบนิเทศออนไลน์สพม.นราธิวาส
// ============================================================

const CONFIG = {
  // Google Apps Script Web App URL
  // แก้ไข URL นี้หลังจาก Deploy Apps Script เป็น Web App
  APPS_SCRIPT_URL: ['localhost', '127.0.0.1'].includes(location.hostname)
    ? `${location.origin}/api`
    : 'https://script.google.com/macros/s/AKfycbzr6Ar78mYnyAKZ_DlO4_WnQqwwFI0VNxZWkqtG4HxNiLy650wYMeVTzfRZqW89Ok3QPg/exec',

  // Google Sheets ID
  SPREADSHEET_ID: '1ojkushQJ6-ueIvmTwhxqOQbsYAXMzVNMKS-UcyKfAQ0',

  // Google Drive Folder ID
  DRIVE_FOLDER_ID: '1FDj-gDOWuvq8A_kugvPrBhPZ8X4pHVUo',

  // System Info
  SCHOOL_NAME: 'สำนักงานเขตพื้นที่การศึกษามัธยมศึกษานราธิวาส',
  SYSTEM_NAME: 'ระบบนิเทศสถานศึกษา',

};
