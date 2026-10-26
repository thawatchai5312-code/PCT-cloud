function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PCT Cloud - Phuluang hospital')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ตรวจสอบและสร้าง/อัปเกรดตารางข้อมูลใน Sheet อัตโนมัติ
function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('News');
  if (!sheet) {
    sheet = ss.insertSheet('News');
    sheet.appendRow(['ID', 'วันที่', 'ประเภท', 'แผนก', 'หัวข้อ', 'เนื้อหา', 'ลิงก์เอกสาร']);
    sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#93c5fd");
  }
  return sheet;
}

// ตรวจสอบและสร้างตารางฐานข้อมูล KPI อัตโนมัติ (เพิ่มคอลัมน์ ผู้รับผิดชอบ)
function initKpiSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('KPI_Data');
  if (!sheet) {
    sheet = ss.insertSheet('KPI_Data');
    sheet.appendRow([
      'ID', 'กลุ่มโรค', 'ชื่อตัวชี้วัด', 'เป้าหมาย', 
      'ปีย้อนหลัง4', 'ปีย้อนหลัง3', 'ปีย้อนหลัง2', 'ปีย้อนหลัง1', 'ปีปัจจุบัน',
      'Q1', 'Q2', 'Q3', 'Q4', 
      'Templateคำอธิบาย', 'ที่มาข้อมูล', 'ผู้รับผิดชอบ'
    ]);
    sheet.getRange("A1:P1").setFontWeight("bold").setBackground("#c7d2fe");
  } else {
    // อัปเกรดตารางเก่าให้มีคอลัมน์ผู้รับผิดชอบถ้ายังไม่มี
    const headers = sheet.getRange("1:1").getValues()[0];
    if (headers.indexOf('ผู้รับผิดชอบ') === -1) {
      sheet.getRange(1, 16).setValue('ผู้รับผิดชอบ').setFontWeight("bold").setBackground("#c7d2fe");
    }
  }
  return sheet;
}

// ดึงข้อมูลตัวเลือกประเภทข่าวสารและแผนก CPG จากแผ่นงาน Config
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['ประเภทข่าวสาร', 'แผนก CPG']);
    sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#64748b");
    
    const defaultTypes = ['ข่าวประชาสัมพันธ์', 'ประกาศด่วนที่สุด', 'กิจกรรมทีมนำ'];
    const defaultDepts = ['อายุกรรม', 'สูตินรีเวช', 'กุมาร', 'ศัลยกรรม', 'ออร์โถปิดิกส์', 'จิตเวช'];
    
    const maxLen = Math.max(defaultTypes.length, defaultDepts.length);
    for (let i = 0; i < maxLen; i++) {
      sheet.appendRow([defaultTypes[i] || '', defaultDepts[i] || '']);
    }
  }
  
  const data = sheet.getDataRange().getValues();
  const types = [];
  const depts = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) types.push(data[i][0].toString());
    if (data[i][1]) depts.push(data[i][1].toString());
  }
  
  return { types: types, depts: depts };
}

// เพิ่มตัวเลือกหมวดหมู่หรือแผนกใหม่ลงในแผ่นงาน Config โดยผู้ดูแลระบบ
function addConfigItem(typeOrDept, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) { getConfig();
  sheet = ss.getSheetByName('Config'); }
  
  const data = sheet.getDataRange().getValues();
  let found = false;
  if (typeOrDept === 'type') {
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) {
        sheet.getRange(i + 1, 1).setValue(value);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([value, '']);
  } else if (typeOrDept === 'dept') {
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow(['', value]);
  }
  
  return getConfig();
}

// ดึงข้อมูลประกาศทั้งหมดจาก Google Sheet
function getNews() {
  const sheet = initSheet();
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const newsList = [];
  for (let i = 1; i < data.length; i++) {
    newsList.push({
      id: (data[i][0] != null) ? data[i][0].toString() : '',
      date: (data[i][1] != null) ? data[i][1].toString() : '',
      type: (data[i][2] != null) ? data[i][2].toString() : '',
      dept: (data[i][3] != null) ? data[i][3].toString() : '',
      title: (data[i][4] != null) ? data[i][4].toString() : '',
      content: (data[i][5] != null) ? data[i][5].toString() : '',
      link: (data[i][6] != null) ? data[i][6].toString() : ''
    });
  }
  return newsList.reverse();
}

// บันทึกข้อมูลข่าวสารหรือแนวทาง CPG ลง Google Sheet
function addNews(type, dept, title, content, link) {
  const sheet = initSheet();
  const today = new Date();
  const thaiMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const dateString = `${today.getDate()} ${thaiMonths[today.getMonth()]} ${today.getFullYear() + 543}`;
  const id = Date.now().toString();
  
  sheet.appendRow([id, dateString, type, dept, title, content, link]);
  return getNews();
}

// ลบประกาศออกจาก Google Sheet ตาม ID
function deleteNews(id) {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getNews();
}

// ดึงรายการข้อมูล KPI ทั้งหมด
function getKpiData() {
  const sheet = initKpiSheet();
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const kpiList = [];
  for (let i = 1; i < data.length; i++) {
    kpiList.push({
      id: data[i][0],
      group: data[i][1],
      title: data[i][2],
      target: data[i][3],
      y4: data[i][4],
      y3: data[i][5],
      y2: data[i][6],
      y1: data[i][7],
      y0: data[i][8],
      q1: data[i][9],
      q2: data[i][10],
      q3: data[i][11],
      q4: data[i][12],
      template: data[i][13],
      source: data[i][14],
      responsible: data[i][15] || ''
    });
  }
  return kpiList;
}

// เพิ่มตัวชี้วัด KPI ใหม่โดยแอดมิน
function addKpiItem(kpiObj) {
  const sheet = initKpiSheet();
  const id = Date.now().toString();
  sheet.appendRow([
    id, kpiObj.group, kpiObj.title, kpiObj.target,
    kpiObj.y4, kpiObj.y3, kpiObj.y2, kpiObj.y1, kpiObj.y0,
    kpiObj.q1, kpiObj.q2, kpiObj.q3, kpiObj.q4,
    kpiObj.template, kpiObj.source, kpiObj.responsible
  ]);
  return getKpiData();
}

// อัปเดตข้อมูล KPI 
function updateKpiItem(kpiObj) {
  const sheet = initKpiSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === kpiObj.id.toString()) {
      sheet.getRange(i + 1, 2, 1, 15).setValues([[
        kpiObj.group, kpiObj.title, kpiObj.target,
        kpiObj.y4, kpiObj.y3, kpiObj.y2, kpiObj.y1, kpiObj.y0,
        kpiObj.q1, kpiObj.q2, kpiObj.q3, kpiObj.q4,
        kpiObj.template, kpiObj.source, kpiObj.responsible
      ]]);
      break;
    }
  }
  return getKpiData();
}

// ลบตัวชี้วัด KPI
function deleteKpiItem(id) {
  const sheet = initKpiSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getKpiData();
}

// --- ส่วนของระบบผู้ใช้งานและการยืนยันสิทธิ์ ---

function initUserSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['ID', 'Username', 'Password', 'Role']);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#fcd34d");
    
    // ใส่ข้อมูลพื้นฐาน
    sheet.appendRow([Date.now().toString(), 'Admin', 'Admin11038', 'Admin']);
    sheet.appendRow(['_SIMPLE_PASSWORD_', 'General', '11038', 'General']);
  }
  return sheet;
}

function verifyUserLogin(username, password) {
  const sheet = initUserSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString() === username.toString() && data[i][2].toString() === password.toString() && data[i][0] !== '_SIMPLE_PASSWORD_') {
      return data[i][3]; // Return Role ('Admin' หรือ 'Identified')
    }
  }
  return null; 
}

function verifySimpleLogin(password) {
  const sheet = initUserSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === '_SIMPLE_PASSWORD_') {
      return data[i][2].toString() === password.toString();
    }
  }
  return false;
}

function getUsersList() {
  const sheet = initUserSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '_SIMPLE_PASSWORD_') {
      users.push({
        id: data[i][0].toString(),
        username: data[i][1].toString(),
        password: data[i][2].toString(),
        role: data[i][3].toString()
      });
    }
  }
  return users;
}

function saveUser(id, username, password, role) {
  const sheet = initUserSheet();
  if (!id) {
    const newId = Date.now().toString();
    sheet.appendRow([newId, username, password, role]);
  } else {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.getRange(i + 1, 2).setValue(username);
        sheet.getRange(i + 1, 3).setValue(password);
        sheet.getRange(i + 1, 4).setValue(role);
        break;
      }
    }
  }
  return getUsersList();
}

function deleteUser(id) {
  const sheet = initUserSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getUsersList();
}

function getSimplePassword() {
  const sheet = initUserSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === '_SIMPLE_PASSWORD_') {
      return data[i][2].toString();
    }
  }
  return '11038';
}

function updateSimplePassword(newPassword) {
  const sheet = initUserSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === '_SIMPLE_PASSWORD_') {
      sheet.getRange(i + 1, 3).setValue(newPassword);
      return true;
    }
  }
  return false;
}