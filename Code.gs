/**
 * ระบบเบิกยา / First Aid Medicine Request System
 * Developer: Pro Gramer Ai Code
 * Company: Thai Semcon
 */

const SPREADSHEET_ID = '1Zz-QofAFnXBKzg_SdK0JOn3o1WmmC9XZXxuPDOPbI_4';
const SHEET_HISTORY = 'FM-OHS-87'; // ประวัติเบิกยา
const SHEET_STOCK = 'FM-OHS-86';   // สต๊อกยา

// ==========================================
// ส่วนรหัสผู้ดูแลระบบ (Admin Access Codes)
// เพิ่ม/แก้ไขรหัส Admin ได้ที่นี่
// ==========================================
const ADMIN_USERS = Object.freeze({
  '2017045': 'Thossapol',
  '2017077': 'Apinya',
  '2023057': 'Jittipong',
  '2018089': 'Pirayut',
  '2023058': 'Jirapan',
  '2025057': 'Suthada',
  '2025053': 'Rattatammanoon'
});

function validateAdminCode(adminCode) {
  const code = String(adminCode || '').trim();
  const name = ADMIN_USERS[code];

  if (!name) {
    return {
      success: false,
      message: 'รหัสพนักงานไม่ถูกต้อง ไม่มีสิทธิ์เข้าถึง'
    };
  }

  return {
    success: true,
    code: code,
    name: name,
    message: 'เข้าสู่ระบบสำเร็จ'
  };
}

// 1. ฟังก์ชันเริ่มต้นสำหรับ Web App
function doGet(e) {
  let template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Thai Semcon - Medicine Request System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// ส่วนที่ 1: การอ่านข้อมูล (Read Data)
// ==========================================

function getBorrowRecords() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HISTORY);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const records = data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
  
  return records.reverse();
}

function getMedicineCategories() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STOCK);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const categories = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) categories.add(data[i][1]);
  }
  return Array.from(categories);
}

function getMedicinesByCategory(category) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STOCK);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const today = new Date();
  today.setHours(0,0,0,0);
  let medicines = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let rowCategory = row[1];
    let remainQty = parseInt(row[5]) || 0;
    let expDateStr = row[8];
    let status = row[11];
    
    let isExpired = false;
    let expParts = expDateStr.split('/');
    if(expParts.length === 3) {
       let expDate = new Date(expParts[2], expParts[1] - 1, expParts[0]);
       if(expDate < today) isExpired = true;
    }

    if (rowCategory === category && remainQty > 0 && !isExpired && status !== 'หมดอายุ') {
      medicines.push({
        id: row[0],
        nameTH: row[2],
        nameEN: row[3],
        remainQty: remainQty,
        unit: row[6],
        mfg: row[7],
        exp: row[8],
        imageUrl: row[13] || '' // ดึงข้อมูลรูปภาพจากคอลัมน์ N (Index 13)
      });
    }
  }
  return medicines;
}

// ==========================================
// ส่วนที่ 2: การบันทึกและจัดการข้อมูล (Write Data)
// ==========================================

function generateRequestId(sheet) {
  const year = new Date().getFullYear();
  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  
  for (let i = 1; i < data.length; i++) {
    let reqId = data[i][1];
    if (reqId && reqId.toString().startsWith(`MED-${year}-`)) {
      let num = parseInt(reqId.split('-')[2]);
      if (num > maxId) maxId = num;
    }
  }
  let nextId = (maxId + 1).toString().padStart(4, '0');
  return `MED-${year}-${nextId}`;
}

function submitMedicineRequest(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    const stockSheet = ss.getSheetByName(SHEET_STOCK);

    const items = Array.isArray(formData.items) && formData.items.length
      ? formData.items
      : [{
          category: formData.category,
          medTH: formData.medTH,
          medEN: formData.medEN,
          qty: formData.qty,
          unit: formData.unit
        }];

    if (!items.length) throw new Error('ไม่พบรายการยาที่ต้องการเบิก');

    const stockData = stockSheet.getDataRange().getValues();

    // รวมรายการยาซ้ำในคำขอเดียวกัน เพื่อป้องกันตัดสต๊อกผิดพลาด
    const requestedMap = {};
    items.forEach(item => {
      const medTH = String(item.medTH || '').trim();
      const qty = parseInt(item.qty, 10) || 0;
      if (!medTH) throw new Error('พบรายการยาไม่สมบูรณ์');
      if (qty <= 0) throw new Error('จำนวนยาต้องมากกว่า 0');

      if (!requestedMap[medTH]) {
        requestedMap[medTH] = {
          category: item.category || '',
          medTH: medTH,
          medEN: item.medEN || '',
          qty: 0,
          unit: item.unit || '',
          stockRowIndex: -1,
          currentStock: 0,
          importQty: 0
        };
      }
      requestedMap[medTH].qty += qty;
    });

    const requestedItems = Object.values(requestedMap);

    // ตรวจสอบสต๊อกทั้งหมดก่อนบันทึกจริง ถ้ารายการใดไม่ผ่านจะไม่ตัดสต๊อกเลย
    requestedItems.forEach(item => {
      for (let i = 1; i < stockData.length; i++) {
        if (String(stockData[i][2]).trim() === item.medTH) {
          item.stockRowIndex = i + 1;
          item.importQty = parseInt(stockData[i][4], 10) || 0;
          item.currentStock = parseInt(stockData[i][5], 10) || 0;
          if (!item.medEN) item.medEN = stockData[i][3] || '';
          if (!item.unit) item.unit = stockData[i][6] || '';
          if (!item.category) item.category = stockData[i][1] || '';
          break;
        }
      }

      if (item.stockRowIndex === -1) throw new Error(`ไม่พบข้อมูลยาในระบบ: ${item.medTH}`);
      if (item.qty > item.currentStock) throw new Error(`จำนวนยาไม่เพียงพอ: ${item.medTH} คงเหลือ ${item.currentStock} ${item.unit}`);
    });

    const reqId = generateRequestId(historySheet);
    const timestamp = new Date();
    const requestDateText = Utilities.formatDate(timestamp, 'GMT+7', 'dd/MM/yyyy HH:mm');

    // ตัดสต๊อกและอัปเดตสถานะ
    requestedItems.forEach(item => {
      const newStock = item.currentStock - item.qty;
      stockSheet.getRange(item.stockRowIndex, 6).setValue(newStock);

      let newStatus = 'พร้อมใช้งาน';
      if (newStock === 0) newStatus = 'หมดสต๊อก';
      else if (item.importQty > 0 && newStock <= (item.importQty * 0.10)) newStatus = 'ใกล้หมด';

      stockSheet.getRange(item.stockRowIndex, 12).setValue(newStatus);
    });

    // บันทึกประวัติ 1 แถวต่อยา 1 ชนิด แต่ใช้ RequestID เดียวกัน เพื่อรู้ว่าเป็นการเบิกครั้งเดียวกัน
    const rows = requestedItems.map(item => ([
      timestamp,
      reqId,
      requestDateText,
      formData.name,
      formData.company,
      formData.gender,
      formData.symptom,
      item.category,
      item.medTH,
      item.medEN,
      item.qty,
      item.unit,
      'System',
      'Completed'
    ]));

    historySheet
      .getRange(historySheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);

    return {
      success: true,
      message: 'บันทึกการเบิกยาสำเร็จ',
      requestId: reqId,
      itemCount: requestedItems.length
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // ignore lock release error
    }
  }
}


// ==========================================
// ส่วนที่ 3: ระบบ Admin (Admin System)
// ==========================================

function getStockDashboard() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const stockSheet = ss.getSheetByName(SHEET_STOCK);
  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  
  const stockData = stockSheet.getDataRange().getDisplayValues();
  const historyData = historySheet.getDataRange().getDisplayValues();
  
  let dashboard = {
    totalItems: stockData.length > 1 ? stockData.length - 1 : 0,
    totalRemaining: 0,
    lowStock: 0,
    expired: 0,
    reqToday: 0,
    reqMonth: 0,
    stockList: []
  };
  
  const today = new Date();
  today.setHours(0,0,0,0);
  
  for (let i = 1; i < stockData.length; i++) {
    let importQty = parseInt(stockData[i][4]) || 0;
    let remain = parseInt(stockData[i][5]) || 0;
    let status = stockData[i][11];
    let expDateStr = stockData[i][8];
    
    dashboard.totalRemaining += remain;
    
    let isExpired = false;
    let expParts = expDateStr.split('/');
    if(expParts.length === 3) {
       let expDate = new Date(expParts[2], expParts[1] - 1, expParts[0]);
       if(expDate < today) isExpired = true;
    }
    
    if (isExpired || status === 'หมดอายุ') {
      dashboard.expired++;
      stockSheet.getRange(i + 1, 12).setValue('หมดอายุ');
      status = 'หมดอายุ';
    } else if (remain === 0) {
      status = 'หมดสต๊อก';
    } else if (importQty > 0 && remain <= (importQty * 0.10)) {
      dashboard.lowStock++;
      status = 'ใกล้หมด';
    } else {
      status = 'พร้อมใช้งาน';
    }
    
    dashboard.stockList.push({
      rowIndex: i + 1,
      category: stockData[i][1],
      medTH: stockData[i][2],
      medEN: stockData[i][3],
      importQty: importQty,
      remainQty: remain,
      unit: stockData[i][6],
      mfg: stockData[i][7],
      exp: stockData[i][8],
      recorder: stockData[i][9] || '-',
      status: status,
      lot: stockData[i][12] || '',
      imageUrl: stockData[i][13] || '' // ดึงข้อมูลรูปลง Dashboard
    });
  }
  
  const currMonth = today.getMonth();
  const currYear = today.getFullYear();
  const todayStr = Utilities.formatDate(today, "GMT+7", "dd/MM/yyyy");
  
  for (let i = 1; i < historyData.length; i++) {
    let reqDateStr = historyData[i][2].split(' ')[0];
    if (reqDateStr === todayStr) dashboard.reqToday++;
    
    let reqDateParts = reqDateStr.split('/');
    if(reqDateParts.length === 3) {
      if (parseInt(reqDateParts[1]) - 1 === currMonth && parseInt(reqDateParts[2]) === currYear) {
        dashboard.reqMonth++;
      }
    }
  }
  
  return dashboard;
}

function addMedicineStock(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_STOCK);
    
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
    const medId = "M" + new Date().getTime(); 
    
    // บันทึกรูปภาพ (data.imageUrl) ในคอลัมน์ N (ลำดับที่ 14)
    sheet.appendRow([
      medId, data.category, data.nameTH, data.nameEN, data.qty, data.qty,
      data.unit, data.mfg, data.exp, data.adminName, timestamp, "พร้อมใช้งาน", data.lot, data.imageUrl
    ]);
    
    return { success: true, message: "เพิ่มรายการยาสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateMedicineStock(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_STOCK);
    
    const row = data.rowIndex;
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
    
    sheet.getRange(row, 2).setValue(data.category);
    sheet.getRange(row, 3).setValue(data.nameTH);
    sheet.getRange(row, 4).setValue(data.nameEN);
    sheet.getRange(row, 5).setValue(data.importQty);
    sheet.getRange(row, 6).setValue(data.remainQty);
    sheet.getRange(row, 7).setValue(data.unit);
    sheet.getRange(row, 8).setValue(data.mfg);
    sheet.getRange(row, 9).setValue(data.exp);
    sheet.getRange(row, 10).setValue(data.adminName);
    sheet.getRange(row, 11).setValue(timestamp);
    sheet.getRange(row, 13).setValue(data.lot);
    sheet.getRange(row, 14).setValue(data.imageUrl); // อัปเดตข้อมูลรูปภาพในคอลัมน์ N
    
    let remQty = parseInt(data.remainQty) || 0;
    let impQty = parseInt(data.importQty) || 0;
    let status = 'พร้อมใช้งาน';
    if (remQty === 0) status = 'หมดสต๊อก';
    else if (impQty > 0 && remQty <= (impQty * 0.10)) status = 'ใกล้หมด';
    
    sheet.getRange(row, 12).setValue(status);
    
    return { success: true, message: "อัปเดตข้อมูลยาสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
