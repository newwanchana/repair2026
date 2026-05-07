// ===== CONFIGURATION =====
const SPREADSHEET_ID = '1nAek9XjJT_KcJVzM8pm3Y8nVZ7GZlLtUeCcxacZRx58';  //ใส่ ID
const DRIVE_FOLDER_ID = '18IfyUFmCWxPIYVcCwlCU_ByhtkSTRWNv';            //ใส่ ID
// Telegram Bot setup:
// 1) สร้าง Bot โดยคุยกับ @BotFather ใน Telegram แล้วนำ Token มาใส่ด้านล่าง
// 2) หา Chat ID / Group ID โดยเพิ่ม bot เข้าแชทหรือกลุ่ม แล้วเรียก Telegram API เช่น getUpdates
const TELEGRAM_API_BASE = 'https://api.telegram.org/bot';
const TELEGRAM_BOT_TOKEN = ''; // ใส่ Token ของ Telegram Bot เช่น '123456789:AAF...'
const TELEGRAM_CHAT_ID = '';   // ใส่ Chat ID หรือ Group ID เช่น '-1001234567890'
const DATA_SHEET_NAME = 'Data';
const EQUIPMENT_SHEET_NAME = 'EquipmentDB';

// ===== ADMIN CREDENTIALS =====
const ADMIN_USERNAME = 'admin'; //username
const ADMIN_PASSWORD = '1234';    //Password

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบแจ้งซ่อม - 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getDataSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    const headers = [
      'backend_id', 'ticket_id', 'requester_name', 'department',
      'equipment', 'description', 'urgency', 'status',
      'assigned_to', 'created_at', 'updated_at', 'image_url',
      'equipment_category'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

const EQUIPMENT_HEADERS = ['equipment_code', 'equipment_name', 'location', 'category', 'note'];

function getEquipmentSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EQUIPMENT_SHEET_NAME);
    sheet.getRange(1, 1, 1, EQUIPMENT_HEADERS.length).setValues([EQUIPMENT_HEADERS]);
    sheet.getRange(1, 1, 1, EQUIPMENT_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Auto-fix: if first cell is NOT 'equipment_code', headers are missing — insert them
  const firstCell = String(sheet.getRange(1, 1).getValue()).trim();
  if (firstCell !== 'equipment_code') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, EQUIPMENT_HEADERS.length).setValues([EQUIPMENT_HEADERS]);
    sheet.getRange(1, 1, 1, EQUIPMENT_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ===== EQUIPMENT DB CRUD =====

function getEquipmentByCode(code) {
  try {
    if (!code) return { success: false, error: 'ไม่พบรหัสอุปกรณ์จาก QR' };
    const sheet = getEquipmentSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, error: 'ยังไม่มีข้อมูลอุปกรณ์ในชีตฐานข้อมูล' };
    const headers = data[0];
    const codeIndex = headers.indexOf('equipment_code');
    if (codeIndex === -1) return { success: false, error: 'โครงสร้างชีตไม่ถูกต้อง' };
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codeIndex]).trim() === String(code).trim()) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
        return { success: true, data: obj };
      }
    }
    return { success: false, error: 'ไม่พบอุปกรณ์รหัส: ' + code };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAllEquipment() {
  try {
    const sheet = getEquipmentSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    const headers = data[0];
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue; // skip empty rows
      const obj = { __rowIndex: i + 1 };
      for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      items.push(obj);
    }
    return { success: true, data: items };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function createEquipment(data) {
  try {
    if (!data.equipment_code || !data.equipment_name) {
      return { success: false, error: 'กรุณากรอกรหัสและชื่ออุปกรณ์' };
    }
    // Check duplicate code
    const sheet = getEquipmentSheet();
    const existing = sheet.getDataRange().getValues();
    const headers = existing[0];
    const codeIdx = headers.indexOf('equipment_code');
    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][codeIdx]).trim() === String(data.equipment_code).trim()) {
        return { success: false, error: 'รหัสอุปกรณ์นี้มีอยู่แล้ว: ' + data.equipment_code };
      }
    }
    const row = [
      sanitizeString(data.equipment_code),
      sanitizeString(data.equipment_name),
      sanitizeString(data.location || ''),
      sanitizeString(data.category || ''),
      sanitizeString(data.note || '')
    ];
    sheet.appendRow(row);
    return { success: true, equipment_code: data.equipment_code };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateEquipment(data) {
  try {
    const sheet = getEquipmentSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const codeIdx = headers.indexOf('equipment_code');
    let rowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][codeIdx]).trim() === String(data.equipment_code).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, error: 'ไม่พบอุปกรณ์' };
    const updates = {
      equipment_code: sanitizeString(data.equipment_code),
      equipment_name: sanitizeString(data.equipment_name),
      location: sanitizeString(data.location || ''),
      category: sanitizeString(data.category || ''),
      note: sanitizeString(data.note || '')
    };
    for (const key in updates) {
      const colIdx = headers.indexOf(key);
      if (colIdx !== -1) sheet.getRange(rowIndex, colIdx + 1).setValue(updates[key]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteEquipment(equipmentCode) {
  try {
    const sheet = getEquipmentSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const codeIdx = headers.indexOf('equipment_code');
    let rowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][codeIdx]).trim() === String(equipmentCode).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, error: 'ไม่พบอุปกรณ์' };
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ===== DRIVE / IMAGE =====

function getImageFolder() {
  if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID !== '<<PUT_FOLDER_ID>>') {        //ใส่ ID
    try { return DriveApp.getFolderById(DRIVE_FOLDER_ID); } catch (e) {}
  }
  const folderName = 'RepairSystem_Images';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function saveImageToDrive(base64Data, ticketId) {
  if (!base64Data || base64Data === '') return '';
  try {
    const matches = base64Data.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) return '';
    const mimeType = matches[1];
    const base64Content = matches[2];
    let extension = 'jpg';
    if (mimeType.includes('png')) extension = 'png';
    else if (mimeType.includes('gif')) extension = 'gif';
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, ticketId + '.' + extension);
    const folder = getImageFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://lh3.googleusercontent.com/d/' + file.getId();
  } catch (e) {
    console.error('Error saving image:', e);
    return '';
  }
}

function deleteImageFromDrive(imageUrl) {
  if (!imageUrl) return;
  try {
    const match = imageUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) DriveApp.getFileById(match[1]).setTrashed(true);
  } catch (e) {}
}

// ===== HELPERS =====

function escapeTelegramHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function sendTelegramNotification(ticketData) {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) return;
  try {
    // รองรับกรณีผู้ใช้เผลอใส่ token เป็นรูปแบบ bot<token>
    const botToken = String(TELEGRAM_BOT_TOKEN).trim().replace(/^bot/i, '');
    const chatId = String(TELEGRAM_CHAT_ID).trim();
    if (!botToken || !chatId) return;
    const categoryMap = {
      building: 'อาคาร',
      location: 'สถานที่',
      classroom: 'อุปกรณ์ห้องเรียน',
      'audio-visual': 'อุปกรณ์โสตทัศนูปกรณ์',
      electrical: 'อุปกรณ์ไฟฟ้า',
      plumbing: 'อุปกรณ์ประปา',
      other: 'อื่นๆ'
    };
    const urgencyMap = {
      low: '🟢 ต่ำ',
      medium: '🟡 ปานกลาง',
      high: '🔴 สูง'
    };
    const message = [
      '<b>🔔 แจ้งซ่อมใหม่!</b>',
      '',
      '<b>🎫 รหัสงาน:</b> ' + escapeTelegramHtml(ticketData.ticket_id),
      '<b>👤 ผู้แจ้ง:</b> ' + escapeTelegramHtml(ticketData.requester_name),
      '<b>🏢 กลุ่มงาน:</b> ' + escapeTelegramHtml(ticketData.department),
      '<b>📦 หมวดหมู่:</b> ' + escapeTelegramHtml(categoryMap[ticketData.equipment_category] || ticketData.equipment_category || 'อื่นๆ'),
      '<b>🔧 อุปกรณ์/พื้นที่:</b> ' + escapeTelegramHtml(ticketData.equipment),
      '<b>📝 รายละเอียด:</b> ' + escapeTelegramHtml(ticketData.description),
      '<b>⚡ ความเร่งด่วน:</b> ' + escapeTelegramHtml(urgencyMap[ticketData.urgency] || ticketData.urgency || '-')
    ].join('\n');
    const imageUrl = String(ticketData.image_url || '').trim();
    const method = imageUrl ? 'sendPhoto' : 'sendMessage';
    const payload = imageUrl ? {
      chat_id: chatId,
      photo: imageUrl,
      caption: message,
      parse_mode: 'HTML'
    } : {
      chat_id: chatId,
      text: message,
      parse_mode: 'HTML'
    };
    const response = UrlFetchApp.fetch(TELEGRAM_API_BASE + botToken + '/' + method, {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true
    });
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    let responseData = null;
    try { responseData = JSON.parse(responseText); } catch (err) { console.error('Telegram response parse error:', responseText); }
    const success = responseCode >= 200 && responseCode < 300 && responseData && responseData.ok === true;
    if (!success) {
      console.error('Telegram API error:', { method: method, status: responseCode, body: responseText });
      if (method === 'sendPhoto') {
        const fallbackResponse = UrlFetchApp.fetch(TELEGRAM_API_BASE + botToken + '/sendMessage', {
          method: 'post',
          payload: { chat_id: chatId, text: message, parse_mode: 'HTML' },
          muteHttpExceptions: true
        });
        const fallbackCode = fallbackResponse.getResponseCode();
        const fallbackText = fallbackResponse.getContentText();
        let fallbackData = null;
        try { fallbackData = JSON.parse(fallbackText); } catch (err) { console.error('Telegram fallback parse error:', fallbackText); }
        if (!(fallbackCode >= 200 && fallbackCode < 300 && fallbackData && fallbackData.ok === true)) {
          console.error('Telegram fallback sendMessage error:', { status: fallbackCode, body: fallbackText });
        }
      }
    }
  } catch (e) {
    console.error('Error sending Telegram notification:', e);
  }
}

function sanitizeString(str) {
  if (typeof str !== 'string') return str;
  return str.replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').trim();
}

function validateTicketData(data) {
  const required = ['requester_name', 'department', 'equipment', 'description', 'urgency'];
  const errors = [];
  for (const f of required) {
    if (!data[f] || String(data[f]).trim() === '') errors.push(f + ' is required');
  }
  if (data.urgency && !['low', 'medium', 'high'].includes(data.urgency)) errors.push('Invalid urgency');
  return errors;
}

function generateBackendId() { return Utilities.getUuid(); }

// ===== TICKET CRUD =====

function createTicket(data) {
  try {
    const errors = validateTicketData(data);
    if (errors.length > 0) return { success: false, error: errors.join(', ') };
    const sheet = getDataSheet();
    const backendId = generateBackendId();
    const ticketId = 'PORMOD-' + new Date().getTime().toString(36).toUpperCase();
    const now = new Date().toISOString();
    let imageUrl = '';
    if (data.image_data && data.image_data !== '') imageUrl = saveImageToDrive(data.image_data, ticketId);
    sheet.appendRow([
      backendId, ticketId,
      sanitizeString(data.requester_name), sanitizeString(data.department),
      sanitizeString(data.equipment), sanitizeString(data.description),
      data.urgency, 'pending', '', now, now, imageUrl,
      sanitizeString(data.equipment_category || '')
    ]);
    sendTelegramNotification({
      ticket_id: ticketId,
      requester_name: data.requester_name,
      department: data.department,
      equipment_category: data.equipment_category || '',
      equipment: data.equipment,
      description: data.description,
      urgency: data.urgency,
      image_url: imageUrl
    });
    return { success: true, ticket_id: ticketId, backend_id: backendId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAllTickets() {
  try {
    const sheet = getDataSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    const headers = data[0];
    const tickets = [];
    for (let i = 1; i < data.length; i++) {
      const ticket = {};
      for (let j = 0; j < headers.length; j++) ticket[headers[j]] = data[i][j];
      ticket.__backendId = ticket.backend_id;
      delete ticket.backend_id;
      tickets.push(ticket);
    }
    return { success: true, data: tickets };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateTicket(data) {
  try {
    const sheet = getDataSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const backendIdIndex = headers.indexOf('backend_id');
    if (backendIdIndex === -1) return { success: false, error: 'Invalid sheet structure' };
    let rowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][backendIdIndex] === data.__backendId) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { success: false, error: 'Ticket not found' };
    const now = new Date().toISOString();
    const updates = { status: data.status, assigned_to: sanitizeString(data.assigned_to || ''), updated_at: now };
    for (const key in updates) {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) sheet.getRange(rowIndex, colIndex + 1).setValue(updates[key]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteTicket(backendId) {
  try {
    const sheet = getDataSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const backendIdIndex = headers.indexOf('backend_id');
    const imageUrlIndex = headers.indexOf('image_url');
    let rowIndex = -1, imageUrl = '';
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][backendIdIndex] === backendId) { rowIndex = i + 1; imageUrl = allData[i][imageUrlIndex] || ''; break; }
    }
    if (rowIndex === -1) return { success: false, error: 'Ticket not found' };
    if (imageUrl) deleteImageFromDrive(imageUrl);
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function verifyAdminLogin(username, password) {
  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) return { success: true };
  return { success: false, error: 'Invalid credentials' };
}

function getTicketCount() {
  try {
    const sheet = getDataSheet();
    return { success: true, count: Math.max(0, sheet.getLastRow() - 1) };
  } catch (e) {
    return { success: false, count: 0, error: e.toString() };
  }
}
