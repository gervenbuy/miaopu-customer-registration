// 若要寫入既有試算表，填入網址 /d/ 與 /edit 之間的試算表 ID。
// 留空後執行 setup()，會自動建立新的試算表。
const SPREADSHEET_ID = '';
const SHEET_NAME = '客戶資料';
const HEADERS = ['填寫時間','大名','手機','LINE ID','店名','職務','店址'];

function setup() {
  const properties = PropertiesService.getScriptProperties();
  let id = SPREADSHEET_ID || properties.getProperty('CUSTOMER_SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('苗圃農產行｜客戶資料登記');
  properties.setProperty('CUSTOMER_SHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  verifyHeaders_(sheet);
  sheet.getRange(1,1,1,7).setBackground('#174b32').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,180);
  sheet.setColumnWidths(2,5,150);
  sheet.setColumnWidth(7,360);
  Logger.log('Google Sheet：' + ss.getUrl());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const p = e && e.parameter ? e.parameter : {};
    const limits = {customerName:80, phone:30, lineId:80, storeName:150, jobTitle:80, address:500};
    const values = {};
    Object.keys(limits).forEach(key => {
      values[key] = String(p[key] || '').trim();
      if (values[key].length > limits[key]) throw new Error('欄位文字過長，請縮短後重試。');
    });
    ['customerName','phone','storeName','jobTitle','address'].forEach(key => {
      if (!values[key]) throw new Error('請完成所有必填欄位。');
    });
    if (!/^[+0-9 ()－-]{6,30}$/.test(values.phone)) throw new Error('請確認手機號碼格式。');
    const id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('CUSTOMER_SHEET_ID');
    if (!id) throw new Error('表單尚未完成設定，請聯絡店家。');
    lock.waitLock(30000);
    const sheet = SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('資料表尚未建立，請聯絡店家。');
    verifyHeaders_(sheet);
    const row = sheet.getLastRow()+1;
    const text = ['customerName','phone','lineId','storeName','jobTitle','address'].map(key => safeText_(values[key]));
    sheet.getRange(row,2,1,6).setNumberFormat('@');
    sheet.getRange(row,1,1,7).setValues([[new Date(),...text]]);
    sheet.getRange(row,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    SpreadsheetApp.flush();
    return resultPage_(true,'我們已收到您的客戶資料，後續將依您提供的聯絡方式與您聯繫。');
  } catch(error) {
    console.error(error);
    const publicMessages = ['欄位文字過長，請縮短後重試。','請完成所有必填欄位。','請確認手機號碼格式。','表單尚未完成設定，請聯絡店家。','資料表尚未建立，請聯絡店家。'];
    return resultPage_(false,publicMessages.includes(error.message) ? error.message : '資料尚未成功儲存，請稍後重試或聯絡店家。');
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
function doGet(){ return resultPage_(false,'請從客戶資料登記問卷填寫並送出。'); }
function verifyHeaders_(sheet){
  if (JSON.stringify(sheet.getRange(1,1,1,7).getValues()[0]) !== JSON.stringify(HEADERS)) throw new Error('資料欄位不符，請聯絡店家確認。');
}
function safeText_(value){return /^[=+@-]/.test(value) ? "'"+value : value;}
function escapeHtml_(text){return String(text).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function resultPage_(success,message){
 const title=success?'登記成功':'尚未完成登記';
 return HtmlService.createHtmlOutput('<!doctype html><html lang="zh-Hant"><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:system-ui,sans-serif;background:#faf7ed;color:#174b32;padding:50px 24px;line-height:1.8}main{max-width:480px;margin:auto;background:white;border-radius:24px;padding:36px;border:1px solid #e2e7d9}small{letter-spacing:2px}h1{font-size:28px}p{color:#667367}</style></head><body><main><small>苗圃農產行</small><h1>'+title+'</h1><p>'+escapeHtml_(message)+'</p><p>'+(success?'您可以關閉此頁。':'請回到問卷頁面確認資料後，再次送出。')+'</p></main></body></html>').setTitle(title);
}
