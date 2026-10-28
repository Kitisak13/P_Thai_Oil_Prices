/**
 * ==============================================================================
 * Thai Oil Price Daily Auto-Sync Backend (Google Apps Script)
 * ==============================================================================
 * ดึงข้อมูลราคาน้ำมันขายปลีกรายวันจากสำนักงานนโยบายและแผนพลังงาน (EPPO)
 * และบันทึกต่อท้าย Google Sheet แบบ Wide Format อัตโนมัติวันละ 2 รอบ (08:00 และ 16:00 น.)
 * 
 * ความต้องการเบื้องต้น:
 * 1. ต้องเปิดใช้งานบริการ "Drive API" ใน Apps Script (Services > + > Drive API)
 * 2. โครงสร้าง Google Sheet: Wide Format (Date, Product, PTT, BCP, Shell, ...)
 * ==============================================================================
 */

// รายชื่อหัวคอลัมน์บริษัทตามลำดับมาตรฐาน
var COMPANY_ORDER = [
  'PTT', 'BCP', 'Shell', 'Esso', 'Chevron', 'IRPC', 'PT', 'Susco', 'Pure', 'SUSCO Dealers'
];

var DETECTION_COMPANIES = [
  'SUSCO Dealers', 'PTT', 'BCP', 'Shell', 'Esso', 'Chevron', 'IRPC', 'PT', 'Susco', 'Pure'
];

var PRODUCT_MAPPING = {
  'Gasohol 95-E10': 'Gasohol 95-E10',
  'Gasohol 95-E20': 'Gasohol 95-E20',
  'Gasohol 95-E85': 'Gasohol 95-E85',
  'Gasohol 91-E10': 'Gasohol 91-E10',
  'ULG 95 RON': 'ULG 95 RON',
  'HSD B7': 'HSD B7',
  'HSD B10': 'HSD B10',
  'HSD B20': 'HSD B20'
};

/**
 * --------------------------------------------------------------------------
 * 1. การตั้งค่าและอ่านตัวแปรจาก Script Properties
 * --------------------------------------------------------------------------
 * ตัวแปรที่รองรับ (ตั้งค่าได้ที่ Project Settings > Script Properties):
 * - TARGET_SHEET_NAME : ชื่อ Sheet/แท็บ ที่ต้องการบันทึกข้อมูล (ค่าเริ่มต้น: แท็บแรก)
 * - SPREADSHEET_ID    : ID ของ Google Sheet (กรณีใช้งานแบบ Standalone Script)
 * - EPPO_URL          : URL หน้าเว็บ EPPO (ค่าเริ่มต้น: https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx)
 * - TIMEZONE          : ไทม์โซน (ค่าเริ่มต้น: Asia/Bangkok)
 * - LINE_NOTIFY_TOKEN : Token สำหรับส่งการแจ้งเตือนเข้า LINE (ทางเลือก)
 */
function getConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return {
    targetSheetName: props['TARGET_SHEET_NAME'] || '',
    spreadsheetId: props['SPREADSHEET_ID'] || '',
    eppoUrl: props['EPPO_URL'] || 'https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx',
    timezone: props['TIMEZONE'] || 'Asia/Bangkok',
    lineNotifyToken: props['LINE_NOTIFY_TOKEN'] || ''
  };
}

/**
 * ดึงออบเจ็กต์ Sheet ปลายทาง (รองรับทั้ง Container-bound และ Standalone)
 */
function getTargetSheet(config) {
  var ss;
  if (config.spreadsheetId) {
    ss = SpreadsheetApp.openById(config.spreadsheetId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error("Cannot open Spreadsheet. If running standalone, please configure 'SPREADSHEET_ID' in Script Properties.");
  }

  if (config.targetSheetName) {
    var sheet = ss.getSheetByName(config.targetSheetName);
    if (!sheet) {
      throw new Error("Sheet with name '" + config.targetSheetName + "' was not found.");
    }
    return sheet;
  }
  return ss.getSheets()[0];
}

/**
 * --------------------------------------------------------------------------
 * 2. ฟังก์ชันตรวจสอบความซ้ำซ้อน (Idempotency Check)
 * --------------------------------------------------------------------------
 * ตรวจสอบจากคอลัมน์ A (Date) ว่ามีวันที่ targetDate (YYYY-MM-DD) แล้วหรือไม่
 */
function isDateAlreadyProcessed(sheet, targetDateStr) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  // อ่านเฉพาะคอลัมน์ A ตั้งแต่แถวที่ 2 จนถึงแถวสุดท้าย
  var datesRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < datesRange.length; i++) {
    var cellVal = datesRange[i][0];
    if (cellVal instanceof Date) {
      var cellDateStr = Utilities.formatDate(cellVal, 'Asia/Bangkok', 'yyyy-MM-dd');
      if (cellDateStr === targetDateStr) return true;
    } else if (cellVal) {
      var str = String(cellVal).trim();
      if (str === targetDateStr || str.indexOf(targetDateStr) === 0) return true;
    }
  }
  return false;
}

/**
 * --------------------------------------------------------------------------
 * 3. ฟังก์ชันหลักสำหรับรันตาม Trigger (08:00 และ 16:00 น.)
 * --------------------------------------------------------------------------
 */
function dailySyncOilPrice() {
  var config = getConfig();
  var now = new Date();
  var dateIso = Utilities.formatDate(now, config.timezone, 'yyyy-MM-dd'); // 2026-10-28
  var dateEppo = Utilities.formatDate(now, config.timezone, 'dd/MM/yyyy'); // 28/10/2026

  console.info("=== Starting Daily Oil Price Sync for: " + dateIso + " (" + dateEppo + ") ===");

  var sheet = getTargetSheet(config);

  // 1. ตรวจสอบ Idempotency
  if (isDateAlreadyProcessed(sheet, dateIso)) {
    console.info("Data for date " + dateIso + " already exists in Google Sheet. Skipping execution.");
    return "SKIPPED_ALREADY_EXISTS";
  }

  // 2. ดึงข้อมูลจาก EPPO
  var result = fetchAndParseEppoData(dateIso, dateEppo, config);
  if (!result || !result.rows || result.rows.length === 0) {
    console.warn("No oil price data available from EPPO for " + dateIso + " (Possible holiday or no price announcement yet).");
    return "NO_DATA_AVAILABLE";
  }

  // 3. บันทึกข้อมูลลง Google Sheet
  var startRow = sheet.getLastRow() + 1;
  var numRows = result.rows.length;
  var numCols = result.rows[0].length;
  sheet.getRange(startRow, 1, numRows, numCols).setValues(result.rows);

  console.info("Successfully appended " + numRows + " rows for " + dateIso + " starting at row " + startRow);

  // 4. บันทึกวันที่ล่าสุดลง Script Properties (เป็นประวัติ)
  PropertiesService.getScriptProperties().setProperty('LAST_SYNCED_DATE', dateIso);

  // 5. ส่งการแจ้งเตือน (กรณีตั้ง LINE_NOTIFY_TOKEN ไว้)
  if (config.lineNotifyToken) {
    sendLineNotify(config.lineNotifyToken, "✅ [Thai Oil Price] อัปเดตราคาน้ำมันวันที่ " + dateIso + " เรียบร้อยแล้ว (" + numRows + " รายการ)");
  }

  return "SUCCESS";
}

/**
 * --------------------------------------------------------------------------
 * 4. เชื่อมต่อ EPPO Web Scraper และดาวน์โหลดไฟล์
 * --------------------------------------------------------------------------
 */
function fetchAndParseEppoData(dateIso, dateEppo, config) {
  var headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
  };

  // Step A: Initial GET to fetch ASP.NET tokens
  console.info("Fetching initial page to extract ASP.NET tokens...");
  var initRes = UrlFetchApp.fetch(config.eppoUrl, {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  });

  if (initRes.getResponseCode() !== 200) {
    throw new Error("Failed to reach EPPO site. HTTP Status: " + initRes.getResponseCode());
  }

  var tokens = extractAspNetTokens(initRes.getContentText());

  // Step B: POST request requesting Excel report for target date
  console.info("Requesting report for " + dateEppo + "...");
  var payload = {
    '__VIEWSTATE': tokens.__VIEWSTATE,
    '__VIEWSTATEGENERATOR': tokens.__VIEWSTATEGENERATOR,
    '__EVENTVALIDATION': tokens.__EVENTVALIDATION,
    'TbxToDate': dateEppo,
    'BtnGenerate.x': '10',
    'BtnGenerate.y': '10'
  };

  var postRes = UrlFetchApp.fetch(config.eppoUrl, {
    method: 'post',
    payload: payload,
    headers: headers,
    muteHttpExceptions: true
  });

  if (postRes.getResponseCode() !== 200) {
    throw new Error("POST request to EPPO failed with HTTP " + postRes.getResponseCode());
  }

  var contentType = postRes.getHeaders()['Content-Type'] || '';
  var blob = postRes.getBlob();
  var bytes = blob.getBytes();

  // ตรวจสอบ Magic bytes ของ BIFF8 OLE2 (0xD0, 0xCF, 0x11, 0xE0) หรือ Content-Type
  var isBiff8 = bytes.length > 8 && (bytes[0] & 0xFF) === 0xD0 && (bytes[1] & 0xFF) === 0xCF && (bytes[2] & 0xFF) === 0x11 && (bytes[3] & 0xFF) === 0xE0;
  var isExcel = contentType.indexOf('application/vnd.ms-excel') !== -1 || isBiff8 || bytes.length > 50000;

  if (!isExcel) {
    return null; // เซิร์ฟเวอร์ส่งกลับเป็นหน้า HTML ปกติ (ไม่มีข้อมูลราคาของวันนี้)
  }

  // Step C: แปลง .xls เป็น Google Sheet ชั่วคราวผ่าน Drive API
  blob.setName('EPPO_RetailOilPrice_on_' + dateIso.replace(/-/g, '') + '.xls');
  var tempFileId = null;

  try {
    tempFileId = convertXlsToTempSpreadsheet(blob);
    var tempSs = SpreadsheetApp.openById(tempFileId);
    var tempSheet = tempSs.getSheets()[0];
    var grid = tempSheet.getDataRange().getValues();

    return parseGridToWideRows(grid, dateIso);
  } finally {
    if (tempFileId) {
      deleteTempFile(tempFileId);
    }
  }
}

/**
 * แยกสกัด Token ของ ASP.NET จาก HTML
 */
function extractAspNetTokens(html) {
  function getVal(idName) {
    var regex1 = new RegExp('id=["\']' + idName + '["\'][^>]*value=["\']([^"\']*)["\']', 'i');
    var m = html.match(regex1);
    if (!m) {
      var regex2 = new RegExp('value=["\']([^"\']*)["\'][^>]*id=["\']' + idName + '["\']', 'i');
      m = html.match(regex2);
    }
    return m ? m[1] : '';
  }

  var vs = getVal('__VIEWSTATE');
  if (!vs) {
    throw new Error("Could not find '__VIEWSTATE' from EPPO response HTML.");
  }

  return {
    __VIEWSTATE: vs,
    __VIEWSTATEGENERATOR: getVal('__VIEWSTATEGENERATOR'),
    __EVENTVALIDATION: getVal('__EVENTVALIDATION')
  };
}

/**
 * --------------------------------------------------------------------------
 * 5. แปลงไฟล์ Excel ไบนารีผ่าน Drive API (Advanced Service)
 * --------------------------------------------------------------------------
 */
function convertXlsToTempSpreadsheet(blob) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error("Drive API Service is not enabled. Please go to Apps Script Editor > Services (+) > Select 'Drive API' and click Add.");
  }

  var tempName = 'temp_eppo_import_' + new Date().getTime();

  // รองรับทั้ง Drive API v3 และ Drive API v2
  try {
    // Drive API v3
    var resource = {
      name: tempName,
      mimeType: MimeType.GOOGLE_SHEETS
    };
    var file = Drive.Files.create(resource, blob);
    return file.id;
  } catch (errV3) {
    try {
      // Fallback: Drive API v2
      var fileMetadata = {
        title: tempName,
        mimeType: MimeType.GOOGLE_SHEETS
      };
      var fileV2 = Drive.Files.insert(fileMetadata, blob, { convert: true });
      return fileV2.id;
    } catch (errV2) {
      throw new Error("Drive API failed to convert .xls file: " + errV2.message);
    }
  }
}

/**
 * ลบไฟล์ชั่วคราวออกจาก Google Drive ทันทีหลังอ่านเสร็จ
 */
function deleteTempFile(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    try {
      if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.remove) {
        Drive.Files.remove(fileId);
      }
    } catch (e2) {
      console.warn("Could not remove temp file " + fileId + ": " + e2.message);
    }
  }
}

/**
 * --------------------------------------------------------------------------
 * 6. Business Logic: คลีนชื่อผลิตภัณฑ์ และตรวจจับ Header คอลัมน์
 * --------------------------------------------------------------------------
 */
function cleanProductName(rawName) {
  var cleaned = String(rawName).replace(/\u00a0/g, ' ').trim();

  // แมปชื่อมาตรฐานภาษาอังกฤษตรงๆ
  for (var key in PRODUCT_MAPPING) {
    if (cleaned.indexOf(key) !== -1) {
      return PRODUCT_MAPPING[key];
    }
  }

  // กลุ่มน้ำมันพรีเมียม (ตรวจสอบอย่างรัดกุม)
  if (cleaned.indexOf('พรีเมียม') !== -1 || cleaned.indexOf('Premium') !== -1) {
    if (cleaned.indexOf('ดีเซล') !== -1 || cleaned.indexOf('HSD') !== -1) {
      return 'HSD Premium';
    }
    if (cleaned.indexOf('เบนซิน') !== -1 || cleaned.indexOf('ULG') !== -1) {
      return 'ULG 95 Premium';
    }
    if (cleaned.indexOf('99') !== -1 || cleaned.indexOf('Gasohol 99') !== -1) {
      return 'Gasohol 99 Premium';
    }
    if (cleaned.indexOf('แก๊สโซฮอล์ 95') !== -1 || cleaned.indexOf('แก๊สโซฮอล ออกเทน 95') !== -1 || cleaned.indexOf('Gasohol 95') !== -1 || cleaned.indexOf('แก๊สโซฮอล์') !== -1 || cleaned.indexOf('แก๊สโซฮอล') !== -1 || cleaned.indexOf('95') !== -1) {
      return 'Gasohol 95 Premium';
    }
  }

  var match = cleaned.match(/\((.*?)\)/);
  if (match && match[1]) {
    return match[1].trim();
  }
  return cleaned;
}

function detectCompanyColumns(grid) {
  var colToCompany = {};
  var maxRow = Math.min(12, grid.length);
  var numCols = grid[0].length;

  for (var col = 1; col < numCols; col++) {
    var textParts = [];
    for (var r = Math.min(8, maxRow); r < maxRow; r++) {
      var cell = grid[r][col];
      if (cell !== undefined && cell !== null && cell !== '') {
        textParts.push(String(cell));
      }
    }
    var combinedText = textParts.join(' ').replace(/\s+/g, ' ').toLowerCase();

    var matchedCompany = null;
    for (var i = 0; i < DETECTION_COMPANIES.length; i++) {
      var comp = DETECTION_COMPANIES[i];
      if (comp === 'PT') {
        // ใช้ Word boundary สำหรับ PT ป้องกันการชนกับคำอื่น
        if (/\bpt\b/i.test(combinedText)) {
          matchedCompany = 'PT';
          break;
        }
      } else if (combinedText.indexOf(comp.toLowerCase()) !== -1) {
        matchedCompany = comp;
        break;
      }
    }

    if (matchedCompany) {
      colToCompany[col] = matchedCompany;
    }
  }

  // Fallback กรณีหาหัวเรื่องไม่เจอ
  if (Object.keys(colToCompany).length < 5) {
    for (var j = 0; j < COMPANY_ORDER.length; j++) {
      if (j + 1 < numCols) {
        colToCompany[j + 1] = COMPANY_ORDER[j];
      }
    }
  }

  return colToCompany;
}

function parseGridToWideRows(grid, dateIso) {
  var colToCompany = detectCompanyColumns(grid);

  var startIdx = -1;
  var endIdx = -1;

  for (var r = 0; r < grid.length; r++) {
    var val = String(grid[r][0] || '');
    if (
      val.indexOf('Gasohol') !== -1 || val.indexOf('ULG') !== -1 || val.indexOf('HSD') !== -1 ||
      val.indexOf('ดีเซล') !== -1 || val.indexOf('เบนซิน') !== -1 || val.indexOf('แก๊สโซฮอล์') !== -1
    ) {
      if (startIdx === -1) startIdx = r;
    }
    if (val.indexOf('(Effective Date)') !== -1 || val.indexOf('มีผลตั้งแต่') !== -1 || val.indexOf('Effective Date') !== -1) {
      endIdx = r;
      break;
    }
  }

  if (startIdx === -1) startIdx = 12;
  if (endIdx === -1) endIdx = Math.min(startIdx + 12, grid.length);

  var outputRows = [];

  for (var idx = startIdx; idx < endIdx; idx++) {
    if (idx >= grid.length) break;
    var rawProduct = String(grid[idx][0] || '').trim();
    if (!rawProduct || rawProduct === 'nan' || rawProduct.indexOf('Effective Date') !== -1 || rawProduct.indexOf('Unit : Baht') !== -1) {
      continue;
    }

    var product = cleanProductName(rawProduct);
    if (!product || product === 'Effective Date' || product === 'Unit : Baht/Litre') {
      continue;
    }

    // แถวผลลัพธ์: [Date, Product, PTT, BCP, Shell, Esso, Chevron, IRPC, PT, Susco, Pure, SUSCO Dealers]
    var row = [dateIso, product];
    var compPriceMap = {};

    for (var c in colToCompany) {
      var compName = colToCompany[c];
      var rawPrice = grid[idx][c];
      if (rawPrice !== undefined && rawPrice !== null && String(rawPrice).trim() !== '' && String(rawPrice).trim() !== '-') {
        var priceNum = parseFloat(rawPrice);
        compPriceMap[compName] = (!isNaN(priceNum) && priceNum > 0) ? priceNum : '';
      } else {
        compPriceMap[compName] = '';
      }
    }

    for (var k = 0; k < COMPANY_ORDER.length; k++) {
      var company = COMPANY_ORDER[k];
      row.push(compPriceMap[company] !== undefined ? compPriceMap[company] : '');
    }

    outputRows.push(row);
  }

  return { rows: outputRows };
}

/**
 * --------------------------------------------------------------------------
 * 7. ระบบตั้งเวลาอัตโนมัติ (Trigger Setup)
 * --------------------------------------------------------------------------
 * รันฟังก์ชันนี้ 1 ครั้ง เพื่อสร้าง Trigger รันทุกวันเวลา 08:00 และ 16:00 น.
 */
function setupDailyTriggers() {
  removeDailyTriggers();

  // รอบเช้า: 08:00 - 09:00 น.
  ScriptApp.newTrigger('dailySyncOilPrice')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Bangkok')
    .create();

  // รอบเย็น: 16:00 - 17:00 น.
  ScriptApp.newTrigger('dailySyncOilPrice')
    .timeBased()
    .atHour(16)
    .everyDays(1)
    .inTimezone('Asia/Bangkok')
    .create();

  console.info("✅ Daily triggers set successfully for 08:00 and 16:00 (Asia/Bangkok).");
}

/**
 * ลบ Trigger เดิมออกทั้งหมดเพื่อป้องกัน Trigger ซ้ำซ้อน
 */
function removeDailyTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailySyncOilPrice') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  console.info("Cleared existing daily triggers.");
}

/**
 * --------------------------------------------------------------------------
 * 8. ฟังก์ชันเสริมสำหรับทดสอบรันด้วยตัวเอง (Manual Testing)
 * --------------------------------------------------------------------------
 */
function testRunToday() {
  console.info("Testing run for today...");
  var res = dailySyncOilPrice();
  console.info("Result: " + res);
}

function sendLineNotify(token, message) {
  try {
    UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: { 'message': message },
      muteHttpExceptions: true
    });
  } catch (e) {
    console.warn("Line Notify error: " + e.message);
  }
}
