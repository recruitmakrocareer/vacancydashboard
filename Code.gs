/**
 * =========================================================
 * 1. ฟังก์ชันเปิดหน้าเว็บ (ผสานทั้ง HTML เดิม และ API ของ React)
 * =========================================================
 */
function doGet(e) {
  // 🟢 หากเป็นการเรียกจาก React (Index เสริม) เพื่อโหลดข้อมูลตอนเปิดเว็บ
  if (e && e.parameter && e.parameter.api === 'react_init') {
    return getReactInitData();
  }
  
  // 🔵 หากเป็นการเปิดหน้าเว็บปกติ (Index หลัก)
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Makro-Operations Recruitment Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * =========================================================
 * 2. ฟังก์ชันอรรถประโยชน์ (Utility)
 * =========================================================
 */
function createHeaderMap(headers) {
  const map = {};
  if (!headers) return map;
  headers.forEach((h, i) => {
    const key = String(h).toLowerCase().replace(/[\s_.]/g, ''); 
    map[key] = i;
  });
  return map;
}

/**
 * =========================================================
 * 3. ฟังก์ชันฝั่งระบบหลัก (Main) เดิมของคุณ (ห้ามลบ)
 * =========================================================
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = { targets: [], applicants: [], staff: [], parttime: [], error: false };
  
  try {
    // 1. MANPOWER
    const qSheet = ss.getSheetByName('Manpower_Status');
    let storeMap = {}; 
    if (qSheet) {
      const data = qSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);
      if (map['storeno'] !== undefined) {
        data.forEach(r => {
          if (!r[map['storeno']]) return;
          let sNo = String(r[map['storeno']]).trim();
          storeMap[sNo] = { 
            name: r[map['storename']] || r[map['name']] || '', 
            format: r[map['format']] || '', 
            area: r[map['area']] || '', 
            od: r[map['od']] || '', 
            focus: r[map['storefocus']] || r[map['focus']] || '' 
          };
          let t = parseInt(r[map['target']] || 0, 10);
          let a = parseInt(r[map['active']] || 0, 10);
          result.targets.push({
            storeNo: sNo,
            storeName: storeMap[sNo].name,
            format: storeMap[sNo].format,
            area: storeMap[sNo].area,
            od: storeMap[sNo].od,
            storeFocus: storeMap[sNo].focus,
            dept: r[map['department']] || r[map['dept']] || '',
            level: r[map['level']] || '',
            target: isNaN(t) ? 0 : t,
            active: isNaN(a) ? 0 : a
          });
        });
      }
    }

    // 2. APPLICANTS
    const appSheet = ss.getSheetByName('Applicant_Tracking');
    if (appSheet) {
      const data = appSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);
      const kApply = map['applydate'] !== undefined ? map['applydate'] : map['appdate']; 
      if (map['storeno'] !== undefined) {
        data.forEach(r => {
          let sNo = String(r[map['storeno']]).trim();
          let meta = storeMap[sNo] || { area: '', format: '' }; 
          result.applicants.push({
            applyDate: r[kApply] || '',
            storeNo: sNo,
            area: meta.area,
            format: meta.format,
            dept: r[map['department']] || '',
            position: r[map['position']] || '',
            level: r[map['level']] || '',
            interviewDate: r[map['interviewdate']] || '',
            status: r[map['interviewstatus']] || '',
            startDate: r[map['startdate']] || '',
            startStatus: r[map['startstatus']] || '',
            source: r[map['applysource']] || r[map['source']] || ''
          });
        });
      }
    }

    // 3. STAFF MOVEMENT
    const sSheet = ss.getSheetByName('Staff_Movement');
    if (sSheet) {
      const data = sSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);
      if (map['storeno'] !== undefined) {
        data.forEach(r => {
           result.staff.push({
             storeNo: String(r[map['storeno']]).trim(),
             dept: r[map['department']] || '',
             level: r[map['level']] || '',
             joinDate: r[map['joindate']] || '',
             termDate: r[map['terminatedate']] || ''
           });
        });
      }
    }

    // 4. PARTTIME
    const pSheet = ss.getSheetByName('Parttime');
    if (pSheet) {
      const data = pSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);
      const kTotal = map['totalparttime'] !== undefined ? map['totalparttime'] : map['total']; 
      if (map['date'] !== undefined && map['storeno'] !== undefined) {
        data.forEach(r => {
           let ptLevel = r[map['level']];
           if (!ptLevel || ptLevel === '') ptLevel = 'Temporary'; 
           result.parttime.push({
             date: r[map['date']],
             storeNo: String(r[map['storeno']]).trim(),
             dept: r[map['department']] || '',
             level: ptLevel,
             temp: parseFloat(r[map['findtemp']] || 0) || 0,
             student: parseFloat(r[map['studentparttime']] || r[map['student']] || 0) || 0,
             elderly: parseFloat(r[map['elderly']] || 0) || 0,
             total: parseFloat(r[kTotal] || 0) || 0
           });
        });
      }
    }
  } catch (e) {
    result.error = true;
    result.message = e.toString();
    Logger.log(e);
  }
  return result;
}

function autoSendDailyReportToOutlook() {
  try {
    var spreadsheetId = '1Wn1gnstzG_2Wi_Tc95cw0rIC1dIoHJvHbghlxceJy9A'; 
    var webhookUrl = "https://4d5e96e05cb6e591aefe0bce82117f.8e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/02f76ccb4ab449dab6c1c7ca9ce3dd0a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ltg01x8lgMKLxv3JIYQL_z5IGtWsLmuUgORuAjojbeY";
    var emailAddress = "ttanawat@cpaxtra.co.th"; 
    
    var file = DriveApp.getFileById(spreadsheetId);
    var ss = SpreadsheetApp.openById(spreadsheetId);
    SpreadsheetApp.flush(); 
    
    // --- 1. ส่วนสร้างตาราง HTML ---
    var top10Sheet = ss.getSheetByName('Top 10 of Vacant (30 Stores)');
    var rawData = top10Sheet.getDataRange().getDisplayValues(); 
    
    var htmlTable = "<table style='border-collapse: collapse; width: 80%; font-family: Tahoma, sans-serif; font-size: 14px; border: 1px solid #ddd;' cellpadding='8'>";
    for (var i = 0; i < rawData.length; i++) {
        var row = rawData[i];
        if (row.join("").trim() === "") continue; 
        
        htmlTable += "<tr>";
        for (var j = 0; j < row.length; j++) {
            if (i === 0) { 
                htmlTable += "<th style='background-color: #ED1C24; color: white; border: 1px solid #ddd; text-align: center;'>" + row[j] + "</th>";
            } else { 
               var align = (j === 0) ? "left" : "center";
               htmlTable += "<td style='border: 1px solid #ddd; text-align: " + align + ";'>" + row[j] + "</td>";
            }
        }
        htmlTable += "</tr>";
    }
    htmlTable += "</table>";

    // --- 2. ส่วนคัดแยกเฉพาะ Sheet ที่ต้องการ ---
    var sheetsToExport = ['Top 10 of Vacant (30 Stores)', 'Breakdown by Position (30)', 'Breakdown by Position (All)', 'Top 10 of Vacant (All Store)']; 
    var tempSs = SpreadsheetApp.create("Temp_Export_Report");
    
    for (var s = 0; s < sheetsToExport.length; s++) {
      var originalSheet = ss.getSheetByName(sheetsToExport[s]);
      if (originalSheet) {
        var originalRange = originalSheet.getDataRange();
        var originalValues = originalRange.getValues(); 
        var copiedSheet = originalSheet.copyTo(tempSs).setName(sheetsToExport[s]);
        copiedSheet.clearContents();
        copiedSheet.getRange(1, 1, originalValues.length, originalValues[0].length).setValues(originalValues);
      }
    }
    
    var defaultSheet = tempSs.getSheetByName('แผ่นงาน 1') || tempSs.getSheetByName('Sheet1');
    if (defaultSheet) { tempSs.deleteSheet(defaultSheet); }
    
    // --- 3. สั่ง Export ไฟล์ชั่วคราวนี้แทน ---
    var token = ScriptApp.getOAuthToken();
    var exportUrl = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx&access_token=" + token;
    var response = UrlFetchApp.fetch(exportUrl, { muteHttpExceptions: true });
    
    if(response.getResponseCode() !== 200) {
      Logger.log("❌ โหลดไฟล์ Excel ไม่สำเร็จ: " + response.getContentText());
      DriveApp.getFileById(tempSs.getId()).setTrashed(true);
      return;
    }

    var fileBase64 = Utilities.base64Encode(response.getBlob().getBytes());
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    var fileName = dateStr + "_Top_10_Store_Vacant.xlsx";
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    
    // --- 4. ส่งไป Power Automate ---
    var payload = {
      "fileName": fileName,
      "fileContent": fileBase64,
      "emailTo": emailAddress,
      "htmlTable": htmlTable
    };
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    };
    
    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("✅ ส่งข้อมูลสำเร็จ (คัดเฉพาะ Sheet ที่ระบุ)!");
    
  } catch (error) {
    Logger.log("❌ Error: " + error.message);
  }
}

/**
 * =========================================================
 * 4. ฟังก์ชันฝั่งระบบเสริม React Pro (เพิ่มใหม่)
 * =========================================================
 */

// โหลดข้อมูลสำหรับ React ตอนเปิดหน้าเว็บ
function getReactInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = { applicants: [], vacancies: [], branches: [], tas: [], positions: [] };
  
  // โหลดผู้สมัคร
  const appSheet = ss.getSheetByName('Applicant_Tracking');
  if (appSheet) {
    const data = appSheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    const map = createHeaderMap(headers);
    
    data.forEach(r => {
      if (r[map['id']]) { // ดึงเฉพาะที่มี ID (มาจากระบบ Pro)
        let intStatus = r[map['interview_status']] || r[map['interviewstatus']];
        let startStatus = r[map['start_status']] || r[map['startstatus']];
        let finalStatus = (intStatus === 'ผ่านสัมภาษณ์' && startStatus === 'รอเริ่มงาน') ? 'รอเริ่มงาน' : intStatus;
        
        result.applicants.push({
          id: r[map['id']],
          branchCode: String(r[map['store_no']] || r[map['storeno']]),
          branchName: r[map['store_name']] || r[map['storename']],
          name: r[map['candidate_name']] || r[map['candidatename']],
          phone: r[map['phone_no']] || r[map['phoneno']],
          position: r[map['position']],
          status: finalStatus,
          startDate: r[map['start_date']] || r[map['startdate']],
          comment: r[map['remark']],
          ta: r[map['ta']],
          resumeName: r[map['resumename']],
          resumeBase64: r[map['resumebase64']],
          timestamp: r[map['timestamp']],
          updatedAt: r[map['updatedat']] || r[map['updated_at']]
        });
      }
    });
  }
  
  // โหลด Vacancies
  let vacSheet = ss.getSheetByName('Vacancies');
  if (!vacSheet) {
    // ถ้าไม่มีชีต Vacancies ให้สร้างอัตโนมัติ
    vacSheet = ss.insertSheet('Vacancies');
    vacSheet.appendRow(['id', 'area', 'branchCode', 'branchName', 'position', 'manualStatus', 'effectiveDate', 'createdAt', 'timestamp']);
  } else {
    const data = vacSheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    const map = createHeaderMap(headers);
    data.forEach(r => {
      if(r[map['id']]) {
        result.vacancies.push({
           id: r[map['id']], area: r[map['area']], branchCode: r[map['branchcode']],
           branchName: r[map['branchname']], position: r[map['position']],
           manualStatus: r[map['manualstatus']], effectiveDate: r[map['effectivedate']],
           createdAt: r[map['createdat']], timestamp: r[map['timestamp']]
        });
      }
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// 🟢 ฟังก์ชัน doPost ตัวหลักสำหรับรับข้อมูล
// (รวมมิตรสำหรับรับ Data จากทั้งระบบเก่าและระบบ React)
function doPost(e) {
  try {
    if (!e.postData) throw new Error("No POST Data");
    const body = JSON.parse(e.postData.contents);
    
    // --- คำสั่งจากระบบ React (Pro) ---
    if (body.action === 'saveAll') return handleReactSaveAll(body);
    if (body.action === 'deleteRecord') return deleteReactRecord(body);
    if (body.action === 'updateMaster') return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
    if (body.action === 'notifyNewVacancy' || body.action === 'sendLineNotify') {
      return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- คำสั่งจากระบบ Main (Vanilla JS) ---
    // (หมายเหตุ: ถ้าใน Project คุณมีไฟล์ Code.gs อื่นที่มี doPost ของระบบหลักอยู่แล้ว 
    // ให้คัดลอกโค้ด if (body.action === '...') ด้านบน ไปใส่ในไฟล์ doPost ตัวเดิมของคุณแทนนะครับ)
    if (body.action && body.payload) {
      // ส่วนนี้เว้นไว้ให้ระบบหลักทำงาน หากมีฟังก์ชันแยก
      // เช่น if(body.action === 'searchFT') return searchFT(body.payload);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// จัดการบันทึกข้อมูลผู้สมัครและตำแหน่งงานว่าง
function handleReactSaveAll(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet = ss.getSheetByName('Applicant_Tracking');
  const vacSheet = ss.getSheetByName('Vacancies');
  
  // อัปเดต Applicant
  if (payload.applicants && payload.applicants.length > 0 && appSheet) {
    const data = appSheet.getDataRange().getValues();
    const headers = data[0];
    const map = createHeaderMap(headers);
    
    payload.applicants.forEach(app => {
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][map['id']] == app.id) { rowIndex = i + 1; break; }
      }
      
      // 🟢 Logic ตัวแปลงสถานะ (Data Mapping)
      let intStatus = app.status;
      let startStatus = '';
      if (app.status === 'รอเริ่มงาน') {
         intStatus = 'ผ่านสัมภาษณ์';
         startStatus = 'รอเริ่มงาน';
      } else if (['ผ่านสัมภาษณ์', 'มาเริ่มงาน', 'ไม่มาเริ่มงาน'].includes(app.status)) {
         intStatus = 'ผ่านสัมภาษณ์';
         startStatus = app.status === 'ผ่านสัมภาษณ์' ? '' : app.status;
      }

      let appDateStr = app.timestamp ? String(app.timestamp).split(' ')[0] : Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
      
      if (rowIndex > -1) {
        // อัปเดตข้อมูลเก่า
        appSheet.getRange(rowIndex, (map['interview_status'] || map['interviewstatus']) + 1).setValue(intStatus);
        appSheet.getRange(rowIndex, (map['start_status'] || map['startstatus']) + 1).setValue(startStatus);
        appSheet.getRange(rowIndex, (map['start_date'] || map['startdate']) + 1).setValue(app.startDate || '');
        appSheet.getRange(rowIndex, map['remark'] + 1).setValue(app.comment || '');
        appSheet.getRange(rowIndex, (map['updatedat'] || map['updated_at']) + 1).setValue(app.updatedAt);
        
        if (app.resumeBase64 && app.resumeBase64 !== 'Uploaded' && !String(app.resumeBase64).startsWith('Error')) {
           appSheet.getRange(rowIndex, map['resumename'] + 1).setValue(app.resumeName);
           appSheet.getRange(rowIndex, map['resumebase64'] + 1).setValue(app.resumeBase64);
        }
      } else {
        // เพิ่มข้อมูลใหม่
        const newRow = new Array(headers.length).fill('');
        newRow[map['app_date'] || map['appdate']] = appDateStr;
        newRow[map['store_no'] || map['storeno']] = app.branchCode;
        newRow[map['store_name'] || map['storename']] = app.branchName;
        newRow[map['candidate_name'] || map['candidatename']] = app.name;
        newRow[map['phone_no'] || map['phoneno']] = app.phone;
        newRow[map['position']] = app.position;
        newRow[map['company']] = 'ADC'; 
        newRow[map['department']] = 'Store';
        newRow[map['apply_source'] || map['applysource']] = 'Walk-in at Store'; // Default
        newRow[map['interview_date'] || map['interviewdate']] = appDateStr;
        newRow[map['interview_status'] || map['interviewstatus']] = intStatus;
        newRow[map['start_date'] || map['startdate']] = app.startDate || '';
        newRow[map['start_status'] || map['startstatus']] = startStatus;
        newRow[map['remark']] = app.comment || '';
        newRow[map['level']] = 'Staff';
        newRow[map['id']] = app.id;
        newRow[map['ta']] = app.ta;
        newRow[map['resumename']] = app.resumeName || '';
        newRow[map['resumebase64']] = app.resumeBase64 || '';
        newRow[map['timestamp']] = app.timestamp;
        newRow[map['updatedat'] || map['updated_at']] = app.updatedAt;
        
        appSheet.appendRow(newRow);
      }
    });
  }
  
  // อัปเดต Vacancies (ทับของเก่าด้วยของใหม่ทั้งหมด)
  if (payload.vacancies && vacSheet) {
     vacSheet.clearContents();
     const headers = ['id', 'area', 'branchCode', 'branchName', 'position', 'manualStatus', 'effectiveDate', 'createdAt', 'timestamp'];
     vacSheet.appendRow(headers);
     payload.vacancies.forEach(v => {
        vacSheet.appendRow([v.id, v.area, v.branchCode, v.branchName, v.position, v.manualStatus, v.effectiveDate, v.createdAt, v.timestamp]);
     });
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
}

// ลบข้อมูล (Delete)
function deleteReactRecord(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = payload.sheetName === 'Applicants' ? 'Applicant_Tracking' : 'Vacancies';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error'})).setMimeType(ContentService.MimeType.JSON);
  
  const data = sheet.getDataRange().getValues();
  const map = createHeaderMap(data[0]);
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][map['id']] == payload.recordId) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
}
