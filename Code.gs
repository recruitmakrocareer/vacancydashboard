/**
 * Serves the HTML file.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Manpower Readiness Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createHeaderMap(headers) {
  const map = {};
  if (!headers) return map;
  headers.forEach((h, i) => {
    const key = String(h).toLowerCase().replace(/[\s_.]/g, ''); 
    map[key] = i;
  });
  return map;
}

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
            level: r[map['level']] || '', // CAPTURE LEVEL
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
            level: r[map['level']] || '', // CAPTURE LEVEL
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
             level: r[map['level']] || '', // CAPTURE LEVEL
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
           if (!ptLevel || ptLevel === '') ptLevel = 'Temporary'; // Default
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

// --- 2. ส่วนคัดแยกเฉพาะ Sheet ที่ต้องการ (เทคนิค Temp File แบบเคลียร์สูตร 100%) ---
    var sheetsToExport = ['Top 10 of Vacant (30 Stores)', 'Breakdown by Position (30)', 'Breakdown by Position (All)', 'Top 10 of Vacant (All Store)']; 
    
    // สร้างไฟล์ชั่วคราว
    var tempSs = SpreadsheetApp.create("Temp_Export_Report");
    
    // วนลูปก๊อปปี้เฉพาะชีตที่ระบุ ไปใส่ในไฟล์ชั่วคราว
    for (var s = 0; s < sheetsToExport.length; s++) {
      var originalSheet = ss.getSheetByName(sheetsToExport[s]);
      
      if (originalSheet) {
        // 2.1 ดึงข้อมูล "ตัวเลข/ข้อความจริงๆ" จากชีตต้นฉบับเตรียมไว้ก่อน
        var originalRange = originalSheet.getDataRange();
        var originalValues = originalRange.getValues(); 
        
        // 2.2 ก๊อปปี้ชีตไปไฟล์ใหม่ (เพื่อเอาหน้าตา สี เส้นตาราง)
        var copiedSheet = originalSheet.copyTo(tempSs).setName(sheetsToExport[s]);
        
        // 🌟 2.3 [หัวใจสำคัญ] สั่งลบข้อมูลและสูตรที่พังทิ้งทั้งหมด! (แต่สีและ Format จะยังอยู่)
        copiedSheet.clearContents();
        
        // 2.4 นำตัวเลข/ข้อความเพียวๆ ที่ดึงไว้ในข้อ 2.1 ไปหยอดกลับคืน
        copiedSheet.getRange(1, 1, originalValues.length, originalValues[0].length).setValues(originalValues);
      }
    }
    
    // ลบหน้า "แผ่นงาน1" (หรือ Sheet1) ที่แถมมาตอนสร้างไฟล์ใหม่ทิ้ง
    var defaultSheet = tempSs.getSheetByName('แผ่นงาน 1') || tempSs.getSheetByName('Sheet1');
    if (defaultSheet) { tempSs.deleteSheet(defaultSheet); }
    // -------------------------------------------------------------------------
    
    // --- 3. สั่ง Export ไฟล์ชั่วคราวนี้แทน ---
    var token = ScriptApp.getOAuthToken();
    var exportUrl = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx&access_token=" + token;
    var response = UrlFetchApp.fetch(exportUrl, { muteHttpExceptions: true });
    
    if(response.getResponseCode() !== 200) {
      Logger.log("❌ โหลดไฟล์ Excel ไม่สำเร็จ: " + response.getContentText());
      // สั่งลบไฟล์ชั่วคราวทิ้งเผื่อกรณี Error
      DriveApp.getFileById(tempSs.getId()).setTrashed(true);
      return;
    }

    var fileBase64 = Utilities.base64Encode(response.getBlob().getBytes());
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    var fileName = dateStr + "_Top_10_Store_Vacant.xlsx";
    
    // 🌟 ส่งไฟล์เสร็จแล้ว ลบไฟล์ชั่วคราวลงถังขยะทันที (สำคัญมาก ไดรฟ์จะได้ไม่เต็ม)
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
