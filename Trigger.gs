// ==========================================
// ⚙️ 1. GLOBAL CONFIGURATION (ตั้งค่าที่นี่)
// ==========================================
const CONFIG = {
  // 🔴 ใส่ ID ของโฟลเดอร์ Google Drive ที่เก็บไฟล์ .xlsx/.csv
  SOURCE_FOLDER_ID: '1Xi1dDoD3fR5Q6AiOg89mZIihGIIQBQTb', 
  
  // 🔴 ใส่ ID ของ Spreadsheet นี้ (ดูจาก URL)
  TARGET_SPREADSHEET_ID: '1Wn1gnstzG_2Wi_Tc95cw0rIC1dIoHJvHbghlxceJy9A',

  // ชื่อชีตต่างๆ (ห้ามเปลี่ยนถ้าไม่จำเป็น)
  SHEET_NAME: {
    RAW: 'sheet1',           // ชีตพักข้อมูลดิบ
    CONFIG: 'Config_Mapping', // ชีตตั้งค่า Mapping
    DEST: 'Staff_Movement',   // ชีตผลลัพธ์ Step 1
    MANPOWER: 'Manpower_Status' // 🟢 เพิ่มชื่อชีต Manpower ตรงนี้
  },

  // ตัวแปรจำค่าเวลาอัปเดต (ไม่ต้องแก้)
  PROP_KEY_LAST_UPDATE: 'LAST_XLSX_UPDATED_MS'
};

// ==========================================
// 🚀 2. PHASE 1: SYNC DATA (Main Trigger)
// ==========================================
function runPhase1_Sync() {
  const lock = LockService.getScriptLock();
  // ป้องกันการรันซ้อนกัน (รอคิว 30 วินาที)
  if (!lock.tryLock(30000)) return;

  try {
    // 2.1 หาไฟล์ล่าสุด (รองรับ Drive API v2)
    let info = getLatestXlsxInfo_(CONFIG.SOURCE_FOLDER_ID);
    if (!info) {
      console.log("ไม่พบไฟล์ในโฟลเดอร์");
      return;
    }

    const props = PropertiesService.getScriptProperties();
    const lastMs = Number(props.getProperty(CONFIG.PROP_KEY_LAST_UPDATE) || 0);

    // 2.2 ถ้าไฟล์ยังเป็นตัวเดิม ให้ข้าม (ไม่ต้องทำซ้ำ)
    if (info.updatedMs <= lastMs) {
      console.log(`ℹ️ ไฟล์ ${info.name} เป็นไฟล์เดิมที่เคยทำไปแล้ว (Skip)`);
      return;
    }

    console.log(`🚀 [Phase 1] พบไฟล์ใหม่: ${info.name} กำลังดึงข้อมูล...`);

    // 2.3 ดึงข้อมูลลง sheet1
    syncFileToSheet1_(info.id);

    // 2.4 บันทึกว่าทำเสร็จแล้ว และตั้งเวลาปลุก Phase 2
    props.setProperty(CONFIG.PROP_KEY_LAST_UPDATE, String(info.updatedMs));

    // สร้าง Trigger ให้รัน Phase 2 ในอีก 10 วินาทีข้างหน้า
    ScriptApp.newTrigger('runPhase2_Process')
      .timeBased()
      .after(1000 * 10) 
      .create();

    console.log(`✅ [Phase 1] Sync ข้อมูลดิบเสร็จสิ้น -> ส่งไม้ต่อให้ Phase 2`);

  } catch (err) {
    console.error('❌ Error Phase 1: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 🚀 3. PHASE 2: PROCESS DATA (Auto Run)
// ==========================================
function runPhase2_Process() {
  console.log(`🚀 [Phase 2] เริ่มประมวลผลข้อมูล...`);
  
  try {
    // 3.1 ลบ Trigger ตัวเองทิ้ง (ใช้แล้วทิ้ง ไม่ให้รก)
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === 'runPhase2_Process') {
        ScriptApp.deleteTrigger(t);
      }
    }

    // 3.2 รัน Logic จัดการ Staff Movement
    importStaffMovement2(); 

    // 🟢 3.3 อัปเดต Manpower Status (Active Count แบบ Snapshot)
    updateManpowerActiveCount();

    console.log(`✅ [Phase 2] เสร็จสมบูรณ์ นำเข้า Staff Movement และอัปเดต Manpower เรียบร้อย!`);

  } catch (err) {
    console.error('❌ Error Phase 2: ' + err.toString());
  }
}

// ==========================================
// 🛠️ 4. LOGIC STEP 1: IMPORT & DATE FIX
// ==========================================
function importStaffMovement2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- อ่าน Config ---
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAME.CONFIG);
  if (!configSheet) { console.error('❌ ไม่พบ Config_Mapping'); return; }
  
  const configData = configSheet.getDataRange().getDisplayValues();
  let positionMap = {};
  for (let i = 1; i < configData.length; i++) {
    let row = configData[i];
    if (row[0]) {
      positionMap[String(row[0]).trim()] = {
        empCategory: String(row[1]).trim(), jobFam: String(row[2]).trim(),
        level: String(row[3]).trim(), dept: String(row[4]).trim()
      };
    }
  }

  // --- เตรียมชีตปลายทาง ---
  const destSheetName = CONFIG.SHEET_NAME.DEST;
  let destSheet = ss.getSheetByName(destSheetName);
  
  const HEADERS = [
    "Emp_ID", "Store_No", "Name", "Company", "Position",
    "Emp_Category", "Job_Family", "Level", "Department",
    "Join_Date", "Terminate_Date", "HR_Status", "Resign_Reason"
  ];

  if (!destSheet) destSheet = ss.insertSheet(destSheetName);
  else destSheet.clearContents(); 

  // --- อ่านข้อมูลดิบจาก sheet1 ---
  let sourceSheet = ss.getSheetByName(CONFIG.SHEET_NAME.RAW);
  if (!sourceSheet) { console.error('❌ ไม่พบ sheet1'); return; }
  
  const data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) return;

  // หา Index Column
  let h = data[0].map(x => String(x).toLowerCase().trim());
  let col = {
    empID: -1, name: -1, status: -1, joinDate: -1, pos: -1, store: -1, comp: -1, terminate: -1,
    reason: 26 
  };
  
  for(let i=0; i<h.length; i++) {
    if(h[i].includes("empl id") || h[i].includes("รหัสพนักงาน") || h[i].includes("emp_id")) col.empID = i;
    else if(h[i].includes("display name") || h[i].includes("ชื่อ")) col.name = i;
    else if(h[i].includes("hr status") || h[i].includes("สถานะ")) col.status = i;
    else if(h[i].includes("join date") || h[i].includes("วันที่เข้างาน")) col.joinDate = i;
    else if(h[i].includes("position") || h[i].includes("ตำแหน่ง")) col.pos = i;
    else if(h[i].includes("st no") || h[i].includes("รหัสสาขา")) col.store = i;
    else if(h[i].includes("company") || h[i].includes("บริษัท")) col.comp = i;
    else if(h[i].includes("termination") || h[i].includes("วันลาออก")) col.terminate = i;
    
    if (i > 0 && h[i-1].includes("reason code") && h[i].includes("descr")) {
      col.reason = i;
    }
  }

  const parseDate = (d) => {
    if (!d) return "";
    if (d instanceof Date) return d;
    const p = String(d).split('/');
    if (p.length === 3) return new Date(p[2], p[0]-1, p[1]);
    return d;
  };

  let finalRows = [];
  let processed = new Set();

  for (let r = 1; r < data.length; r++) {
    let row = data[r];
    let empID = String(row[col.empID] || "").trim();
    
    if (!empID || processed.has(empID)) continue;

    let position = String(row[col.pos]).trim();
    let mapData = positionMap[position];

    if (mapData) {
      processed.add(empID);
      let storeRaw = row[col.store];
      let storeClean = (storeRaw !== "" && storeRaw !== null) ? String(parseInt(storeRaw, 10)) : "";

      let reasonText = "";
      if (col.reason >= 0 && col.reason < row.length) {
        reasonText = String(row[col.reason] || "").trim();
      }

      finalRows.push([
        empID, storeClean, row[col.name], row[col.comp], position,
        mapData.empCategory, mapData.jobFam, mapData.level, mapData.dept,
        parseDate(row[col.joinDate]),
        parseDate(row[col.terminate]),
        row[col.status],
        reasonText
      ]);
    }
  }

  if (finalRows.length > 0) {
    destSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold").setBackground("#d9d9d9");
    destSheet.getRange(2, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
    let lastRow = finalRows.length + 1;
    destSheet.getRange(2, 1, lastRow - 1, HEADERS.length).setNumberFormat("@"); 
    destSheet.getRange(2, 10, lastRow - 1, 2).setNumberFormat("dd/mm/yyyy"); 
  }
  console.log(`✅ Import Staff Movement: ${finalRows.length} รายการ`);
}

// ==========================================
// 🟢 5. LOGIC STEP 2: อัปเดตยอด Active ใน Manpower Status (Snapshot)
// ==========================================
function updateManpowerActiveCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 5.1 อ่านข้อมูล Staff Movement เพื่อหักยอด
  const staffSheet = ss.getSheetByName(CONFIG.SHEET_NAME.DEST);
  if (!staffSheet) return;
  const staffData = staffSheet.getDataRange().getValues();
  if (staffData.length < 2) return;

  const stH = staffData[0].map(h => String(h).trim().toLowerCase());
  const cStore = stH.indexOf("store_no");
  const cDept = stH.indexOf("department");
  const cLevel = stH.indexOf("level");
  const cStatus = stH.indexOf("hr_status");

  // นับจำนวนพนักงานที่เป็น Active
  let activeCounts = {};
  for (let i = 1; i < staffData.length; i++) {
    let status = String(staffData[i][cStatus]).trim().toLowerCase();
    if (status === 'active') {
      let key = String(staffData[i][cStore]).trim() + "|" + 
                String(staffData[i][cDept]).trim().toUpperCase() + "|" + 
                String(staffData[i][cLevel]).trim().toUpperCase();
      activeCounts[key] = (activeCounts[key] || 0) + 1;
    }
  }

  // 5.2 อ่านข้อมูล Manpower Status
  const mpSheet = ss.getSheetByName(CONFIG.SHEET_NAME.MANPOWER);
  if (!mpSheet) { console.error('❌ ไม่พบชีต Manpower_Status'); return; }
  let mpData = mpSheet.getDataRange().getValues();
  if (mpData.length < 2) return;

  const mpH = mpData[0].map(h => String(h).trim().toLowerCase());
  const mStore = mpH.indexOf("store_no");
  const mDept = mpH.indexOf("department");
  const mLevel = mpH.indexOf("level");
  const mActive = mpH.indexOf("active"); // Column K (ดัชนี 10)
  const mDate = mpH.indexOf("effective_date"); // Column L (ดัชนี 11)

  if (mStore === -1 || mDept === -1 || mLevel === -1 || mActive === -1 || mDate === -1) {
    console.error('❌ คอลัมน์ใน Manpower_Status ไม่ครบ (ต้องมี Store_No, Department, Level, Active, Effective_Date)');
    return;
  }

  // ตัวแปลงวันที่ให้กลายเป็นตัวเลข (ไว้เทียบว่าอันไหนใหม่สุด)
  const parseDateToMs = (d) => {
    if (!d) return 0;
    if (d instanceof Date) return d.getTime();
    let s = String(d).split(' ')[0]; 
    let p = s.split('/');
    if (p.length === 3) return new Date(p[2], p[1]-1, p[0]).getTime(); // แปลงจาก DD/MM/YYYY
    return 0;
  };

  // 5.3 หาบรรทัดที่เป็นข้อมูลล่าสุด (Latest Snapshot) ของแต่ละ Store/Dept/Level
  let latestRowMap = {}; // เก็บค่าว่าแต่ละ Key แถวไหนใหม่สุด
  for (let i = 1; i < mpData.length; i++) {
    let key = String(mpData[i][mStore]).trim() + "|" + 
              String(mpData[i][mDept]).trim().toUpperCase() + "|" + 
              String(mpData[i][mLevel]).trim().toUpperCase();
              
    let ms = parseDateToMs(mpData[i][mDate]);
    
    // ถ้ายังไม่มี หรือวันที่ใหม่กว่า/เท่ากับอันเดิม ให้ยึดบรรทัดนี้เป็นล่าสุด
    if (!latestRowMap[key] || ms >= latestRowMap[key].maxMs) {
      latestRowMap[key] = { rowIndex: i, maxMs: ms };
    }
  }

  // 5.4 อัปเดต Column Active ลงใน "เฉพาะบรรทัดล่าสุด"
  for (let key in latestRowMap) {
    let r = latestRowMap[key].rowIndex;
    let newCount = activeCounts[key] || 0; // ถ้าไม่เจอใน Staff Movement แสดงว่าเหลือ 0
    mpData[r][mActive] = newCount;
  }

  // 5.5 เขียนข้อมูลทับกลับลงไป (รักษารูปแบบเดิมไว้)
  mpSheet.getRange(1, 1, mpData.length, mpData[0].length).setValues(mpData);
  console.log(`✅ อัปเดตยอด Active ลง Manpower_Status สำเร็จ! (เฉพาะข้อมูล Snapshot ล่าสุด)`);
}

// ==========================================
// 📂 HELPER FUNCTIONS (Drive API v2 Compatible)
// ==========================================
function getLatestXlsxInfo_(folderId) {
  const q = `'${folderId}' in parents and trashed = false and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'text/csv' or title contains '.xlsx' or title contains '.csv')`;
  const resp = Drive.Files.list({ q: q, fields: "items(id,title,modifiedDate)", orderBy: "modifiedDate desc", maxResults: 1, supportsAllDrives: true, includeItemsFromAllDrives: true });
  const files = resp.items || [];
  if (files.length === 0) return null;
  const f = files[0];
  return { id: f.id, name: f.title, updatedMs: new Date(f.modifiedDate).getTime() };
}

function syncFileToSheet1_(fileId) {
  const resource = { title: 'TEMP_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
  const converted = Drive.Files.copy(resource, fileId, { supportsAllDrives: true });
  try {
    const tempSs = SpreadsheetApp.openById(converted.id);
    const tempSheet = tempSs.getSheets()[0];
    const values = tempSheet.getDataRange().getValues();
    const targetSs = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
    let targetSheet = targetSs.getSheetByName(CONFIG.SHEET_NAME.RAW);
    if (!targetSheet) targetSheet = targetSs.insertSheet(CONFIG.SHEET_NAME.RAW);
    targetSheet.clear();
    if (values.length > 0) targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  } finally {
    Drive.Files.remove(converted.id);
  }
}