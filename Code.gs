/**
 * Serves the HTML file.
 *  Main dashboard = index
 *  Turnover page  = Turnover
 *
 *  🌟 Updates (v2):
 *  1. รองรับคอลัมน์ "Subregion" (เดิมชื่อ Area) ของ Manpower_Status — มี fallback
 *     ในกรณีชีตเก่ายังใช้ชื่อ Area อยู่ จะใช้ค่านั้นต่อแทน Subregion
 *  2. Performance: อ่านเฉพาะคอลัมน์ที่จำเป็นจาก Applicant_Tracking
 *     (ตัด resumeBase64, resumeName, base64 อื่นๆ ออก) ลดขนาด payload ลงอย่างมาก
 *  3. ลด JSON payload ทำให้ Cache เก็บได้ครอบคลุมขึ้น
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'turnover') {
    return HtmlService.createTemplateFromFile('Turnover')
      .evaluate()
      .setTitle('Turnover Analytics Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (e && e.parameter && e.parameter.page === 'aec') {
    return HtmlService.createTemplateFromFile('AEC')
      .evaluate()
      .setTitle('AEC / Foreign Worker Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('HR Analytics Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function createHeaderMap(headers) {
  const map = {};
  if (!headers) return map;
  headers.forEach((h, i) => {
    if (h === null || h === undefined) return;
    const key = String(h).toLowerCase().replace(/[\s_.]/g, '');
    map[key] = i;
  });
  return map;
}

/**
 * 🌟 ดึงข้อมูล Dashboard ทั้งหมด (ปรับปรุงประสิทธิภาพ + รองรับ Subregion)
 */
function getDashboardData(forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'HR_DASHBOARD_DATA_V4'; // v4: เพิ่ม nationality field ใน staff

  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      Logger.log("✅ Load data from Cache (Fast)");
      return cachedData;
    }
  }

  Logger.log("⏳ Load data from Spreadsheet (First time or Cache expired)");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = { targets: [], applicants: [], staff: [], parttime: [], error: false };

  try {
    // ============================================================
    // 1. MANPOWER_STATUS (รองรับทั้ง Subregion และ Area เพื่อ backward compat)
    // ============================================================
    const qSheet = ss.getSheetByName('Manpower_Status');
    let storeMap = {};
    if (qSheet) {
      const data = qSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);

      const keyEffective = map['effectivedate'] !== undefined ? map['effectivedate'] : map['date'];
      // 🌟 ใช้ Subregion ถ้ามี ถ้าไม่มี fallback ไป Area
      const keySubregion = map['subregion'] !== undefined ? map['subregion'] : map['area'];
      // 🌟 v3: Regional CEO column (ใหม่)
      const keyRegionalCEO = map['regionalceo'];

      if (map['storeno'] !== undefined) {
        data.forEach(r => {
          if (!r[map['storeno']]) return;
          let sNo = String(r[map['storeno']]).trim();
          storeMap[sNo] = {
            name: r[map['storename']] || r[map['name']] || '',
            format: r[map['format']] || '',
            subregion: keySubregion !== undefined ? (r[keySubregion] || '') : '',
            od: r[map['od']] || '',
            regionalCEO: keyRegionalCEO !== undefined ? (r[keyRegionalCEO] || '') : '', // 🌟
            focus: r[map['storefocus']] || r[map['focus']] || ''
          };

          let t = parseInt(r[map['target']] || 0, 10);
          let a = parseInt(r[map['active']] || 0, 10);

          result.targets.push({
            storeNo: sNo,
            storeName: storeMap[sNo].name,
            format: storeMap[sNo].format,
            subregion: storeMap[sNo].subregion,
            od: storeMap[sNo].od,
            regionalCEO: storeMap[sNo].regionalCEO, // 🌟
            storeFocus: storeMap[sNo].focus,
            dept: r[map['department']] || r[map['dept']] || '',
            level: r[map['level']] || '',
            target: isNaN(t) ? 0 : t,
            active: isNaN(a) ? 0 : a,
            effectiveDate: keyEffective !== undefined ? r[keyEffective] : ''
          });
        });
      }
    }

    // ============================================================
    // 2. APPLICANT_TRACKING (อ่านเฉพาะคอลัมน์ที่ใช้ - ลด payload)
    // ============================================================
    const appSheet = ss.getSheetByName('Applicant_Tracking');
    if (appSheet) {
      const data = appSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);
      const kApply = map['applydate'] !== undefined ? map['applydate'] : map['appdate'];
      if (map['storeno'] !== undefined) {
        data.forEach(r => {
          let sNo = String(r[map['storeno']]).trim();
          if (!sNo) return; // ข้ามแถวว่าง
          let meta = storeMap[sNo] || { subregion: '', format: '' };
          // 🌟 ดึงเฉพาะ field ที่ใช้จริง (ตัด resumeBase64 ฯลฯ ออก)
          result.applicants.push({
            applyDate: r[kApply] || '',
            storeNo: sNo,
            subregion: meta.subregion, // 🌟
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

    // ============================================================
    // 3. STAFF_MOVEMENT
    // ============================================================
    const sSheet = ss.getSheetByName('Staff_Movement');
    if (sSheet) {
      const data = sSheet.getDataRange().getDisplayValues();
      const headers = data.shift();
      const map = createHeaderMap(headers);
      if (map['storeno'] !== undefined) {
        data.forEach(r => {
           result.staff.push({
             empId: r[map['employeeid']] || r[map['empid']] || r[map['id']] || '',
             storeNo: String(r[map['storeno']]).trim(),
             dept: r[map['department']] || r[map['dept']] || '',
             level: r[map['level']] || '',
             joinDate: r[map['joindate']] || '',
             termDate: r[map['terminatedate']] || r[map['termdate']] || '',
             status: r[map['status']] || r[map['employeestatus']] || r[map['hrstatus']] || '',
             resignReason: r[map['resignreason']] || r[map['resign_reason']] || '',
             nationality: r[map['nationality']] || ''
           });
        });
      }
    }

    // ============================================================
    // 4. PARTTIME
    // ============================================================
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

    try {
      const payloadString = JSON.stringify(result);
      // GAS Cache รองรับ ~100KB แต่ลดเหลือ ~95KB เพื่อความปลอดภัย
      if (payloadString.length < 95000) {
        cache.put(cacheKey, payloadString, 900); // 15 นาที
        Logger.log("✅ Cached payload size: " + payloadString.length + " bytes");
      } else {
        Logger.log("⚠️ Data > 95KB limit (" + payloadString.length + " bytes). Skipping cache.");
      }
    } catch(err) { Logger.log("⚠️ Cache error: " + err.message); }

    return JSON.stringify(result);

  } catch (e) {
    result.error = true;
    result.message = e.toString();
    Logger.log("❌ getDashboardData Error: " + e.toString());
    return JSON.stringify(result);
  }
}

/**
 * 🌟 ฟังก์ชันสำหรับล้าง cache แบบบังคับ (ใช้กรณีต้องการดูข้อมูลล่าสุดทันที)
 */
function clearDashboardCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('HR_DASHBOARD_DATA_V3');
  cache.remove('HR_DASHBOARD_DATA_V4');
  return "Cache cleared. Next call will reload from spreadsheet.";
}
