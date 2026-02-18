function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Manpower Readiness Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- 1. MANPOWER DATA (Master Source for Store Info) ---
  const qSheet = ss.getSheetByName('Manpower_Status');
  let targets = [];
  let storeMap = {}; // Lookup object: StoreNo -> {Area, Format, OD, etc}

  if (qSheet) {
    const qData = qSheet.getDataRange().getDisplayValues();
    let h = qData[0];
    
    // Default indices based on your CSV structure
    let cStore=0, cName=1, cFmt=2, cArea=3, cOD=4, cFocus=5, cDept=6, cLevel=7, cTgt=8, cActive=9;

    // specific header mapping to be safe
    for(let i=0; i<h.length; i++) {
      let s = String(h[i]).toLowerCase().trim().replace(/ /g, '_'); 
      if(s.includes('store_no')) cStore=i;
      if(s.includes('name')) cName=i;
      if(s.includes('format')) cFmt=i;
      if(s.includes('area')) cArea=i;
      if(s.includes('od')) cOD=i;
      if(s.includes('focus')) cFocus=i; 
      if(s.includes('department')) cDept=i;
      if(s.includes('target')) cTgt=i;
      if(s === 'active' || s.includes('active')) cActive=i; 
    }

    for (let i = 1; i < qData.length; i++) {
      let r = qData[i];
      if (!r[cStore]) continue;
      
      let sNo = String(r[cStore]).replace(/\.0$/, "").trim();
      let storeName = r[cName];
      let fmt = String(r[cFmt]).trim();
      let area = String(r[cArea]).trim();
      let od = String(r[cOD]).trim();
      let focus = String(r[cFocus]).trim();

      // SAVE STORE INFO TO MAP (This fixes the missing data in other sheets)
      if (!storeMap[sNo]) {
        storeMap[sNo] = {
            name: storeName,
            format: fmt,
            area: area,
            od: od,
            focus: focus
        };
      }
      
      let tgtVal = parseFloat(r[cTgt]);
      let actVal = parseFloat(r[cActive]);
      let t = isNaN(tgtVal) ? 0 : tgtVal;
      let a = isNaN(actVal) ? 0 : actVal;

      targets.push({
        storeNo: sNo,
        storeName: storeName,
        format: fmt,
        area: area,
        od: od,
        storeFocus: focus,
        dept: r[cDept],
        target: t,
        active: a,
        gap: t - a
      });
    }
  }

  // --- 2. APPLICANT TRACKING (Fixed Missing Columns) ---
  const appSheet = ss.getSheetByName('Applicant_Tracking');
  let applicants = [];
  
  if (appSheet) {
    const aData = appSheet.getDataRange().getDisplayValues();
    let h = aData[0];

    // Corrected Indices for Applicant_Tracking.csv
    let cApply=0, cStore=1, cName=2, cCand=3, cPhone=4, cPos=5, cComp=6, cDept=7, cSource=8, cInterview=9, cStatus=10, cStartDate=11, cStartStatus=12;

    for(let i=0; i<h.length; i++) {
      let s = String(h[i]).toLowerCase().trim().replace(/ /g, '_');
      if(s.includes('store_no')) cStore = i;
      if(s.includes('department')) cDept = i;
      if(s.includes('position')) cPos = i;
      if(s.includes('interview_date')) cInterview = i;
      if(s.includes('start_date')) cStartDate = i; 
      if(s.includes('apply_date') || s === 'app_date') cApply = i;
      if(s.includes('interview_status')) cStatus = i; 
      if(s.includes('start_status')) cStartStatus = i; 
      if(s.includes('source')) cSource = i;
    }

    for (let i = 1; i < aData.length; i++) {
      let r = aData[i];
      let sNo = String(r[cStore]).replace(/\.0$/, "").trim();
      
      // LOOKUP MISSING INFO from Manpower Map
      let meta = storeMap[sNo] || { area: 'N/A', format: 'N/A', name: r[cName] || 'N/A' };

      applicants.push({
        applyDate: r[cApply],
        storeNo: sNo,
        area: meta.area,       // Injected from map
        format: meta.format,   // Injected from map
        dept: String(r[cDept]).trim(),
        position: String(r[cPos]).trim(),
        interviewDate: r[cInterview],
        status: String(r[cStatus]),
        startDate: r[cStartDate], 
        startStatus: String(r[cStartStatus]),
        source: String(r[cSource]).trim()
      });
    }
  }

  // --- 3. STAFF MOVEMENT ---
  const sSheet = ss.getSheetByName('Staff_Movement');
  let staff = [];
  
  if (sSheet) {
    const sData = sSheet.getDataRange().getDisplayValues();
    let h = sData[0];
    
    // Corrected Indices for Staff_Movement.csv
    let cStore=1, cDept=8, cJoin=9, cTerm=10;
    
    for(let i=0; i<h.length; i++) {
      let s = String(h[i]).toLowerCase().trim().replace(/ /g, '_');
      if(s.includes('join_date')) cJoin = i;
      if(s.includes('terminate_date')) cTerm = i;
      if(s === 'store_no' || s.includes('store_no')) cStore = i;
      if(s.includes('department')) cDept = i;
    }

    for (let i = 1; i < sData.length; i++) {
      staff.push({ 
        joinDate: sData[i][cJoin], 
        termDate: sData[i][cTerm],
        storeNo: String(sData[i][cStore]).replace(/\.0$/, "").trim(),
        dept: String(sData[i][cDept]).trim()
      });
    }
  }

  // --- 4. PART-TIME ---
  const pSheet = ss.getSheetByName('Parttime');
  let parttime = [];
  
  if (pSheet) {
    const pData = pSheet.getDataRange().getDisplayValues();
    let h = pData[0];
    
    let cDate=0, cStore=1, cDept=3, cTemp=4, cStudent=5, cElderly=6, cTotal=7;

    for(let i=0; i<h.length; i++) {
      let s = String(h[i]).toLowerCase().trim().replace(/ /g, '_');
      if(s.includes('date')) cDate = i;
      if(s.includes('store_no')) cStore = i;
      if(s.includes('department')) cDept = i;
      if(s.includes('temp')) cTemp = i;
      if(s.includes('student')) cStudent = i;
      if(s.includes('elderly')) cElderly = i;
      if(s.includes('total')) cTotal = i;
    }

    for (let i = 1; i < pData.length; i++) {
      let r = pData[i];
      if (!r[cDate] && !r[cStore]) continue;

      parttime.push({
        date: r[cDate],
        storeNo: String(r[cStore]).replace(/\.0$/, "").trim(),
        dept: String(r[cDept]).trim(),
        temp: parseFloat(r[cTemp]) || 0,
        student: parseFloat(r[cStudent]) || 0,
        elderly: parseFloat(r[cElderly]) || 0,
        total: parseFloat(r[cTotal]) || 0
      });
    }
  }

  return { targets: targets, applicants: applicants, staff: staff, parttime: parttime, error: false };
}
