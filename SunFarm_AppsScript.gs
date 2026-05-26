// ═══════════════════════════════════════════════════════════════════════════
// SunFarm Breeding OS — Google Apps Script
// ═══════════════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1ZFyaOQvaKFOqiy-F9BiQt4lf-_0OHc-wzE_sj0BuIhY';

const SHEET_COLUMNS = {
  '01_ไข่รายวัน': [
    'Record_Date','Mating_ID','Family_Code','Line_Name',
    'Eggs_Collected','Dirty_Eggs','Cracked_Eggs','Eggs_Usable',
    'Notes','Recorded_By','Timestamp'
  ],
  '02_ผลฟักไข่': [
    'Hatch_Lot_ID','Mating_ID','Family_Code','Line_Name',
    'Set_Date','Hatch_Date','Eggs_Set','Infertile','Fertile_Eggs',
    'Chicks_Hatched','Dead_In_Shell','Culled',
    'Fertility_pct','SF_pct','DIS_pct',
    'Notes','Hatchery_Staff','Timestamp'
  ],
  '03_ทะเบียนไก่': [
    'Bird_ID','Hatch_Lot_ID','Family_Code','Line_Name','Generation',
    'Sex','Hatch_Date','Ring_Color','Ring_Number','Pen_Location',
    'Status','Notes','Timestamp'
  ],
  '04_แผนผสมพันธุ์': [
    'Mating_ID','Line_Name','Family_Code','Sire_ID','Dam_IDs',
    'Sire_Count','Dam_Count','Mating_Type',
    'Start_Date','End_Date','Pen_Location','Status','Notes','Timestamp'
  ],
  '05_คัดเลือก': [
    'Bird_ID','Family_Code','Sex',
    'Weight_Score','Conformation_Score','Health_Score','Production_Score',
    'Final_Score','Decision','Notes','Selected_By','Selection_Date','Timestamp'
  ]
};

// ── CORS Helper ───────────────────────────────────────────────────────────────
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helper: Sheet rows → array of objects ────────────────────────────────────
function sheetToObjects(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── GET: ดึงข้อมูลทุก sheet สำหรับ Dashboard ────────────────────────────────
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;

    // ทดสอบการเชื่อมต่อ
    if (action === 'test' || !action) {
      return corsOutput({ status: 'ok', message: 'SunFarm API ready' });
    }

    // ดึงข้อมูลทั้งหมดสำหรับ Dashboard
    if (action === 'dashboard') {
      const eggs    = sheetToObjects('01_ไข่รายวัน');
      const hatch   = sheetToObjects('02_ผลฟักไข่');
      const birds   = sheetToObjects('03_ทะเบียนไก่');
      const mating  = sheetToObjects('04_แผนผสมพันธุ์');
      const select  = sheetToObjects('05_คัดเลือก');

      // สรุป KPI
      const totalEggs    = eggs.reduce((s, r) => s + (Number(r.Eggs_Collected) || 0), 0);
      const totalChicks  = hatch.reduce((s, r) => s + (Number(r.Chicks_Hatched) || 0), 0);
      const totalEggsSet = hatch.reduce((s, r) => s + (Number(r.Eggs_Set) || 0), 0);
      const totalFertile = hatch.reduce((s, r) => s + (Number(r.Fertile_Eggs) || 0), 0);
      const totalDIS     = hatch.reduce((s, r) => s + (Number(r.Dead_In_Shell) || 0), 0);
      const activeBirds  = birds.filter(r => r.Status === 'Active').length;
      const selectedBirds= birds.filter(r => r.Status === 'Selected').length;

      // %SE, %SF, %DIS รวม
      const se_pct = totalEggsSet > 0 ? (totalFertile / totalEggsSet * 100).toFixed(1) : 0;
      const sf_pct = totalFertile > 0 ? (totalChicks / totalFertile * 100).toFixed(1) : 0;
      const dis_pct= totalFertile > 0 ? (totalDIS / totalFertile * 100).toFixed(1) : 0;

      // แยกตาม Family สำหรับ chart
      const byFamily = {};
      hatch.forEach(r => {
        const fam = r.Family_Code || r.Mating_ID || 'Unknown';
        if (!byFamily[fam]) byFamily[fam] = { eggs: 0, chicks: 0, fertile: 0, dis: 0, count: 0 };
        byFamily[fam].eggs    += Number(r.Eggs_Set) || 0;
        byFamily[fam].chicks  += Number(r.Chicks_Hatched) || 0;
        byFamily[fam].fertile += Number(r.Fertile_Eggs) || 0;
        byFamily[fam].dis     += Number(r.Dead_In_Shell) || 0;
        byFamily[fam].count++;
      });

      // แยกตาม Line สำหรับ chart เปรียบเทียบ
      const byLine = {};
      hatch.forEach(r => {
        const line = r.Line_Name || 'Unknown';
        if (!byLine[line]) byLine[line] = { eggs: 0, chicks: 0, fertile: 0, dis: 0 };
        byLine[line].eggs    += Number(r.Eggs_Set) || 0;
        byLine[line].chicks  += Number(r.Chicks_Hatched) || 0;
        byLine[line].fertile += Number(r.Fertile_Eggs) || 0;
        byLine[line].dis     += Number(r.Dead_In_Shell) || 0;
      });

      // ไข่รายวัน แยกตามวัน (30 วันล่าสุด)
      const eggsByDate = {};
      eggs.forEach(r => {
        const d = r.Record_Date ? String(r.Record_Date).split('T')[0] : '';
        if (!d) return;
        if (!eggsByDate[d]) eggsByDate[d] = 0;
        eggsByDate[d] += Number(r.Eggs_Collected) || 0;
      });

      // Unique lines และ families
      const lines    = [...new Set(hatch.map(r => r.Line_Name).filter(Boolean))];
      const families = [...new Set(hatch.map(r => r.Family_Code).filter(Boolean))];

      return corsOutput({
        status: 'ok',
        updated: new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
        kpi: {
          total_eggs: totalEggs,
          total_chicks: totalChicks,
          total_eggs_set: totalEggsSet,
          active_birds: activeBirds,
          selected_birds: selectedBirds,
          total_families: families.length,
          total_lines: lines.length,
          se_pct, sf_pct, dis_pct
        },
        by_family: byFamily,
        by_line: byLine,
        eggs_by_date: eggsByDate,
        raw: {
          eggs: eggs.slice(-100),    // 100 rows ล่าสุด
          hatch: hatch.slice(-100),
          birds: birds.slice(-200),
          mating,
          select: select.slice(-100)
        }
      });
    }

    return corsOutput({ status: 'error', message: 'Unknown action: ' + action });

  } catch(err) {
    return corsOutput({ status: 'error', message: err.message });
  }
}

// ── POST: รับข้อมูลจาก HTML Form ─────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const cols = SHEET_COLUMNS[data.sheet];
    if (!cols) throw new Error('ไม่พบ sheet: ' + data.sheet);
    const sheet = getOrCreateSheet(data.sheet);
    data.Timestamp = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
    const row = cols.map(col => data[col] !== undefined ? data[col] : '');
    sheet.appendRow(row);
    return corsOutput({ status: 'success', sheet: data.sheet, rows: sheet.getLastRow() });
  } catch (err) {
    return corsOutput({ status: 'error', message: err.message });
  }
}

// ── HELPER: หา Sheet หรือสร้างใหม่ ──────────────────────────────────────────
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const cols = SHEET_COLUMNS[sheetName];
    if (cols) {
      const headerRange = sheet.getRange(1, 1, 1, cols.length);
      headerRange.setValues([cols]);
      headerRange.setBackground('#7F0000');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}
