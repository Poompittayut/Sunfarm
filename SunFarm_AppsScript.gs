const SS_ID = '1ZFyaOQvaKFOqiy-F9BiQt4lf-_0OHc-wzE_sj0BuIhY';

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var result;

  if (action === 'dashboard') {
    result = getDashboardData();
  } else {
    result = { error: 'unknown action' };
  }

  var output = ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return output;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheetName = data.sheet;
    if (!sheetName) return ok('no sheet');

    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return ok('sheet not found: ' + sheetName);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    delete data.sheet;

    data['Timestamp'] = new Date().toISOString();
    var row = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    sheet.appendRow(row);
    return ok('saved');
  } catch(err) {
    return ok('error: ' + err.message);
  }
}

function ok(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDashboardData() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

  function sheetToObjects(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var headers = vals[0];
    return vals.slice(1).filter(function(r) { return r[0] !== ''; }).map(function(r) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = r[i]; });
      return obj;
    });
  }

  var eggs   = sheetToObjects('01_ไข่รายวัน');
  var hatch  = sheetToObjects('02_ผลฟักไข่');
  var birds  = sheetToObjects('03_ทะเบียนไก่');
  var mating = sheetToObjects('04_แผนผสมพันธุ์');
  var select = sheetToObjects('05_คัดเลือก');

  // KPI
  var totalEggs  = eggs.reduce(function(s,r){ return s + (Number(r['Eggs_Collected'])||0); }, 0);
  var totalUsable= eggs.reduce(function(s,r){ return s + (Number(r['Eggs_Usable'])||0); }, 0);
  var hSet  = hatch.reduce(function(s,r){ return s + (Number(r['Eggs_Set'])||0); }, 0);
  var hFert = hatch.reduce(function(s,r){ return s + (Number(r['Fertile_Eggs'])||0); }, 0);
  var hChk  = hatch.reduce(function(s,r){ return s + (Number(r['Chicks_Hatched'])||0); }, 0);
  var hDIS  = hatch.reduce(function(s,r){ return s + (Number(r['Dead_In_Shell'])||0); }, 0);

  var families = new Set(hatch.map(function(r){ return r['Family_Code']; }).filter(Boolean)).size;
  var activeB  = birds.filter(function(r){ return r['Status']==='Active'; }).length;
  var lines    = new Set(hatch.map(function(r){ return r['Line_Name']; }).filter(Boolean)).size;

  // eggs_by_date
  var ebd = {};
  eggs.forEach(function(r) {
    var d = String(r['Record_Date']).trim();
    if (!d) return;
    // normalize d-m-yyyy → yyyy-mm-dd
    var parts = d.split('-');
    if (parts.length === 3 && parts[0].length <= 2) {
      d = parts[2] + '-' + ('0'+parts[1]).slice(-2) + '-' + ('0'+parts[0]).slice(-2);
    }
    ebd[d] = (ebd[d]||0) + (Number(r['Eggs_Collected'])||0);
  });

  // by_line
  var bl = {};
  hatch.forEach(function(r) {
    var l = r['Line_Name'] || 'Unknown';
    if (!bl[l]) bl[l] = { eggs:0, fertile:0, chicks:0, dis:0 };
    bl[l].eggs   += Number(r['Eggs_Set'])||0;
    bl[l].fertile+= Number(r['Fertile_Eggs'])||0;
    bl[l].chicks += Number(r['Chicks_Hatched'])||0;
    bl[l].dis    += Number(r['Dead_In_Shell'])||0;
  });

  return {
    updated: now,
    kpi: {
      total_eggs:     totalEggs,
      total_usable:   totalUsable,
      total_chicks:   hChk,
      se_pct:  hSet>0  ? Math.round(hFert/hSet*1000)/10 : 0,
      sf_pct:  hFert>0 ? Math.round(hChk/hFert*1000)/10 : 0,
      dis_pct: hFert>0 ? Math.round(hDIS/hFert*1000)/10 : 0,
      total_families: families,
      active_birds:   activeB,
      total_lines:    lines
    },
    raw: { eggs:eggs, hatch:hatch, birds:birds, mating:mating, select:select },
    eggs_by_date: ebd,
    by_line: bl
  };
}
