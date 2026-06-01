// ================================================================
//  Code.gs  –  COD Response Monitoring Dashboard
//  Google Apps Script Backend
//
//  Sheet: history
//    REF ID | month | date_reminder | last_update | id | name |
//    hub | amount | notes | case_status | rek_pnp | lead | ass_lead |
//    case_close_reason | tanggal_janji_bayar
//
//  Sheet: notes_history
//    ref_id | timestamp | notes | status_changed_to | updated_by |
//    case_close_reason
// ================================================================

var HIST_SHEET  = 'history';
var NOTES_SHEET = 'notes_history';

var HIST_HEADERS = [
  'REF ID','month','date_reminder','last_update','id','name',
  'hub','amount','notes','case_status','rek_pnp','lead','ass_lead',
  'case_close_reason'
];
var NOTES_HEADERS = [
  'ref_id','timestamp','notes','status_changed_to','updated_by','case_close_reason'
];

// ── Serve Dashboard HTML ──────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('dashboard')
    .setTitle('COD Response Monitoring')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Auto-Create Sheets If Missing ────────────────────────────
function setupSheets() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName(HIST_SHEET);
  if (!hist) {
    hist = ss.insertSheet(HIST_SHEET);
    hist.appendRow(HIST_HEADERS);
    hist.getRange(1,1,1,HIST_HEADERS.length)
        .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    hist.setFrozenRows(1);
    hist.setColumnWidth(1,200); hist.setColumnWidth(6,160); hist.setColumnWidth(9,300);
    Logger.log('Sheet "'+HIST_SHEET+'" created.');
  }
  var notes = ss.getSheetByName(NOTES_SHEET);
  if (!notes) {
    notes = ss.insertSheet(NOTES_SHEET);
    notes.appendRow(NOTES_HEADERS);
    notes.getRange(1,1,1,NOTES_HEADERS.length).setFontWeight('bold').setBackground('#e8eaed');
    notes.setFrozenRows(1);
    Logger.log('Sheet "'+NOTES_SHEET+'" created.');
  }
}

// ── Utility ──────────────────────────────────────────────────
function pa(v) {
  if (!v && v !== 0) return 0;
  return parseFloat(String(v).replace(/,/g,'')) || 0;
}
function fd(d) {
  if (!d) return '';
  var dt = d instanceof Date ? d : new Date(d);
  return isNaN(dt.getTime()) ? String(d)
    : Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd MMM yyyy');
}
function fdt(d) {
  if (!d) return '';
  var dt = d instanceof Date ? d : new Date(d);
  return isNaN(dt.getTime()) ? String(d)
    : Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}
function colMap(headers) {
  var m = {};
  headers.forEach(function(h,i){ m[String(h).trim()] = i; });
  return m;
}

// Hitung DPD (Days Past Due) dari tanggal janji bayar ke hari ini
function calcDPD(dateVal) {
  if (!dateVal) return null;
  var dt = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(dt.getTime())) return null;
  var today = new Date(); today.setHours(0,0,0,0);
  dt = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  return Math.round((today - dt) / (1000*60*60*24));
}

// Cari kolom secara case-insensitive
function findColIdx(headers, name) {
  var trimmed = String(name).trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === trimmed) return i;
  }
  return -1;
}

function getHistSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HIST_SHEET);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(HIST_SHEET); }
  return sh;
}
function curMonth() {
  var n = new Date();
  return ['January','February','March','April','May','June','July',
          'August','September','October','November','December'][n.getMonth()]
          +' '+n.getFullYear();
}
function passFilters(row, COL, f, fStart, fEnd) {
  if (!row[COL['REF ID']]) return false;
  if (f.month   && String(row[COL['month']]      ||'').trim()!==f.month)   return false;
  if (f.lead    && String(row[COL['lead']]       ||'').trim()!==f.lead)    return false;
  if (f.assLead && String(row[COL['ass_lead']]   ||'').trim()!==f.assLead) return false;
  if (f.status  && String(row[COL['case_status']]||'').trim()!==f.status)  return false;
  if (fStart||fEnd) {
    var dr=row[COL['date_reminder']];
    var d=dr instanceof Date?dr:new Date(dr);
    if (!isNaN(d.getTime())) {
      if (fStart&&d<fStart) return false;
      if (fEnd  &&d>fEnd)   return false;
    }
  }
  return true;
}

// =================================================================
//  getFilterOptions()
// =================================================================
function getFilterOptions() {
  setupSheets();
  var sh=getHistSheet(), data=sh.getDataRange().getValues();
  if (data.length<=1) return {months:[],leads:[],allAssLeads:[],statuses:[],leadToAssLeads:{},currentMonth:curMonth()};
  var COL=colMap(data[0]);
  var months={},leads={},allAL={},statuses={},lMap={};
  for (var i=1;i<data.length;i++) {
    var row=data[i];
    if (!row[COL['REF ID']]) continue;
    var m=String(row[COL['month']]      ||'').trim();
    var l=String(row[COL['lead']]       ||'').trim();
    var al=String(row[COL['ass_lead']]  ||'').trim();
    var s=String(row[COL['case_status']]||'').trim();
    if(m) months[m]=true; if(l) leads[l]=true;
    if(al) allAL[al]=true; if(s) statuses[s]=true;
    if(l&&al){if(!lMap[l])lMap[l]={};lMap[l][al]=true;}
  }
  var l2al={};
  Object.keys(lMap).forEach(function(l){l2al[l]=Object.keys(lMap[l]).sort();});
  return {months:Object.keys(months).sort(),leads:Object.keys(leads).sort(),
          allAssLeads:Object.keys(allAL).sort(),statuses:Object.keys(statuses).sort(),
          leadToAssLeads:l2al,currentMonth:curMonth()};
}

// =================================================================
//  getKPIData(filters)
// =================================================================
function getKPIData(filters) {
  setupSheets();
  var sh=getHistSheet(),data=sh.getDataRange().getValues();
  if (data.length<=1) return {byLead:[],totals:_emptyTotals()};
  var COL=colMap(data[0]),f=filters||{};
  var fS=f.start?new Date(f.start):null, fE=f.end?new Date(f.end):null;
  if (fE) fE.setHours(23,59,59);
  var byLead={};
  for (var i=1;i<data.length;i++) {
    var row=data[i];
    if (!passFilters(row,COL,f,fS,fE)) continue;
    var al=String(row[COL['ass_lead']]||'').trim(), lg=String(row[COL['lead']]||'').trim();
    var amt=pa(row[COL['amount']]), s=String(row[COL['case_status']]||'').trim(), rek=pa(row[COL['rek_pnp']]);
    if (!byLead[al]) byLead[al]={lead:al,leadGroup:lg,
      reminderCase:0,respondedCase:0,notRespondedCase:0,
      reminderAmt:0,respondedAmt:0,notRespondedAmt:0,
      recoveryAmt:0,outstandingAmt:0,rekPnpOutstanding:0,
      onInvCount:0,closeCount:0,onInvAmt:0,closeAmt:0};
    var b=byLead[al];
    b.reminderCase++;b.reminderAmt+=amt;
    if(s!=='Open'){b.respondedCase++;b.respondedAmt+=amt;}
    else{b.notRespondedCase++;b.notRespondedAmt+=amt;}
    if(s==='Case Close'){b.recoveryAmt+=amt;b.closeCount++;b.closeAmt+=amt;}
    if(s==='On Investigation'){b.outstandingAmt+=amt;b.rekPnpOutstanding+=rek;b.onInvCount++;b.onInvAmt+=amt;}
  }
  var arr=Object.keys(byLead).map(function(k){
    var b=byLead[k];
    b.pctResponse     =b.reminderCase>0?b.respondedCase/b.reminderCase:0;
    b.pctNotRespAmt   =b.reminderAmt>0?b.notRespondedAmt/b.reminderAmt:0;
    b.finalOutstanding=b.outstandingAmt-b.rekPnpOutstanding;
    b.finalRecovery   =b.recoveryAmt+b.rekPnpOutstanding;
    b.pctRespRecovery =b.respondedAmt>0?b.finalRecovery/b.respondedAmt:0;
    return b;
  });
  arr.sort(function(a,b){return b.reminderCase-a.reminderCase;});
  var T=_emptyTotals();
  arr.forEach(function(b){for(var k in T)if(k in b)T[k]+=b[k];});
  T.pctResponse    =T.reminderCase>0?T.respondedCase/T.reminderCase:0;
  T.pctNotRespAmt  =T.reminderAmt>0?T.notRespondedAmt/T.reminderAmt:0;
  T.pctRespRecovery=T.respondedAmt>0?T.finalRecovery/T.respondedAmt:0;
  return {byLead:arr,totals:T};
}
function _emptyTotals() {
  return {reminderCase:0,respondedCase:0,notRespondedCase:0,
    reminderAmt:0,respondedAmt:0,notRespondedAmt:0,
    recoveryAmt:0,outstandingAmt:0,rekPnpOutstanding:0,
    finalOutstanding:0,finalRecovery:0,onInvCount:0,closeCount:0,
    onInvAmt:0,closeAmt:0,pctResponse:0,pctNotRespAmt:0,pctRespRecovery:0};
}

// =================================================================
//  getCases(filters)
//  [FIX] last_update & tanggal_janji_bayar kini pakai findColIdx
//  (case-insensitive) agar tidak gagal karena mismatch header
// =================================================================
function getCases(filters) {
  setupSheets();
  var sh      = getHistSheet();
  var data    = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var COL     = colMap(headers);
  var f       = filters || {};

  var fS = f.start ? new Date(f.start) : null;
  var fE = f.end   ? new Date(f.end)   : null;
  if (fE) fE.setHours(23, 59, 59);

  // [FIX] Gunakan findColIdx agar lookup tahan terhadap beda casing
  //       atau karakter tersembunyi pada header sheet
  var iLastUpd  = findColIdx(headers, 'last_update');
  var iJanjiByr = findColIdx(headers, 'tanggal_janji_bayar');

  var fSrch = (f.search || '').toLowerCase();
  var cases = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!passFilters(row, COL, f, fS, fE)) continue;

    var refId  = String(row[COL['REF ID']]           || '').trim();
    var al     = String(row[COL['ass_lead']]          || '').trim();
    var lead   = String(row[COL['lead']]              || '').trim();
    var cid    = String(row[COL['id']]                || '').trim();
    var name   = String(row[COL['name']]              || '').trim();
    var hub    = String(row[COL['hub']]               || '').trim();
    var notes  = String(row[COL['notes']]             || '').trim();
    var status = String(row[COL['case_status']]       || '').trim();
    var month  = String(row[COL['month']]             || '').trim();
    var closeR = String(row[COL['case_close_reason']] || '').trim();

    if (fSrch && (refId+' '+al+' '+cid+' '+name+' '+hub+' '+notes)
                   .toLowerCase().indexOf(fSrch) === -1) continue;

    // [FIX] Ambil nilai kolom via index yang sudah di-resolve
    var lastUpdVal  = iLastUpd  >= 0 ? row[iLastUpd]  : '';
    var janjiByrVal = iJanjiByr >= 0 ? row[iJanjiByr] : '';

    cases.push({
      refId             : refId,
      assLead           : al,
      lead              : lead,
      courierId         : cid,
      name              : name,
      hub               : hub,
      amount            : pa(row[COL['amount']]),
      notes             : notes,
      status            : status,
      rekPnp            : pa(row[COL['rek_pnp']]),
      dateReminder      : fd(row[COL['date_reminder']]),
      lastUpdate        : fd(lastUpdVal),       // [FIX]
      tanggalJanjiBayar : fd(janjiByrVal),       // [FIX]
      dpd               : calcDPD(janjiByrVal),  // [FIX]
      month             : month,
      caseCloseReason   : closeR
    });
  }

  return cases;
}

// =================================================================
//  getCaseDetail(refId)
//  [FIX] last_update & tanggal_janji_bayar kini pakai findColIdx
// =================================================================
function getCaseDetail(refId) {
  setupSheets();
  var sh      = getHistSheet();
  var data    = sh.getDataRange().getValues();
  if (data.length <= 1) return null;

  var headers = data[0];
  var COL     = colMap(headers);

  // [FIX] Gunakan findColIdx agar lookup tahan terhadap beda casing
  var iLastUpd  = findColIdx(headers, 'last_update');
  var iJanjiByr = findColIdx(headers, 'tanggal_janji_bayar');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL['REF ID']] || '').trim() !== refId) continue;

    var row = data[i];

    // [FIX] Ambil nilai kolom via index yang sudah di-resolve
    var lastUpdVal  = iLastUpd  >= 0 ? row[iLastUpd]  : '';
    var janjiByrVal = iJanjiByr >= 0 ? row[iJanjiByr] : '';

    return {
      refId             : refId,
      assLead           : String(row[COL['ass_lead']]          || '').trim(),
      lead              : String(row[COL['lead']]              || '').trim(),
      courierId         : String(row[COL['id']]                || '').trim(),
      name              : String(row[COL['name']]              || '').trim(),
      hub               : String(row[COL['hub']]               || '').trim(),
      amount            : pa(row[COL['amount']]),
      notes             : String(row[COL['notes']]             || '').trim(),
      status            : String(row[COL['case_status']]       || '').trim(),
      rekPnp            : pa(row[COL['rek_pnp']]),
      dateReminder      : fd(row[COL['date_reminder']]),
      lastUpdate        : fd(lastUpdVal),        // [FIX]
      month             : String(row[COL['month']]             || '').trim(),
      caseCloseReason   : String(row[COL['case_close_reason']] || '').trim(),
      tanggalJanjiBayar : fd(janjiByrVal),        // [FIX]
      dpd               : calcDPD(janjiByrVal),   // [FIX]
      notesHistory      : getNoteHistory_(refId)
    };
  }

  return null;
}

function getNoteHistory_(refId) {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(NOTES_SHEET);
  if (!sh) return [];
  var data=sh.getDataRange().getValues();
  if (data.length<=1) return [];
  var COL=colMap(data[0]),out=[];
  for (var i=1;i<data.length;i++) {
    if (String(data[i][COL['ref_id']]||'').trim()!==refId) continue;
    out.push({timestamp      :fdt(data[i][COL['timestamp']]),
              notes          :String(data[i][COL['notes']]            ||'').trim(),
              statusChanged  :String(data[i][COL['status_changed_to']]||'').trim(),
              updatedBy      :String(data[i][COL['updated_by']]       ||'').trim(),
              caseCloseReason:String(data[i][COL['case_close_reason']]||'').trim()});
  }
  return out.reverse();
}

// =================================================================
//  updateCase(refId, newStatus, newNotes, updatedBy, closeReason, tanggalJanjiBayar)
// =================================================================
function updateCase(refId, newStatus, newNotes, updatedBy, closeReason, tanggalJanjiBayar) {
  var sh      = getHistSheet();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];

  function col(name) { return findColIdx(headers, name); }

  var colRefId    = col('REF ID');
  var colStatus   = col('case_status');
  var colNotes    = col('notes');
  var colLastUpd  = col('last_update');
  var colCloseR   = col('case_close_reason');
  var colJanjiByr = col('tanggal_janji_bayar');

  if (colRefId < 0 || colStatus < 0) {
    return {success:false, error:'Kolom wajib tidak ditemukan. Cek header sheet history.'};
  }

  var rowIdx = -1, oldStatus = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colRefId]||'').trim() === refId) {
      rowIdx    = i + 1;
      oldStatus = String(data[i][colStatus]||'').trim();
      break;
    }
  }
  if (rowIdx === -1) return {success:false, error:'Case tidak ditemukan: '+refId};

  var statusToSet = newStatus || '';
  if (oldStatus === 'Open' && newNotes && !statusToSet) statusToSet = 'On Investigation';

  var now       = new Date();
  var hasChange = !!(statusToSet || newNotes);

  if (statusToSet && colStatus >= 0) {
    sh.getRange(rowIdx, colStatus + 1).setValue(statusToSet);
  }

  if (newNotes && colNotes >= 0) {
    var existing = String(sh.getRange(rowIdx, colNotes + 1).getValue()||'').trim();
    var entry    = '['+fdt(now)+'] '+newNotes;
    sh.getRange(rowIdx, colNotes + 1).setValue(existing ? existing+'\n'+entry : entry);
  }

  var finalStatus = statusToSet || oldStatus;
  var reasonToSet = '';
  if (finalStatus === 'Case Close' && closeReason && colCloseR >= 0) {
    reasonToSet = closeReason;
    sh.getRange(rowIdx, colCloseR + 1).setValue(closeReason);
  }

  if (tanggalJanjiBayar && colJanjiByr >= 0) {
    var jbDate = new Date(tanggalJanjiBayar);
    if (!isNaN(jbDate.getTime())) {
      var jbCell = sh.getRange(rowIdx, colJanjiByr + 1);
      jbCell.setValue(jbDate);
      jbCell.setNumberFormat('dd MMM yyyy');
    }
  }

  if (hasChange) {
    if (colLastUpd >= 0) {
      var lastUpdCell = sh.getRange(rowIdx, colLastUpd + 1);
      lastUpdCell.setValue(now);
      lastUpdCell.setNumberFormat('dd MMM yyyy');
      Logger.log('last_update updated at col ' + (colLastUpd+1) + ' row ' + rowIdx);
    } else {
      Logger.log('WARNING: kolom last_update tidak ditemukan. Header sheet: ' + JSON.stringify(headers));
    }
  }

  SpreadsheetApp.flush();

  addNoteHist_(refId, newNotes, finalStatus, updatedBy||'Dashboard', reasonToSet);

  return {success:true, newStatus:finalStatus};
}

function addNoteHist_(refId, notes, statusChanged, updatedBy, closeReason) {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(NOTES_SHEET);
  if (!sh) {
    sh=ss.insertSheet(NOTES_SHEET);
    sh.appendRow(NOTES_HEADERS);
    sh.getRange(1,1,1,NOTES_HEADERS.length).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
  }
  sh.appendRow([refId, new Date(), notes||'', statusChanged||'', updatedBy||'Dashboard', closeReason||'']);
}

// =================================================================
//  deleteCaseRow(refId)
//  Hapus 1 baris dari sheet history berdasarkan REF ID
// =================================================================
function deleteCaseRow(refId) {
  if (!refId) return { success: false, error: 'REF ID tidak boleh kosong' };

  var sh   = getHistSheet();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { success: false, error: 'Sheet kosong' };

  var colRefId = findColIdx(data[0], 'REF ID');
  if (colRefId < 0) return { success: false, error: 'Kolom REF ID tidak ditemukan' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colRefId] || '').trim() === refId) {
      sh.deleteRow(i + 1);
      SpreadsheetApp.flush();
      Logger.log('deleteCaseRow: baris ' + (i+1) + ' (REF ID: ' + refId + ') berhasil dihapus.');
      return { success: true };
    }
  }

  return { success: false, error: 'Case tidak ditemukan: ' + refId };
}

// ── Debug helper ─────────────────────────────────────────────
function getDashboardStats() {
  setupSheets();
  var sh=getHistSheet();
  return {historyRows:Math.max(0,sh.getLastRow()-1),
    sheetName:SpreadsheetApp.getActiveSpreadsheet().getName(),
    timezone:Session.getScriptTimeZone(),
    columns:sh.getLastRow()>0?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]:[]};
}
