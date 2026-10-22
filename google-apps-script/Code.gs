/**
 * Student Driver Log — Google Sheets backend
 * Deploy as Web app: Execute as Me; Who has access: Anyone.
 * WARNING: Anyone with the deployment URL can read/write/delete all records.
 * Do not put sensitive personal information in this sheet.
 */
const TAB = 'Sessions';
const HEADERS = ['id','date','period','minutes','supervisor','notes','createdAt'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) sh = ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}
function output_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function sessions_() {
  const rows = sheet_().getDataRange().getValues();
  return rows.slice(1).filter(r => r[0]).map(r => ({
    id:String(r[0]),date:String(r[1]),period:String(r[2]),
    minutes:Number(r[3]),supervisor:String(r[4]||''),
    notes:String(r[5]||''),createdAt:String(r[6]||'')
  }));
}
function doGet(e) {
  let result;
  try { result = {ok:true,sessions:sessions_()}; }
  catch(err) { result = {ok:false,error:String(err)}; }
  const callback = String((e && e.parameter && e.parameter.callback) || '');
  if (/^drivelogCallback_[A-Za-z0-9_]+$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return output_(result);
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const sh = sheet_();
    if (req.action === 'add') {
      const s = req.session || {};
      const minutes = Number(s.minutes);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s.date||'')) ||
          !['day','night'].includes(s.period) ||
          !Number.isInteger(minutes) || minutes < 1 || minutes > 1440)
        throw new Error('Invalid driving session');
      const id = /^[a-zA-Z0-9_-]{8,80}$/.test(String(s.id||'')) ? String(s.id) : Utilities.getUuid();
      if (sessions_().some(row => row.id === id)) return output_({ok:true,id});
      sh.appendRow([id, String(s.date), s.period, minutes,
        String(s.supervisor||'').slice(0,100),
        String(s.notes||'').slice(0,500),new Date().toISOString()]);
      return output_({ok:true,id});
    }
    if (req.action === 'delete') {
      const id = String(req.id||'');
      const ids = sh.getRange(2,1,Math.max(1,sh.getLastRow()-1),1).getValues();
      const idx = ids.findIndex(r=>String(r[0])===id);
      if(idx<0) throw new Error('Session not found');
      sh.deleteRow(idx+2);
      return output_({ok:true});
    }
    throw new Error('Unknown action');
  } catch(err) { return output_({ok:false,error:String(err)}); }
  finally { lock.releaseLock(); }
}
