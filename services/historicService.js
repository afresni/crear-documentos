/** =========================
 *  HISTORIC SERVICE (idempotent)
 *  ========================= */
(function (global) {
  if (global.HistoricService) return;

  function ensureHeader_(sh){
    if (sh.getLastRow() === 0) {
      const headers = ['timestampISO','user','ambit','tipus','curs','numero','docId','name','url','folderUrl','status'];
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }

  function append(rec, SHEET_ID, HISTORIC_SHEET){
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(HISTORIC_SHEET) || ss.insertSheet(HISTORIC_SHEET);
    ensureHeader_(sh);
    const row = [
      rec.timestampISO, rec.user, rec.ambit, rec.tipus, rec.curs, rec.numero,
      rec.docId, rec.name, rec.url, rec.folderUrl, rec.status || ''
    ];
    sh.appendRow(row);
  }

  function listRecent(SHEET_ID, HISTORIC_SHEET, limit){
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(HISTORIC_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];
    const values = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    values.sort((a,b)=> new Date(b[0]) - new Date(a[0]));
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    return values.slice(0, limit).map(r => Object.fromEntries(headers.map((h,i)=>[h,r[i]])));
  }

  function search(SHEET_ID, HISTORIC_SHEET, params){
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(HISTORIC_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const values = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    let rows = values.map(r => Object.fromEntries(headers.map((h,i)=>[h,r[i]])));

    if (params.ambit) rows = rows.filter(x => String(x.ambit) === String(params.ambit));
    if (params.tipus) rows = rows.filter(x => String(x.tipus) === String(params.tipus));
    if (params.text)  rows = rows.filter(x => String(x.name||'').toLowerCase().includes(String(params.text).toLowerCase()));

    const hasMin = !!params.dateMin;
    const hasMax = !!params.dateMax;
    let minTS, maxTS;
    if (hasMin) minTS = new Date(params.dateMin + 'T00:00:00');
    if (hasMax) maxTS = new Date(params.dateMax + 'T23:59:59');
    if (hasMin || hasMax){
      rows = rows.filter(x=>{
        const t = new Date(x.timestampISO);
        if (isNaN(t.getTime())) return false;
        if (hasMin && t < minTS) return false;
        if (hasMax && t > maxTS) return false;
        return true;
      });
    }

    rows.sort((a,b)=> new Date(b.timestampISO) - new Date(a.timestampISO));
    const limit = Number(params.limit || 200);
    return rows.slice(0, limit);
  }

  global.HistoricService = { append, listRecent, search };
})(this);
