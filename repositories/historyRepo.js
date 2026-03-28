/** =========================
 *  HISTORY REPO
 *  ========================= */
const HistoryRepo = (function(){
  const HIST_SHEET = 'HISTORIC';
  const HEADERS = [
    'timestampISO','user','ambit','tipus','dataReunio','horaReunio',
    'docId','url','folderUrl','name','status'
  ];

  function open_(){
    const { SHEET_ID } = getConfig_();
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(HIST_SHEET);
    if (!sh) sh = ss.insertSheet(HIST_SHEET);
    if (sh.getLastRow() === 0){
      sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    } else {
      const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      if (!headers.includes('status')){
        sh.getRange(1,headers.length+1,1,1).setValues([['status']]);
      }
    }
    return { ss, sh };
  }

  function append(rec){
    const { sh } = open_();
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    if (!('status' in rec)) rec.status = 'CREADO';
    const row = headers.map(k=>rec[k] ?? '');
    sh.appendRow(row);
  }

  function listRecent(limit){
    const { sh } = open_();
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const rows = Math.min(limit||10, lastRow-1);
    const range = sh.getRange(lastRow-rows+1, 1, rows, sh.getLastColumn());
    const vals = range.getValues().reverse();
    return vals.map(r => Object.fromEntries(headers.map((h,i)=>[h, r[i]])));
  }

  function search(params){
    const { sh } = open_();
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const vals = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();

    const fromTs = params?.from ? new Date(params.from).getTime() : -Infinity;
    const toTs   = params?.to   ? new Date(params.to).getTime()   :  Infinity;
    const type   = params?.tipus ? String(params.tipus) : '';
    const ambit  = params?.ambit ? String(params.ambit) : '';
    const q      = params?.q ? String(params.q).toLowerCase().trim() : '';

    const out = [];
    for (const r of vals){
      const rec = Object.fromEntries(headers.map((h,i)=>[h, r[i]]));
      if (type && rec.tipus !== type) continue;
      if (ambit && rec.ambit !== ambit) continue;

      const t = rec.dataReunio ? new Date(rec.dataReunio).getTime() : NaN;
      if (isNaN(t)) continue;
      if (t < fromTs || t > toTs) continue;

      if (q){
        const hay = (rec.name||'').toString().toLowerCase().includes(q) ||
                    (rec.tipus||'').toString().toLowerCase().includes(q) ||
                    (rec.ambit||'').toString().toLowerCase().includes(q);
        if (!hay) continue;
      }
      out.push(rec);
    }
    out.sort((a,b)=> new Date(b.dataReunio) - new Date(a.dataReunio));
    return out;
  }

  function markFinalByDocId(docId, updates){
    const { sh } = open_();
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));
    const data = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();
    for (let i=0;i<data.length;i++){
      if (String(data[i][idx.docId]) === String(docId)){
        data[i][idx.status] = 'FINAL';
        if (updates?.folderUrl) data[i][idx.folderUrl] = updates.folderUrl;
        sh.getRange(i+2,1,1,sh.getLastColumn()).setValues([data[i]]);
        return Object.fromEntries(headers.map((h,j)=>[h,data[i][j]]));
      }
    }
    throw new Error('DOC_NO_ENCONTRADO');
  }

  return { append, listRecent, search, markFinalByDocId, HEADERS };
})();
