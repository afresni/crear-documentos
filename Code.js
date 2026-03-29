/** =========================
 *  ROUTER + ORQUESTACIÓN
 *  ========================= */
const PROJECT_VERSION = 'v1';
const PROJECT_NAME_UI = 'Documents de reunió - Actes/Ordre del dia/Instruccions';

function doGet() {
  return HtmlService.createTemplateFromFile('ui/index')
    .evaluate()
    .setTitle(PROJECT_NAME_UI)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(path){ return HtmlService.createTemplateFromFile(path).getRawContent(); }

/* ======== CONFIG ======== */
function getConfig_() {
  const p = PropertiesService.getScriptProperties();
  const SHEET_ID        = p.getProperty('SHEET_ID')        || '__RELLENAR__';
  const LOG_SHEET       = p.getProperty('LOG_SHEET')       || 'LOGS';
  const HISTORIC_SHEET  = p.getProperty('HISTORIC_SHEET')  || 'HISTORIC';
  const ADMIN_EMAIL     = p.getProperty('ADMIN_EMAIL')     || 'admin@tu-dominio.com';
  return { SHEET_ID, LOG_SHEET, HISTORIC_SHEET, ADMIN_EMAIL };
}

/* ======== LOGGING (no bloquea si falla) ======== */
function log_(action, ok, message, extra) {
  try {
    const {SHEET_ID, LOG_SHEET} = getConfig_();
    if (!SHEET_ID || SHEET_ID === '__RELLENAR__') return;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
    const row = [
      new Date().toISOString(),
      (function(){ try { return Session.getActiveUser().getEmail() || ''; } catch(_){ return ''; } })(),
      action,
      ok ? 'OK' : 'ERR',
      message || '',
      JSON.stringify(extra || {})
    ];
    sh.appendRow(row);
  } catch(e) {
    console.error('[log_] failed', e);
  }
}

/* ======== BOOT (sin roles) ======== */
function api_boot(){
  try {
    const meta = {
      version: PROJECT_VERSION,
      tz: Session.getScriptTimeZone(),
      user: (function(){ try { return Session.getActiveUser().getEmail() || ''; } catch(_){ return ''; } })()
    };
    const catalog = {
      tipus: ConfigRepo.getTipusList(),               // ["ACTES","ORDRE DEL DIA","INSTRUCCIONS"]
      ambits: ConfigRepo.getAmbitsList(),             // [{key:"CLAUSTRE", folderId:"..."}...]
      ambitsAmbInstruccions: ConfigRepo.getAmbitsAmbInstruccions()
    };
    log_('BOOT', true, 'OK');
    return { ok:true, meta, catalog };
  } catch(e){
    log_('BOOT', false, e.message);
    return { ok:false, code:'BOOT_ERR', message:e.message };
  }
}

/* ======== CREATE DOC ======== */
function api_createDoc(payload){
  const lock = LockService.getScriptLock();
 const locked = lock.tryLock(30000);
  try {
    if (!locked) {
      throw new Error('LOCK_TIMEOUT');
    }
    ValidationService.validateNonce(payload?.nonce);
    const clean = ValidationService.validateCreatePayload(payload);

    // Ausents: cogemos de clean si el validador lo devuelve, o del payload en su defecto
    const ausents = ((clean && clean.ausents) || payload.ausents || '').trim();

    const res = DocService.createFromTemplate(
      clean.ambit,
      clean.tipus,
      clean.dataReunio,
      clean.horaReunio,
      clean.llocReunio,
      clean.assistents,
      ausents              // 🔹 NUEVO PARÁMETRO
    ); // { url, docId, folderUrl, name, cursFormat, cursShort, numero, existed }

    // Guardar histórico (si tienes hoja HISTORIC configurada)
    const { SHEET_ID, HISTORIC_SHEET } = getConfig_();
    if (SHEET_ID && SHEET_ID !== '__RELLENAR__') {
      HistoricService.append({
        timestampISO: new Date().toISOString(),
        user: (function(){ 
          try { 
            return Session.getActiveUser().getEmail() || ''; 
          } catch(_){ 
            return ''; 
          } 
        })(),
        ambit: clean.ambit,
        tipus: clean.tipus,
        curs: res.cursFormat,
        numero: res.numero,
        docId: res.docId,
        name: res.name,
        url: res.url,
        folderUrl: res.folderUrl,
        status: res.existed ? 'EXISTING' : 'CREATED'
      }, SHEET_ID, HISTORIC_SHEET);
    }

    log_(
      'CREATE_DOC',
      true,
      res.existed ? 'YA_EXISTIA' : 'CREADO',
      { ambit: clean.ambit, tipus: clean.tipus, url: res.url }
    );

    return {
      ok: true,
      message: res.existed ? 'El document ja existia' : 'Document creat',
      url: res.url,
      folderUrl: res.folderUrl,
      existed: res.existed
    };

  } catch(e){
    log_('CREATE_DOC', false, e.message, { payload });
    try {
      const { ADMIN_EMAIL } = getConfig_();
      if (ADMIN_EMAIL && ADMIN_EMAIL !== 'admin@tu-dominio.com') {
        MailApp.sendEmail(
          ADMIN_EMAIL,
          `❌ ERROR CREATE_DOC: ${e.message}`,
          JSON.stringify(payload, null, 2)
        );
      }
    } catch(_){}
    return { ok:false, code:'CREATE_ERR', message:e.message };
  } finally {
    try{ lock.releaseLock(); }catch(_){}
  }
}

/* ======== LIST RECENT (últimos N del HISTORIC) ======== */
function api_listRecent(limit){
  try {
    const { SHEET_ID, HISTORIC_SHEET } = getConfig_();
    if (!SHEET_ID || SHEET_ID === '__RELLENAR__') return { ok:true, data: [] };
    const data = HistoricService.listRecent(SHEET_ID, HISTORIC_SHEET, Number(limit||10));
    return { ok:true, data };
  } catch(e){
    log_('LIST_RECENT', false, e.message);
    return { ok:false, code:'LIST_RECENT_ERR', message:e.message };
  }
}

/* ======== SEARCH ======== */
function api_search(params){
  try {
    const { SHEET_ID, HISTORIC_SHEET } = getConfig_();
    if (!SHEET_ID || SHEET_ID === '__RELLENAR__') return { ok:true, data: [] };
    const res = HistoricService.search(SHEET_ID, HISTORIC_SHEET, params||{});
    return { ok:true, data: res };
  } catch(e){
    log_('SEARCH', false, e.message, { params });
    return { ok:false, code:'SEARCH_ERR', message:e.message };
  }
}
/* ======== PEOPLE (grups i persones) ======== */
/* Ara llegeix PROFESSORAT de la BD_CENTRE_GESTIÓ_INTERNA (DB_CENTRE_ID)
   i ho adapta al format: { groups:[...], byGroup:{ Grup:[{nom,email,grups}] }, all:[...] } */
function api_people(){
  try {
    const data = DBService.getProfesAsPeople_();
    // (opcional) log_: descomenta si vols registrar ús
    // log_('PEOPLE', true, 'OK', { groups: data.groups.length, people: data.all.length });
    return { ok:true, groups: data.groups, byGroup: data.byGroup, all: data.all };
  } catch (e){
    try { log_('PEOPLE', false, e.message); } catch(_){}
    return { ok:false, code:'PEOPLE_ERR', message:e.message };
  }
}