/** =========================
 *  DOCUMENT SERVICE (idempotent)
 *  ========================= */
(function (global) {
  if (global.DocService) return; // ja existeix
  
  function formatAcademicYear_(d){
    const any = d.getFullYear();
    const mes = d.getMonth() + 1;
    const anyInici = (mes >= 9) ? any : any - 1;
    const anyFi = anyInici + 1;
    return {
      cursFormat: `CURS ${anyInici}-${anyFi}`,
      cursShort: `${String(anyInici).slice(-2)}${String(anyFi).slice(-2)}`
    };
  }
  
  function sanitizeName_(s){
    return String(s).replace(/[\s\-\/\\:*?"<>|]/g, '');
  }
  
  /**
   * ✅ VERSIÓN CORREGIDA: Lee el número máximo desde HISTORIC
   * Si no hay registros previos, empieza en 1
   */
  function nextNumeroSmart_(key, ambitFolder, cursFormat, tipus, prefix, ambitSanitized, cursShort){
    const { SHEET_ID, HISTORIC_SHEET } = getConfig_();
    
    // Si no hay Excel configurado, empezar en 1
    if (!SHEET_ID || SHEET_ID === '__RELLENAR__') {
      return 1;
    }
    
    // ✅ Leer del Excel HISTORIC
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sh = ss.getSheetByName(HISTORIC_SHEET);
      
      // Si no existe la hoja, empezar en 1
      if (!sh) {
        return 1;
      }
      
      const lastRow = sh.getLastRow();
      
      // Si solo hay encabezados o está vacía, empezar en 1
      if (lastRow <= 1) {
        return 1;
      }
      
      // Obtener todos los datos
      const data = sh.getDataRange().getValues();
      
      // Si solo hay encabezados, empezar en 1
      if (data.length <= 1) {
        return 1;
      }
      
      // Obtener índices de columnas
      const headers = data[0];
      const colAmbit = headers.indexOf('ambit');
      const colTipus = headers.indexOf('tipus');
      const colCurs = headers.indexOf('curs');
      const colNumero = headers.indexOf('numero');
      
      // Si no se encuentran las columnas necesarias, empezar en 1
      if (colAmbit < 0 || colTipus < 0 || colCurs < 0 || colNumero < 0) {
        console.warn('Columnas no encontradas en HISTORIC');
        return 1;
      }
      
      // Buscar el número máximo para este ambit+tipus+curs
      let maxNum = 0;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rAmbit = String(row[colAmbit] || '').trim();
        const rTipus = String(row[colTipus] || '').trim();
        const rCurs = String(row[colCurs] || '').trim();
        const rNum = parseInt(row[colNumero]) || 0;
        
        // Comparar: mismo ambit (sanitizado), mismo tipus, mismo curso
        if (sanitizeName_(rAmbit) === ambitSanitized && rTipus === tipus && rCurs === cursFormat) {
          if (rNum > maxNum) {
            maxNum = rNum;
          }
        }
      }
      
      // Si no se encontró nada, empezar en 1
      return maxNum + 1;
      
    } catch(e) {
      console.error('Error leyendo HISTORIC:', e);
      // En caso de error, empezar en 1
      return 1;
    }
  }
  
  /**
   * Crea un document a partir de la plantilla del tipus.
   *  - ambit, tipus, dataReunio, horaReunio, llocReunio, assistents, ausents
   */
  function createFromTemplate(ambit, tipus, dataReunio, horaReunio, llocReunio, assistents, ausents){
    const tplId = ConfigRepo.getPlantillaIdByTipus(tipus);
    const ambitFolderId = ConfigRepo.getAmbitFolderIdByKey(ambit);
    if (!tplId) {
      throw new Error(`No hi ha plantilla configurada per al tipus '${tipus}'.`);
    }
    if (!ambitFolderId) {
      throw new Error(`No hi ha carpeta configurada per a l'àmbit '${ambit}'.`);
    }
    const d = new Date(dataReunio);
    if (isNaN(d.getTime())) {
      throw new Error(`La data de reunió '${dataReunio}' no és vàlida.`);
    }
    const { cursFormat, cursShort } = formatAcademicYear_(d);
    if (!tplId) {
      throw new Error(`No hi ha plantilla configurada per al tipus '${tipus}'.`);
    }
    if (!ambitFolderId) {
      throw new Error(`No hi ha carpeta configurada per a l'àmbit '${ambit}'.`);
    }
    const key = `${tipus}_${ambit}_${cursShort}`;
    const prefix = ConfigRepo.getPrefixByTipus(tipus);
    const ambitSan = sanitizeName_(ambit);
    const numero = nextNumeroSmart_(key, ambitFolder, cursFormat, tipus, prefix, ambitSan, cursShort);
    const name = `${prefix}${String(numero).padStart(2,'0')}${ambitSan}${cursShort}`;
    const finalName = name.substring(0, 255);
    const cursFolder  = DriveRepo.ensureFolder(ambitFolder, cursFormat);
    const finalFolder = DriveRepo.ensureFolder(cursFolder, tipus);
    
    // Si ja existeix un document amb el mateix nom, retornem el mateix (idempotent)
    const existing = DriveRepo.findFileByExactName(finalFolder, finalName);
    if (existing) {
      return {
        url: existing.getUrl(),
        docId: existing.getId(),
        folderUrl: finalFolder.getUrl(),
        name: finalName,
        cursFormat, cursShort, numero,
        existed: true
      };
    }
    
    // Copiar plantilla
    let tplFile;
    try {
      tplFile = DriveApp.getFileById(tplId);
    } catch (e) {
      throw new Error(`No s'ha pogut obrir la plantilla del tipus '${tipus}' (ID: ${tplId}). Revisa ID i permisos de Drive.`);
    }
    const newFile = tplFile.makeCopy(finalName);
    finalFolder.addFile(newFile);
    try { DriveRepo.removeFromRootIfPresent(newFile); } catch(_){}
    
    // Omplir el document
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    const formattedDate = Utilities.formatDate(
      d,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );
    body.replaceText('{{AMBIT}}', ambit || '');
    body.replaceText('{{DATA}}', formattedDate || '');
    body.replaceText('{{HORA}}', horaReunio || '');
    body.replaceText('{{LLOC}}', llocReunio || '');
    body.replaceText('{{ASSISTENTS}}', assistents || '');
    body.replaceText('{{AUSENTS}}', ausents || '');
    body.replaceText('{{CURS}}', cursFormat || '');
    doc.saveAndClose();
    
    return {
      url: newFile.getUrl(),
      docId: newFile.getId(),
      folderUrl: finalFolder.getUrl(),
      name: finalName,
      cursFormat, cursShort, numero,
      existed: false
    };
  }
  
  global.DocService = { createFromTemplate };
  
})(this);