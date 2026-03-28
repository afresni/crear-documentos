/** =========================
 *  VALIDATION LAYER
 *  ========================= */
const ValidationService = (function(){
  function validateNonce(nonce){
    if (!nonce || String(nonce).length < 10) throw new Error('NONCE_INVALIDO');
  }

  function validateCreatePayload(p){
    const obj = Object.assign({}, p);

    // ✅ Campos SIEMPRE requeridos (sin 'assistents')
    const required = ['ambit','tipus','dataReunio','horaReunio','llocReunio','nonce'];
    required.forEach(k=>{
      if (obj[k] === undefined || obj[k] === null || obj[k] === '')
        throw new Error(`FALTA_${k}`);
    });

    // ✅ NUEVO: Asistentes SOLO obligatorio para ACTES
    if (obj.tipus === 'ACTES') {
      if (!obj.assistents || String(obj.assistents).trim() === '') {
        throw new Error('FALTA_assistents');
      }
    }

    // Tipus válidos
    const TIPUS = ConfigRepo.getTipusList();
    if (!TIPUS.includes(obj.tipus)) throw new Error('TIPUS_INVALID');

    // Àmbit válido
    const ambitsKeys = ConfigRepo.getAmbitsList().map(a=>a.key);
    if (!ambitsKeys.includes(obj.ambit)) throw new Error('AMBIT_INVALID');

    // Regla: INSTRUCCIONS solo en ciertos ámbitos
    if (obj.tipus === 'INSTRUCCIONS') {
      const allow = ConfigRepo.getAmbitsAmbInstruccions();
      if (!allow.includes(obj.ambit)) throw new Error('AMBIT_SENSE_INSTRUCCIONS');
    }

    // Fecha válida (permitimos formatos comunes, lo convertirá DocService)
    const d = new Date(obj.dataReunio);
    if (isNaN(d.getTime())) throw new Error('DATA_REUNIO_INVALIDA');

    // Hora simple hh:mm (opcionalmente más estricto)
    if (!/^\d{1,2}:\d{2}$/.test(String(obj.horaReunio))) throw new Error('HORA_INVALIDA');

    // Longitudes razonables
    if (String(obj.llocReunio).length > 200) throw new Error('LLOC_MOLT_LLARG');
    
    // ✅ Solo validar longitud si hay asistentes (opcional para ORDRE/INSTRUCCIONS)
    if (obj.assistents && String(obj.assistents).length > 2000) {
      throw new Error('ASSISTENTS_MOLT_LLARG');
    }

    return obj;
  }

  return { validateNonce, validateCreatePayload };
})();