/** =========================
 *  PEOPLE REPOSITORY (guarded)
 *  Lee la pestaña PERSONES con cabeceras:
 *   - Nom
 *   - Email (opcional)
 *   - Grups  (separados por coma o punto y coma, p. ej. "Claustre, Primària")
 *  Devuelve: { groups:[...], byGroup:{ Grup:[{nom,email,grups}] }, all:[...] }
 *  ========================= */
if (typeof PeopleRepo === 'undefined') {
  var PeopleRepo = (function(){

    function _readSheet_(sh){
      const v = sh.getDataRange().getValues();
      if (!v || v.length < 2) return { head: [], rows: [] };
      return { head: v[0].map(String), rows: v.slice(1) };
    }

    /** Public: list(sheetId, peopleSheetName) */
    function list(sheetId, peopleSheetName){
      const ss = SpreadsheetApp.openById(sheetId);
      const sh = ss.getSheetByName(peopleSheetName);
      if (!sh) return { groups: [], byGroup: {}, all: [] };

      const { head, rows } = _readSheet_(sh);
      const iNom   = head.indexOf('Nom');
      const iEmail = head.indexOf('Email'); // opcional
      const iGrups = head.indexOf('Grups');

      const all = [];
      const byGroup = {};
      const groupsSet = new Set();

      rows.forEach(r=>{
        const nom   = String((iNom>=0 ? r[iNom]   : '') || '').trim();
        const email = String((iEmail>=0 ? r[iEmail]: '') || '').trim();
        const grupsRaw = String((iGrups>=0 ? r[iGrups]: '') || '').trim();

        if (!nom && !email) return; // fila vacía

        const grups = grupsRaw
          ? grupsRaw.split(/[,;]+/).map(s=>s.trim()).filter(Boolean)
          : [];

        const person = { nom, email, grups };
        all.push(person);

        if (!grups.length){
          const g = '(Sense grup)';
          groupsSet.add(g);
          if (!byGroup[g]) byGroup[g] = [];
          byGroup[g].push(person);
        } else {
          grups.forEach(g=>{
            groupsSet.add(g);
            if (!byGroup[g]) byGroup[g] = [];
            byGroup[g].push(person);
          });
        }
      });

      const groups = Array.from(groupsSet).sort((a,b)=>a.localeCompare(b, 'ca'));
      groups.forEach(g=>{
        byGroup[g].sort((a,b)=>a.nom.localeCompare(b.nom, 'ca'));
      });

      return { groups, byGroup, all };
    }

    return { list };
  })();
}
