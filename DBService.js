/** =========================
 *  DB SERVICE - BD_CENTRE_GESTIÓ_INTERNA
 *  Lee PROFESSORAT y lo adapta al formato de People:
 *    { groups:[...], byGroup:{ Grup:[{nom,email,grups}] }, all:[...] }
 *  ========================= */
const DBService = (() => {
  // ID de la BD central (BD_CENTRE_GESTIÓ_INTERNA)
  const DB_ID = PropertiesService.getScriptProperties().getProperty('DB_CENTRE_ID');

  function getDb_() {
    if (!DB_ID) {
      throw new Error("No s'ha configurat la propietat DB_CENTRE_ID al projecte.");
    }
    return SpreadsheetApp.openById(DB_ID);
  }

  function getSheet_(name) {
    const ss = getDb_();
    const sh = ss.getSheetByName(name);
    if (!sh) {
      throw new Error(`No s'ha trobat la fulla '${name}' a la BD_CENTRE_GESTIÓ_INTERNA.`);
    }
    return sh;
  }

  /**
   * Traduce la columna Etapa a nom de grup
   *   INF  -> Infantil
   *   PRI  -> Primària
   *   ESO  -> Secundària
   *   CENTRE -> Centre
   */
  function etapaToGroup_(etapa) {
    const e = String(etapa || '').toUpperCase().trim();
    if (!e) return null;
    if (e === 'INF')    return 'Infantil';
    if (e === 'PRI')    return 'Primària';
    if (e === 'ESO')    return 'Secundària';
    if (e === 'CENTRE') return 'Centre';
    return e; // por si afegeixes altres valors en el futur
  }

  /**
   * Lee la fulla PROFESSORAT:
   *  ID_PROF | Nom | Llinatges | Email | Etapa | Departament | Rol | Actiu
   *
   * Devuelve en format People:
   *  {
   *    groups:  [ 'Infantil','Primària','Secundària','Docent','Coordinador TIC', ... ],
   *    byGroup: { 'Infantil': [ {nom,email,grups}, ... ], ... },
   *    all:     [ {nom,email,grups}, ... ]
   *  }
   *
   * NOTA: La columna Rol es pot posar amb combinacions:
   *   "Coordinador TIC / Docent", "Docent/Cap d’estudis secundaria", etc.
   *   Aquí es separa per "/", "," o ";" i es crea un grup per a cada part:
   *   -> "Coordinador TIC", "Docent", "Cap d’estudis secundaria", ...
   */
  function getProfesAsPeople_() {
    const sh = getSheet_('PROFESSORAT');
    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) {
      return { groups: [], byGroup: {}, all: [] };
    }

    const header = values[0].map(String);
    const rows   = values.slice(1);

    const idx = {
      nom:        header.indexOf('Nom'),
      llinatges:  header.indexOf('Llinatges'),
      email:      header.indexOf('Email'),
      etapa:      header.indexOf('Etapa'),
      rol:        header.indexOf('Rol'),
      actiu:      header.indexOf('Actiu')
    };

    const all = [];
    const byGroup = {};
    const groupsSet = new Set();

    rows.forEach(r => {
      const actiuVal = idx.actiu >= 0 ? String(r[idx.actiu] || '').toLowerCase().trim() : 'sí';
      if (actiuVal && actiuVal !== 'sí') return; // només profes actius

      const nom        = idx.nom       >= 0 ? String(r[idx.nom]       || '').trim() : '';
      const llinatges  = idx.llinatges >= 0 ? String(r[idx.llinatges] || '').trim() : '';
      const email      = idx.email     >= 0 ? String(r[idx.email]     || '').trim() : '';
      const etapa      = idx.etapa     >= 0 ? String(r[idx.etapa]     || '').trim() : '';
      const rol        = idx.rol       >= 0 ? String(r[idx.rol]       || '').trim() : '';

      if (!nom && !email) return; // fila buida

      const fullName = llinatges ? `${nom} ${llinatges}` : nom;

      // Construïm l'array de grups: etapa + rols desglossats
      const grups = [];

      // Grup per etapa
      const gEtapa = etapaToGroup_(etapa);
      if (gEtapa) grups.push(gEtapa);

      // Grups per rol: separem per "/", "," o ";" per evitar combinats
      if (rol) {
        rol
          .split(/[\/,;]+/)
          .map(s => s.trim())
          .filter(Boolean)
          .forEach(part => grups.push(part));
      }

      // Si no hi ha cap grup definit, el posem a "(Sense grup)"
      if (!grups.length) {
        grups.push('(Sense grup)');
      }

      const person = { nom: fullName, email, grups };

      all.push(person);
      grups.forEach(g => {
        groupsSet.add(g);
        if (!byGroup[g]) byGroup[g] = [];
        byGroup[g].push(person);
      });
    });

    const groups = Array.from(groupsSet).sort((a, b) => a.localeCompare(b, 'ca'));
    groups.forEach(g => {
      byGroup[g].sort((a, b) => a.nom.localeCompare(b.nom, 'ca'));
    });

    return { groups, byGroup, all };
  }

  return {
    getProfesAsPeople_
  };
})();
