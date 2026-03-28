/** =========================
 *  ROLES REPOSITORY
 *  ========================= */
const RolesRepo = (function(){
  const ROLES_SHEET = 'ROLES'; // Email | Rol
  const VALID_ROLES = ['view','create','edit','delete','admin'];

  function getSheet_(){
    const { SHEET_ID } = getConfig_();
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return ss.getSheetByName(ROLES_SHEET) || null;
  }

  function normalizeRole_(r){
    r = String(r||'').toLowerCase().trim();
    if (VALID_ROLES.includes(r)) return r;
    return '';
  }

  function getPropertiesRoles_(){
    const p = PropertiesService.getScriptProperties();
    const map = {
      view: (p.getProperty('ROLE_VIEW')||'').split(',').map(s=>s.trim()).filter(Boolean),
      create: (p.getProperty('ROLE_CREATE')||'').split(',').map(s=>s.trim()).filter(Boolean),
      edit: (p.getProperty('ROLE_EDIT')||'').split(',').map(s=>s.trim()).filter(Boolean),
      delete: (p.getProperty('ROLE_DELETE')||'').split(',').map(s=>s.trim()).filter(Boolean),
      admin: (p.getProperty('ROLE_ADMIN')||'').split(',').map(s=>s.trim()).filter(Boolean),
    };
    return map;
  }

  function getAll(){
    const sheet = getSheet_();
    const propRoles = getPropertiesRoles_();
    const res = { view:new Set(propRoles.view), create:new Set(propRoles.create),
                  edit:new Set(propRoles.edit), delete:new Set(propRoles.delete), admin:new Set(propRoles.admin) };

    if (sheet){
      const last = sheet.getLastRow();
      if (last >= 2){
        const values = sheet.getRange(2,1,last-1,2).getValues(); // Email | Rol
        values.forEach(([email, rol])=>{
          email = String(email||'').trim();
          const r = normalizeRole_(rol);
          if (email && r){ res[r].add(email); }
        });
      }
    }
    return {
      view: Array.from(res.view),
      create: Array.from(res.create),
      edit: Array.from(res.edit),
      delete: Array.from(res.delete),
      admin: Array.from(res.admin),
    };
  }

  function userHasRole(email, role){
    if (!email) return false;
    const all = getAll();
    if (all.admin.includes(email)) return true;
    if (role === 'view'){
      return all.view.includes(email) || all.create.includes(email) || all.edit.includes(email) || all.delete.includes(email);
    }
    return (all[role]||[]).includes(email);
  }

  return { getAll, userHasRole };
})();
