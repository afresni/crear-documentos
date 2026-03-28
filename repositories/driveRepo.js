/** =========================
 *  DRIVE REPOSITORY (idempotent)
 *  ========================= */
(function (global) {
  if (global.DriveRepo) return; // ya existe, no redeclarar

  function ensureFolder(parentFolder, name){
    const it = parentFolder.getFoldersByName(name);
    if (it.hasNext()) return it.next();
    return parentFolder.createFolder(name);
  }

  function tryGetFolder(parentFolder, name){
    const it = parentFolder.getFoldersByName(name);
    return it.hasNext() ? it.next() : null;
  }

  function removeFromRootIfPresent(file){
    const root = DriveApp.getRootFolder();
    const it = root.getFilesByName(file.getName());
    while (it.hasNext()){
      const f = it.next();
      if (f.getId() === file.getId()){
        root.removeFile(file);
        return;
      }
    }
  }

  function findFileByExactName(folder, exactName){
    const it = folder.getFilesByName(exactName);
    return it.hasNext() ? it.next() : null;
  }

  function findMaxNumero(ambitFolder, cursFormat, tipus, prefix, ambitSanitized, cursShort){
    const cursFolder = tryGetFolder(ambitFolder, cursFormat);
    if (!cursFolder) return 0;
    const finalFolder = tryGetFolder(cursFolder, tipus);
    if (!finalFolder) return 0;

    const pattern = new RegExp('^' + escapeRegExp(prefix) + '(\\d{2})' + escapeRegExp(ambitSanitized) + escapeRegExp(cursShort) + '$');
    let max = 0;
    const it = finalFolder.getFiles();
    while (it.hasNext()){
      const f = it.next();
      const m = String(f.getName()).match(pattern);
      if (m && m[1]){
        const n = parseInt(m[1],10);
        if (!isNaN(n)) max = Math.max(max, n);
      }
    }
    return max;
  }

  function escapeRegExp(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  global.DriveRepo = {
    ensureFolder,
    tryGetFolder,
    removeFromRootIfPresent,
    findFileByExactName,
    findMaxNumero
  };
})(this);
