/** =========================
 *  CONFIG REPOSITORY
 *  Centraliza catálogos e IDs
 *  ========================= */
const ConfigRepo = (function(){

  // Puedes mover estos mapas a una hoja / Properties si quieres gestionarlos sin tocar código.
  const PLANTILLES = {
    "ACTES": "1xzeyb-QqbFxYc4CyXoRHZTxy2wnb4HFovtXKYqYDOd4",
    "ORDRE DEL DIA": "1s25bvG6nOT0RNPbtzgUGDnpWLgydXAzgtGNufOkQ-P0",
    "INSTRUCCIONS": "15nHybcjQN7KgqzKxeCW07y9wzzh8Ccx8aoEcb6GDKyI"
  };

  const PREFIX = {
    "ACTES": "A",
    "ORDRE DEL DIA": "O",
    "INSTRUCCIONS": "I"
  };

  const AMBITS = {
    "CLAUSTRE": "0AMl43wSUR4FCUk9PVA",
    "COMISSIÓ - MEDI AMBIENT": "0ANm57z8-LkSlUk9PVA",
    "COMISSIÓ - TIC": "0APy3M1F_Ui32Uk9PVA",
    "DEPARTAMENT - ART I MÚSICA": "0ABpm3pODJSF_Uk9PVA",
    "DEPARTAMENT - CIÈNCIES": "0AG2e48N31Vh2Uk9PVA",
    "DEPARTAMENT - LLENGÜES I HUMANITATS": "0ACmRak0LgFJFUk9PVA",
    "DEPARTAMENT - ORIENTACIÓ": "0AE7YRRw8yfDhUk9PVA",
    "DEPARTAMENT - PASTORAL": "0AJBt5WipugU1Uk9PVA",
    "EDUCACIÓ - INFANTIL": "0AGq6qGlZy2QXUk9PVA",
    "EDUCACIÓ - PRIMÀRIA": "0AKDvpAba01qZUk9PVA",
    "EDUCACIÓ - SECUNDÀRIA": "0AKRL0eC5Z2cJUk9PVA",
    "EQUIP - COMUNICACIÓ": "0AOHgLuPXlfy3Uk9PVA",
    "FOTOS I VÍDEOS": "0AKjk8cuglO6yUk9PVA",
    "INCIDÈNCIES": "0AOK9POO7f72uUk9PVA"
  };

  const AMBITS_AMB_INSTRUCCIONS = [
    "CLAUSTRE",
    "EDUCACIÓ - INFANTIL",
    "EDUCACIÓ - PRIMÀRIA",
    "EDUCACIÓ - SECUNDÀRIA"
  ];

  function getTipusList(){ return Object.keys(PLANTILLES); }
  function getAmbitsList(){
    return Object.keys(AMBITS).map(k=>({ key:k, folderId: AMBITS[k] }));
  }
  function getAmbitFolderIdByKey(key){ return AMBITS[key]; }
  function getPlantillaIdByTipus(t){ return PLANTILLES[t]; }
  function getPrefixByTipus(t){ return PREFIX[t] || ''; }
  function getAmbitsAmbInstruccions(){ return AMBITS_AMB_INSTRUCCIONS.slice(); }

  return {
    getTipusList, getAmbitsList,
    getAmbitFolderIdByKey, getPlantillaIdByTipus, getPrefixByTipus,
    getAmbitsAmbInstruccions
  };
})();
