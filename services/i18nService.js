/** =========================
 *  I18N SERVICE (UserProperties)
 *  ========================= */
const I18nService = (function(){
  const KEY = 'USER_LANG'; // 'ca' | 'es'

  function getUserLang(email){
    try{
      const up = PropertiesService.getUserProperties();
      const lang = (up.getProperty(KEY) || '').trim();
      return lang || 'ca';
    } catch(e){
      return 'ca';
    }
  }

  function setUserLang(email, lang){
    const v = String(lang || '').trim().toLowerCase();
    if (!v || !/^([a-z]{2})(-[A-Z]{2})?$/.test(v)) throw new Error('LANG_INVALID');
    const up = PropertiesService.getUserProperties();
    up.setProperty(KEY, v);
    return v;
  }

  return { getUserLang, setUserLang };
})();
