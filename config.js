/* FolioSpace · shared client configuration. NEVER place secrets in this file. */
(() => {
  'use strict';
  const CONFIG = Object.freeze({
    APP_NAME: 'FolioSpace',
    // Set to the /exec Web App deployment URL (not /dev).
    GAS_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwo9fb7HOmV2G8pIfve83qHC1-tKuqHTQUC-_LxngA43NK72HKXpKb16WFajpsx4Xjg/exec',
    ENV: 'production',
    REQUEST_TIMEOUT_MS: 45000,
    CACHE_VERSION: '1.0',
    SESSION_KEY: 'foliospace.session.v1',
    NON_SENSITIVE_CACHE_TTL_MS: 90 * 1000,
    ROUTES: Object.freeze({ login: 'index.html', verify: 'verify.html', dashboard: 'main.html' })
  });
  const current = new Map();
  const ensureConfig = () => {
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(CONFIG.GAS_WEB_APP_URL)) {
      throw new Error('API belum dikonfigurasi. Isi GAS_WEB_APP_URL di config.js dengan URL /exec.');
    }
    if (!/^https?:\/\//.test(location.origin)) throw new Error('Buka website melalui localhost atau HTTPS, bukan file://');
  };
  const apiOrigin = origin => origin === 'https://script.google.com' ||
    /^https:\/\/[a-z0-9-]+-script\.googleusercontent\.com$/i.test(origin) ||
    origin === 'https://script.googleusercontent.com';
  const session = {
    get: () => { try { return sessionStorage.getItem(CONFIG.SESSION_KEY) || ''; } catch { return ''; } },
    set: value => { try { if(value) sessionStorage.setItem(CONFIG.SESSION_KEY,value); } catch {} },
    clear: () => { try { sessionStorage.removeItem(CONFIG.SESSION_KEY); } catch {} }
  };
  /**
   * Cross-origin FORM POST -> Apps Script HtmlService -> window.top.postMessage.
   * No fetch CORS bypass and no JSONP for credentials. Not a guaranteed universal
   * transport: test iframe/CSP support with the actual deployed GAS endpoint.
   */
  function api(operation,data={},opts={}) {
    try { ensureConfig(); if(navigator.onLine===false)throw new Error('Perangkat sedang offline. Sambungkan internet dan coba lagi.'); } catch(error) { return Promise.reject(error); }
    const requestId = crypto.randomUUID().replace(/-/g,''), frameName='fs_'+requestId;
    const timeout = Math.min(120000, Math.max(5000, opts.timeout || CONFIG.REQUEST_TIMEOUT_MS));
    return new Promise((resolve,reject) => {
      const frame=document.createElement('iframe');frame.name=frameName;frame.title='FolioSpace secure transport';
      frame.tabIndex=-1;frame.setAttribute('aria-hidden','true');frame.style.cssText='position:absolute;width:0;height:0;border:0;opacity:0;pointer-events:none';
      const form=document.createElement('form');form.method='POST';form.action=CONFIG.GAS_WEB_APP_URL;form.target=frameName;
      form.acceptCharset='UTF-8';form.hidden=true;form.style.display='none';
      const fields={requestId,replyOrigin:location.origin,payload:JSON.stringify({operation,data,session:opts.anonymous?'':session.get()})};
      Object.entries(fields).forEach(([key,value])=>{const el=document.createElement('input');el.name=key;el.value=value;el.type='hidden';form.append(el);});
      let timer,finished=false;
      const cleanup=()=>{if(finished)return;finished=true;clearTimeout(timer);window.removeEventListener('message',onMessage);frame.remove();form.remove();current.delete(requestId);};
      const onMessage=event=>{
        if(!apiOrigin(event.origin))return;
        const msg=event.data;
        if(!msg || msg.foliospace!==true || msg.requestId!==requestId)return;
        cleanup();
        if(msg.success)resolve(msg.data);
        else {const error=new Error(msg.error?.message||'Permintaan gagal.');error.code=msg.error?.code||'SERVER_ERROR';if(error.code==='UNAUTHORIZED'&&!opts.anonymous){session.clear();cache.clearAll();if(!location.search.includes('p='))location.replace('index.html');}reject(error);}
      };
      window.addEventListener('message',onMessage);
      document.body.append(frame,form);
      current.set(requestId,{operation,started:performance.now()});
      timer=setTimeout(()=>{cleanup();const error=new Error('Tidak menerima respons API. Periksa deployment Apps Script, APP_ORIGIN, dan apakah browser memblokir iframe.');error.code='TIMEOUT';reject(error);},timeout);
      try { form.submit(); } catch(error) {cleanup();reject(error);}
    });
  }
  function cacheKey(userId){return 'fs.cache.'+CONFIG.CACHE_VERSION+'.'+userId;}
  const cache={
    get(userId){try{const r=JSON.parse(sessionStorage.getItem(cacheKey(userId))||'null');return r&&Date.now()-r.ts<CONFIG.NON_SENSITIVE_CACHE_TTL_MS?r.data:null;}catch{return null;}},
    set(userId,data){try{const safe={...data};delete safe.user;sessionStorage.setItem(cacheKey(userId),JSON.stringify({ts:Date.now(),data:safe}));}catch{}},
    clear(userId){try{if(userId)sessionStorage.removeItem(cacheKey(userId));}catch{}},
    clearAll(){try{Object.keys(sessionStorage).filter(k=>k.startsWith('fs.cache.')).forEach(k=>sessionStorage.removeItem(k));}catch{}}
  };
  const url = value => { try { const u=new URL(String(value));return /^https?:$/.test(u.protocol) && !u.username && !u.password ? u.href : null; } catch { return null; } };
  const escape = value => String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  window.FolioSpace = Object.freeze({CONFIG,api,session,cache,url,escape,requestStats:current});
})();
