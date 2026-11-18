/**
 * FolioSpace · Google Apps Script backend (V8 runtime)
 * Deploy as Web App: execute as owner; access anyone.
 * Script Properties: SPREADSHEET_ID, APP_ORIGIN, APP_NAME (optional),
 * PBKDF2_ITERATIONS (optional), EXTRA_ORIGINS (optional, comma-separated).
 * Form POST -> HtmlService postMessage bridge (no fake no-cors / sensitive JSONP).
 * Important: TEST the iframe transport in your deployed environment before public launch.
 */
'use strict';
const FS = Object.freeze({
  VERSION: '1.0.0', SESSION_MS: 7 * 24 * 3600 * 1000,
  VERIFY_MS: 24 * 3600 * 1000, RESET_MS: 40 * 60 * 1000,
  MAX_PORTFOLIOS: 5000, MAX_IMPORT: 150, MAX_EVENTS: 30000,
  HEADERS: {
    Users: ['id','name','email','passwordHash','salt','iterations','verified','createdAt','updatedAt'],
    Portfolios: ['id','userId','title','description','url','categoryId','type','platform','thumbnail','tags','status','visibility','featured','displayOrder','createdAt','updatedAt'],
    Categories: ['id','userId','name','color','icon','displayOrder','createdAt','updatedAt'],
    Collections: ['id','userId','name','description','visibility','displayOrder','createdAt','updatedAt'],
    CollectionItems: ['id','userId','collectionId','portfolioId','position','createdAt'],
    PublicProfiles: ['userId','slug','headline','bio','location','avatar','socialLinks','contactEmail','isPublic','layout','gridColumns','headerStyle','accent','showDescriptions','showCategories','showCollections','updatedAt'],
    AnalyticsEvents: ['id','userId','portfolioId','collectionId','type','at','visitorHash'],
    Sessions: ['id','userId','tokenHash','expiresAt','revokedAt','createdAt'],
    VerificationTokens: ['id','userId','tokenHash','expiresAt','usedAt','createdAt'],
    PasswordResetTokens: ['id','userId','tokenHash','expiresAt','usedAt','createdAt'],
    UserSettings: ['userId','layout','sort','dateFormat','language','motion','updatedAt']
  }
});
function property_(key, fallback) { return PropertiesService.getScriptProperties().getProperty(key) || fallback || ''; }
function appOrigin_() { return property_('APP_ORIGIN').replace(/\/$/, ''); }
function allowedOrigin_(origin) {
  const allowed = [appOrigin_()].concat(property_('EXTRA_ORIGINS').split(',').map(x => x.trim()));
  return !!origin && allowed.includes(origin) && /^https?:\/\/[^\s/]+(?::\d+)?$/.test(origin);
}
function now_() { return new Date().toISOString(); }
function uid_() { return Utilities.getUuid(); }
function token_() { return uid_().replace(/-/g,'') + uid_().replace(/-/g,''); }
function digest_(value) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value)).map(x => ('0'+(x&255).toString(16)).slice(-2)).join(''); }
function equals_(a,b) { a=String(a||''); b=String(b||''); let n=a.length^b.length; for(let i=0;i<Math.max(a.length,b.length);i++) n |= (a.charCodeAt(i)||0)^(b.charCodeAt(i)||0); return n===0; }
function err_(code, msg) { const e = new Error(msg); e.code=code; throw e; }
function assert_(condition, msg, code) { if (!condition) err_(code || 'VALIDATION_ERROR', msg); }
function str_(s, n) { return String(s == null?'':s).trim().slice(0, n || 2000); }
function bool_(v) { return v === true || v === 'true' || v === 1 || v === '1'; }
function parse_(v, def) { try { return JSON.parse(String(v||'')); } catch(e) { return def; } }
function safeUrl_(u) { const s=str_(u,2048); assert_(/^https?:\/\/[a-z0-9.-]+(?::\d{2,5})?(?:[\/?#][^\s<>\"']*)?$/i.test(s) && !/[\u0000-\u001f]/.test(s),'URL tidak valid (hanya http/https).'); return s; }
function email_(v) { const s = str_(v, 240).toLowerCase(); assert_(/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s),'Format email tidak valid.'); return s; }
function validPass_(v) { assert_(typeof v==='string' && v.length>=10 && v.length<=128 && /[a-z]/i.test(v) && /\d/.test(v),'Password minimal 10 karakter, mengandung huruf dan angka.'); }
function reply_(id, origin, response) {
  const output = { foliospace: true, requestId: str_(id,100), ...response };
  // Escape HTML-sensitive characters and JS line separators; origin MUST be exact configured origin.
  const jsData = JSON.stringify(output).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script>window.top.postMessage(' + jsData + ',' + JSON.stringify(origin) + ');<\/script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doGet(e) { return HtmlService.createHtmlOutput('<!doctype html><html><body><h3>FolioSpace API</h3><p>POST bridge endpoint. Open the configured frontend to use the application.</p></body></html>'); }
function doPost(e) {
  const p = (e && e.parameter) || {};
  const id = str_(p.requestId,100), origin = str_(p.replyOrigin,300);
  if(!allowedOrigin_(origin)) return HtmlService.createHtmlOutput('Origin not allowed.');
  try {
    assert_(p.payload && p.payload.length<300000,'Payload terlalu besar.','PAYLOAD_TOO_LARGE');
    const request = JSON.parse(p.payload);
    assert_(request && typeof request.operation === 'string','Operation tidak valid.');
    const data = dispatch_(request.operation, request.data||{}, request.session||'');
    return reply_(id,origin,{success:true,data:data,error:null,message:'OK'});
  } catch(ex) {
    const code = ex && ex.code || 'SERVER_ERROR';
    if(code==='SERVER_ERROR') console.error('FolioSpace error:',String(ex && ex.stack || ex));
    return reply_(id,origin,{success:false,data:null,error:{code:code,message:code==='SERVER_ERROR'?'Terjadi kendala pada server.':String(ex.message||'Operation failed.')},message:'Error'});
  }
}
function db_() {
  const id = property_('SPREADSHEET_ID'); assert_(id,'SPREADSHEET_ID belum diatur.','NOT_CONFIGURED');
  return SpreadsheetApp.openById(id);
}
function setupDatabase() {
  const book = db_();
  Object.keys(FS.HEADERS).forEach(name=>{
    let sheet=book.getSheetByName(name);
    if(!sheet) sheet=book.insertSheet(name);
    const h=FS.HEADERS[name];
    if(sheet.getLastRow()===0) { sheet.getRange(1,1,1,h.length).setValues([h]); sheet.setFrozenRows(1); sheet.getRange(1,1,1,h.length).setBackground('#e5f8ec').setFontWeight('bold'); }
    else { const actual=sheet.getRange(1,1,1,h.length).getValues()[0]; assert_(actual.join('|')===h.join('|'), 'Header berbeda pada '+name,'SCHEMA_MISMATCH'); }
  });
  return 'FolioSpace database ready: '+Object.keys(FS.HEADERS).length+' sheets';
}
function sheet_(name) { const sh=db_().getSheetByName(name); assert_(sh,'Sheet '+name+' belum ada; jalankan setupDatabase().','DB_NOT_READY'); return sh; }
function rows_(name) {
  const sh=sheet_(name), last=sh.getLastRow();
  if(last<2) return [];
  const headers=FS.HEADERS[name], data=sh.getRange(2,1,last-1,headers.length).getValues();
  return data.map((cells,i)=>{ const obj={_row:i+2}; headers.forEach((h,j)=>obj[h]=cells[j] instanceof Date?cells[j].toISOString():cells[j]); return obj; }).filter(x=>x[headers[0]]!=='');
}
function publicRow_(r) { const obj={...r}; delete obj._row; return obj; }
function safeCell_(v) { return typeof v==='string' && /^[=+@\-\t\r]/.test(v) ? '\''+v : v; }
function append_(name,objs) {
  if(!objs.length) return;
  const sh=sheet_(name), headers=FS.HEADERS[name];
  const vals=objs.map(r=>headers.map(h=>safeCell_(r[h]===undefined?'':r[h])));
  sh.getRange(sh.getLastRow()+1,1,vals.length,headers.length).setValues(vals);
}
function write_(name,row) { const headers=FS.HEADERS[name]; sheet_(name).getRange(row._row,1,1,headers.length).setValues([headers.map(k=>safeCell_(row[k]===undefined?'':row[k]))]); }
function delete_(name, rows) { const sh=sheet_(name); rows.map(x=>x._row).sort((a,b)=>b-a).forEach(n=>sh.deleteRow(n)); }
function lock_(fn) { const lock=LockService.getScriptLock(); assert_(lock.tryLock(25000),'Server sedang sibuk. Coba lagi.','BUSY'); try { return fn(); } finally { lock.releaseLock(); } }
function take_(arr,id) { const r=arr.find(x=>String(x.id)===String(id)); assert_(r,'Data tidak ditemukan.','NOT_FOUND'); return r; }
function own_(table,id,userId) { const r=take_(rows_(table).filter(x=>x.userId===userId),id); return r; }
function session_(token) {
  assert_(typeof token==='string' && /^[0-9a-f]{64}$/.test(token),'Silakan login kembali.','UNAUTHORIZED');
  const h=digest_(token), now=Date.now();
  const s=rows_('Sessions').find(x=>equals_(String(x.tokenHash),h) && !x.revokedAt && Date.parse(x.expiresAt)>now);
  assert_(s,'Sesi berakhir. Silakan login kembali.','UNAUTHORIZED');
  const u=rows_('Users').find(x=>x.id===s.userId && bool_(x.verified));
  assert_(u,'Akun tidak tersedia atau belum diverifikasi.','UNAUTHORIZED');
  return u;
}
function me_(u) { return {id:u.id,name:u.name,email:u.email,verified:bool_(u.verified),createdAt:u.createdAt}; }
function rate_(scope,key,max,ttl) { const cache=CacheService.getScriptCache(), name=scope+':'+digest_(key).slice(0,28), n=Number(cache.get(name)||0)+1; cache.put(name,String(n),ttl); assert_(n<=max,'Terlalu banyak percobaan. Silakan coba nanti.','RATE_LIMITED'); }
function hashPassword_(pass,salt,iterations) { return pbkdf2_(pass,salt,iterations); }
function makeCredentials_(pass) { const salt=token_().slice(0,32), iterations=Math.max(100000,Math.min(400000, Number(property_('PBKDF2_ITERATIONS','120000'))||120000)); return {salt,iterations,passwordHash:hashPassword_(pass,salt,iterations)}; }
function newToken_(sheet,userId,expires) { const v=token_(); append_(sheet,[{id:uid_(),userId,tokenHash:digest_(v),expiresAt:new Date(Date.now()+expires).toISOString(),usedAt:'',createdAt:now_()}]); return v; }
function verificationMail_(u,v) {
  const link=appOrigin_()+'/verify.html#token='+encodeURIComponent(v);
  MailApp.sendEmail({to:u.email,subject:'Verifikasi email '+property_('APP_NAME','FolioSpace'),body:'Halo '+u.name+',\n\nVerifikasi akun '+property_('APP_NAME','FolioSpace')+' Anda melalui tautan berikut (berlaku 24 jam):\n'+link+'\n\nJika bukan Anda, abaikan pesan ini.'});
}
function resetMail_(u,v) {
  const link=appOrigin_()+'/index.html#reset='+encodeURIComponent(v);
  MailApp.sendEmail({to:u.email,subject:'Reset password '+property_('APP_NAME','FolioSpace'),body:'Halo '+u.name+',\n\nBuat password baru melalui tautan berikut (berlaku 40 menit):\n'+link+'\n\nJika bukan Anda, abaikan pesan ini.'});
}
function register_(d) {
  const name=str_(d.name,90), em=email_(d.email); validPass_(d.password);
  assert_(name.length>=2,'Nama minimal 2 karakter.'); rate_('register',em,5,3600);
  return lock_(()=>{
    assert_(!rows_('Users').some(x=>x.email===em),'Email sudah terdaftar.','ALREADY_EXISTS');
    const c=makeCredentials_(d.password), id=uid_(), tm=now_(), slug=(name.toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'')||'portfolio')+'-'+id.slice(0,6);
    const u={id,name,email:em,...c,verified:false,createdAt:tm,updatedAt:tm};
    append_('Users',[u]);
    append_('PublicProfiles',[{userId:id,slug,headline:'',bio:'',location:'',avatar:'',socialLinks:'[]',contactEmail:'',isPublic:true,layout:'grid',gridColumns:3,headerStyle:'soft',accent:'#16865D',showDescriptions:true,showCategories:true,showCollections:true,updatedAt:tm}]);
    append_('UserSettings',[{userId:id,layout:'grid',sort:'newest',dateFormat:'id-ID',language:'id',motion:'auto',updatedAt:tm}]);
    try { verificationMail_(u,newToken_('VerificationTokens',id,FS.VERIFY_MS)); }
    catch(e) { console.error('Verification mail error',String(e)); err_('EMAIL_SEND_FAILED','Akun tersimpan, tetapi email belum terkirim. Gunakan kirim ulang verifikasi.'); }
    return {email:em,message:'Cek inbox email untuk verifikasi.'};
  });
}
function verify_(d) { const t=str_(d.token,150), h=digest_(t); assert_(t.length===64,'Token tidak valid.');
  return lock_(()=>{ const x=rows_('VerificationTokens').find(r=>equals_(r.tokenHash,h)&&!r.usedAt); assert_(x && Date.parse(x.expiresAt)>Date.now(),'Tautan verifikasi tidak valid atau kedaluwarsa.','TOKEN_EXPIRED'); const u=rows_('Users').find(r=>r.id===x.userId); assert_(u,'Akun tidak ditemukan.'); x.usedAt=now_(); write_('VerificationTokens',x); u.verified=true;u.updatedAt=now_();write_('Users',u); return {email:u.email}; });
}
function resend_(d) { const em=email_(d.email);rate_('resend',em,3,3600);const u=rows_('Users').find(r=>r.email===em&&!bool_(r.verified));if(u) verificationMail_(u,newToken_('VerificationTokens',u.id,FS.VERIFY_MS));return {message:'Jika akun membutuhkan verifikasi, email akan dikirim.'}; }
function login_(d) { const em=email_(d.email);rate_('login',em,12,900);const u=rows_('Users').find(x=>x.email===em);assert_(u,'Email atau password salah.','INVALID_CREDENTIALS');
  assert_(equals_(hashPassword_(String(d.password||''),String(u.salt),Number(u.iterations)),String(u.passwordHash)),'Email atau password salah.','INVALID_CREDENTIALS');
  assert_(bool_(u.verified),'Verifikasi email terlebih dahulu.','EMAIL_NOT_VERIFIED');
  const t=token_(); append_('Sessions',[{id:uid_(),userId:u.id,tokenHash:digest_(t),expiresAt:new Date(Date.now()+FS.SESSION_MS).toISOString(),revokedAt:'',createdAt:now_()}]);return {session:t,user:me_(u)};
}
function forgot_(d) { const em=email_(d.email);rate_('forgot',em,4,3600);const u=rows_('Users').find(x=>x.email===em&&bool_(x.verified));if(u) {try { resetMail_(u,newToken_('PasswordResetTokens',u.id,FS.RESET_MS)); }catch(e){console.error('Reset mail error',String(e));}}return {message:'Jika email terdaftar, tautan reset akan dikirim.'}; }
function reset_(d) { const token=str_(d.token,150);validPass_(d.password);assert_(token.length===64,'Token invalid.'); const h=digest_(token);return lock_(()=>{const r=rows_('PasswordResetTokens').find(x=>equals_(x.tokenHash,h)&&!x.usedAt);assert_(r && Date.parse(r.expiresAt)>Date.now(),'Tautan reset kedaluwarsa.','TOKEN_EXPIRED');const u=rows_('Users').find(x=>x.id===r.userId);assert_(u,'Akun tidak ditemukan.');Object.assign(u,makeCredentials_(d.password),{updatedAt:now_()});write_('Users',u);r.usedAt=now_();write_('PasswordResetTokens',r);rows_('Sessions').filter(x=>x.userId===u.id&&!x.revokedAt).forEach(x=>{x.revokedAt=now_();write_('Sessions',x);});return {message:'Password berhasil diperbarui.'};}); }
function logout_(u,session,all) { const hash=digest_(session);rows_('Sessions').filter(x=>x.userId===u.id &&(all||equals_(x.tokenHash,hash))&&!x.revokedAt).forEach(r=>{r.revokedAt=now_();write_('Sessions',r);});return {ok:true}; }
function changePassword_(u,d) { validPass_(d.password);assert_(equals_(hashPassword_(String(d.oldPassword||''),u.salt,Number(u.iterations)),u.passwordHash),'Password saat ini salah.','INVALID_CREDENTIALS');return lock_(()=>{Object.assign(u,makeCredentials_(d.password),{updatedAt:now_()});write_('Users',u);return {ok:true};}); }
function portfolioInput_(d,u,existing,categorySet) {
  const title=str_(d.title,150);assert_(title.length>=2,'Judul minimal 2 karakter.');
  const url=safeUrl_(d.url), cat=d.categoryId?String(d.categoryId):'';
  if(cat) assert_(categorySet ? categorySet.has(cat) : rows_('Categories').some(x=>x.userId===u.id&&x.id===cat),'Kategori tidak ditemukan.');
  const tags=(Array.isArray(d.tags)?d.tags:String(d.tags||'').split(',')).map(x=>str_(x,30)).filter(Boolean).slice(0,12);
  const status=['published','draft','archived'].includes(d.status)?d.status:'draft';
  const visibility=d.visibility==='private'?'private':'public';
  return {title,description:str_(d.description,1600),url,categoryId:cat,type:str_(d.type,65),platform:str_(d.platform,65),thumbnail:d.thumbnail?safeUrl_(d.thumbnail):'',tags:JSON.stringify(tags),status,visibility,featured:bool_(d.featured),displayOrder:Number.isFinite(Number(d.displayOrder))?Number(d.displayOrder):existing&&existing.displayOrder||Date.now(),updatedAt:now_()};
}
function portfolios_(u) { return rows_('Portfolios').filter(x=>x.userId===u.id).map(x=>({ ...publicRow_(x),tags:parse_(x.tags,[]),featured:bool_(x.featured)})); }
function portfolioCreate_(u,d) { return lock_(()=>{const list=rows_('Portfolios').filter(x=>x.userId===u.id);assert_(list.length<FS.MAX_PORTFOLIOS,'Batas portofolio tercapai.');const row={id:uid_(),userId:u.id,...portfolioInput_(d,u),createdAt:now_()};append_('Portfolios',[row]);return {...publicRow_(row),tags:parse_(row.tags,[])};}); }
function portfolioUpdate_(u,d) { return lock_(()=>{const row=own_('Portfolios',d.id,u.id);Object.assign(row,portfolioInput_({...row,...d},u,row));write_('Portfolios',row);return {...publicRow_(row),tags:parse_(row.tags,[])};}); }
function portfolioDelete_(u,id) { return lock_(()=>{const row=own_('Portfolios',id,u.id);delete_('CollectionItems',rows_('CollectionItems').filter(x=>x.userId===u.id&&x.portfolioId===id));delete_('Portfolios',[row]);return {id};}); }
function portfolioReorder_(u,d) {return lock_(()=>{assert_(Array.isArray(d.ids)&&d.ids.length<=FS.MAX_PORTFOLIOS,'Urutan tidak valid.');const list=rows_('Portfolios').filter(x=>x.userId===u.id);assert_(list.length===d.ids.length&&new Set(d.ids).size===d.ids.length&&list.every(x=>d.ids.includes(x.id)),'Urutan harus memuat setiap portofolio akun tepat satu kali.');list.forEach(x=>{x.displayOrder=d.ids.indexOf(x.id);x.updatedAt=now_();write_('Portfolios',x);});return {ok:true};});}
function portfolioBulk_(u,d) { assert_(Array.isArray(d.ids)&&d.ids.length<=100&&d.ids.length>0,'Pilih maksimal 100 portofolio.');assert_(['delete','archive','restore','publish','feature','unfeature'].includes(d.action),'Aksi tidak valid.');return lock_(()=>{
  const targets=rows_('Portfolios').filter(x=>x.userId===u.id&&d.ids.includes(x.id));assert_(targets.length===d.ids.length,'Ada portofolio yang tidak ditemukan.');
  if(d.action==='delete') {const ids=new Set(d.ids);delete_('CollectionItems',rows_('CollectionItems').filter(x=>x.userId===u.id&&ids.has(x.portfolioId)));delete_('Portfolios',targets);}
  else targets.forEach(x=>{if(d.action==='archive')x.status='archived';if(d.action==='restore')x.status='draft';if(d.action==='publish')x.status='published';if(d.action==='feature')x.featured=true;if(d.action==='unfeature')x.featured=false;x.updatedAt=now_();write_('Portfolios',x);});
  return {count:targets.length};}); }
function import_(u,d) { assert_(Array.isArray(d.items)&&d.items.length>0&&d.items.length<=FS.MAX_IMPORT,'Import maksimal 150 item per batch.');return lock_(()=>{
  const current=rows_('Portfolios').filter(x=>x.userId===u.id), existing=new Set(current.map(x=>String(x.url).toLowerCase()));assert_(current.length+d.items.length<=FS.MAX_PORTFOLIOS,'Batas portofolio tercapai.');
  const items=[],errors=[],categorySet=new Set(rows_('Categories').filter(x=>x.userId===u.id).map(x=>x.id));d.items.forEach((x,i)=>{try{const input=portfolioInput_(x,u,null,categorySet);if(existing.has(input.url.toLowerCase())){errors.push({row:i+1,error:'URL duplikat'});return;}existing.add(input.url.toLowerCase());items.push({id:uid_(),userId:u.id,...input,createdAt:now_()});}catch(e){errors.push({row:i+1,error:e.message});}});append_('Portfolios',items);return {imported:items.length,errors};}); }
function categoryInput_(d) {const name=str_(d.name,70);assert_(name.length>=2,'Nama kategori minimal 2 karakter.');const color=/^#[0-9a-fA-F]{6}$/.test(d.color)?d.color:'#42B883';return {name,color,icon:str_(d.icon,40)||'folder',updatedAt:now_()}; }
function categoryCreate_(u,d) {return lock_(()=>{const input=categoryInput_(d);assert_(!rows_('Categories').some(x=>x.userId===u.id&&x.name.toLowerCase()===input.name.toLowerCase()),'Kategori sudah ada.');const row={id:uid_(),userId:u.id,...input,displayOrder:Date.now(),createdAt:now_()};append_('Categories',[row]);return publicRow_(row);}); }
function categoryUpdate_(u,d) {return lock_(()=>{const row=own_('Categories',d.id,u.id);Object.assign(row,categoryInput_(d));write_('Categories',row);return publicRow_(row);}); }
function categoryDelete_(u,id) {return lock_(()=>{const row=own_('Categories',id,u.id);rows_('Portfolios').filter(x=>x.userId===u.id&&x.categoryId===id).forEach(x=>{x.categoryId='';write_('Portfolios',x);});delete_('Categories',[row]);return {id};}); }
function collectionCreate_(u,d) {return lock_(()=>{const name=str_(d.name,110);assert_(name.length>=2,'Nama koleksi minimal 2 karakter.');const row={id:uid_(),userId:u.id,name,description:str_(d.description,350),visibility:d.visibility==='private'?'private':'public',displayOrder:Date.now(),createdAt:now_(),updatedAt:now_()};append_('Collections',[row]);return publicRow_(row);}); }
function collectionUpdate_(u,d) {return lock_(()=>{const r=own_('Collections',d.id,u.id);const name=str_(d.name,110);assert_(name.length>=2,'Nama koleksi terlalu pendek.');Object.assign(r,{name,description:str_(d.description,350),visibility:d.visibility==='private'?'private':'public',updatedAt:now_()});write_('Collections',r);return publicRow_(r);}); }
function collectionDelete_(u,id) {return lock_(()=>{const r=own_('Collections',id,u.id);delete_('CollectionItems',rows_('CollectionItems').filter(x=>x.userId===u.id&&x.collectionId===id));delete_('Collections',[r]);return {id};}); }
function collectionItem_(u,d,remove) {return lock_(()=>{own_('Collections',d.collectionId,u.id);own_('Portfolios',d.portfolioId,u.id);const list=rows_('CollectionItems').filter(x=>x.userId===u.id&&x.collectionId===d.collectionId),current=list.find(x=>x.portfolioId===d.portfolioId);if(remove){if(current)delete_('CollectionItems',[current]);}else if(!current){append_('CollectionItems',[{id:uid_(),userId:u.id,collectionId:d.collectionId,portfolioId:d.portfolioId,position:Date.now(),createdAt:now_()}]);}return {ok:true};}); }
function collectionReorder_(u,d) {return lock_(()=>{own_('Collections',d.collectionId,u.id);assert_(Array.isArray(d.ids)&&d.ids.length<=500,'Urutan tidak valid.');const items=rows_('CollectionItems').filter(x=>x.userId===u.id&&x.collectionId===d.collectionId);assert_(items.length===d.ids.length && items.every(x=>d.ids.includes(x.portfolioId)),'Isi koleksi berbeda.');items.forEach(x=>{x.position=d.ids.indexOf(x.portfolioId);write_('CollectionItems',x);});return {ok:true};}); }
function profile_(userId) {const r=rows_('PublicProfiles').find(x=>x.userId===userId);assert_(r,'Profil tidak ditemukan.','NOT_FOUND');return { ...publicRow_(r),socialLinks:parse_(r.socialLinks,[]),isPublic:bool_(r.isPublic),showDescriptions:bool_(r.showDescriptions),showCategories:bool_(r.showCategories),showCollections:bool_(r.showCollections)}; }
function updateProfile_(u,d) {return lock_(()=>{const r=rows_('PublicProfiles').find(x=>x.userId===u.id);assert_(r,'Profil tidak ada.');const slug=str_(d.slug,55).toLowerCase();assert_(/^[a-z0-9](?:[a-z0-9-]{1,53})?[a-z0-9]$/.test(slug),'Slug 3–55 karakter: huruf kecil, angka, dan tanda hubung.');assert_(!rows_('PublicProfiles').some(x=>x.slug===slug&&x.userId!==u.id),'Slug sudah dipakai.','ALREADY_EXISTS');
 const socials=(Array.isArray(d.socialLinks)?d.socialLinks:[]).slice(0,8).map(x=>({label:str_(x.label,35),url:safeUrl_(x.url)}));
 Object.assign(r,{slug,headline:str_(d.headline,140),bio:str_(d.bio,1600),location:str_(d.location,80),avatar:d.avatar?safeUrl_(d.avatar):'',socialLinks:JSON.stringify(socials),contactEmail:d.contactEmail?email_(d.contactEmail):'',isPublic:bool_(d.isPublic),layout:['grid','list'].includes(d.layout)?d.layout:'grid',gridColumns:[2,3,4].includes(Number(d.gridColumns))?Number(d.gridColumns):3,headerStyle:['soft','minimal'].includes(d.headerStyle)?d.headerStyle:'soft',accent:/^#[0-9a-fA-F]{6}$/.test(d.accent)?d.accent:'#16865D',showDescriptions:bool_(d.showDescriptions),showCategories:bool_(d.showCategories),showCollections:bool_(d.showCollections),updatedAt:now_()});write_('PublicProfiles',r);return profile_(u.id);}); }
function publicProfile_(d) { const slug=str_(d.slug,60).toLowerCase();assert_(slug,'Username profil wajib diisi.');const pf=rows_('PublicProfiles').find(x=>x.slug===slug&&bool_(x.isPublic));assert_(pf,'Profil tidak ditemukan atau privat.','NOT_FOUND');const u=rows_('Users').find(x=>x.id===pf.userId && bool_(x.verified));assert_(u,'Profil tidak tersedia.','NOT_FOUND');const portfolios=rows_('Portfolios').filter(x=>x.userId===u.id&&x.status==='published'&&x.visibility==='public').map(x=>({id:x.id,title:x.title,description:x.description,url:x.url,categoryId:x.categoryId,type:x.type,platform:x.platform,thumbnail:x.thumbnail,tags:parse_(x.tags,[]),featured:bool_(x.featured),displayOrder:x.displayOrder,updatedAt:x.updatedAt}));
  const ids=new Set(portfolios.map(x=>x.id)),collections=rows_('Collections').filter(x=>x.userId===u.id&&x.visibility==='public').map(publicRow_),setIds=new Set(collections.map(x=>x.id)),items=rows_('CollectionItems').filter(x=>x.userId===u.id&&setIds.has(x.collectionId)&&ids.has(x.portfolioId)).sort((a,b)=>Number(a.position)-Number(b.position)).map(publicRow_);
  return {user:{name:u.name},profile:profile_(u.id),portfolios,categories:rows_('Categories').filter(x=>x.userId===u.id).map(publicRow_),collections,collectionItems:items}; }
function settings_(u) {const r=rows_('UserSettings').find(x=>x.userId===u.id);return r?publicRow_(r):{};}
function settingsUpdate_(u,d) {return lock_(()=>{const r=rows_('UserSettings').find(x=>x.userId===u.id);assert_(r,'Settings tidak ditemukan.');r.layout=['grid','list'].includes(d.layout)?d.layout:'grid';r.sort=['newest','oldest','az','clicked','custom'].includes(d.sort)?d.sort:'newest';r.dateFormat=['id-ID','en-US'].includes(d.dateFormat)?d.dateFormat:'id-ID';r.language='id';r.motion=['auto','reduce'].includes(d.motion)?d.motion:'auto';r.updatedAt=now_();write_('UserSettings',r);if(d.name){u.name=str_(d.name,90);u.updatedAt=now_();write_('Users',u);}return {settings:publicRow_(r),user:me_(u)};}); }
function track_(d) {
  const type=d.type;assert_(['page_view','portfolio_view','external_click'].includes(type),'Event tidak valid.');const slug=str_(d.slug,60).toLowerCase();assert_(slug,'Profil diperlukan.');
  const pf=rows_('PublicProfiles').find(x=>x.slug===slug&&bool_(x.isPublic));assert_(pf,'Profil tidak ditemukan.','NOT_FOUND');
  const p=d.portfolioId?rows_('Portfolios').find(x=>x.id===d.portfolioId&&x.userId===pf.userId&&x.status==='published'&&x.visibility==='public'):null;
  if(type!=='page_view')assert_(p,'Portofolio tidak tersedia.');
  const visitor=str_(d.visitor,64);assert_(/^[a-zA-Z0-9_-]{16,64}$/.test(visitor),'Visitor token tidak valid.');
  const key=digest_(visitor+slug+(p?p.id:'')+type);const cache=CacheService.getScriptCache();if(cache.get(key))return {recorded:false};
  cache.put(key,'1',type==='page_view'?300:15);
  // Sheets is not an analytics warehouse: keep a bounded rolling log.
  const sh=sheet_('AnalyticsEvents');if(sh.getLastRow()>FS.MAX_EVENTS+1) sh.deleteRows(2, Math.min(500,sh.getLastRow()-FS.MAX_EVENTS-1));
  append_('AnalyticsEvents',[{id:uid_(),userId:pf.userId,portfolioId:p?p.id:'',collectionId:'',type,at:now_(),visitorHash:key.slice(0,24)}]);return {recorded:true};
}
function analytics_(u) {const events=rows_('AnalyticsEvents').filter(x=>x.userId===u.id);const views=events.filter(x=>x.type==='page_view').length,clicks=events.filter(x=>x.type==='external_click').length,details=events.filter(x=>x.type==='portfolio_view').length;
  const per={},byDay={};events.forEach(e=>{if(e.portfolioId){per[e.portfolioId]=per[e.portfolioId]||{views:0,clicks:0};if(e.type==='portfolio_view')per[e.portfolioId].views++;if(e.type==='external_click')per[e.portfolioId].clicks++;}const day=String(e.at).slice(0,10);byDay[day]=byDay[day]||{views:0,clicks:0};if(e.type==='page_view')byDay[day].views++;if(e.type==='external_click')byDay[day].clicks++;});const portfolioRows=rows_('Portfolios').filter(p=>p.userId===u.id);const byCategory={};portfolioRows.forEach(p=>{const key=p.categoryId||'uncategorized';byCategory[key]=byCategory[key]||{views:0,clicks:0};byCategory[key].views+=(per[p.id]?.views||0);byCategory[key].clicks+=(per[p.id]?.clicks||0);});const byCollection={};const members=rows_('CollectionItems').filter(x=>x.userId===u.id);members.forEach(item=>{byCollection[item.collectionId]=byCollection[item.collectionId]||{views:0,clicks:0};byCollection[item.collectionId].views+=(per[item.portfolioId]?.views||0);byCollection[item.collectionId].clicks+=(per[item.portfolioId]?.clicks||0);});return {pageViews:views,detailViews:details,linkClicks:clicks,perPortfolio:per,byCategory,byCollection,byDay:Object.keys(byDay).sort().slice(-30).map(date=>({date,...byDay[date]})),totalEvents:events.length}; }
function bootstrap_(u) {return {user:me_(u),portfolios:portfolios_(u),categories:rows_('Categories').filter(x=>x.userId===u.id).map(publicRow_),collections:rows_('Collections').filter(x=>x.userId===u.id).map(publicRow_),collectionItems:rows_('CollectionItems').filter(x=>x.userId===u.id).map(publicRow_),profile:profile_(u.id),settings:settings_(u),analytics:analytics_(u)};}
function deleteAccount_(u,d) {assert_(String(d.confirm||'')==='DELETE','Masukkan DELETE untuk mengonfirmasi.');return lock_(()=>{Object.keys(FS.HEADERS).forEach(table=>{const field=table==='Users'?'id':'userId';delete_(table,rows_(table).filter(x=>x[field]===u.id));});return {deleted:true};}); }
function dispatch_(op,d,session) {
  switch(op) {
    case 'app.ping': return {version:FS.VERSION,time:now_()};
    case 'auth.register': return register_(d);
    case 'auth.verify': return verify_(d);
    case 'auth.resendVerification': return resend_(d);
    case 'auth.login': return login_(d);
    case 'auth.forgotPassword': return forgot_(d);
    case 'auth.resetPassword': return reset_(d);
    case 'public.getProfile': return publicProfile_(d);
    case 'analytics.track': return track_(d);
  }
  const u=session_(session);
  switch(op) {
    case 'auth.me': return me_(u);
    case 'auth.logout': return logout_(u,session,false);
    case 'auth.logoutAll': return logout_(u,session,true);
    case 'auth.deleteAccount': return deleteAccount_(u,d);
    case 'dashboard.bootstrap': return bootstrap_(u);
    case 'portfolio.list': return portfolios_(u);
    case 'portfolio.get': return publicRow_(own_('Portfolios',d.id,u.id));
    case 'portfolio.create': return portfolioCreate_(u,d);
    case 'portfolio.update': return portfolioUpdate_(u,d);
    case 'portfolio.delete': return portfolioDelete_(u,d.id);
    case 'portfolio.archive': return portfolioBulk_(u,{ids:[d.id],action:'archive'});
    case 'portfolio.restore': return portfolioBulk_(u,{ids:[d.id],action:'restore'});
    case 'portfolio.bulk': return portfolioBulk_(u,d);
    case 'portfolio.import': return import_(u,d);
    case 'portfolio.reorder': return portfolioReorder_(u,d);
    case 'category.list': return rows_('Categories').filter(x=>x.userId===u.id).map(publicRow_);
    case 'category.create': return categoryCreate_(u,d);
    case 'category.update': return categoryUpdate_(u,d);
    case 'category.delete': return categoryDelete_(u,d.id);
    case 'collection.list': return rows_('Collections').filter(x=>x.userId===u.id).map(publicRow_);
    case 'collection.create': return collectionCreate_(u,d);
    case 'collection.update': return collectionUpdate_(u,d);
    case 'collection.delete': return collectionDelete_(u,d.id);
    case 'collection.addPortfolio': return collectionItem_(u,d,false);
    case 'collection.removePortfolio': return collectionItem_(u,d,true);
    case 'collection.reorder': return collectionReorder_(u,d);
    case 'public.updateProfile': return updateProfile_(u,d);
    case 'analytics.summary': return analytics_(u);
    case 'settings.get': return settings_(u);
    case 'settings.update': return settingsUpdate_(u,d);
    case 'settings.changePassword': return changePassword_(u,d);
    default: err_('UNKNOWN_OPERATION','Operasi tidak dikenal: '+str_(op,80));
  }
}
/** Portable PBKDF2-HMAC-SHA256 implementation. Runtime/iteration count tradeoff:
 * benchmark in GAS and use a dedicated identity provider for truly high-scale or sensitive apps.
 * No bcrypt/Argon2 primitives are built into Apps Script.
 */
const SHA_K=[0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
function bytes_(s) {return Utilities.newBlob(String(s)).getBytes().map(x=>x&255);}
function sha256_(arr) {
  const data=arr.slice(), len=data.length, bit=len*8;data.push(0x80);while(data.length%64!==56)data.push(0);data.push(0,0,0,0,(bit/16777216)&255,(bit>>>16)&255,(bit>>>8)&255,bit&255);
  const H=[0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19], W=new Int32Array(64);
  const R=(x,n)=>(x>>>n)|(x<<(32-n));
  for(let k=0;k<data.length;k+=64){for(let i=0;i<16;i++)W[i]=(data[k+4*i]<<24)|(data[k+4*i+1]<<16)|(data[k+4*i+2]<<8)|data[k+4*i+3];
    for(let i=16;i<64;i++){const a=W[i-15],b=W[i-2];W[i]=(W[i-16]+(R(a,7)^R(a,18)^(a>>>3))+W[i-7]+(R(b,17)^R(b,19)^(b>>>10)))|0;}
    let [a,b,c,d,e,f,g,h]=H;
    for(let i=0;i<64;i++){const t1=(h+(R(e,6)^R(e,11)^R(e,25))+((e&f)^(~e&g))+SHA_K[i]+W[i])|0,t2=((R(a,2)^R(a,13)^R(a,22))+((a&b)^(a&c)^(b&c)))|0;
      h=g;g=f;f=e;e=(d+t1)|0;d=c;c=b;b=a;a=(t1+t2)|0;}
    H[0]=(H[0]+a)|0;H[1]=(H[1]+b)|0;H[2]=(H[2]+c)|0;H[3]=(H[3]+d)|0;H[4]=(H[4]+e)|0;H[5]=(H[5]+f)|0;H[6]=(H[6]+g)|0;H[7]=(H[7]+h)|0;
  }
  const out=[];for(let x of H)out.push((x>>>24)&255,(x>>>16)&255,(x>>>8)&255,x&255);return out;
}
function hmac_(key,message) {let k=key.length>64?sha256_(key):key.slice();while(k.length<64)k.push(0);const inner=k.map(x=>x^0x36),outer=k.map(x=>x^0x5c);return sha256_(outer.concat(sha256_(inner.concat(message))));}
function pbkdf2_(password,salt,iterations) {
  const key=bytes_(password), salted=bytes_(salt).concat([0,0,0,1]);let u=hmac_(key,salted),out=u.slice();
  for(let i=1;i<iterations;i++){u=hmac_(key,u);for(let j=0;j<32;j++)out[j]^=u[j];}
  return out.map(v=>('0'+v.toString(16)).slice(-2)).join('');
}
/** Optional quick backend self-check; run manually after setupDatabase(). */
function testBackend() {
  const pbkdf=pbkdf2_('password','salt',1);
  assert_(pbkdf==='120fb6cffcf8b32c43e7225256c4f837a86548c92ccc35480805987cb70be17b','PBKDF2 self-test failed.');
  const origin=appOrigin_();assert_(!!origin,'APP_ORIGIN belum diatur.');
  return {ok:true,sheets:Object.keys(FS.HEADERS).length,hashTest:pbkdf.slice(0,10),origin};
}
