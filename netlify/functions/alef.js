const XLSX = require('xlsx');
const guard = require('./lib/guard');

// ============================================================
//  محرّك استعلام بيانات منصة ألف — آمن (البحث في الخادم فقط)
//  غيّر هذا السطر فقط لو تغيّر رابط ملف البيانات:
// ============================================================
const SHARE_URL = process.env.ALEF_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQCZyBQZVBylTJOE9hd0etxHAX-tpgxFTQdC8kTa3VfaXaE?e=j6OPIH";
// ============================================================

const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
const cell=(ws,r,c)=>{const x=ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:null;};
const norm=s=>clean(s).toLowerCase();
const LAT=s=>String(s==null?'':s).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
const digits=s=>LAT(s).replace(/\D/g,''); // أرقام فقط (لرقم الهوية) — تُقبل الأرقام العربية أيضاً

// كاش بسيط داخل الذاكرة (تنزيل مرة كل 5 دقائق) + حدّ للمحاولات
let CACHE={buf:null, at:0};

const UA={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36','Accept':'*/*'};
const isZip=b=>b&&b.length>3&&b[0]===0x50&&b[1]===0x4B;
function shareToken(u){return 'u!'+Buffer.from(u,'utf8').toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');}
async function tryFetch(u){try{const r=await fetch(u,{redirect:'follow',headers:UA});const b=Buffer.from(await r.arrayBuffer());return isZip(b)?b:null;}catch(e){return null;}}
async function download(){
  if(CACHE.buf && (Date.now()-CACHE.at)<300000) return CACHE.buf; // 5 دقائق
  const m=SHARE_URL.match(/^(https:\/\/[^/]+)\/:[a-z]:\/[a-z]\/(personal\/[^/]+)\/([^/?#]+)/i);
  const urls=[];
  if(m) urls.push(m[1]+'/'+m[2]+'/_layouts/15/download.aspx?share='+m[3]);
  urls.push(SHARE_URL+(SHARE_URL.includes('?')?'&':'?')+'download=1');
  urls.push('https://api.onedrive.com/v1.0/shares/'+shareToken(SHARE_URL)+'/root/content');
  for(const u of urls){const b=await tryFetch(u); if(b){CACHE={buf:b,at:Date.now()}; return b;}}
  return null;
}

// يكتشف صف الترويسة وأعمدة UAEID / FirstName_EN / Email / Password / (الاسم للتأكيد)
function locate(ws){
  const R=XLSX.utils.decode_range(ws['!ref']);
  let hr=-1, H={};
  for(let r=1;r<=Math.min(R.e.r+1,8);r++){
    const tmp={};
    for(let c=1;c<=R.e.c+1;c++){ const h=norm(cell(ws,r,c));
      if(/uaeid|هوية|الهوية/.test(h)) tmp.id=c;
      else if(/firstname|first_name|الاسم الأول|first/.test(h)) tmp.first=c;
      else if(/password|كلمة المرور|مرور/.test(h)) tmp.pass=c;
      else if(/email|بريد/.test(h)) tmp.email=c;
      else if(/grade|الصف/.test(h)) tmp.grade=c;
      else if(/section/.test(h)) tmp.section=c;
      else if(/verify|رمز|birth|ميلاد|dob|mobile|phone|جوال|هاتف/.test(h)){ tmp.verify=c; tmp.verifyKind=verifyKind(h); }   // عمود تحقّق ثانٍ اختياري
    }
    if(tmp.id && tmp.email){ hr=r; H=tmp; break; }
  }
  return {R, hr, H};
}

// ---- تحقّق ثانٍ اختياري: يُفعَّل تلقائياً إن أضفتَ عموداً (سنة الميلاد/آخر 4 أرقام الجوال/رمز) في ملف ألف ----
function verifyKind(h){ return /birth|ميلاد|dob/.test(h)?'birth':(/mobile|phone|جوال|هاتف/.test(h)?'phone':'code'); }
const VERIFY_LABEL={birth:'سنة ميلاد الطالب (مثال: 2012)',phone:'آخر 4 أرقام من جوال ولي الأمر',code:'رمز التحقق'};
const latin=s=>String(s==null?'':s).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
function verifyOk(cellVal, input, kind){
  const c=latin(cellVal), v=latin(input);
  if(kind==='birth'){ const y=(c.match(/(19|20)\d{2}/)||[])[0]; const yi=(v.match(/(19|20)\d{2}/)||[])[0]; return !!y&&y===yi; }
  if(kind==='phone'){ const a=c.replace(/\D/g,'').slice(-4), b=v.replace(/\D/g,'').slice(-4); return a.length===4&&a===b; }
  return clean(c).toLowerCase()===clean(v).toLowerCase() && clean(c)!=='';
}
function open(buf){ const wb=XLSX.read(buf,{type:'buffer'}); const ws=wb.Sheets[wb.SheetNames[0]]; if(!ws||!ws['!ref']) return null; const L=locate(ws); return {ws,...L}; }
// معلومات غير حسّاسة: هل يلزم تحقّق ثانٍ؟ وكم صفاً في الملف؟ (للصفحة وللمراقبة)
function info(buf){
  const o=open(buf); if(!o) return {status:'nofile'};
  const {ws,R,hr,H}=o; if(hr<0||!H.id||!H.email||!H.pass) return {status:'badformat'};
  let rows=0; for(let r=hr+1;r<=R.e.r+1;r++) if(digits(cell(ws,r,H.id)).length>=10) rows++;
  return {status:'ok', rows, verify: H.verify?{needed:true,kind:H.verifyKind,label:VERIFY_LABEL[H.verifyKind]}:{needed:false}};
}
// كل فشل (رقم غير موجود / اسم لا يطابق / تحقّق خاطئ) يُرجع «nomatch» نفسه، فلا يُعرف هل الرقم موجود
function lookup(buf, idInput, nameInput, verifyInput){
  const o=open(buf); if(!o) return {status:'nofile'};
  const {ws,R,hr,H}=o;
  if(hr<0 || !H.id || !H.email || !H.pass){ return {status:'badformat'}; }
  const idQ=digits(idInput), nameQ=norm(nameInput);
  if(idQ.length<10 || !nameQ) return {status:'invalid'};
  for(let r=hr+1;r<=R.e.r+1;r++){
    if(digits(cell(ws,r,H.id))!==idQ) continue;
    if(H.first && norm(cell(ws,r,H.first))!==nameQ) return {status:'nomatch'};
    if(H.verify){ const cv=cell(ws,r,H.verify); if(clean(cv) && !verifyOk(cv,verifyInput,H.verifyKind)) return {status:'nomatch'}; }
    return {status:'ok', data:{ name: clean(cell(ws,r,H.first)), email: clean(cell(ws,r,H.email)), password: clean(cell(ws,r,H.pass)) }};
  }
  return {status:'nomatch'};
}

function ipOf(event){
  const h=(event&&event.headers)||{};
  return (h['x-nf-client-connection-ip']||h['client-ip']||h['x-forwarded-for']||'0').split(',')[0].trim();
}
const J=(code,obj)=>({statusCode:code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify(obj)});

async function handle(event, deps){
  const method=(event&&event.httpMethod)||'GET', hd=(event&&event.headers)||{}, ip=ipOf(event), ua=/mobile|android|iphone|ipad/i.test(hd['user-agent']||'')?'m':'d';
  let q;
  if(method==='POST'){ try{ q=JSON.parse(event.body||'{}'); }catch(e){ return J(400,{status:'invalid'}); } } else q=(event&&event.queryStringParameters)||{};
  try{
    // نقطتان بلا بيانات شخصية: info (هل يلزم تحقق ثانٍ) وping (للمراقبة)
    if(method==='GET' && (q.info||q.ping)){
      const buf=await deps.download(); if(!buf) return J(502,{status:'nofile'});
      const i=info(buf); if(i.status!=='ok') return J(i.status==='badformat'?500:502,{status:i.status});
      return J(200, q.ping?{status:'ok',rows:i.rows}:{status:'ok',verify:i.verify});
    }
    if(method!=='POST') return J(405,{status:'method'});   // لا استعلام بالرابط: الرقم والاسم لا يُرسَلان إلا في جسم POST
    const idD=digits(q.id), name=clean(q.name);
    if(idD.length<10 || !name) return J(400,{status:'invalid'});
    const g=await guard.check(deps.store, idD, ip);
    const base={k:'alef',id:guard.mask(idD),idh:g.idh.slice(0,6),iph:g.iph.slice(0,6),m:method,ua};
    if(!g.ok){ await guard.log(deps.store,Object.assign({r:g.reason==='locked'?'locked':'rate'},base)); return J(429,{status:g.reason}); }
    const buf=await deps.download(); if(!buf) return J(502,{status:'nofile'});
    const res=(deps.lookup||lookup)(buf, idD, name, q.v);
    if(res.status==='ok'){ await guard.log(deps.store,Object.assign({r:'ok'},base)); return J(200,{status:'ok',data:res.data}); }
    if(res.status==='nomatch'){ await guard.fail(deps.store,g); await guard.log(deps.store,Object.assign({r:'fail'},base)); return J(404,{status:'nomatch'}); }
    return J(res.status==='invalid'?400:(res.status==='badformat'?500:502),{status:res.status});
  }catch(e){ return J(500,{status:'error'}); }
}
exports.handler = async (event) => handle(event,{download, store:guard.open(event)});
exports._handle=handle; exports._lookup=lookup; exports._info=info;
