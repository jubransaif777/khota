const XLSX = require('xlsx');

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
const digits=s=>String(s==null?'':s).replace(/\D/g,''); // أرقام فقط (لرقم الهوية)

// كاش بسيط داخل الذاكرة (تنزيل مرة كل 5 دقائق) + حدّ للمحاولات
let CACHE={buf:null, at:0};
const RATE={}; // ip -> {count, resetAt}

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
    }
    if(tmp.id && tmp.email){ hr=r; H=tmp; break; }
  }
  return {R, hr, H};
}

function lookup(buf, idInput, nameInput){
  const wb=XLSX.read(buf,{type:'buffer'});
  const ws=wb.Sheets[wb.SheetNames[0]];
  if(!ws || !ws['!ref']) return {status:'nofile'};
  const {R,hr,H}=locate(ws);
  if(hr<0 || !H.id || !H.email || !H.pass){ return {status:'badformat'}; }
  const idQ=digits(idInput);
  const nameQ=norm(nameInput);
  if(idQ.length<10 || !nameQ) return {status:'invalid'};
  for(let r=hr+1;r<=R.e.r+1;r++){
    const id=digits(cell(ws,r,H.id));
    if(id!==idQ) continue;
    // تطابق الرقم — تحقّق من الاسم الأول
    const first=norm(cell(ws,r,H.first));
    if(H.first && first!==nameQ){ return {status:'namemismatch'}; }
    return {status:'ok', data:{
      name: clean(cell(ws,r,H.first)),
      email: clean(cell(ws,r,H.email)),
      password: clean(cell(ws,r,H.pass))
    }};
  }
  return {status:'notfound'};
}

function ipOf(event){
  const h=(event&&event.headers)||{};
  return (h['x-nf-client-connection-ip']||h['client-ip']||h['x-forwarded-for']||'0').split(',')[0].trim();
}
function limited(ip){
  const now=Date.now(); const rec=RATE[ip];
  if(!rec || now>rec.resetAt){ RATE[ip]={count:1, resetAt:now+3600000}; return false; }
  rec.count++; return rec.count>20; // 20 محاولة/ساعة
}

exports.handler = async (event) => {
  const q=(event&&event.queryStringParameters)||{};
  const id=q.id||'', name=q.name||'';
  const J=(code,obj)=>({statusCode:code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify(obj)});
  if(!id && !name) return J(400,{status:'invalid'});
  if(limited(ipOf(event))) return J(429,{status:'ratelimited'});
  try{
    const buf=await download();
    if(!buf) return J(502,{status:'nofile'});
    const res=lookup(buf, id, name);
    const map={ok:200,notfound:404,namemismatch:404,invalid:400,badformat:500,nofile:502};
    return J(map[res.status]||400, res.status==='ok'?{status:'ok',data:res.data}:{status:res.status});
  }catch(e){ return J(500,{status:'error'}); }
};
exports._lookup=lookup;
