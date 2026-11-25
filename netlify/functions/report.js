const XLSX = require('xlsx');
const guard = require('./lib/guard');

// ============================================================
//  محرّك استعلام التقرير الشهري لولي الأمر — آمن (بحث في الخادم فقط)
// ============================================================
const SHARE_URL = process.env.REPORT_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQDWatQFXbUpTZuWcfKbG_0tAbAKwwfqofXVFUrnWnwiNwI?e=fZVhNi";

const SHEETS = [
  // مواد المجموعة A — المواد الأساسية
  ['report_islamic','التربية الإسلامية','IS','🕌','#1F6E4E','A'],
  ['report_arabic','اللغة العربية','AR','📖','#0E5A54','A'],
  ['report_english','اللغة الإنجليزية','E','🔤','#2A6F97','A'],
  ['report_social','الدراسات الاجتماعية','SS','🌍','#3A5A8C','A'],
  ['report_math','الرياضيات','MA','➗','#4B4B8F','A'],
  ['report_science','العلوم','SC','🔬','#2E7D6B','A'],
  // مواد المجموعة B — النشاط والمهارات
  ['report_ai','الذكاء الاصطناعي','AI','🤖','#6B4F9E','B'],
  ['report_pe','التربية البدنية','PE','⚽','#4F7A3A','B'],
  ['report_aa','الفنون السمعية','AA','🎵','#4A6D8C','B'],
  ['report_va','الفنون البصرية','VA','🎨','#B5546A','B'],
  ['report_dr','المسرح','DR','🎭','#7A4A78','B'],
];

const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
const norm=s=>clean(s).toLowerCase();
const LAT=s=>String(s==null?'':s).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
const digits=s=>LAT(s).replace(/\D/g,''); // أرقام فقط (لرقم الهوية) — تُقبل الأرقام العربية أيضاً
const cell=(ws,r,c)=>{const x=ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:null;};

// يكتشف صف الترويسة تلقائيًا (يتحمّل صفوفًا فارغة أو تعليمات في الأعلى)
function locateHeader(ws){
  const R=XLSX.utils.decode_range(ws['!ref']);
  for(let r=1;r<=Math.min(R.e.r+1,8);r++){
    let hasId=false, hasLevel=false; const H={};
    for(let c=1;c<=R.e.c+1;c++){ const h=norm(cell(ws,r,c));
      if(/uaeid|هوية/.test(h)){H.id=c;hasId=true;}
      else if(/اسم/.test(h)) H.name=c;
      else if(/شعبة/.test(h)) H.section=c;
      else if(/رقم\s*الطالب|studentid/i.test(h)) H.studentNo=c;
      else if(/درجات|اختبار/.test(h)) H.examScores=c;
      else if(/تقرير/.test(h)) H.rnum=c;
      else if(/مستوى/.test(h)){H.level=c;hasLevel=true;}
      else if(/قوة/.test(h)) H.strengths=c;
      else if(/تحسين/.test(h)) H.improve=c;
      else if(/ملاحظ/.test(h)) H.notes=c;
    }
    if(hasId && hasLevel) return {R, hr:r, H};
  }
  return {R, hr:-1, H:{}};
}

function parseSheet(ws){
  if(!ws || !ws['!ref']) return [];
  const {R,hr,H}=locateHeader(ws);
  if(hr<0 || !H.id) return [];
  const out=[];
  for(let r=hr+1;r<=R.e.r+1;r++){
    const id=digits(cell(ws,r,H.id));
    if(!id) continue;
    const name=clean(cell(ws,r,H.name));
    const notes=H.notes?clean(cell(ws,r,H.notes)):'';
    if(/مثال توضيحي/.test(notes)) continue; // تجاهل صف المثال إن نُسي
    out.push({
      id, name,
      section: H.section?clean(cell(ws,r,H.section)):'',
      studentNo: H.studentNo?clean(cell(ws,r,H.studentNo)):'',
      examScores: H.examScores?clean(cell(ws,r,H.examScores)):'',
      reportNum: H.rnum?(clean(cell(ws,r,H.rnum))||'1'):'1',
      level: H.level?clean(cell(ws,r,H.level)):'',
      strengths: H.strengths?clean(cell(ws,r,H.strengths)):'',
      improvement: H.improve?clean(cell(ws,r,H.improve)):'',
      notes: notes && !/مثال توضيحي/.test(notes) ? notes : '',
    });
  }
  return out;
}

function lookup(buf, idInput, nameInput){
  const wb=XLSX.read(buf,{type:'buffer'});
  const idQ=digits(idInput), nameQ=norm(nameInput);
  if(idQ.length<10 || !nameQ) return {status:'invalid'};
  let found=false, nameOk=false, studentName='', section='', studentNo='';
  const items=[];
  for(const [sheetKey, label, code, emoji, color, group] of SHEETS){
    const rows=parseSheet(wb.Sheets[sheetKey]);
    for(const row of rows){
      if(row.id!==idQ) continue;
      found=true;
      const first=norm((row.name||'').split(/\s+/)[0]||'');
      if(first!==nameQ) continue;
      nameOk=true; studentName=row.name; section=row.section||section; studentNo=row.studentNo||studentNo;
      items.push({subject:label, code, emoji, color, group, reportNum:row.reportNum,
        examScores:row.examScores, level:row.level, strengths:row.strengths, improvement:row.improvement, notes:row.notes});
    }
  }
  if(!found||!nameOk) return {status:'nomatch'};   // رد واحد لأي فشل، فلا يُعرف هل الرقم موجود
  return {status:'ok', data:{name:studentName, section, studentNo, items}};
}

// تنزيل + كاش + حدّ محاولات (نفس نمط محرّك ألف)
let CACHE={buf:null, at:0};
const UA={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36','Accept':'*/*'};
const isZip=b=>b&&b.length>3&&b[0]===0x50&&b[1]===0x4B;
function shareToken(u){return 'u!'+Buffer.from(u,'utf8').toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');}
async function tryFetch(u){try{const r=await fetch(u,{redirect:'follow',headers:UA});const b=Buffer.from(await r.arrayBuffer());return isZip(b)?b:null;}catch(e){return null;}}
async function download(){
  if(CACHE.buf && (Date.now()-CACHE.at)<300000) return CACHE.buf;
  const m=SHARE_URL.match(/^(https:\/\/[^/]+)\/:[a-z]:\/[a-z]\/(personal\/[^/]+)\/([^/?#]+)/i);
  const urls=[];
  if(m) urls.push(m[1]+'/'+m[2]+'/_layouts/15/download.aspx?share='+m[3]);
  urls.push(SHARE_URL+(SHARE_URL.includes('?')?'&':'?')+'download=1');
  urls.push('https://api.onedrive.com/v1.0/shares/'+shareToken(SHARE_URL)+'/root/content');
  for(const u of urls){const b=await tryFetch(u); if(b){CACHE={buf:b,at:Date.now()}; return b;}}
  return null;
}
function ipOf(event){ const h=(event&&event.headers)||{}; return (h['x-nf-client-connection-ip']||h['client-ip']||h['x-forwarded-for']||'0').split(',')[0].trim(); }

const J=(code,obj)=>({statusCode:code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify(obj)});
function countRows(buf){ const wb=XLSX.read(buf,{type:'buffer'}); let n=0; for(const [sheetKey] of SHEETS){ try{ n+=parseSheet(wb.Sheets[sheetKey]).length; }catch(e){} } return n; }

async function handle(event, deps){
  const method=(event&&event.httpMethod)||'GET', hd=(event&&event.headers)||{}, ip=ipOf(event), ua=/mobile|android|iphone|ipad/i.test(hd['user-agent']||'')?'m':'d';
  let q;
  if(method==='POST'){ try{ q=JSON.parse(event.body||'{}'); }catch(e){ return J(400,{status:'invalid'}); } } else q=(event&&event.queryStringParameters)||{};
  try{
    if(method==='GET' && q.ping){ const buf=await deps.download(); if(!buf) return J(502,{status:'nofile'}); return J(200,{status:'ok',rows:countRows(buf)}); }   // للمراقبة: عدد الصفوف فقط
    if(method!=='POST') return J(405,{status:'method'});   // لا استعلام بالرابط: الرقم والاسم لا يُرسَلان إلا في جسم POST
    const idD=digits(q.id), name=clean(q.name);
    if(idD.length<10 || !name) return J(400,{status:'invalid'});
    const g=await guard.check(deps.store, idD, ip);
    const base={k:'report',id:guard.mask(idD),idh:g.idh.slice(0,6),iph:g.iph.slice(0,6),m:method,ua};
    if(!g.ok){ await guard.log(deps.store,Object.assign({r:g.reason==='locked'?'locked':'rate'},base)); return J(429,{status:g.reason}); }
    const buf=await deps.download(); if(!buf) return J(502,{status:'nofile'});
    const res=(deps.lookup||lookup)(buf, idD, name);
    if(res.status==='ok'){ await guard.log(deps.store,Object.assign({r:'ok'},base)); return J(200,res); }
    if(res.status==='nomatch'){ await guard.fail(deps.store,g); await guard.log(deps.store,Object.assign({r:'fail'},base)); return J(404,{status:'nomatch'}); }
    return J(res.status==='invalid'?400:502,{status:res.status});
  }catch(e){ return J(500,{status:'error'}); }
}
exports.handler = async (event) => handle(event,{download, store:guard.open(event)});
// قائمة الطلاب (رقم الهوية + الاسم + الشعبة) من ملف التقرير — لتسهيل إصدار الشهادات من لوحة الإدارة
function roster(buf){
  const wb=XLSX.read(buf,{type:'buffer'}), by=new Map();
  for(const [sheetKey] of SHEETS){
    for(const row of parseSheet(wb.Sheets[sheetKey])){
      if(row.id.length<10||!row.name) continue;
      const cur=by.get(row.id); if(!cur) by.set(row.id,{id:row.id,name:row.name,section:row.section||''}); else if(!cur.section&&row.section) cur.section=row.section;
    }
  }
  return [...by.values()].sort((a,b)=>a.name.localeCompare(b.name,'ar'));
}
exports._handle=handle; exports._lookup=lookup; exports._download=download; exports._roster=roster;
