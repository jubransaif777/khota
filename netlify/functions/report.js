const XLSX = require('xlsx');

// ============================================================
//  محرّك استعلام التقرير الشهري لولي الأمر — آمن (بحث في الخادم فقط)
// ============================================================
const SHARE_URL = process.env.REPORT_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQDWatQFXbUpTZuWcfKbG_0tAbAKwwfqofXVFUrnWnwiNwI?e=fZVhNi";

const SHEETS = [
  ['report_islamic','التربية الإسلامية','IS','🕌','#1F6E4E'],
  ['report_arabic','اللغة العربية','AR','📖','#0E5A54'],
  ['report_english','اللغة الإنجليزية','E','🔤','#2A6F97'],
  ['report_social','الدراسات الاجتماعية','SS','🌍','#3A5A8C'],
  ['report_math','الرياضيات','MA','➗','#4B4B8F'],
  ['report_science','العلوم','SC','🔬','#2E7D6B'],
];

const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
const norm=s=>clean(s).toLowerCase();
const digits=s=>String(s==null?'':s).replace(/\D/g,'');
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
  for(const [sheetKey, label, code, emoji, color] of SHEETS){
    const rows=parseSheet(wb.Sheets[sheetKey]);
    for(const row of rows){
      if(row.id!==idQ) continue;
      found=true;
      const first=norm((row.name||'').split(/\s+/)[0]||'');
      if(first!==nameQ) continue;
      nameOk=true; studentName=row.name; section=row.section||section; studentNo=row.studentNo||studentNo;
      items.push({subject:label, code, emoji, color, reportNum:row.reportNum,
        level:row.level, strengths:row.strengths, improvement:row.improvement, notes:row.notes});
    }
  }
  if(!found) return {status:'notfound'};
  if(!nameOk) return {status:'namemismatch'};
  return {status:'ok', data:{name:studentName, section, studentNo, items}};
}

// تنزيل + كاش + حدّ محاولات (نفس نمط محرّك ألف)
let CACHE={buf:null, at:0};
const RATE={};
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
function limited(ip){ const now=Date.now(); const rec=RATE[ip];
  if(!rec || now>rec.resetAt){ RATE[ip]={count:1, resetAt:now+3600000}; return false; }
  rec.count++; return rec.count>20; }

exports.handler = async (event) => {
  const q=(event&&event.queryStringParameters)||{};
  const J=(code,obj)=>({statusCode:code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify(obj)});
  if(!q.id && !q.name) return J(400,{status:'invalid'});
  if(limited(ipOf(event))) return J(429,{status:'ratelimited'});
  try{
    const buf=await download();
    if(!buf) return J(502,{status:'nofile'});
    const res=lookup(buf, q.id, q.name);
    const map={ok:200,notfound:404,namemismatch:404,invalid:400};
    return J(map[res.status]||400, res.status==='ok'?res:{status:res.status});
  }catch(e){ return J(500,{status:'error'}); }
};
exports._lookup=lookup;
