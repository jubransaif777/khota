const XLSX = require('xlsx');

// ============================================================
//  محرّك صفحة تعاون المعلمين — محمي بكلمة سرّ خاصة به (منفصلة عن /admin)
// ============================================================
const SHARE_URL = process.env.REPORT_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQDWatQFXbUpTZuWcfKbG_0tAbAKwwfqofXVFUrnWnwiNwI?e=fZVhNi";
const STAFF_PASSWORD = process.env.REPORT_STAFF_PASSWORD || "";

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
const digits=s=>String(s==null?'':s).replace(/\D/g,'');
const cell=(ws,r,c)=>{const x=ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:null;};

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
    const notes=H.notes?clean(cell(ws,r,H.notes)):'';
    if(/مثال توضيحي/.test(notes)) continue;
    out.push({
      id, name:clean(cell(ws,r,H.name)),
      section: H.section?clean(cell(ws,r,H.section)):'',
      studentNo: H.studentNo?clean(cell(ws,r,H.studentNo)):'',
      examScores: H.examScores?clean(cell(ws,r,H.examScores)):'',
      reportNum: H.rnum?(parseInt(clean(cell(ws,r,H.rnum)),10)||1):1,
      level: H.level?clean(cell(ws,r,H.level)):'',
      strengths: H.strengths?clean(cell(ws,r,H.strengths)):'',
      improvement: H.improve?clean(cell(ws,r,H.improve)):'',
      notes: notes,
    });
  }
  return out;
}

// يبني: قائمة الشعب + لكل شعبة قائمة طلاب، كل طالب فيه بيانات آخر تقرير لكل مادة
function build(buf){
  const wb=XLSX.read(buf,{type:'buffer'});
  const students={}; // id -> {id,name,section, subjects:{code:{...,reportNum}}}
  for(const [sheetKey, label, code, emoji, color, group] of SHEETS){
    const rows=parseSheet(wb.Sheets[sheetKey]);
    for(const row of rows){
      if(!students[row.id]) students[row.id]={id:row.id,name:row.name,section:row.section,studentNo:'',subjects:{}};
      const s=students[row.id];
      if(row.name) s.name=row.name; if(row.section) s.section=row.section; if(row.studentNo) s.studentNo=row.studentNo;
      const cur=s.subjects[code];
      if(!cur || row.reportNum>=cur.reportNum){
        s.subjects[code]={label,emoji,color,group,reportNum:row.reportNum,
          examScores:row.examScores, level:row.level, strengths:row.strengths, improvement:row.improvement, notes:row.notes};
      }
    }
  }
  const list=Object.values(students);
  const sections=[...new Set(list.map(s=>s.section).filter(Boolean))].sort();
  const bySection={}; sections.forEach(sec=>{ bySection[sec]=list.filter(s=>s.section===sec).sort((a,b)=>a.name.localeCompare(b.name,'ar')); });
  return {subjectOrder:SHEETS.map(([,label,code,emoji,color,group])=>({code,label,emoji,color,group})), sections, bySection};
}

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
  rec.count++; return rec.count>40; }

exports.handler = async (event) => {
  const q=(event&&event.queryStringParameters)||{};
  const J=(code,obj)=>({statusCode:code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify(obj)});
  if(!STAFF_PASSWORD) return J(500,{status:'nopass'});
  if((q.pw||'')!==STAFF_PASSWORD) return J(401,{status:'denied'});
  if(limited(ipOf(event))) return J(429,{status:'ratelimited'});
  try{
    const buf=await download();
    if(!buf) return J(502,{status:'nofile'});
    return J(200,{status:'ok', ...build(buf)});
  }catch(e){ return J(500,{status:'error'}); }
};
exports._build=build;
