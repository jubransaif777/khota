const XLSX = require('xlsx');
const guard = require('./lib/guard');

// ============================================================
//  محرّك حجز الاختبارات — يقرأ ملف الاختبارات (تخطيط الأيام كالأسبوع 4)
//  غيّر هذا السطر فقط لو تغيّر رابط مشاركة ملف الاختبارات:
// ============================================================
const SHARE_URL = process.env.EXAMS_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQBvcMAZLQmlTbzr1iHrcQ-vAbvgRqhsfDYrwJ7fAVNhuc0?e=KzSx9W";
// ============================================================
const SCHOOL="مدرسة المعيريض للحلقة الثانية بنين";

const SECTIONS=['5G1','5A1','5A2','6G1','6G2','6A1','6A2','7G1','7G2','7A1','7A2','8G1','8G2','8A1','8A2'];
const DAYS=[['MON','الاثنين'],['TUE','الثلاثاء'],['WED','الأربعاء'],['THU','الخميس'],['FRI','الجمعة']];
const SUBJ={IS:'إسلامية',AR:'عربي',SS:'دراسات اجتماعية',E:'إنجليزي',MA:'رياضيات',SC:'علوم',
  PE:'تربية بدنية',DT:'تصميم تكنولوجي',DR:'مسرح',VA:'فنية',AA:'موسيقى',AI:'ذكاء اصطناعي'};
const AR_MONTHS=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];

const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();

// مؤشّر الأسبوع الحالي: بتوقيت الإمارات (لا توقيت الخادم UTC)، ويتحوّل يوم السبت منتصف الليل
// استعداداً للأسبوع القادم (السبت+2 = الإثنين). upcoming=true حين يكون الأسبوع المعروض لم يبدأ بعد.
function weekPointer(weeks, nowMs){
  const dub=new Date(new Date(nowMs==null?Date.now():nowMs).toLocaleString('en-US',{timeZone:'Asia/Dubai'}));
  const DAY=86400000, today=Date.UTC(dub.getFullYear(),dub.getMonth(),dub.getDate()), ref=today+2*DAY;
  const t=iso=>{ const [y,m,d]=String(iso).split('-').map(Number); return Date.UTC(y,m-1,d); };
  let cur=0; weeks.forEach((w,i)=>{ if(w.start && t(w.start)<=ref) cur=i; });
  const upcoming=!!(weeks[cur] && weeks[cur].start && t(weeks[cur].start)>today);
  return {cur, upcoming};
}

// قراءة تاريخ بداية الأسبوع بأي صيغة شائعة: نص (2-11-2026 / 02/11/2026 / 27-31/10/2026 / 29-03 / 9-10 / 2026 / 2026-11-02)
// أو تاريخ إكسل حقيقي، وبأرقام عربية أو إنجليزية. rangeFirst: يفضّل صيغة المدى (أيام من-إلى).
function parseStartDate(v, rangeFirst){
  const mk=(y,m,d)=>{ y=+y;m=+m;d=+d; if(!(m>=1&&m<=12&&d>=1&&d<=31&&y>=2000&&y<=2100)) return null; return new Date(y,m-1,d); };
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==='number' && v>20000 && v<80000){ const p=XLSX.SSF.parse_date_code(v); return p?mk(p.y,p.m,p.d):null; }
  const s=String(v==null?'':v).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\s+/g,' ').trim();
  if(!s) return null;
  const SEP='[\\/\\-.]', DASH='[-–—]';
  const iso=new RegExp('(\\d{4})\\s*'+SEP+'\\s*(\\d{1,2})\\s*'+SEP+'\\s*(\\d{1,2})');
  const cross=new RegExp('(\\d{1,2})\\s*'+DASH+'\\s*\\d{1,2}\\s*'+SEP+'\\s*(\\d{1,2})\\s*'+DASH+'\\s*\\d{1,2}\\s*'+SEP+'\\s*(\\d{4})');
  const range=new RegExp('(\\d{1,2})\\s*'+DASH+'\\s*\\d{1,2}\\s*'+SEP+'\\s*(\\d{1,2})\\s*'+SEP+'\\s*(\\d{4})');
  const single=new RegExp('(\\d{1,2})\\s*'+SEP+'\\s*(\\d{1,2})\\s*'+SEP+'\\s*(\\d{4})');
  let m;
  if((m=s.match(iso))) return mk(m[1],m[2],m[3]);
  const order=rangeFirst?['cross','range','single']:['single','cross','range'];
  for(const t of order){
    if(t==='cross' && (m=s.match(cross))) return mk(m[3],m[2],m[1]);
    if(t==='range' && (m=s.match(range))) return mk(m[3],m[2],m[1]);
    if(t==='single' && (m=s.match(single))) return mk(m[3],m[2],m[1]);
  }
  return null;
}
const cell=(ws,r,c)=>{const x=ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:null;};
const isSection=s=>/^[5-9][GA]\d{1,2}$/.test(s);   // v3.4: أي رقم شعبة (5G3…) لا ١ و٢ فقط
function dayToken(s){s=String(s);
  const m=s.toUpperCase().match(/MON|TUE|WED|THU|FRI/); if(m) return m[0];
  if(/الاثن|الإثن/.test(s)) return 'MON'; if(/الثلاث/.test(s)) return 'TUE';
  if(/الأربع|الاربع/.test(s)) return 'WED'; if(/الخميس/.test(s)) return 'THU';
  if(/الجمع/.test(s)) return 'FRI'; return null;}

function fmtRange(start){
  const e=new Date(start.getTime()+4*864e5);
  const sd=start.getDate(), ed=e.getDate();
  return start.getMonth()===e.getMonth()?`${sd} – ${ed} ${AR_MONTHS[start.getMonth()]}`
    :`${sd} ${AR_MONTHS[start.getMonth()]} – ${ed} ${AR_MONTHS[e.getMonth()]}`;
}
// تصحيح تاريخ الأسبوع المقلوب (يوم/شهر): يحدث حين يُكتب 2/11/2026 فيقرؤه إكسل بترتيب الشهر أولاً (11 فبراير).
// نقارن كل أسبوع بتسلسل الأسابيع المجاورة (≈ +7 أيام لكل أسبوع)، ونقلب اليوم والشهر إن كان القلب يوافق التسلسل والأصل لا يوافقه.
function repairStartDates(weeks){
  const DAY=86400000, tol=21*DAY;
  const t=d=>Date.UTC(d.getFullYear(),d.getMonth(),d.getDate());
  weeks.forEach(w=>{
    if(!w._d) return;
    const day=w._d.getDate(), mon=w._d.getMonth()+1;
    if(day>12 || day===mon) return;                               // لا يمكن القلب (اليوم لا يصلح شهراً) أو لا فرق
    const alt=new Date(w._d.getFullYear(), day-1, mon);
    const near=weeks.filter(x=>x!==w && x._d && Math.abs(x.num-w.num)<=3);
    if(near.length<2) return;                                     // لا مرجع كافٍ: لا نغيّر شيئاً
    const offs=near.map(x=>t(x._d)-7*DAY*x.num).sort((a,b)=>a-b); // الوسيط يتحمّل أسبوعاً أو اثنين خاطئين
    const med=offs[Math.floor(offs.length/2)];
    const expected=med+7*DAY*w.num;
    if(Math.abs(t(w._d)-expected)>tol && Math.abs(t(alt)-expected)<=tol) w._d=alt;
  });
}
function parseDate(s){
  const d=parseStartDate(s,true); if(!d) return null;
  if(d.getFullYear()<2026) d.setFullYear(2026); // تصحيح أمان: القالب القديم يكتب 2025 والصحيح 2026
  return d;
}

// يقرأ بطاقات الاختبار من تبويب cards إن وُجد: مفتاح = week|DAY|section|code
function readCards(wb){
  const name=wb.SheetNames.find(n=>/^(cards|exam[_ ]?cards)$/i.test(n.trim()));
  const map={};
  if(!name) return map;
  const ws=wb.Sheets[name]; if(!ws) return map; const ref=ws['!ref']; if(!ref) return map;
  const R=XLSX.utils.decode_range(ref);
  // اكتشاف صف الترويسة تلقائيًا (قد توجد صفوف فارغة في الأعلى)
  let hr=-1;
  for(let r=1;r<=Math.min(R.e.r+1,8);r++){
    for(let c=1;c<=R.e.c+1;c++){ const t=clean(cell(ws,r,c)).toLowerCase();
      if(/week|أسبوع/.test(t)||/subject/.test(t)||/class/.test(t)){ hr=r; break; } }
    if(hr>0) break;
  }
  if(hr<0) return map;
  const H={};
  for(let c=1;c<=R.e.c+1;c++){ const h=clean(cell(ws,hr,c)); const hl=h.toLowerCase();
    // ترتيب دقيق: "الصفحات" تحوي "صف" فنفحص pages قبل class
    if(/week|أسبوع/.test(hl)) H.week=c;
    else if(/day|يوم/.test(hl)) H.day=c;
    else if(/subject|رمز|مادة/.test(hl)) H.code=c;
    else if(/page|صفح/.test(hl)) H.pages=c;
    else if(/class|شعب|صف/.test(hl)) H.sec=c;
    else if(/content|عنوان|مقرر/.test(hl)) H.title=c;
    else if(/type|نوع/.test(hl)) H.type=c;
    else if(/resource|مصادر|مذاكر|تعلم/.test(hl)) H.src=c;
    else if(/score|درج|مدة/.test(hl)) H.marks=c; }
  for(let r=hr+1;r<=R.e.r+1;r++){
    const wk=parseInt(clean(cell(ws,r,H.week)),10); const dy=dayToken(cell(ws,r,H.day))||clean(cell(ws,r,H.day)).toUpperCase();
    const sec=clean(cell(ws,r,H.sec)).toUpperCase(); const code=clean(cell(ws,r,H.code)).toUpperCase();
    if(!wk||!dy||!sec||!code) continue;
    const key=wk+'|'+dy+'|'+sec+'|'+code;
    map[key]={title:clean(cell(ws,r,H.title)),pages:clean(cell(ws,r,H.pages)),
      type:clean(cell(ws,r,H.type)),marks:clean(cell(ws,r,H.marks)),sources:clean(cell(ws,r,H.src))};
  }
  return map;
}

// v3.4 — التخطيط المبسّط: الشعب صفوف، والأيام أعمدة (اثنين…جمعة)، وفي الخانة رمز المادة (أو أكثر مفصولة بفاصلة).
// يُكتشف بوجود ٣ أيام أو أكثر في صف واحد؛ وإلا يُقرأ التخطيط القديم (كتلتان وأعمدة Exam 1/2).
function parseDayCols(ws,maxRow){
  let hr=0, map={};
  for(let r=1;r<=Math.min(maxRow,30);r++){ const m={}; let n=0;
    for(let c=2;c<=12;c++){ const t=clean(cell(ws,r,c)); if(!t||t.length>30) continue; const d=dayToken(t); if(d&&!m[d]){ m[d]=c; n++; } }
    if(n>=3){ hr=r; map=m; break; } }
  if(!hr) return null;
  const bookings={}; SECTIONS.forEach(x=>bookings[x]={});
  for(let r=hr+1;r<=maxRow;r++){
    const lab=clean(cell(ws,r,1)).toUpperCase().replace(/\s+/g,''); if(!isSection(lab)) continue;
    for(const [d,c] of Object.entries(map)){
      const codes=clean(cell(ws,r,c)).toUpperCase().split(/[\s,،;+\/&]+/).filter(Boolean);
      if(codes.length){ if(!bookings[lab]) bookings[lab]={}; bookings[lab][d]=(bookings[lab][d]||[]).concat(codes); }
    }
  }
  // رقم الأسبوع والتاريخ: بالعنوان («رقم الأسبوع» / «التاريخ») والقيمة تحته
  let wk='', start=null;
  for(let r=1;r<=Math.min(maxRow,10);r++) for(let c=1;c<=14;c++){ const t=clean(cell(ws,r,c));
    if(!wk && /^رقم الأسبوع|^week\s*no/i.test(t)) wk=clean(cell(ws,r+1,c));
    if(!start && /^التاريخ|^date/i.test(t)) start=parseDate(cell(ws,r+1,c)); }
  return {bookings,wk,start};
}

function parseWeek(ws){
  if(!ws) return null;
  const ref=ws['!ref']; if(!ref) return null;
  const maxRow=XLSX.utils.decode_range(ref).e.r+1;
  const dc=parseDayCols(ws,maxRow);
  if(dc) return {num:dc.wk?+dc.wk:null, start:dc.start, bookings:dc.bookings};
  const wknum=clean(cell(ws,6,13)) || '';
  const start=parseDate(cell(ws,6,10));
  const bookings={}; SECTIONS.forEach(s=>bookings[s]={});
  // كتلتان: يسار (صف=1,اختبار=2,3) ويمين (صف=8,اختبار=9,10)
  const blocks=[{cc:1,e1:2,e2:3,dcols:[1,2,3,4,5,6,7]},{cc:8,e1:9,e2:10,dcols:[8,9,10,11,12,13,14]}];
  for(const b of blocks){
    let day=null;
    for(let r=1;r<=maxRow;r++){
      let dt=null;
      for(const dc of b.dcols){ dt=dayToken(clean(cell(ws,r,dc))); if(dt) break; }
      if(dt){ day=dt; continue; }
      const lab=clean(cell(ws,r,b.cc));
      if(day && isSection(lab)){
        if(!bookings[lab]) bookings[lab]={};
        const codes=[clean(cell(ws,r,b.e1)),clean(cell(ws,r,b.e2))].map(x=>x.toUpperCase()).filter(Boolean);
        if(codes.length){ bookings[lab][day]=(bookings[lab][day]||[]).concat(codes); }
      }
    }
  }
  return {num:wknum?+wknum:null, start, bookings};
}

function build(buf, withAudit){
  const wb=XLSX.read(buf,{type:'buffer'});
  let cards={}; try{ cards=readCards(wb); }catch(e){ cards={}; }
  const weeks=[]; const usedKeys=new Set();
  // v3.4 — في الملف الموحّد تُسمّى أوراق الاختبارات «exam 1، exam 2…» (أوراق week للخطة)؛ في الملف القديم تبقى week
  const EXAM_RX=/^(exams?|اختبار)/i;
  const examSheets=wb.SheetNames.filter(n=>EXAM_RX.test(n.trim()));
  const useRx=examSheets.length ? EXAM_RX : /^week/i;
  const parsed=[]; const secSet=new Set(SECTIONS);
  for(const name of wb.SheetNames){
    if(!useRx.test(name.trim())) continue;
    let p=null; try{ p=parseWeek(wb.Sheets[name]); }catch(e){ p=null; }
    if(!p) continue;
    parsed.push(p); Object.keys(p.bookings).forEach(k=>secSet.add(k));
  }
  // ترتيب الشعب: الصف ثم عام قبل متقدم ثم الرقم (الافتراضية دائماً حاضرة فلا يتغيّر شيء في الملف القديم)
  const secKey=c=>{ const m=c.match(/^(\d)([GA])(\d+)$/); return m?(+m[1])*1000+(m[2]==='G'?0:100)+(+m[3]):1e9; };
  const SECS=[...secSet].sort((a,b)=>secKey(a)-secKey(b));
  for(const p of parsed){
    const num=p.num||weeks.length+1;
    // بناء أيام كل شعبة مع أسماء المواد والبطاقات
    const sections={};
    for(const sec of SECS){
      const days=DAYS.map(([tok,alabel])=>{
        const codes=((p.bookings[sec]||{})[tok]||[]);
        const exams=codes.map(code=>{
          const key=num+'|'+tok+'|'+sec+'|'+code;
          usedKeys.add(key);
          const card=cards[key]||null;
          return {code, subject:SUBJ[code]||code, card};
        });
        return {day:tok, label:alabel, exams};
      });
      sections[sec]=days;
    }
    weeks.push({num, label:'الأسبوع '+num, _d:p.start||null, range:'', start:null, sections});
  }
  weeks.sort((a,b)=>a.num-b.num);
  repairStartDates(weeks);
  const iso=d=>d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);
  weeks.forEach(w=>{ w.range=w._d?fmtRange(w._d):''; w.start=w._d?iso(w._d):null; delete w._d; });
  const {cur,upcoming}=weekPointer(weeks);
  const out={school:SCHOOL, currentWeekIndex:cur, currentIsUpcoming:upcoming, sectionOrder:SECS, dayOrder:DAYS, weeks};
  if(withAudit){
    const dayLabel={}; DAYS.forEach(([t,l])=>dayLabel[t]=l);
    const orphans=[];
    for(const [key,card] of Object.entries(cards)){
      if(usedKeys.has(key)) continue;
      const [wkS,day,section,code]=key.split('|');
      orphans.push({week:+wkS, day, dayLabel:dayLabel[day]||day, section, code, subject:SUBJ[code]||code, card});
    }
    out.audit={orphans};
  }
  return out;
}

const UA={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36','Accept':'*/*'};
const isZip=b=>b&&b.length>3&&b[0]===0x50&&b[1]===0x4B;
function shareToken(u){return 'u!'+Buffer.from(u,'utf8').toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');}
async function tryFetch(u){try{const r=await fetch(u,{redirect:'follow',headers:UA});const b=Buffer.from(await r.arrayBuffer());return isZip(b)?b:null;}catch(e){return null;}}
async function download(){
  const m=SHARE_URL.match(/^(https:\/\/[^/]+)\/:[a-z]:\/[a-z]\/(personal\/[^/]+)\/([^/?#]+)/i);
  const urls=[];
  if(m) urls.push(m[1]+'/'+m[2]+'/_layouts/15/download.aspx?share='+m[3]);
  urls.push(SHARE_URL+(SHARE_URL.includes('?')?'&':'?')+'download=1');
  urls.push('https://api.onedrive.com/v1.0/shares/'+shareToken(SHARE_URL)+'/root/content');
  for(const u of urls){const b=await tryFetch(u); if(b) return b;}
  return null;
}
exports.handler=async(event)=>{
  try{
    // v3.3.1: كلمة السرّ لا تُقبل في الرابط أبداً. التدقيق الإداري عبر POST + البوابة الموحّدة فقط.
    let audit=false;
    if(String((event&&event.httpMethod)||'GET').toUpperCase()==='POST'){
      let b={}; try{b=JSON.parse((event&&event.body)||'{}');}catch(e){return {statusCode:400,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:'bad'})};}
      const gt=await guard.gate(event,b&&b.pw,process.env.ADMIN_PASSWORD,'exams');
      if(!gt.ok) return {statusCode:gt.code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify({error:gt.error})};
      audit=true;
    }
    const buf=await download();
    if(!buf) return {statusCode:502,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:'تعذّر تنزيل ملف الاختبارات'})};
    return {statusCode:200,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':audit?'no-store':'public, max-age=300'},body:JSON.stringify(build(buf,audit))};
  }catch(e){return {statusCode:500,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:String(e&&e.message||e)})};}
};
exports._build=build;
exports._weekPointer=weekPointer;
