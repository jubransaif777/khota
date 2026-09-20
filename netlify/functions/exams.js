const XLSX = require('xlsx');

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
  PE:'تربية بدنية',DT:'تصميم تكنولوجي',DR:'مسرح',VA:'فنية',AA:'موسيقى'};
const AR_MONTHS=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];

const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
const cell=(ws,r,c)=>{const x=ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:null;};
const isSection=s=>/^[5-8][GA][12]$/.test(s);
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
function parseDate(s){
  const m=clean(s).match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/);
  if(!m) return null;
  let y=+m[3]; if(y<2026) y=2026; // تصحيح أمان: الملف يكتب 2025 والصحيح 2026
  return new Date(y, +m[2]-1, +m[1]);
}

// يقرأ بطاقات الاختبار من تبويب cards إن وُجد: مفتاح = week|DAY|section|code
function readCards(wb){
  const name=wb.SheetNames.find(n=>/^(cards|exam[_ ]?cards)$/i.test(n.trim()));
  const map={};
  if(!name) return map;
  const ws=wb.Sheets[name]; if(!ws) return map; const ref=ws['!ref']; if(!ref) return map;
  const R=XLSX.utils.decode_range(ref);
  // ترويسة الأعمدة (الصف الأول)
  const H={};
  for(let c=1;c<=R.e.c+1;c++){ const h=clean(cell(ws,1,c)); const hl=h.toLowerCase();
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
  for(let r=2;r<=R.e.r+1;r++){
    const wk=clean(cell(ws,r,H.week)); const dy=dayToken(cell(ws,r,H.day))||clean(cell(ws,r,H.day)).toUpperCase();
    const sec=clean(cell(ws,r,H.sec)); const code=clean(cell(ws,r,H.code)).toUpperCase();
    if(!wk||!dy||!sec||!code) continue;
    const key=wk+'|'+dy+'|'+sec+'|'+code;
    map[key]={title:clean(cell(ws,r,H.title)),pages:clean(cell(ws,r,H.pages)),
      type:clean(cell(ws,r,H.type)),marks:clean(cell(ws,r,H.marks)),sources:clean(cell(ws,r,H.src))};
  }
  return map;
}

function parseWeek(ws){
  if(!ws) return null;
  const ref=ws['!ref']; if(!ref) return null;
  const maxRow=XLSX.utils.decode_range(ref).e.r+1;
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
        const codes=[clean(cell(ws,r,b.e1)),clean(cell(ws,r,b.e2))].map(x=>x.toUpperCase()).filter(Boolean);
        if(codes.length){ bookings[lab][day]=(bookings[lab][day]||[]).concat(codes); }
      }
    }
  }
  return {num:wknum?+wknum:null, start, bookings};
}

function build(buf){
  const wb=XLSX.read(buf,{type:'buffer'});
  let cards={}; try{ cards=readCards(wb); }catch(e){ cards={}; }
  const weeks=[];
  for(const name of wb.SheetNames){
    if(!/^week/i.test(name)) continue;
    let p=null; try{ p=parseWeek(wb.Sheets[name]); }catch(e){ p=null; }
    if(!p) continue;
    const num=p.num||weeks.length+1;
    // بناء أيام كل شعبة مع أسماء المواد والبطاقات
    const sections={};
    for(const sec of SECTIONS){
      const days=DAYS.map(([tok,alabel])=>{
        const codes=(p.bookings[sec][tok]||[]);
        const exams=codes.map(code=>{
          const key=num+'|'+tok+'|'+sec+'|'+code;
          const card=cards[key]||null;
          return {code, subject:SUBJ[code]||code, card};
        });
        return {day:tok, label:alabel, exams};
      });
      sections[sec]=days;
    }
    weeks.push({num, label:'الأسبوع '+num,
      range:p.start?fmtRange(p.start):'', start:p.start?p.start.toISOString().slice(0,10):null,
      sections});
  }
  weeks.sort((a,b)=>a.num-b.num);
  const today=new Date(); today.setHours(0,0,0,0);
  let cur=0; weeks.forEach((w,i)=>{ if(w.start && new Date(w.start)<=today) cur=i; });
  return {school:SCHOOL, currentWeekIndex:cur, sectionOrder:SECTIONS, dayOrder:DAYS, weeks};
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
exports.handler=async()=>{
  try{
    const buf=await download();
    if(!buf) return {statusCode:502,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:'تعذّر تنزيل ملف الاختبارات'})};
    return {statusCode:200,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'public, max-age=300'},body:JSON.stringify(build(buf))};
  }catch(e){return {statusCode:500,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:String(e&&e.message||e)})};}
};
exports._build=build;
