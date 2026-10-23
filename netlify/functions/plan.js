const XLSX = require('xlsx');

// ============================================================
//  رابط مشاركة ملف الإكسل (OneDrive). غيّر هذا السطر فقط لو تغيّر الرابط:
// ============================================================
const SHARE_URL = process.env.PLAN_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQDfAXk8MM1VQb-ewqBaf2K2AckPddsEiesRhWLxTm__Zo0?rtime=MN-AdZkM30g";
// ============================================================

const SCHOOL = "مدرسة المعيريض للحلقة الثانية بنين";
const TERM   = "الفصل الأول ٢٦-٢٧";

const SECTIONS = [['5G','الخامس العام'],['5A','الخامس المتقدم'],
  ['6G','السادس العام'],['6A','السادس المتقدم'],
  ['7G','السابع العام'],['7A','السابع المتقدم'],
  ['8G','الثامن العام'],['8A','الثامن المتقدم']];
const COL = {'5G':3,'5A':4,'6G':5,'6A':6,'7G':7,'7A':8,'8G':9,'8A':10};
const C_TEACHER=1, C_SUBJECT=2, C_NOTE=11, C_EXAM=12, C_REM=13;
const AR_MONTHS=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو',
  'أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];

const adig=x=>String(x).replace(/[0-9]/g,d=>'٠١٢٣٤٥٦٧٨٩'[+d]);
const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();

// ====== أسماء المعلمين: تطبيع + تبويب names + دمج تلقائي + «الشعب الحالية» ======
const RECENT_WEEKS = 2;   // «الشعب الحالية» للمعلم = ما أدخله في آخر أسبوعين أدخل فيهما شيئاً
const _norm=s=>String(s||'').toLowerCase().replace(/[\u064B-\u0652\u0640]/g,'').replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/[^\p{L}\p{N}\s]/gu,' ').replace(/\s+/g,' ').trim();
const _toks=s=>_norm(s).split(' ').filter(Boolean), _compact=s=>_norm(s).replace(/\s/g,'');

// تبويب اختياري باسم names في ملف الخطة: (الاسم في الملف) | (الاسم المعتمد) | (الحالة)
//   • مرادف: اكتب الاسم المعتمد  → كل كتابة للاسم الأول تُقرأ كالاسم المعتمد.
//   • متوقّف: اكتب أي كلمة في عمود الحالة (متوقف/غادر/نُقل) → يُستثنى من المتابعة ومن مُحدِّد المعلم.
function parseNamesTab(wb){
  const out={alias:[],retired:[]}; const sn=wb.SheetNames.find(n=>/^names?$/i.test(String(n).trim())); if(!sn) return out;
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
  rows.forEach((r,i)=>{ const a=clean(r[0]), b=clean(r[1]), c=clean(r[2]); if(!a) return; if(i<2 && /^(الاسم|اسم|name)/i.test(a)) return;
    if(c) out.retired.push(a); else if(b && _compact(b)!==_compact(a)) out.alias.push([a,b]); });
  return out;
}

// توحيد أسماء الشخص الواحد بصيغ مختلفة. القواعد (داخل المادة الواحدة): مرادف يدوي من names ← دائماً · تطابق بلا مسافات ← دمج ·
// اسم أقصر يطابق الأطول في أول كلمة وآخر كلمة ← دمج · الاسمان في أسبوع واحد ← لا دمج (شخصان)، يُعرض كاقتراح · تشابه أضعف ← اقتراح فقط.
function buildNameMap(recs, alias){
  const explicit=new Map(); (alias||[]).forEach(([from,to])=>explicit.set(_compact(from),to));
  const ex=n=>explicit.get(_compact(n))||n;
  const bySub=new Map();
  recs.forEach(r=>{ if(!r.teacher||!r.subject) return; const t=ex(r.teacher);
    let o=bySub.get(r.subject); if(!o){ o={names:new Map(),weeks:new Map()}; bySub.set(r.subject,o); }
    const n=o.names.get(t)||{count:0,last:0}; n.count++; n.last=Math.max(n.last,r.week||0); o.names.set(t,n);
    if(!o.weeks.has(r.week)) o.weeks.set(r.week,new Set()); o.weeks.get(r.week).add(t); });
  const autoMap=new Map(), suspects=[];
  for(const [sub,o] of bySub){
    const names=[...o.names.keys()]; if(names.length<2) continue;
    const together=(a,b)=>{ for(const s of o.weeks.values()) if(s.has(a)&&s.has(b)) return true; return false; };
    const parent={}; names.forEach(n=>parent[n]=n);
    const find=x=>{ while(parent[x]!==x){ parent[x]=parent[parent[x]]; x=parent[x]; } return x; };
    const union=(a,b)=>{ parent[find(a)]=find(b); };
    for(let i=0;i<names.length;i++) for(let j=i+1;j<names.length;j++){
      const a=names[i], b=names[j];
      if(_compact(a)===_compact(b)){ union(a,b); continue; }
      const ta=_toks(a), tb=_toks(b), S=ta.length<=tb.length?ta:tb, L=ta.length<=tb.length?tb:ta;
      if(!(S.length>=2 && S.length<L.length && S.every(w=>L.includes(w)))) continue;
      if(together(a,b)){ suspects.push({subject:sub,a,b,sameWeek:true}); continue; }
      if(S[0]===L[0] && S[S.length-1]===L[L.length-1]) union(a,b); else suspects.push({subject:sub,a,b,sameWeek:false});
    }
    const groups={}; names.forEach(n=>{ const k=find(n); (groups[k]=groups[k]||[]).push(n); });
    for(const g of Object.values(groups)){
      if(g.length<2) continue;
      g.sort((x,y)=>_toks(x).length-_toks(y).length || o.names.get(y).count-o.names.get(x).count || o.names.get(y).last-o.names.get(x).last || x.length-y.length);
      g.slice(1).forEach(n=>autoMap.set(sub+'||'+n,g[0]));
    }
  }
  const map=new Map(), merges=[], seen=new Set();
  recs.forEach(r=>{ if(!r.teacher||!r.subject) return; const k=r.subject+'||'+r.teacher; if(seen.has(k)) return; seen.add(k);
    const n2=ex(r.teacher), c=autoMap.get(r.subject+'||'+n2)||n2;
    if(c!==r.teacher){ map.set(k,c); merges.push({subject:r.subject,from:r.teacher,to:c,manual:n2!==r.teacher}); } });
  return {map,merges,suspects:suspects.filter(s=>!autoMap.has(s.subject+'||'+s.a)&&!autoMap.has(s.subject+'||'+s.b))};
}

// «الشعب الحالية» لمعلم/مادة: اتحاد ما أدخله في آخر K أسبوع أدخل فيه شيئاً (حتى الأسبوع uptoIdx). يتجاوز تأخّره، ويُسقط الشعب التي ترك تدريسها.
function currentSections(timeline, uptoIdx, K){
  const act=timeline.filter(t=>t.i<=uptoIdx && t.codes.size>0).slice(-(K||RECENT_WEEKS)); const u=new Set(); act.forEach(t=>t.codes.forEach(c=>u.add(c))); return u;
}

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

function meta(name){
  const n=name.toLowerCase();
  if(name.includes('إسلام')) return ['🕌','#1F6E4E'];
  if(name.includes('عرب') && !name.includes('رياض')) return ['📖','#0E5A54'];
  if(n.includes('engl')||name.includes('إنجليز')) return ['🔤','#2A6F97'];
  if(name.includes('دراس')||name.includes('اجتماع')) return ['🌍','#3A5A8C'];
  if(n.includes('math')||name.includes('رياضي')) return ['➗','#4B4B8F'];
  if(n.includes('scien')||name.includes('علوم')) return ['🔬','#2E7D6B'];
  if(name.includes('تصميم')||n.includes('ccdi')||name.includes('تكنولوج')) return ['⚙️','#8A6D3B'];
  if(name.includes('بدني')||name.includes('صحي')||name.includes('رياضة')) return ['⚽','#4F7A3A'];
  if(name.includes('فنون')||name.includes('دراما')||name.includes('مسرح')) return ['🎨','#7A4A78'];
  return ['📘','#556B6A'];
}
function fixSubject(s){ s=clean(s).replace('Engliah','English'); return s==='الدراسات'?'الدراسات الاجتماعية':s; }
function fmtRange(start){
  const end=new Date(start.getTime()+4*86400000);
  const sd=adig(start.getDate()), ed=adig(end.getDate());
  if(start.getMonth()===end.getMonth()) return `${sd} – ${ed} ${AR_MONTHS[start.getMonth()]}`;
  return `${sd} ${AR_MONTHS[start.getMonth()]} – ${ed} ${AR_MONTHS[end.getMonth()]}`;
}
function cell(ws,r1,c1){ const cel=ws[XLSX.utils.encode_cell({r:r1-1,c:c1-1})]; return cel?cel.v:null; }

function pivot(buf){
  const wb=XLSX.read(buf,{type:'buffer'});
  const weeks=[];
  for(const name of wb.SheetNames){
    if(!/^week/i.test(name)) continue;
    const ws=wb.Sheets[name]; const ref=ws['!ref']; if(!ref) continue;
    const maxRow=XLSX.utils.decode_range(ref).e.r+1;
    const title=clean(cell(ws,3,2));
    const mNum=title.match(/الأسبوع\s*(\d+)/); const num=mNum?+mNum[1]:weeks.length+1;
    const start=parseStartDate(title,false);
    const sections={}; SECTIONS.forEach(([c])=>sections[c]=[]);
    // إعلانات الأسبوع: نصٌّ في عمود (ملاحظات/موعد/تذكير) لا يحمله أي درس.
    //  • صفّ بلا اسم معلم ولا مادة  → إعلان عام لكل الأولياء دائماً.
    //  • صفّ معلم/مادة بلا دروس      → إعلان عام فقط إن كان الأسبوع كله بلا دروس (عطلة)؛ وإلا يُحجب (غالباً أُدخل في العمود الخطأ) ويظهر للإدارة في لوحة المتابعة.
    const general=[], stray=[], roster=[];
    for(let r=5;r<=maxRow;r++){
      const subject=fixSubject(cell(ws,r,C_SUBJECT));
      const teacher=clean(cell(ws,r,C_TEACHER));
      const note=clean(cell(ws,r,C_NOTE)), exam=clean(cell(ws,r,C_EXAM)), rem=clean(cell(ws,r,C_REM));
      const texts=()=>[note, rem?'تذكير: '+rem:'', exam?'موعد اختبار: '+exam:''];
      if(!subject && !teacher){ texts().forEach(t=>general.push(t)); continue; }
      const [emoji,tint]=meta(subject);
      let placed=false; const codesFilled=[];
      for(const [code] of SECTIONS){
        const lesson=clean(cell(ws,r,COL[code]));
        if(!lesson) continue;
        const e={subject,teacher,lesson,emoji,tint};
        if(note) e.note=note; if(rem) e.reminder=rem; if(exam) e.exam=exam;
        sections[code].push(e); placed=true; codesFilled.push(code);
      }
      if(teacher && subject) roster.push({subject,teacher,codes:codesFilled});
      if(!placed) texts().forEach(t=>stray.push(t));
    }
    const weekHasLessons=Object.values(sections).some(a=>a.length>0);
    const notices=[];
    (weekHasLessons?general:general.concat(stray)).forEach(t=>{ t=clean(t); if(t.length>1 && !/^[-–—.\s]+$/.test(t) && !notices.includes(t)) notices.push(t); });
    weeks.push({num,label:'الأسبوع '+adig(num), notices, _roster:roster,
      range:start?fmtRange(start):'', start:start?start.toISOString().slice(0,10):null, sections});
  }
  // قراءة إعلان إدارة المدرسة من تبويب announcement (A1 عنوان، A2 نص) — يُقرأ ولو كان مخفيًّا
  let announcement=null;
  const annName=wb.SheetNames.find(n=>/^announcement$/i.test(n.trim()));
  if(annName){
    const aw=wb.Sheets[annName];
    const title=clean(cell(aw,1,1));
    const body =clean(cell(aw,2,1));
    if(body) announcement={title:title||'إعلان', body};
  }
  // قراءة تبويب الروابط المفيدة links (title/url/category — بالعربية أو الإنجليزية)
  let links=[];
  const lname=wb.SheetNames.find(n=>/^links$/i.test(n.trim()));
  if(lname){
    const lw=wb.Sheets[lname];
    if(lw && lw['!ref']){
      const LR=XLSX.utils.decode_range(lw['!ref']);
      // اكتشاف صف الترويسة
      let hr=-1;
      for(let r=1;r<=Math.min(LR.e.r+1,8);r++){
        for(let c=1;c<=LR.e.c+1;c++){ const t=clean(cell(lw,r,c)).toLowerCase();
          if(/url|رابط|link/.test(t)||/title|اسم|عنوان|وصف/.test(t)){ hr=r; break; } }
        if(hr>0) break;
      }
      const H={};
      if(hr>0){
        for(let c=1;c<=LR.e.c+1;c++){ const h=clean(cell(lw,hr,c)).toLowerCase();
          if(/url|رابط|link/.test(h)) H.url=c;
          else if(/categ|تصنيف|قسم|فئة/.test(h)) H.cat=c;
          else if(/desc|وصف/.test(h)) H.desc=c;
          else if(/title|اسم|عنوان/.test(h)) H.title=c; }
        for(let r=hr+1;r<=LR.e.r+1;r++){
          const url=clean(cell(lw,r,H.url));
          const title=clean(cell(lw,r,H.title));
          if(!url && !title) continue;
          if(!url) continue; // بلا رابط لا فائدة
          links.push({title:title||url, url,
            category:H.cat?clean(cell(lw,r,H.cat)):'', desc:H.desc?clean(cell(lw,r,H.desc)):''});
        }
      }
    }
  }
  // قراءة تبويب الإعلانات المهمة news
  let news=[];
  const nname=wb.SheetNames.find(n=>/^news$/i.test(n.trim()));
  if(nname){
    const nw=wb.Sheets[nname];
    if(nw && nw['!ref']){
      const NR=XLSX.utils.decode_range(nw['!ref']);
      let hr=-1;
      for(let r=1;r<=Math.min(NR.e.r+1,8);r++){
        for(let c=1;c<=NR.e.c+1;c++){ const t=clean(cell(nw,r,c)).toLowerCase();
          if(/title|عنوان/.test(t)||/body|نص|محتوى/.test(t)||/date|تاريخ/.test(t)){ hr=r; break; } }
        if(hr>0) break;
      }
      if(hr>0){
        const H={};
        for(let c=1;c<=NR.e.c+1;c++){ const h=clean(cell(nw,hr,c)).toLowerCase();
          if(/date|تاريخ/.test(h)) H.date=c;
          else if(/pin|مثبّت|مثبت|تثبيت/.test(h)) H.pin=c;
          else if(/title|عنوان/.test(h)) H.title=c;
          else if(/body|نص|محتوى|تفاصيل/.test(h)) H.body=c; }
        for(let r=hr+1;r<=NR.e.r+1;r++){
          const title=clean(cell(nw,r,H.title)); const body=clean(cell(nw,r,H.body));
          if(!title && !body) continue;
          let dv=cell(nw,r,H.date), diso='';
          if(dv instanceof Date) diso=dv.toISOString().slice(0,10);
          else { const ds=clean(dv); const m=ds.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/)||ds.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
            if(m){ if(m[1].length===4) diso=m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
                   else diso=m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2); } }
          const pin=/^(نعم|yes|true|1|مثبّت|مثبت|✓|y)$/i.test(clean(cell(nw,r,H.pin)));
          news.push({date:diso, title:title||'إعلان', body, pinned:pin});
        }
        // ترتيب: المثبّت أولاً ثم الأحدث تاريخًا
        news.sort((a,b)=>{ if(a.pinned!==b.pinned) return a.pinned?-1:1; return (b.date||'').localeCompare(a.date||''); });
      }
    }
  }
  weeks.sort((a,b)=>a.num-b.num);
  const NT=parseNamesTab(wb);
  const recs=[]; weeks.forEach(w=>{ Object.values(w.sections).forEach(a=>a.forEach(e=>recs.push({week:w.num,subject:e.subject,teacher:e.teacher}))); w._roster.forEach(r=>recs.push({week:w.num,subject:r.subject,teacher:r.teacher})); });
  const nm=buildNameMap(recs,NT.alias); const canon=(sub,t)=>nm.map.get(sub+'||'+t)||t;
  weeks.forEach(w=>{ Object.values(w.sections).forEach(a=>a.forEach(e=>{ e.teacher=canon(e.subject,e.teacher); })); w._roster.forEach(r=>{ r.teacher=canon(r.subject,r.teacher); }); });
  const {cur,upcoming}=weekPointer(weeks);
  // المعلمون «الحاليون» لكل (شعبة، مادة): لهم صفّ في أسبوع اليوم + يدرّسون الشعبة في آخر أسبوعين أدخلوا فيهما شيئاً + ليسوا متوقّفين
  const retKeys=new Set(NT.retired.map(_compact)), isRet=n=>retKeys.has(_compact(n));
  const timeline={}; weeks.forEach((w,i)=>{ const per={}; w._roster.forEach(r=>{ const k=r.subject+'||'+r.teacher; const st=per[k]||(per[k]=new Set()); r.codes.forEach(c=>st.add(c)); }); Object.entries(per).forEach(([k,st])=>{ (timeline[k]=timeline[k]||[]).push({i,codes:st}); }); });
  const rosterNow=new Set((weeks[cur]?weeks[cur]._roster:[]).map(r=>r.subject+'||'+r.teacher)), useRoster=rosterNow.size>0;
  const activeTeachers={}; SECTIONS.forEach(([c])=>activeTeachers[c]={});
  Object.entries(timeline).forEach(([k,tl])=>{ const p=k.indexOf('||'), subject=k.slice(0,p), teacher=k.slice(p+2);
    if(isRet(teacher) || (useRoster && !rosterNow.has(k))) return;
    currentSections(tl,weeks.length-1).forEach(c=>{ const arr=activeTeachers[c][subject]||(activeTeachers[c][subject]=[]); if(!arr.includes(teacher)) arr.push(teacher); }); });
  const retired=[...new Set(recs.map(r=>canon(r.subject,r.teacher)).filter(t=>t&&isRet(t)))];
  weeks.forEach(w=>{ delete w._roster; });
  return {school:SCHOOL,term:TERM,currentWeekIndex:cur, currentIsUpcoming:upcoming, activeTeachers, retired, announcement, links, news,
    sectionOrder:SECTIONS.map(([code,label])=>({code,label})), weeks};
}

const UA={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36','Accept':'*/*'};
const isZip=b=> b && b.length>3 && b[0]===0x50 && b[1]===0x4B; // "PK"

function shareToken(url){
  const b64=Buffer.from(url,'utf8').toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');
  return 'u!'+b64;
}

async function tryFetch(url){
  try{
    const r=await fetch(url,{redirect:'follow',headers:UA});
    const ct=r.headers.get('content-type')||'';
    const buf=Buffer.from(await r.arrayBuffer());
    return {ok:r.ok && isZip(buf), status:r.status, ct, buf,
            head: isZip(buf)?'[xlsx]':buf.slice(0,80).toString('utf8').replace(/\s+/g,' ')};
  }catch(e){ return {ok:false,status:0,ct:'',buf:null,head:'ERR '+(e&&e.message||e)}; }
}

async function download(){
  const attempts=[];
  const urls=[];
  // 1) download.aspx — الأنسب لحسابات SharePoint المؤسسية (anyone-with-link)
  const m=SHARE_URL.match(/^(https:\/\/[^/]+)\/:[a-z]:\/[a-z]\/(personal\/[^/]+)\/([^/?#]+)/i);
  if(m) urls.push(['download.aspx', m[1]+'/'+m[2]+'/_layouts/15/download.aspx?share='+m[3]]);
  // 2) download=1 على رابط المشاركة
  urls.push(['download=1', SHARE_URL+(SHARE_URL.includes('?')?'&':'?')+'download=1']);
  // 3) واجهة وان درايف للمشاركات
  urls.push(['onedrive-shares', 'https://api.onedrive.com/v1.0/shares/'+shareToken(SHARE_URL)+'/root/content']);
  for(const [name,url] of urls){
    const a=await tryFetch(url);
    attempts.push({method:name,status:a.status,contentType:a.ct,head:a.head});
    if(a.ok) return {buf:a.buf, attempts};
  }
  return {buf:null, attempts};
}

exports.handler = async (event) => {
  const debug = event && event.queryStringParameters && event.queryStringParameters.debug;
  try{
    const {buf,attempts}=await download();
    if(!buf){
      return { statusCode:502, headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},
        body: JSON.stringify({error:'تعذّر تنزيل ملف الإكسل من مايكروسوفت',attempts},null,2) };
    }
    const data=pivot(buf);
    if(debug) data._attempts=attempts;
    return { statusCode:200,
      headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'public, max-age=300'},
      body: JSON.stringify(data) };
  }catch(e){
    return { statusCode:500, headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},
      body: JSON.stringify({error:String(e&&e.message||e)}) };
  }
};
exports._pivot = pivot;
exports._weekPointer = weekPointer;
exports._buildNameMap = buildNameMap;
exports._currentSections = currentSections;
