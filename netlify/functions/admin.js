const XLSX = require('xlsx');

// ============================================================
//  دالة لوحة متابعة المعلمين — محمية بكلمة سرّ (متغيّر بيئة)
//  رابط ملف الإكسل: نفس متغيّر رابط الخطة
// ============================================================
const SHARE_URL = process.env.PLAN_SHARE_URL ||
  "https://emiratesschoolsese-my.sharepoint.com/:x:/g/personal/gubran_algumaei_moe_sch_ae/IQDfAXk8MM1VQb-ewqBaf2K2AckPddsEiesRhWLxTm__Zo0?rtime=MN-AdZkM30g";
// كلمة السرّ تُضبط في Netlify كمتغيّر بيئة باسم ADMIN_PASSWORD
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "";

const SECTIONS = [['5G','الخامس العام'],['5A','الخامس المتقدم'],
  ['6G','السادس العام'],['6A','السادس المتقدم'],
  ['7G','السابع العام'],['7A','السابع المتقدم'],
  ['8G','الثامن العام'],['8A','الثامن المتقدم']];
const COL = {'5G':3,'5A':4,'6G':5,'6A':6,'7G':7,'7A':8,'8G':9,'8A':10};
const C_TEACHER=1, C_SUBJECT=2, C_NOTE=11, C_EXAM=12, C_REM=13;

const adig=x=>String(x).replace(/[0-9]/g,d=>'٠١٢٣٤٥٦٧٨٩'[+d]);
const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();

// توحيد أسماء المعلمين: الاسم نفسه بصيغ مختلفة (اختصار الاسم الرباعي/الخماسي، أو اختلاف المسافات) = معلّم واحد.
// القواعد (داخل المادة الواحدة فقط): تطابق بلا مسافات ← دمج · اسم أقصر يطابق الأطول في أول كلمة وآخر كلمة ← دمج
// · الاسمان في أسبوع واحد للمادة نفسها ← شخصان مختلفان (لا دمج) · تشابه أضعف ← لا دمج، ويُعرض كاقتراح للإدارة.
function buildNameMap(recs){
  const norm=s=>String(s||'').toLowerCase().replace(/[\u064B-\u0652\u0640]/g,'').replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/[^\p{L}\p{N}\s]/gu,' ').replace(/\s+/g,' ').trim();
  const toks=s=>norm(s).split(' ').filter(Boolean), compact=s=>norm(s).replace(/\s/g,'');
  const bySub=new Map();
  recs.forEach(r=>{ if(!r.teacher||!r.subject) return;
    let o=bySub.get(r.subject); if(!o){ o={names:new Map(),weeks:new Map()}; bySub.set(r.subject,o); }
    const n=o.names.get(r.teacher)||{count:0,last:0}; n.count++; n.last=Math.max(n.last,r.week||0); o.names.set(r.teacher,n);
    if(!o.weeks.has(r.week)) o.weeks.set(r.week,new Set()); o.weeks.get(r.week).add(r.teacher); });
  const map=new Map(), merges=[], suspects=[];
  for(const [sub,o] of bySub){
    const names=[...o.names.keys()]; if(names.length<2) continue;
    const together=(a,b)=>{ for(const s of o.weeks.values()) if(s.has(a)&&s.has(b)) return true; return false; };
    const parent={}; names.forEach(n=>parent[n]=n);
    const find=x=>{ while(parent[x]!==x){ parent[x]=parent[parent[x]]; x=parent[x]; } return x; };
    const union=(a,b)=>{ parent[find(a)]=find(b); };
    for(let i=0;i<names.length;i++) for(let j=i+1;j<names.length;j++){
      const a=names[i], b=names[j];
      if(compact(a)===compact(b)){ union(a,b); continue; }
      const ta=toks(a), tb=toks(b), S=ta.length<=tb.length?ta:tb, L=ta.length<=tb.length?tb:ta;
      if(!(S.length>=2 && S.length<L.length && S.every(w=>L.includes(w)))) continue;
      if(together(a,b)) continue;
      if(S[0]===L[0] && S[S.length-1]===L[L.length-1]) union(a,b); else suspects.push({subject:sub,a,b});
    }
    const groups={}; names.forEach(n=>{ const k=find(n); (groups[k]=groups[k]||[]).push(n); });
    for(const g of Object.values(groups)){
      if(g.length<2) continue;
      g.sort((x,y)=>toks(x).length-toks(y).length || o.names.get(y).count-o.names.get(x).count || o.names.get(y).last-o.names.get(x).last || x.length-y.length);
      g.slice(1).forEach(n=>{ map.set(sub+'||'+n,g[0]); merges.push({subject:sub,from:n,to:g[0]}); });
    }
  }
  // اقتراح لا يتكرر لزوج دُمج أصلاً
  const merged=new Set(merges.map(m=>m.subject+'||'+m.from+'||'+m.to));
  return {map,merges,suspects:suspects.filter(s=>!map.has(s.subject+'||'+s.a)&&!map.has(s.subject+'||'+s.b))};
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
const cell=(ws,r,c)=>{const x=ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:null;};
function fixSubject(s){s=clean(s).replace('Engliah','English');return s==='الدراسات'?'الدراسات الاجتماعية':s;}

const UA={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36','Accept':'*/*'};
const isZip=b=>b&&b.length>3&&b[0]===0x50&&b[1]===0x4B;
function shareToken(url){return 'u!'+Buffer.from(url,'utf8').toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');}
async function tryFetch(url){try{const r=await fetch(url,{redirect:'follow',headers:UA});const b=Buffer.from(await r.arrayBuffer());return{ok:r.ok&&isZip(b),buf:b};}catch(e){return{ok:false,buf:null};}}
async function download(){
  const m=SHARE_URL.match(/^(https:\/\/[^/]+)\/:[a-z]:\/[a-z]\/(personal\/[^/]+)\/([^/?#]+)/i);
  const urls=[];
  if(m) urls.push(m[1]+'/'+m[2]+'/_layouts/15/download.aspx?share='+m[3]);
  urls.push(SHARE_URL+(SHARE_URL.includes('?')?'&':'?')+'download=1');
  urls.push('https://api.onedrive.com/v1.0/shares/'+shareToken(SHARE_URL)+'/root/content');
  for(const u of urls){const a=await tryFetch(u);if(a.ok)return a.buf;}
  return null;
}

// يحسب حالة الإدخال لكل معلّم/مادة عبر الشعب في أسبوع محدّد
function follow(buf){
  const wb=XLSX.read(buf,{type:'buffer'});
  const weekSheets=wb.SheetNames.filter(n=>/^week/i.test(n));
  const weeks=[];
  const AR_MONTHS=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
  for(const name of weekSheets){
    const ws=wb.Sheets[name];const ref=ws['!ref'];if(!ref)continue;
    const maxRow=XLSX.utils.decode_range(ref).e.r+1;
    const title=clean(cell(ws,3,2));
    const mNum=title.match(/الأسبوع\s*(\d+)/);const num=mNum?+mNum[1]:weeks.length+1;
    const start=parseStartDate(title,false);
    let range='';
    if(start){const e=new Date(start.getTime()+4*864e5);
      range=start.getMonth()===e.getMonth()?`${adig(start.getDate())} – ${adig(e.getDate())} ${AR_MONTHS[start.getMonth()]}`
        :`${adig(start.getDate())} ${AR_MONTHS[start.getMonth()]} – ${adig(e.getDate())} ${AR_MONTHS[e.getMonth()]}`;}
    const rows=[], strayRows=[];
    for(let r=5;r<=maxRow;r++){
      const subject=fixSubject(cell(ws,r,C_SUBJECT));
      const teacher=clean(cell(ws,r,C_TEACHER));
      if(!teacher && !subject) continue;
      { let any=0; for(const [code] of SECTIONS) if(clean(cell(ws,r,COL[code]))) any++;      // صفّ تدريس بلا دروس: نحتفظ بنصوصه للتنبيه
        if(!any) strayRows.push({teacher,subject,note:clean(cell(ws,r,C_NOTE)),exam:clean(cell(ws,r,C_EXAM)),rem:clean(cell(ws,r,C_REM))}); }
      if(!teacher || /^شاغر/.test(teacher)) continue; // تجاهل الصفوف الشاغرة
      const cells={}; let filled=0, teaches=0;
      for(const [code] of SECTIONS){
        const v=clean(cell(ws,r,COL[code]));
        // نعتبر الشعبة \"مطلوبة\" لهذا المعلّم إن كان يدرّسها في أي أسبوع — نحسبها لاحقًا
        cells[code]=v?1:0; if(v){filled++;}
      }
      rows.push({teacher,subject,cells,filled});
    }
    // نصوص كُتبت في عمود (ملاحظات/موعد/تذكير) بلا دروس في خانات الشعب، في أسبوع فيه دروس: لا تظهر لأولياء الأمور فيُنبَّه عليها
    const strays=[]; const okT=t=>t.length>1 && !/^[-–—.\s]+$/.test(t);
    if(rows.some(x=>x.filled>0)) strayRows.forEach(x=>{
      if(okT(x.note)) strays.push({teacher:x.teacher,subject:x.subject,kind:'note',text:x.note});
      if(okT(x.exam)) strays.push({teacher:x.teacher,subject:x.subject,kind:'exam',text:x.exam});
      if(okT(x.rem))  strays.push({teacher:x.teacher,subject:x.subject,kind:'reminder',text:x.rem});
    });
    weeks.push({num,label:'الأسبوع '+adig(num),range,start:start?start.toISOString().slice(0,10):null,rows,strays});
  }
  weeks.sort((a,b)=>a.num-b.num);
  // توحيد أسماء المعلمين (نفس القواعد في plan.js) حتى لا يظهر معلّم واحد كمعلّمين: أحدهما «متأخّر» خطأً
  const recs=[]; weeks.forEach(w=>w.rows.forEach(r=>recs.push({week:w.num,subject:r.subject,teacher:r.teacher})));
  const nm=buildNameMap(recs);
  if(nm.map.size) weeks.forEach(w=>{ w.rows.forEach(r=>{ const c=nm.map.get(r.subject+'||'+r.teacher); if(c) r.teacher=c; });
    (w.strays||[]).forEach(x=>{ const c=nm.map.get(x.subject+'||'+x.teacher); if(c) x.teacher=c; }); });
  // الشعب التي "يدرّسها" كل معلّم/مادة = اتحاد الشعب المعبأة له عبر كل الأسابيع
  const taughtKey={}; // teacher||subject -> Set(codes)
  for(const w of weeks) for(const row of w.rows){
    const k=row.teacher+'||'+row.subject; taughtKey[k]=taughtKey[k]||new Set();
    for(const [code] of SECTIONS) if(row.cells[code]) taughtKey[k].add(code);
  }
  // لكل أسبوع: لكل معلّم/مادة، المطلوب = الشعب التي يدرّسها، المنجز = المعبأ هذا الأسبوع
  for(const w of weeks){
    for(const row of w.rows){
      const req=[...(taughtKey[row.teacher+'||'+row.subject]||new Set())];
      row.required=req;
      row.done=req.filter(c=>row.cells[c]);
      row.status = req.length===0 ? 'none' : (row.done.length===0 ? 'missing' : (row.done.length<req.length ? 'partial' : 'done'));
    }
    const active=w.rows.filter(r=>r.required.length>0);
    w.summary={total:active.length,
      done:active.filter(r=>r.status==='done').length,
      partial:active.filter(r=>r.status==='partial').length,
      missing:active.filter(r=>r.status==='missing').length};
  }
  const {cur,upcoming}=weekPointer(weeks);
  return {sectionOrder:SECTIONS.map(([code,label])=>({code,label})),currentWeekIndex:cur,currentIsUpcoming:upcoming,weeks,nameMerges:nm.merges,nameSuspects:nm.suspects};
}

exports.handler = async (event) => {
  const pw = (event && event.queryStringParameters && event.queryStringParameters.pw) || '';
  if(!ADMIN_PASSWORD) return {statusCode:500,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:'لم تُضبط كلمة السرّ (ADMIN_PASSWORD) في إعدادات Netlify.'})};
  if(pw!==ADMIN_PASSWORD) return {statusCode:401,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify({error:'كلمة السرّ غير صحيحة'})};
  try{
    const buf=await download();
    if(!buf) return {statusCode:502,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:'تعذّر تنزيل ملف الإكسل'})};
    return {statusCode:200,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'},body:JSON.stringify(follow(buf))};
  }catch(e){return {statusCode:500,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:String(e&&e.message||e)})};}
};
exports._follow = follow;
exports._weekPointer = weekPointer;
