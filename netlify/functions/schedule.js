const XLSX = require('xlsx');

// ============================================================
//  محرّك جدول الحصص — يقرأ ملف الجدول (معلم × يوم × حصة) ويقلبه إلى جدول لكل شعبة
//  رابط مشاركة الملف يُضبط كمتغيّر بيئة في Netlify باسم SCHEDULE_SHARE_URL
//  (أو ضعه هنا مكان النص الفارغ):
// ============================================================
const SHARE_URL = process.env.SCHEDULE_SHARE_URL || "";
const SCHOOL = "مدرسة المعيريض للحلقة الثانية بنين";

const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
const normAr=s=>clean(s).toLowerCase().replace(/[\u064B-\u0652\u0640]/g,'')
  .replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه');

const DAY_LABEL={MON:'الإثنين',TUE:'الثلاثاء',WED:'الأربعاء',THU:'الخميس',FRI:'الجمعة',SAT:'السبت',SUN:'الأحد'};
function dayToken(s){
  const t=normAr(s); if(!t || t.length>14) return null;
  if(/اثن|^mon/.test(t)) return 'MON';
  if(/ثلاث|^tue/.test(t)) return 'TUE';
  if(/اربع|^wed/.test(t)) return 'WED';
  if(/خميس|^thu/.test(t)) return 'THU';
  if(/جمع|^fri/.test(t)) return 'FRI';
  if(/سبت|^sat/.test(t)) return 'SAT';
  if(/^الاحد$|^احد$|^sun/.test(t)) return 'SUN';
  return null;
}
const isSection=s=>/^\d{1,2}[A-Za-z]{1,2}\d{0,2}$/.test(s);
const periodNum=s=>{ const m=clean(s).match(/(\d+)/); return m?+m[1]:null; };

// نمط كل مادة: أيقونة + لون (نفس ألوان بقية صفحات الموقع) + اسم مختصر للخلية
function styleFor(subject){
  const t=normAr(subject);
  if(/اسلام/.test(t)) return {short:'إسلامية',emoji:'🕌',color:'#1F6E4E'};
  if(/عرب/.test(t)) return {short:'عربي',emoji:'📖',color:'#0E5A54'};
  if(/انجليز|english/.test(t)) return {short:'إنجليزي',emoji:'🔤',color:'#2A6F97'};
  if(/اجتماع|دراسات|social/.test(t)) return {short:'دراسات',emoji:'🌍',color:'#3A5A8C'};
  if(/ذكاء|\bai\b|ccdi/.test(t)) return {short:'ذكاء اصطناعي',emoji:'🤖',color:'#6B4F9E'};
  if(/بدن|رياضيه|physical|\bpe\b/.test(t)) return {short:'بدنية',emoji:'⚽',color:'#4F7A3A'};
  if(/رياضيات|math/.test(t)) return {short:'رياضيات',emoji:'➗',color:'#4B4B8F'};
  if(/علوم|science/.test(t)) return {short:'علوم',emoji:'🔬',color:'#2E7D6B'};
  if(/سمع|موسيق|music/.test(t)) return {short:'سمعية',emoji:'🎵',color:'#4A6D8C'};
  if(/بصر|فنون|فني|\bart\b/.test(t)) return {short:'بصرية',emoji:'🎨',color:'#B5546A'};
  if(/مسرح|drama/.test(t)) return {short:'مسرح',emoji:'🎭',color:'#7A4A78'};
  if(/تصميم|تقني|design/.test(t)) return {short:'تصميم',emoji:'⚙️',color:'#8A6D3B'};
  return {short:clean(subject),emoji:'📘',color:'#556B6A'};
}

function timeText(v){
  let s=clean(v); if(!s) return '';
  if(/^0?\.\d+$/.test(s)){ const mins=Math.round(parseFloat(s)*1440); return Math.floor(mins/60)+':'+('0'+(mins%60)).slice(-2); }
  const m=s.match(/(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(AM|PM|am|pm|ص|م)?/);
  if(!m) return s;
  let h=+m[1]; const mi=m[2]; const ap=m[3];
  if(ap){ const pm=/PM|pm|م/.test(ap); if(pm && h<12) h+=12; if(!pm && h===12) h=0; }
  return h+':'+mi;
}

function rowsOf(ws){ return XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,blankrows:true}); }
function sheetNamed(wb,re){ const n=wb.SheetNames.find(x=>re.test(x.trim())); return n?wb.Sheets[n]:null; }

function parseTeachersTab(wb){
  const ws=sheetNamed(wb,/^teachers?$/i); const map={};
  if(!ws||!ws['!ref']) return map;
  const rows=rowsOf(ws);
  let hr=-1, cN=0, cS=1;
  for(let r=0;r<Math.min(rows.length,8);r++){
    let n=-1,s=-1;
    rows[r].forEach((v,c)=>{ const t=normAr(v); if(!t||t.length>40) return; if(/معلم|teacher/.test(t)) n=c; else if(/ماده|subject/.test(t)) s=c; });
    if(n>=0 && s>=0 && n!==s){ hr=r; cN=n; cS=s; break; }
  }
  if(hr<0) return map;
  for(let r=hr+1;r<rows.length;r++){
    const name=clean(rows[r][cN]), sub=clean(rows[r][cS]);
    if(name && sub) map[normAr(name)]=sub;
  }
  return map;
}

function parseTimesTab(wb){
  const out={default:{},byDay:{}};
  const ws=sheetNamed(wb,/^times?$/i);
  if(!ws||!ws['!ref']) return out;
  const rows=rowsOf(ws);
  let hr=-1, cD=-1, cP=-1, cF=-1, cT=-1;
  for(let r=0;r<Math.min(rows.length,8);r++){
    let d=-1,p=-1,f=-1,t=-1;
    rows[r].forEach((v,c)=>{ const x=normAr(v); if(!x||x.length>30) return;
      if(/يوم|day/.test(x)) d=c; else if(/حصه|period/.test(x)) p=c;
      else if(/^من|بدايه|start|from/.test(x)) f=c; else if(/^الي|^الى|نهايه|end|^to/.test(x)) t=c; });
    if(p>=0 && f>=0 && t>=0){ hr=r; cD=d; cP=p; cF=f; cT=t; break; }
  }
  if(hr<0) return out;
  for(let r=hr+1;r<rows.length;r++){
    const row=rows[r]; if(row.join(' ').includes('مثال توضيحي')) continue;
    const p=periodNum(row[cP]); const from=timeText(row[cF]), to=timeText(row[cT]);
    if(!p || !from || !to) continue;
    const dt=cD>=0?dayToken(row[cD]):null;
    if(dt){ (out.byDay[dt]=out.byDay[dt]||{})[p]={from,to}; } else { out.default[p]={from,to}; }
  }
  return out;
}

function sortSections(a,b){
  const key=s=>{ const m=s.match(/^(\d+)([A-Za-z]+)(\d*)$/); if(!m) return [99,9,0,s];
    const track=/^g/i.test(m[2])?0:/^a/i.test(m[2])?1:2; return [+m[1],track,+(m[3]||0),s]; };
  const A=key(a),B=key(b); for(let i=0;i<3;i++) if(A[i]!==B[i]) return A[i]-B[i]; return 0;
}

function build(buf){
  const wb=XLSX.read(buf,{type:'buffer'});
  const subjOf=parseTeachersTab(wb);
  const times=parseTimesTab(wb);

  // تبويب الجدول: schedule إن وُجد، وإلا أول تبويب ليس teachers/times
  let ws=sheetNamed(wb,/^schedule$/i);
  if(!ws){ const n=wb.SheetNames.find(x=>!/^(teachers?|times?)$/i.test(x.trim())); ws=n?wb.Sheets[n]:null; }
  if(!ws||!ws['!ref']) return {status:'badformat'};
  const rows=rowsOf(ws);

  // صف الأيام: أول صف فيه يومان مختلفان على الأقل
  let hr=-1;
  for(let r=0;r<Math.min(rows.length,12);r++){
    const set=new Set(rows[r].map(dayToken).filter(Boolean));
    if(set.size>=2){ hr=r; break; }
  }
  if(hr<0 || hr+1>=rows.length) return {status:'badformat'};
  const head=rows[hr], per=rows[hr+1];
  const width=Math.max(head.length,per.length);
  const colDay={}, colPer={}; const dayOrder=[];
  let cur=null;
  for(let c=0;c<width;c++){
    const t=dayToken(head[c]); if(t){ cur=t; if(!dayOrder.includes(t)) dayOrder.push(t); }
    const p=periodNum(per[c]);
    if(cur && p){ colDay[c]=cur; colPer[c]=p; }
  }
  const validCols=Object.keys(colDay).map(Number);
  if(!validCols.length) return {status:'badformat'};
  // عمود اسم المعلم
  let nameCol=Math.max(0,Math.min(...validCols)-1);
  for(const r of [hr,hr+1]){ rows[r].forEach((v,c)=>{ if(/معلم|teacher/i.test(clean(v))) nameCol=c; }); }

  const teacherNames=[]; const grid={}; const used={};
  const add=(sec,day,p,tname)=>{
    ((grid[sec]=grid[sec]||{})[day]=grid[sec][day]||{});
    (grid[sec][day][p]=grid[sec][day][p]||[]).push(tname);
    used[day]=Math.max(used[day]||0,p);
  };
  for(let r=hr+2;r<rows.length;r++){
    const name=clean(rows[r][nameCol]); if(!name || /^(المجموع|total)/i.test(name)) continue;
    let any=false;
    for(const c of validCols){
      const cellTxt=clean(rows[r][c]); if(!cellTxt) continue;
      for(const code of cellTxt.split(/[\/,،;\s]+/).filter(isSection)){ add(code.toUpperCase(),colDay[c],colPer[c],name); any=true; }
    }
    if(any && !teacherNames.includes(name)) teacherNames.push(name);
  }

  // جدول المعلمين الموحّد
  const unmapped=[];
  const teachers=teacherNames.map(name=>{
    const vacant=/^(شاغر|vacan)/i.test(name);
    const subject=subjOf[normAr(name)]||'';
    if(!subject && !vacant) unmapped.push(name);
    const st=subject?styleFor(subject):{short:vacant?'حصة':name.split(' ')[0],emoji:vacant?'📘':'👤',color:'#6B7A78'};
    return {name:vacant?'':name, subject, short:st.short, emoji:st.emoji, color:st.color, vacant};
  });
  const idx={}; teacherNames.forEach((n,i)=>idx[n]=i);

  const sections=Object.keys(grid).sort(sortSections);
  const days=dayOrder.filter(d=>used[d]).map(d=>({token:d,label:DAY_LABEL[d],periods:used[d]}));
  const outGrid={};
  for(const sec of sections){
    outGrid[sec]={};
    for(const d of days){
      const arr=[];
      for(let p=1;p<=d.periods;p++){
        const list=(((grid[sec]||{})[d.token]||{})[p]||[]).map(n=>idx[n]);
        arr.push(list.length?list:null);
      }
      outGrid[sec][d.token]=arr;
    }
  }
  return {status:'ok', school:SCHOOL, sections, days, teachers, grid:outGrid, times, warnings:{unmapped}};
}

// ---------- تنزيل الملف من SharePoint (نفس أسلوب بقية المحرّكات) ----------
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

exports.handler = async () => {
  const J=(code,obj,cache)=>({statusCode:code,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':cache||'no-store'},body:JSON.stringify(obj)});
  if(!SHARE_URL) return J(200,{status:'nourl'});
  try{
    const buf=await download();
    if(!buf) return J(502,{status:'nofile'});
    const data=build(buf);
    if(data.status!=='ok') return J(500,data);
    return J(200,data,'public, max-age=300');
  }catch(e){ return J(500,{status:'error'}); }
};
exports._build=build;
