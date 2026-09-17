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
    const mD=title.match(/(\d{2})-(\d{2})-(\d{4})/);
    const start=mD? new Date(+mD[3],+mD[2]-1,+mD[1]) : null;
    const sections={}; SECTIONS.forEach(([c])=>sections[c]=[]);
    for(let r=5;r<=maxRow;r++){
      const subject=fixSubject(cell(ws,r,C_SUBJECT));
      const teacher=clean(cell(ws,r,C_TEACHER));
      if(!subject && !teacher) continue;
      const note=clean(cell(ws,r,C_NOTE)), exam=clean(cell(ws,r,C_EXAM)), rem=clean(cell(ws,r,C_REM));
      const [emoji,tint]=meta(subject);
      for(const [code] of SECTIONS){
        const lesson=clean(cell(ws,r,COL[code]));
        if(!lesson) continue;
        const e={subject,teacher,lesson,emoji,tint};
        if(note) e.note=note; if(rem) e.reminder=rem; if(exam) e.exam=exam;
        sections[code].push(e);
      }
    }
    weeks.push({num,label:'الأسبوع '+adig(num),
      range:start?fmtRange(start):'', start:start?start.toISOString().slice(0,10):null, sections});
  }
  weeks.sort((a,b)=>a.num-b.num);
  const today=new Date(); today.setHours(0,0,0,0);
  let cur=0; weeks.forEach((w,i)=>{ if(w.start && new Date(w.start)<=today) cur=i; });
  return {school:SCHOOL,term:TERM,currentWeekIndex:cur,
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
