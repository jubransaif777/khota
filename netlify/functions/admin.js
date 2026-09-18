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
const C_TEACHER=1, C_SUBJECT=2;

const adig=x=>String(x).replace(/[0-9]/g,d=>'٠١٢٣٤٥٦٧٨٩'[+d]);
const clean=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
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
    const mD=title.match(/(\d{2})-(\d{2})-(\d{4})/);
    const start=mD?new Date(+mD[3],+mD[2]-1,+mD[1]):null;
    let range='';
    if(start){const e=new Date(start.getTime()+4*864e5);
      range=start.getMonth()===e.getMonth()?`${adig(start.getDate())} – ${adig(e.getDate())} ${AR_MONTHS[start.getMonth()]}`
        :`${adig(start.getDate())} ${AR_MONTHS[start.getMonth()]} – ${adig(e.getDate())} ${AR_MONTHS[e.getMonth()]}`;}
    const rows=[];
    for(let r=5;r<=maxRow;r++){
      const subject=fixSubject(cell(ws,r,C_SUBJECT));
      const teacher=clean(cell(ws,r,C_TEACHER));
      if(!teacher && !subject) continue;
      if(!teacher || /^شاغر/.test(teacher)) continue; // تجاهل الصفوف الشاغرة
      const cells={}; let filled=0, teaches=0;
      for(const [code] of SECTIONS){
        const v=clean(cell(ws,r,COL[code]));
        // نعتبر الشعبة "مطلوبة" لهذا المعلّم إن كان يدرّسها في أي أسبوع — نحسبها لاحقًا
        cells[code]=v?1:0; if(v){filled++;}
      }
      rows.push({teacher,subject,cells,filled});
    }
    weeks.push({num,label:'الأسبوع '+adig(num),range,start:start?start.toISOString().slice(0,10):null,rows});
  }
  weeks.sort((a,b)=>a.num-b.num);
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
  const today=new Date();today.setHours(0,0,0,0);let cur=0;
  weeks.forEach((w,i)=>{if(w.start&&new Date(w.start)<=today)cur=i;});
  return {sectionOrder:SECTIONS.map(([code,label])=>({code,label})),currentWeekIndex:cur,weeks};
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
