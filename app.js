import {initializeApp} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import {getAuth,onAuthStateChanged,signInWithEmailAndPassword,createUserWithEmailAndPassword,sendPasswordResetEmail,sendEmailVerification,updateProfile,signOut,GoogleAuthProvider,signInWithPopup} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import {initializeFirestore,persistentLocalCache,persistentMultipleTabManager,doc as fdoc,collection as fcol,getDoc,getDocs,setDoc,updateDoc,deleteDoc,onSnapshot as fsnap,query as fquery,orderBy as forder,limit as flimit} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import {firebaseConfig,OWNER_EMAIL,ENV} from './firebase-config.js';
const APP_VERSION='2.0.10292107-test';


const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const XMLNS_XML='http://www.w3.org/XML/1998/namespace';
const STATUSES=['DONE','ONGOING','NOT DONE'];
const BLOCKERS=[
  ['Scaffold not ready','Scaffolding'],
  ['Permit / isolation','Permits'],
  ['Crane / lifting not available','Lifting'],
  ['Lifting plan / AP','Lifting'],
  ['QC hold / inspection','QC'],
  ['Material not on site','Materials'],
  ['Access / area clash','Area coordination'],
  ['Manpower','Resourcing'],
  ['Drawing / engineering query','Engineering'],
  ['Time ran out','Planning'],
  ['Other','Other']
];
const STORE_KEY='dabsNight.v1';

let doc=null;      // parsed DABS {fileName, sections:[{title,trades,items}]}
let bytes=null;    // original file Uint8Array
let st={};         // per item state {status:{col:val}, blockers:[], note:'', extra:[cols], open:bool}
let prepBy='';
let filter={view:'all',level:null};
let added=[];      // activities added on the night, not on the DABS
const itemsOf=si=>[...doc.sections[si].items,...added.filter(a=>a.sec===si)];
const allItems=()=>doc.sections.flatMap((s,i)=>itemsOf(i));

/* ---------- utils ---------- */
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function colNum(c){let n=0;for(const ch of c)n=n*26+ch.charCodeAt(0)-64;return n}
function colName(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26)}return s}
function splitRef(r){const m=/^([A-Z]+)(\d+)$/.exec(r);return [colNum(m[1]),+m[2]]}
const parseXml=s=>new DOMParser().parseFromString(s,'application/xml');
const serXml=d=>new XMLSerializer().serializeToString(d);
const kids=(el,name)=>Array.from(el.children).filter(c=>c.localName===name);
const kid=(el,name)=>kids(el,name)[0]||null;
function toast(msg){const t=$('#toast');t.textContent=msg;t.hidden=false;clearTimeout(toast._t);toast._t=setTimeout(()=>t.hidden=true,2200)}
function b64enc(u8){let s='';for(let i=0;i<u8.length;i+=0x8000)s+=String.fromCharCode.apply(null,u8.subarray(i,i+0x8000));return btoa(s)}
function b64dec(b){const s=atob(b);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u}
function fmtLevel(v){if(typeof v==='number'){const s=v.toFixed(2);return v>0?'+'+s:s}const t=String(v??'').trim();return t||'—'}
function rowStatus(it){
  const s=st[it.id];const vals=allCols(it).map(c=>s.status[c]||null);
  if(!vals.length)return null;
  if(vals.every(v=>v==='DONE'))return 'DONE';
  if(vals.every(v=>v==='NOT DONE'))return 'NOT DONE';
  if(vals.some(v=>v==='DONE'||v==='ONGOING'))return 'ONGOING';
  if(vals.every(v=>v===null))return null;
  return 'NOT DONE';
}
const allCols=it=>[...it.trades.map(t=>t.col),...(st[it.id]?.extra||[])];

/* ---------- xlsx reading ---------- */
async function sheetPaths(zip){
  const wb=parseXml(await zip.file('xl/workbook.xml').async('string'));
  const sheet=wb.getElementsByTagNameNS(NS,'sheet')[0];
  const rid=sheet.getAttributeNS(RNS,'id');
  const rels=parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  let target=null;
  for(const r of rels.getElementsByTagNameNS(PKG,'Relationship'))if(r.getAttribute('Id')===rid)target=r.getAttribute('Target');
  if(!target)throw new Error('Could not find the first worksheet in this file.');
  const path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
  return {path,name:sheet.getAttribute('name')};
}
async function readShared(zip){
  const f=zip.file('xl/sharedStrings.xml');if(!f)return [];
  const d=parseXml(await f.async('string'));
  return Array.from(d.getElementsByTagNameNS(NS,'si')).map(si=>{
    let s='';for(const t of si.getElementsByTagNameNS(NS,'t')){if(t.parentNode.localName==='rPh')continue;s+=t.textContent}return s});
}
function cellValue(c,ss){
  const t=c.getAttribute('t');
  if(t==='inlineStr'){let s='';for(const x of c.getElementsByTagNameNS(NS,'t'))s+=x.textContent;return s}
  const v=kid(c,'v');if(!v)return null;
  if(t==='s')return ss[+v.textContent]??'';
  if(t==='str'||t==='e')return v.textContent;
  if(t==='b')return v.textContent==='1';
  const n=Number(v.textContent);return isNaN(n)?v.textContent:n;
}
async function parseDabs(u8,fileName){
  const zip=await JSZip.loadAsync(u8);
  const {path}=await sheetPaths(zip);
  const ss=await readShared(zip);
  const sx=parseXml(await zip.file(path).async('string'));
  const grid=new Map();let maxRow=0;
  for(const c of sx.getElementsByTagNameNS(NS,'c')){
    const [col,row]=splitRef(c.getAttribute('r'));const v=cellValue(c,ss);
    if(v!==null&&v!=='')grid.set(row+','+col,v);if(row>maxRow)maxRow=row;
  }
  const own=(r,c)=>{const v=grid.get(r+','+c);return v==null?'':v};
  const merged=new Map(grid);
  for(const m of sx.getElementsByTagNameNS(NS,'mergeCell')){
    const [a,b]=m.getAttribute('ref').split(':');if(!b)continue;
    const [c1,r1]=splitRef(a),[c2,r2]=splitRef(b);const v=grid.get(r1+','+c1);
    if(v==null)continue;for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)merged.set(r+','+c,v);
  }
  const mv=(r,c)=>{const v=merged.get(r+','+c);return v==null?'':v};
  const txt=v=>String(v??'').trim();
  // header rows
  const headers=[];
  for(let r=1;r<=maxRow;r++){
    for(let c=1;c<=60;c++){if(txt(own(r,c)).toLowerCase()==='activity description'){headers.push({row:r,actCol:c});break}}
  }
  if(!headers.length)throw new Error('No "Activity Description" header found. Is this a DABS handover sheet?');
  const sections=[];
  headers.forEach((h,i)=>{
    const cols={};const hdr=[];
    for(let c=1;c<=60;c++){const t=txt(own(h.row,c));if(t)hdr.push([c,t])}
    for(const [c,t] of hdr){const l=t.toLowerCase();
      if(l==='level')cols.level=c;else if(l==='shift')cols.shift=c;else if(l.startsWith('black'))cols.bh=c;
      else if(l.includes('remark')||l.includes('blocker'))cols.remark=c;}
    cols.act=h.actCol;
    const before=hdr.filter(([c])=>c<h.actCol&&![cols.level,cols.shift,cols.bh].includes(c));
    cols.tag=before.length?before[before.length-1][0]:h.actCol-1;
    const tagLabel=before.length?before[before.length-1][1]:'Tag';
    const trades=hdr.filter(([c])=>c>h.actCol&&c!==cols.remark&&(!cols.remark||c<cols.remark)).map(([c,t])=>({col:c,name:t}));
    if(!cols.remark)cols.remark=(trades.length?trades[trades.length-1].col:h.actCol)+1;
    let title='';let titleRow=h.row;
    for(let r=h.row-1;r>=Math.max(1,h.row-4);r--){const vals=[];for(let c=1;c<=20;c++){const t=txt(own(r,c));if(t)vals.push(t)}
      if(vals.length){title=vals[0];titleRow=r;for(let c=1;c<=20;c++)if(txt(own(r,c))){h.titleCol=c;break}break}}
    h.titleRow=titleRow;
    sections.push({title:title||('Section '+(i+1)),header:h.row,titleRow,titleCol:h.titleCol||cols.level||1,hdr,cols,trades,tagLabel,items:[]});
  });
  sections.forEach((s,i)=>{
    const end=i+1<headers.length?headers[i+1].titleRow-1:maxRow;
    for(let r=s.header+1;r<=end;r++){
      const act=txt(own(r,s.cols.act)),tag=txt(own(r,s.cols.tag));
      if(!act&&!tag)continue;
      const trades=[];const pre={};
      for(const t of s.trades){const v=txt(own(r,t.col)).toUpperCase();
        if(v==='X'||STATUSES.includes(v)){trades.push(t);if(STATUSES.includes(v))pre[t.col]=v}}
      s.items.push({id:i+'-'+r,row:r,sec:i,level:fmtLevel(s.cols.level?own(r,s.cols.level):''),
        shift:txt(own(r,s.cols.shift)),bh:txt(mv(r,s.cols.bh)),tag,act,dayRemark:txt(own(r,s.cols.remark)),trades,pre});
    }
  });
  const shortName=t=>t.replace(/^dabs\s*/i,'').replace(/\s+\d.*$/,'').trim()||t;
  sections.forEach(s=>s.short=shortName(s.title));
  return {fileName,sections};
}

/* ---------- state ---------- */
function blankState(it){return {status:{...(it.pre||{})},blockers:[],note:'',extra:[],hist:[]}}
function initState(){for(const it of allItems())if(!st[it.id])st[it.id]=blankState(it)}

/* ---------- where the data lives (test copy uses its own collections) ---------- */
const P=n=>(ENV==='live'?'':'test_')+n;
const LIVE=n=>n;

/* ---------- shared store (Firebase, works offline and syncs when signal returns) ---------- */
let db=null,authUser=null,canWrite=false,isAdmin=false,ready=false;
let nights=[],olderNights=[],cur=null,unsubItems=null,unsubNights=null,opening=false;
const NIGHTS_ON_OPEN=30;
const meta={currentId:null,prepBy:''};
const LOCAL='dabsNight.pwa.'+ENV;
function loadLocal(){try{Object.assign(meta,JSON.parse(localStorage.getItem(LOCAL)||'{}'))}catch(e){}}
function saveLocal(){try{localStorage.setItem(LOCAL,JSON.stringify(meta))}catch(e){}}
const myName=()=>(authUser&&(authUser.displayName||(authUser.email||'').split('@')[0]))||'';
const allNights=()=>{const m=new Map();for(const n of [...olderNights,...nights])m.set(n.id,n);return [...m.values()]};
const byDate=()=>allNights().sort((a,b)=>b.loadedAt-a.loadedAt);
const latest=()=>nights.slice().sort((a,b)=>b.loadedAt-a.loadedAt)[0]||null;
const fmtTime=t=>new Date(t).toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'});
const fmtDT=t=>new Date(t).toLocaleString('en-GB',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'});
const MONTHS=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
function nightLabel(d,loadedAt){
  const m=/(\d{1,2})(?:st|nd|rd|th)?[\s\-.]*([A-Za-z]{3})/.exec(d.sections.map(s=>s.title).join(' '));
  if(m)return m[1]+' '+m[2][0].toUpperCase()+m[2].slice(1,3).toLowerCase();
  return new Date(loadedAt).toLocaleDateString('en-GB',{day:'numeric',month:'short'});
}
/* The night's date, read from the DABS title ("28th-Oct Night"); the year is the one closest to the upload day. */
function dateKeyFrom(text,around){
  const m=/(\d{1,2})(?:st|nd|rd|th)?[\s\-.]*([A-Za-z]{3})/.exec(String(text||''));
  if(!m)return null;const mon=MONTHS.indexOf(m[2].toLowerCase());const day=+m[1];
  if(mon<0||day<1||day>31)return null;
  const base=new Date(around||Date.now());let best=null;
  for(const y of [base.getFullYear()-1,base.getFullYear(),base.getFullYear()+1]){
    const t=new Date(y,mon,day).getTime();if(!best||Math.abs(t-base)<Math.abs(best-base))best=t}
  const d=new Date(best);return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
const docDateKey=(d,around)=>dateKeyFrom(d.sections.map(s=>s.title).join(' '),around);
const nightDateKey=n=>n.dateKey||dateKeyFrom(n.label,n.loadedAt);
function hashStr(s){let h=0x811c9dc5;for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619)}return (h>>>0).toString(36)}
const tradeNames=it=>{const sec=doc.sections[it.sec];return allCols(it).map(c=>{const t=(sec.trades.find(x=>x.col===c)||it.trades.find(x=>x.col===c));return t?t.name:''}).filter(Boolean)};
function compactItems(){return allItems().map(it=>{const s=st[it.id]||{};return {t:it.tag,a:it.act,l:it.level,s:rowStatus(it),b:s.blockers||[],x:!!it.added,tr:tradeNames(it),n:(s.note||'').slice(0,160),dr:(it.dayRemark||'').slice(0,120)}})}

/* ---------- sync status, with the time of the last change anyone made ---------- */
let localPending=false;
function lastEdit(){let t=0;for(const k in st){const a=st[k]&&st[k].at;if(a>t)t=a}return t}
function setSync(s){const el=$('#sync');if(!el)return;el.dataset.s=s;
  const le=lastEdit();const tail=le&&doc?' · last edit '+fmtTime(le):'';
  el.textContent=({ok:'Synced',saving:'Sending…',offline:'Saved on this phone · sends when signal returns',ro:'View only',load:'Connecting…'}[s]||'')+(s==='load'?'':tail)}
function updateSync(){if(!authUser)return;if(!canWrite){setSync('ro');return}setSync(localPending?(navigator.onLine?'saving':'offline'):'ok')}
window.addEventListener('online',updateSync);window.addEventListener('offline',updateSync);

/* ---------- writing one activity; every change is added to its history ---------- */
const timers={},dirty=new Set(),base={};
function tradeLabel(it,col){const sec=doc.sections[it.sec];const t=sec.trades.find(x=>x.col===col)||(it.trades||[]).find(x=>x.col===col);return t?t.name:colName(col)}
function describeChanges(it,b,s){
  const out=[];b=b||{status:{},blockers:[],note:''};
  for(const col of new Set([...Object.keys(b.status||{}),...Object.keys(s.status||{})])){
    const o=(b.status||{})[col]||null,n=(s.status||{})[col]||null;
    if(o!==n)out.push(tradeLabel(it,+col)+': '+(o||'—')+' → '+(n||'—'))}
  for(const x of s.blockers||[])if(!(b.blockers||[]).includes(x))out.push('Blocker added: '+x);
  for(const x of b.blockers||[])if(!(s.blockers||[]).includes(x))out.push('Blocker removed: '+x);
  if((b.note||'')!==(s.note||''))out.push((b.note?'Note changed':'Note added')+': '+(s.note||'(cleared)').slice(0,80));
  return out;
}
function writeItem(id,delay){
  if(!cur||!canWrite)return;const nid=cur.id;
  dirty.add(id);clearTimeout(timers[id]);
  timers[id]=setTimeout(()=>{
    dirty.delete(id);const s=st[id];const it=findItem(id);if(!s||!it||!cur||cur.id!==nid)return;
    const ch=describeChanges(it,base[id],s);
    const now=Date.now();
    if(ch.length){s.hist=[...(s.hist||[]),{at:now,by:authUser.uid,byName:myName(),ch}].slice(-25)}
    const body={status:s.status,blockers:s.blockers,note:s.note,extra:s.extra,open:!!s.open,hist:s.hist||[],by:authUser.uid,byName:myName(),at:now};
    if(it.added)body.added={sec:it.sec,level:it.level,tag:it.tag,act:it.act,trades:it.trades,created:it.created||now,createdByName:it.createdByName||myName()};
    s.by=body.by;s.byName=body.byName;s.at=body.at;
    base[id]=JSON.parse(JSON.stringify({status:s.status,blockers:s.blockers,note:s.note}));
    db.doc(P('nights')+'/'+nid+'/items/'+id).set(body).catch(e=>toast(e&&e.code==='permission-denied'?'Not saved: your access was changed. Ask the night shift lead.':'Not saved ('+((e&&e.code)||'error')+')'));
    localPending=true;updateSync();scheduleSummary();
  },delay||0);
}
let sumT=null;
function monthKey(t){const d=new Date(t);return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')}
function scheduleSummary(){clearTimeout(sumT);sumT=setTimeout(()=>{
  if(!cur||!canWrite)return;const items=compactItems();
  db.doc(P('nights')+'/'+cur.id).update({summary:counts(),items,updatedAt:Date.now()}).catch(()=>{});
  writeTagHist(cur,items);
},1500)}
function writeTagHist(n,items){
  const entry={label:n.label||'',at:n.loadedAt,dateKey:nightDateKey(n)||'',items:items||n.items||[]};
  db.doc(P('tagHist')+'/'+monthKey(n.loadedAt)).merge({nights:{[n.id]:entry}}).catch(()=>{});
  tagHist.set(n.id,entry);histIdx=null;
}

function applyItemDoc(id,d){
  if(d.added){
    const info={id,sec:d.added.sec,added:true,level:d.added.level,shift:'Night',bh:'',tag:d.added.tag||'',act:d.added.act||'',dayRemark:'',trades:d.added.trades||[],pre:{},created:d.added.created,createdByName:d.added.createdByName};
    const a=added.find(x=>x.id===id);if(a)Object.assign(a,info);else added.push(info);
  }
  if(dirty.has(id))return false;      // this phone has a newer edit waiting to be written
  const focused=document.activeElement&&document.activeElement.id==='note-'+id;
  st[id]={status:{...(d.status||{})},blockers:[...(d.blockers||[])],note:focused&&st[id]?st[id].note:(d.note||''),extra:[...(d.extra||[])],open:!!d.open,hist:[...(d.hist||[])],by:d.by||null,byName:d.byName||'',at:d.at||0};
  base[id]={status:{...(d.status||{})},blockers:[...(d.blockers||[])],note:d.note||''};
  return true;
}
function subscribeItems(){
  if(unsubItems)unsubItems();
  const nid=cur.id;let first=true;
  unsubItems=db.collection(P('nights')+'/'+nid+'/items').onSnapshot(snap=>{
    if(!cur||cur.id!==nid)return;
    localPending=!!(snap.metadata&&snap.metadata.hasPendingWrites);
    let structural=first;const changed=[];
    for(const ch of snap.docChanges()){
      const id=ch.doc.id;
      if(ch.type==='removed'){
        if(added.some(a=>a.id===id)){added=added.filter(a=>a.id!==id);delete st[id];structural=true}
        else{const it=findItem(id);if(it&&!dirty.has(id)){st[id]=blankState(it);delete base[id];changed.push(id)}}
        continue;
      }
      const d=ch.doc.data()||{};
      if(d.added&&!added.some(a=>a.id===id))structural=true;
      if(applyItemDoc(id,d))changed.push(id);
    }
    first=false;initState();updateSync();
    if(structural)renderAll();else{changed.forEach(id=>rerenderCard(id,true));renderTop()}
  },e=>{toast('Live updates stopped ('+(e.code||'error')+'). Reopen the app.')},true);
}
function showNoAccess(){
  $('#main').innerHTML=`<section class="empty"><h2>No access to the team's DABS</h2>
    <p>${isAdmin?'Firebase is refusing access. In the Firebase console open <strong>Firestore Database → Rules</strong>, paste the contents of <strong>firestore.rules</strong>, press <strong>Publish</strong> and reopen this app.':'Your access may have changed. Ask the night shift lead, then reopen the app.'}</p>
    <p>Signed in as <strong>${esc(authUser&&authUser.email||'')}</strong>.</p>
    <button class="btn alt" type="button" id="naOut">Sign out</button></section>`;
  $('#naOut').onclick=fbSignOut;
}
let backfilled=new Set();
function subscribeNights(){
  if(unsubNights)return;
  unsubNights=db.collection(P('nights')).orderBy('loadedAt','desc').limit(NIGHTS_ON_OPEN).onSnapshot(async snap=>{
    nights=snap.docs.map(d=>Object.assign({id:d.id},d.data()));histIdx=null;ready=true;
    if(canWrite)for(const n of nights)if(n.items&&n.items.length&&!tagHist.has(n.id)&&!backfilled.has(n.id)&&tagHistReady){backfilled.add(n.id);writeTagHist(n,n.items)}
    if(opening)return;
    if(!cur){const want=nights.find(n=>n.id===meta.currentId)||latest();
      if(want){try{await openNight(want)}catch(e){renderAll();toast(e.message)}}else renderAll();return}
    const c=allNights().find(n=>n.id===cur.id);
    if(!c){if(unsubItems)unsubItems();unsubItems=null;cur=null;doc=null;bytes=null;st={};added=[];
      const l=latest();if(l)await openNight(l);else renderAll();toast('That night was deleted by an admin');return}
    const fileChanged=(c.fileId||c.id)!==(cur.fileId||cur.id);
    Object.assign(cur,c);
    if(fileChanged){try{await openNight(cur);toast('The DABS for '+cur.label+' was replaced by '+(cur.replacedByName||'an admin'))}catch(e){toast(e.message)}return}
    renderTop();renderBanner();if(sheetView==='archive')openArchive();
  },e=>{if(e.code==='permission-denied'){showNoAccess();return}
    $('#main').innerHTML=`<p class="err">Could not load the team's DABS (${esc(e.code||'error')}). Reopen the app.</p>`});
}
let tagHist=new Map(),tagHistReady=false,unsubTagHist=null;
function subscribeTagHist(){
  if(unsubTagHist)return;
  unsubTagHist=db.collection(P('tagHist')).onSnapshot(snap=>{
    tagHist=new Map();for(const d of snap.docs){const ns=(d.data()||{}).nights||{};for(const k in ns)tagHist.set(k,ns[k])}
    tagHistReady=true;histIdx=null;
    if(canWrite)for(const n of nights)if(n.items&&n.items.length&&!tagHist.has(n.id)&&!backfilled.has(n.id)){backfilled.add(n.id);writeTagHist(n,n.items)}
    if(doc)renderMain();
  },()=>{});
}
async function loadOlderNights(){
  const s=await db.collection(P('nights')).orderBy('loadedAt','desc').limit(500).get();
  olderNights=s.docs.map(d=>Object.assign({id:d.id},d.data()));
}
function stopListening(){for(const u of [unsubItems,unsubNights,unsubTagHist,unsubConcerns,unsubPeople,unsubSettings])if(u)u();
  unsubItems=unsubNights=unsubTagHist=unsubConcerns=unsubPeople=unsubSettings=null;nights=[];olderNights=[];cur=null;doc=null;bytes=null;st={};added=[];ready=false;tagHistReady=false}
async function openNight(n){
  opening=true;
  try{
    const f=await db.doc(P('files')+'/'+(n.fileId||n.id)).get();
    if(!f.exists)throw new Error('The DABS file for '+(n.label||'this night')+' is not on this phone yet. Open it once with signal.');
    const u8=b64dec(f.data().b64);const d=await parseDabs(u8,n.fileName);
    if(unsubItems){unsubItems();unsubItems=null}
    bytes=u8;doc=d;added=[];st={};dirty.clear();for(const k in base)delete base[k];initState();
    cur=n;meta.currentId=n.id;saveLocal();filter={view:'all',level:null};histIdx=null;
    renderAll();subscribeItems();
  }finally{opening=false}
}

/* ---------- loading a DABS: one per night ---------- */
async function findExistingNight(key){
  if(!key)return null;
  const lock=await db.doc(P('dates')+'/'+key).get();
  if(lock.exists){const nid=lock.data().nightId;
    let n=allNights().find(x=>x.id===nid);
    if(!n){const nd=await db.doc(P('nights')+'/'+nid).get();if(nd.exists)n=Object.assign({id:nid},nd.data())}
    if(n)return n}
  return allNights().find(x=>nightDateKey(x)===key)||null;
}
async function uploadDabs(f){
  const u8=new Uint8Array(await f.arrayBuffer());const d=await parseDabs(u8,f.name);const b64=b64enc(u8);
  if(b64.length>900000)throw new Error('This file is too large to share (over 650 KB). Remove pictures or extra sheets and try again.');
  const now=Date.now();const key=docDateKey(d,now);
  const fileId='f'+hashStr(b64)+u8.length.toString(36);
  if(!navigator.onLine)throw new Error('No signal. Loading a DABS needs a connection; try again near the cabins or on Wi-Fi.');
  const ex=await findExistingNight(key);
  if(ex){
    await openNight(ex);
    const sameFile=(ex.fileId||ex.id)===fileId;
    openSheet(`<h3>${esc(ex.label)} DABS is already loaded</h3>
      <p>Loaded by <strong>${esc(ex.createdByName||'a teammate')}</strong> at ${esc(fmtDT(ex.loadedAt))}. ${sameFile?'This is the same file.':'Only one DABS can be loaded per night.'}</p>
      ${isAdmin&&!sameFile?`<p>If the day shift sent a corrected DABS, you can replace it. Statuses, blockers, notes and added activities are kept for jobs with the same tag.</p>
        <div class="grid2"><button class="btn alt" type="button" id="exKeep">Keep current</button><button class="btn" type="button" id="exRep">Replace DABS</button></div>`
        :`<button class="btn" type="button" id="exKeep">OK</button>`}`);
    $('#exKeep').onclick=closeSheet;
    const rp=$('#exRep');if(rp)rp.onclick=()=>confirmReplace(ex,u8,d,f.name,b64,fileId);
    return;
  }
  if(!canWrite)throw new Error('You can view the DABS but not add files. Ask the night shift lead for Contributor access.');
  const nid=key?'n'+key:'n'+fileId;
  if(key){
    try{await db.doc(P('dates')+'/'+key).create({nightId:nid,by:authUser.uid,byName:myName(),at:now})}
    catch(e){const again=await findExistingNight(key);if(again){await openNight(again);toast(again.label+' DABS was just loaded by '+(again.createdByName||'a teammate'));return}throw e}
  }
  const m={fileName:f.name,fileId,dateKey:key||'',label:nightLabel(d,now),loadedAt:now,updatedAt:now,createdBy:authUser.uid,createdByName:myName(),summary:null,items:[]};
  await db.doc(P('files')+'/'+fileId).set({b64,fileName:f.name});
  await db.doc(P('nights')+'/'+nid).set(m);
  await openNight(Object.assign({id:nid},m));
  scheduleSummary();
  toast(d.sections.reduce((k,s)=>k+s.items.length,0)+' activities shared with the team');
}
const normKey=(t,a)=>String(t||'').toUpperCase().replace(/\s+/g,'')+'|'+String(a||'').toLowerCase().replace(/\s+/g,' ').trim();
function confirmReplace(n,u8,d,fileName,b64,fileId){
  const k=docDateKey(d,n.loadedAt),nk=nightDateKey(n);
  if(k&&nk&&k!==nk){openSheet(`<h3>Different night</h3><p>This file is for <strong>${esc(nightLabel(d,Date.now()))}</strong>, but the night on screen is <strong>${esc(n.label)}</strong>. Load it as a new night instead.</p><button class="btn" type="button" id="cOk">OK</button>`);$('#cOk').onclick=closeSheet;return}
  openSheet(`<h3>Replace the ${esc(n.label)} DABS?</h3>
    <p>The whole team will switch to <strong>${esc(fileName)}</strong>. Statuses, blockers, notes and history are moved to jobs with the same tag and activity. Jobs that are no longer on the new DABS but already have entries are kept as added activities, so nothing entered tonight is lost.</p>
    <div class="grid2"><button class="btn alt" type="button" id="rpNo">Cancel</button><button class="btn danger" type="button" id="rpYes">Replace</button></div>`);
  $('#rpNo').onclick=closeSheet;
  $('#rpYes').onclick=async()=>{const b=$('#rpYes');b.disabled=true;b.textContent='Replacing…';
    try{await replaceDabs(n,u8,d,fileName,b64,fileId);closeSheet();toast('DABS replaced for the whole team')}
    catch(e){toast('Replace failed ('+(e.code||e.message||'error')+')');closeSheet()}};
}
async function replaceDabs(n,u8,d,fileName,b64,fileId){
  const oldF=await db.doc(P('files')+'/'+(n.fileId||n.id)).get();
  const oldDoc=await parseDabs(b64dec(oldF.data().b64),n.fileName);
  const oldItems=new Map();oldDoc.sections.forEach((s,si)=>s.items.forEach(it=>oldItems.set(it.id,it)));
  const newIdx=new Map();d.sections.forEach((s,si)=>s.items.forEach(it=>{const k=normKey(it.tag,it.act);if(!newIdx.has(k))newIdx.set(k,it)}));
  const snap=await db.collection(P('nights')+'/'+n.id+'/items').get();
  const writes=[],deletes=[];const used=new Set();
  for(const docu of snap.docs){
    const data=docu.data()||{};if(data.added){continue}
    const old=oldItems.get(docu.id);if(!old){continue}
    const tgt=newIdx.get(normKey(old.tag,old.act));
    const hasData=Object.values(data.status||{}).some(Boolean)||(data.blockers||[]).length||data.note;
    if(tgt&&!used.has(tgt.id)){used.add(tgt.id);
      if(tgt.id!==docu.id){writes.push([tgt.id,data]);deletes.push(docu.id)}}
    else if(hasData){
      const aid='r'+hashStr(docu.id+n.id)+Date.now().toString(36);
      const sec=doc.sections[old.sec]||doc.sections[0];
      writes.push([aid,Object.assign({},data,{added:{sec:Math.min(old.sec,d.sections.length-1),level:old.level,tag:old.tag,act:old.act,trades:old.trades.map(t=>({col:t.col,name:t.name})),created:Date.now(),createdByName:myName()+' (from replaced DABS)'}})]);
      deletes.push(docu.id);
    }else deletes.push(docu.id);
  }
  await db.doc(P('files')+'/'+fileId).set({b64,fileName});
  for(const [id,data] of writes)await db.doc(P('nights')+'/'+n.id+'/items/'+id).set(data);
  for(const id of deletes)if(!writes.some(w=>w[0]===id))await db.doc(P('nights')+'/'+n.id+'/items/'+id).delete();
  await db.doc(P('nights')+'/'+n.id).update({fileId,fileName,replacedAt:Date.now(),replacedByName:myName(),updatedAt:Date.now()});
  n.fileId=fileId;n.fileName=fileName;n.replacedByName=myName();
  await openNight(n);scheduleSummary();
}

/* ---------- history of a tag across the whole year ---------- */
let histIdx=null;
function tokens(tag){
  const out=[];let prefix='';
  for(let p of String(tag||'').toUpperCase().split(/&|,|\+|\bAND\b/)){
    p=p.replace(/\s+/g,'');if(!p||!/\d/.test(p))continue;
    const m=/^(.*?--)(.+)$/.exec(p);
    if(m){prefix=m[1];out.push(p)}else if(prefix&&/^[A-Z]{1,3}\d+$/.test(p))out.push(prefix+p);else out.push(p);
  }
  return out;
}
function buildHist(){
  histIdx=new Map();if(!cur)return;
  const src=new Map(tagHist);
  for(const n of allNights())if(!src.has(n.id)&&n.items)src.set(n.id,{label:n.label,at:n.loadedAt,items:n.items});
  for(const [id,n] of src){if(id===cur.id||n.at>=cur.loadedAt)continue;
    for(const it of n.items||[])for(const tk of tokens(it.t)){
      const arr=histIdx.get(tk)||[];arr.push({id,label:n.label||'',at:n.at,a:it.a,s:it.s,b:it.b||[]});histIdx.set(tk,arr)}}
}
function historyOf(tag){
  if(!histIdx)buildHist();const seen=new Map();
  for(const tk of tokens(tag))for(const h of histIdx.get(tk)||[]){const k=h.id+'|'+h.a;if(!seen.has(k))seen.set(k,h)}
  return [...seen.values()].sort((a,b)=>b.at-a.at);
}
const nightsNotDone=h=>new Set(h.filter(x=>x.s!=='DONE').map(x=>x.id)).size;

/* ---------- rendering ---------- */
function counts(){
  const c={DONE:0,ONGOING:0,'NOT DONE':0,unset:0,total:0};
  for(const it of allItems()){const r=rowStatus(it);c.total++;if(r)c[r]++;else c.unset++}
  return c;
}
function renderTop(){
  $('#fileName').textContent=doc?(cur&&cur.label?cur.label+' · ':'')+doc.fileName:(authUser?'No DABS loaded':'');
  $('#btnMenu').hidden=!authUser||!ready;
  $('#btnNew').textContent=doc?'New DABS':'Load DABS';$('#btnNew').hidden=!authUser||!canWrite;updateConcernBadge();updateSync();
  $('#progressBox').hidden=!doc;$('#filters').hidden=!doc;$('#actions').hidden=!doc;
  if(!doc)return;
  const c=counts();const pct=n=>c.total?(n/c.total*100)+'%':'0';
  $('#bar').innerHTML=`<span style="width:${pct(c.DONE)};background:var(--ok)"></span><span style="width:${pct(c.ONGOING)};background:var(--go)"></span><span style="width:${pct(c['NOT DONE'])};background:var(--no)"></span>`;
  $('#counts').innerHTML=`<span><i class="dot" style="background:var(--ok)"></i>DONE <b>${c.DONE}</b></span><span><i class="dot" style="background:var(--go)"></i>ONGOING <b>${c.ONGOING}</b></span><span><i class="dot" style="background:var(--no)"></i>NOT DONE <b>${c['NOT DONE']}</b></span><span><i class="dot" style="background:var(--unset)"></i>Not updated <b>${c.unset}</b></span><span>of <b>${c.total}</b> activities</span>`;
}
function renderFilters(){
  const levels=[...new Set(allItems().map(i=>i.level))].sort((a,b)=>parseFloat(a)-parseFloat(b));
  const f=[['all','All'],['open','Open items'],...doc.sections.map((s,i)=>['sec'+i,s.short])];
  $('#filters').innerHTML=f.map(([k,l])=>`<button class="chip" type="button" data-view="${k}" aria-pressed="${filter.view===k}">${esc(l)}</button>`).join('')
   +levels.map(l=>`<button class="chip" type="button" data-level="${esc(l)}" aria-pressed="${filter.level===l}">Lvl ${esc(l)}</button>`).join('');
}
function visible(it){
  if(filter.level!==null&&it.level!==filter.level)return false;
  if(filter.view==='open'){const r=rowStatus(it);return r!=='DONE'}
  if(filter.view.startsWith('sec'))return it.sec===+filter.view.slice(3);
  return true;
}
function cardHtml(it){
  const s=st[it.id];const sec=doc.sections[it.sec];const rs=rowStatus(it);
  const cls=rs?'s-'+rs.replace(' ',''):'';
  const cols=allCols(it);
  const trades=cols.map(col=>{const t=sec.trades.find(x=>x.col===col);const v=s.status[col]||null;
    return `<div class="trade"><span class="tn">${esc(t?t.name:colName(col))}</span><div class="seg" role="group" aria-label="${esc(t?t.name:'')} status">${
      STATUSES.map(x=>`<button type="button" data-act="st" data-id="${it.id}" data-col="${col}" data-v="${x}" class="${v===x?'on':''}" aria-pressed="${v===x}">${x}</button>`).join('')}</div></div>`}).join('');
  const missing=sec.trades.filter(t=>!cols.includes(t.col));
  const needBlk=s.open||s.blockers.length||s.note||rs==='NOT DONE'||rs==='ONGOING';
  const blk=needBlk?`<div class="blk"><span class="lbl">Blocker / reason</span><div class="bchips">${
    BLOCKERS.map(([b])=>`<button type="button" class="bchip" data-act="blk" data-id="${it.id}" data-b="${esc(b)}" aria-pressed="${s.blockers.includes(b)}">${esc(b)}</button>`).join('')
    }</div><textarea id="note-${it.id}" data-act="note" data-id="${it.id}" placeholder="Details for the day shift (who, what is needed)">${esc(s.note)}</textarea></div>`:'';
  return `<article class="card ${cls}" id="card-${it.id}">
    <div class="head"><div class="lvl"><small>LEVEL</small>${esc(it.level)}</div>
      <div class="tagbox">${it.added?'<span class="badge">Added on night</span>':''}<div class="tag">${esc(it.tag||'—')}</div><div class="act">${esc(it.act)}</div>
</div></div>
    ${it.dayRemark?`<div class="daynote"><em>Day shift</em>${esc(it.dayRemark)}</div>`:''}
    ${histHtml(it)}
    ${s.by?`<div class="who">Updated by ${esc(authUser&&s.by===authUser.uid?'you':(s.byName||'a teammate'))}${s.at?' · '+fmtTime(s.at):''}</div>`:''}
    ${(s.hist||[]).length?`<details class="hist chist"><summary><span>Change history</span><span class="hc">${s.hist.length}</span></summary><ul>${s.hist.slice().reverse().map(h=>`<li><span class="hd">${esc(fmtTime(h.at))}</span><span class="ha"><strong>${esc(h.byName||'Someone')}</strong> · ${esc((h.ch||[]).join('; '))}</span></li>`).join('')}</ul></details>`:''}
    ${trades||'<div class="meta">No trade marked on the DABS. Add the trade that does this job.</div>'}
    ${blk}
    <div class="row-tools">
      ${!needBlk?`<button class="link" type="button" data-act="openblk" data-id="${it.id}">+ Blocker / note</button>`:''}
      ${it.added?`<button class="link danger" type="button" data-act="remove" data-id="${it.id}">Remove</button>`:''}
      ${missing.length?`<select data-act="addtrade" data-id="${it.id}" aria-label="Add trade"><option value="">+ Add trade</option>${missing.map(t=>`<option value="${t.col}">${esc(t.name)}</option>`).join('')}</select>`:''}
    </div>
  </article>`;
}
function histHtml(it){
  const h=historyOf(it.tag);if(!h.length)return '';
  const nd=nightsNotDone(h);const tone=s=>s==='DONE'?'ok':s==='ONGOING'?'go':s==='NOT DONE'?'no':'un';
  return `<details class="hist"><summary><span class="hl">${nd?`Carried over · ${nd} earlier ${nd===1?'night':'nights'} not done`:'Seen on earlier DABS'}</span><span class="hc">${h.length}</span></summary>
    <ul>${h.slice(0,6).map(x=>`<li><span class="hd">${esc(x.label)}</span><span class="hs ${tone(x.s)}">${esc(x.s||'NOT UPDATED')}</span><span class="ha">${esc(x.a)}${x.b.length?' · '+esc(x.b.join(', ')):''}</span></li>`).join('')}</ul></details>`;
}
function renderMain(){
  const m=$('#main');
  renderBanner();
  if(!doc){
    if(!authUser)return;
    if(!ready){m.innerHTML=`<p class="loading">Loading the team's DABS…</p>`;return}
    if(!canWrite){m.innerHTML=`<section class="empty"><h2>No DABS shared yet</h2><p>When the night shift lead loads tonight's DABS it appears here.</p></section>`;return}
    m.innerHTML=`<section class="empty">
    <h2>Share tonight's DABS with the team</h2>
    <p>Load the Excel the day shift sent. Everyone on the night shift then marks jobs from their own phone, even without signal. Only one DABS can be loaded per night.</p>
    <label class="drop" for="fileInput"><strong>Choose DABS .xlsx</strong><br><span style="color:var(--muted);font-size:13px">Shared with your approved team.</span></label>
    <ol class="steps"><li>Tap DONE, ONGOING or NOT DONE per trade.</li><li>Pick the blocker and add a short note.</li><li>Keep the concern list up to date, then Export Excel.</li></ol>
  </section>`;return}
  m.innerHTML=doc.sections.map((sec,si)=>{
    const items=itemsOf(si).filter(visible);
    const addBtn=`<button class="addact" type="button" data-act="addnew" data-sec="${si}">+ Add activity to ${esc(sec.short)}</button>`;
    if(filter.view.startsWith('sec')&&filter.view!=='sec'+si)return '';
    if(!items.length&&filter.view==='open')return '';
    const bh=[...new Set(sec.items.map(i=>i.bh).filter(Boolean))].map(b=>b.replace(/\s*\n\s*(?=\()/g,' ').replace(/\s*\n\s*/g,' · ')).join(' · ');
    return `<section class="sec"><h2>${esc(sec.title)}</h2>${bh?`<p class="meta" style="margin:-4px 0 10px">Black hat: ${esc(bh)}</p>`:''}<div class="list">${items.map(cardHtml).join('')}${addBtn}</div></section>`;
  }).join('')||`<p style="color:var(--muted)">Nothing matches this filter.</p>`;
}
function renderBanner(){
  const b=$('#banner');const lt=latest();
  if(!(doc&&cur&&lt&&lt.id!==cur.id)){b.innerHTML='';return}
  b.innerHTML=`<div class="older"><span>You are viewing an older night (${esc(cur.label)}). Changes are saved to that night.</span><button class="btn alt" type="button" id="goLatest">Open ${esc(lt.label)}</button></div>`;
  $('#goLatest').onclick=async()=>{try{await openNight(latest());window.scrollTo(0,0)}catch(e){toast(e.message)}};
}
function renderAll(){renderTop();if(doc)renderFilters();renderMain()}
function rerenderCard(id,remote){
  const it=findItem(id);const el=$('#card-'+id);if(!el||!it)return;
  const focusNote=document.activeElement&&document.activeElement.id==='note-'+id;
  if(remote&&focusNote)return;
  const open=[...el.querySelectorAll('details')].map(d=>d.open);
  el.outerHTML=cardHtml(it);renderTop();
  const ne=$('#card-'+id);if(ne)[...ne.querySelectorAll('details')].forEach((d,i)=>{if(open[i])d.open=true});
  if(focusNote){const t=$('#note-'+id);t.focus();t.setSelectionRange(t.value.length,t.value.length)}
}
function findItem(id){return allItems().find(i=>i.id===id)}

/* ---------- sheets that stay in sync while open ---------- */
let sheetView=null;
function openView(name,html){sheetView=name;openSheet(html)}
const normTrade=t=>{let s=String(t||'').toLowerCase().replace(/[^a-z]/g,'');if(s==='generallabourer'||s==='labourer'||s==='genop')s='generaloperative';return s};

/* ---------- settings: fixed briefing details + which crafts sign each DABS trade's DAB ---------- */
const DEFAULT_SETTINGS={
  projectName:'HM1 W12 Works',company:'ENKA UK',workLocation:'HM1',startTime:'07:00pm',finishTime:'5:20am',
  icChannel:'CH6',tmicChannel:'CH6',muster:'East Canteen',fap:'FAP11',rvp:'RVP F',
  otherTeams:'Rigging/Scaffolding teams will be working in the same areas, communicate with other teams in the area and make them aware of what your task is, use barriers to segregate the area and clear signage to warn others in the vicinity.\nMaintain clear communication, good housekeeping and cable management.',
  siteNews:'Radiography in areas on -5/0 Level, be aware of barriers around the controlled vicinity – always obey barriers if in doubt ask.\nLifting operations, be aware of your surroundings, do not cross barriers, communicate with other work parties.',
  mapping:{
    Plating:['Plater','Mechanical Working Chargehand (Plater)'],
    Piping:['Pipe Fitter','Pipe Welder','Pipe Welder Working Chargehand','Post Weld Heat Treatment Operator','Post Weld Heat Treatment Supervisor (Blue Hat)','Materials Controller','General Labourer','Firewatcher'],
    Lifting:['Rigger','Rigger Working Chargehand','Telescopic Handler Operator'],
    Scaffolding:['Scaffolder Basic','Scaffolder Advanced','Scaffold Working Chargehand','Scaffold Inspector'],
    QC:['QC Inspector']
  }
};
let settings=JSON.parse(JSON.stringify(DEFAULT_SETTINGS)),unsubSettings=null;
function subscribeSettings(){
  if(unsubSettings)return;
  unsubSettings=db.doc(P('settings')+'/main').onSnapshot(s=>{
    const d=s.exists?s.data():{};settings=Object.assign(JSON.parse(JSON.stringify(DEFAULT_SETTINGS)),d||{});
    settings.mapping=Object.assign({},DEFAULT_SETTINGS.mapping,(d&&d.mapping)||{});
    if(sheetView==='briefing')refreshBriefingPeople();
  },()=>{});
}
function dabsTrades(){const s=new Set();if(doc)for(const sec of doc.sections)for(const t of sec.trades)s.add(t.name);for(const k of Object.keys(settings.mapping||{}))s.add(k);return [...s]}
function openSettings(){
  const S=settings;const people=[...new Set(peopleList.map(p=>p.trade).filter(Boolean))];
  const allCrafts=[...new Set([...people,...Object.values(S.mapping||{}).flat()])].sort();
  const f=(id,label,v)=>`<label for="st_${id}">${label}<input id="st_${id}" value="${esc(v||'')}"></label>`;
  openView('settings',`<h3>Settings</h3><p>These details fill every Daily Activity Briefing. Only admins can change them.</p>
    <form id="stForm" novalidate style="display:flex;flex-direction:column;gap:10px">
      <div class="grid2">${f('projectName','Project name',S.projectName)}${f('company','Company',S.company)}</div>
      <div class="grid2">${f('workLocation','Work location',S.workLocation)}${f('muster','Nearest internal muster',S.muster)}</div>
      <div class="grid2">${f('startTime','Start time',S.startTime)}${f('finishTime','Finish time',S.finishTime)}</div>
      <div class="grid2">${f('icChannel','IC radio channel',S.icChannel)}${f('tmicChannel','TMIC / FPOS channel',S.tmicChannel)}</div>
      <div class="grid2">${f('fap','Fire Assembly Point (FAP)',S.fap)}${f('rvp','RVP point',S.rvp)}</div>
      <label for="st_otherTeams">Default: other teams working in our areas<textarea id="st_otherTeams" rows="3">${esc(S.otherTeams)}</textarea></label>
      <label for="st_siteNews">Default: what is happening on site<textarea id="st_siteNews" rows="3">${esc(S.siteNews)}</textarea></label>
      <h4 style="margin:6px 0 0">Who signs each DAB</h4>
      <p style="margin:0">Pick the crafts from the people list that belong on each DABS trade's sign-in sheet.</p>
      ${dabsTrades().map(T=>`<fieldset class="mapf"><legend>${esc(T)}</legend><div class="bchips">${allCrafts.map(c=>`<button type="button" class="bchip pick" data-map="${esc(T)}" data-craft="${esc(c)}" aria-pressed="${(S.mapping[T]||[]).some(x=>normTrade(x)===normTrade(c))}">${esc(c)}</button>`).join('')}</div></fieldset>`).join('')}
      <label for="st_newCraft">Add a craft name<input id="st_newCraft" placeholder="e.g. Fire Watcher"></label>
      <div class="grid2"><button class="btn alt" type="button" id="stClose">Close</button><button class="btn" type="submit"${isAdmin?'':' disabled'}>Save settings</button></div>
    </form>`);
  const map=JSON.parse(JSON.stringify(S.mapping||{}));
  $('#stClose').onclick=closeSheet;
  $('#stForm').onclick=e=>{const b=e.target.closest('.pick[data-map]');if(!b||!isAdmin)return;const T=b.dataset.map,c=b.dataset.craft;
    map[T]=map[T]||[];const i=map[T].findIndex(x=>normTrade(x)===normTrade(c));if(i>=0)map[T].splice(i,1);else map[T].push(c);b.setAttribute('aria-pressed',String(i<0))};
  $('#st_newCraft').onkeydown=e=>{if(e.key!=='Enter')return;e.preventDefault();const v=e.target.value.trim();if(!v)return;
    for(const fs of document.querySelectorAll('.mapf .bchips')){const T=fs.querySelector('[data-map]')?.dataset.map||fs.closest('fieldset').querySelector('legend').textContent;
      fs.insertAdjacentHTML('beforeend',`<button type="button" class="bchip pick" data-map="${esc(T)}" data-craft="${esc(v)}" aria-pressed="false">${esc(v)}</button>`)}e.target.value=''};
  $('#stForm').onsubmit=async e=>{e.preventDefault();if(!isAdmin)return;
    const out={mapping:map};for(const k of ['projectName','company','workLocation','muster','startTime','finishTime','icChannel','tmicChannel','fap','rvp','otherTeams','siteNews'])out[k]=$('#st_'+k).value;
    try{await db.doc(P('settings')+'/main').set(out);toast('Settings saved');closeSheet()}catch(err){toast('Not saved ('+(err.code||'error')+')')}};
}

/* ---------- people ---------- */
let peopleList=[],unsubPeople=null;
function subscribePeople(){
  if(unsubPeople)return;
  unsubPeople=db.collection(P('people')).onSnapshot(s=>{peopleList=s.docs.map(d=>Object.assign({id:d.id},d.data())).filter(p=>p.active!==false);
    if(sheetView==='people')renderPeopleList();if(sheetView==='briefing')refreshBriefingPeople()},()=>{});
}
function parsePeopleLines(text,defTrade){
  const out=[];
  for(let line of String(text||'').split(/\r?\n/)){
    line=line.trim();if(!line)continue;
    let parts=line.includes('\t')?line.split('\t'):line.includes(',')?line.split(','):line.split(/\s{2,}/);
    parts=parts.map(x=>x.trim()).filter(Boolean);if(!parts.length)continue;
    if(/^name/i.test(parts[0])&&parts.length>1)continue;   // header row from Excel
    let name=parts[0],sap='',trade='';
    for(const p of parts.slice(1)){if(!sap&&/^\d{4,}$/.test(p))sap=p;else if(!trade)trade=p}
    if(parts.length===1){const m=/^(.*?)\s+(\d{5,})$/.exec(name);if(m){name=m[1];sap=m[2]}}
    trade=trade||defTrade||'';
    if(normTrade(trade)==='generaloperative')trade='General Operative';
    out.push({name,sap,trade});
  }
  return out;
}
function openPeople(){
  const crafts=[...new Set([...peopleList.map(p=>p.trade),...Object.values(settings.mapping||{}).flat()].filter(Boolean))].sort();
  openView('people',`<h3>Personel</h3>
    <div id="pplSum" class="pplsum"></div>
    ${isAdmin?`<details class="pplAdd"><summary>Add people</summary>
      <form id="pplOne" novalidate class="pplform">
        <label for="pp_name">Name<input id="pp_name" autocomplete="off"></label>
        <div class="grid2"><label for="pp_sap">SAP no. (badge)<input id="pp_sap" inputmode="numeric" autocomplete="off"></label>
        <label for="pp_trade">Trade<input id="pp_trade" list="craftList" autocomplete="off"></label></div>
        <button class="btn" type="submit">Add person</button></form>
      <form id="pplBulk" novalidate class="pplform">
        <label for="pp_bulk">Paste a list (one person per line: name, SAP no., trade). Copying rows from Excel works.<textarea id="pp_bulk" rows="5" placeholder="Mark Hicks\t561850\tRigger Chargehand"></textarea></label>
        <label for="pp_def">Trade for lines without one<input id="pp_def" list="craftList" autocomplete="off"></label>
        <p class="nc" id="pp_prev"></p>
        <button class="btn alt" type="submit">Add list</button></form>
      <datalist id="craftList">${crafts.map(c=>`<option value="${esc(c)}"></option>`).join('')}</datalist></details>`:`<p class="err">You are signed in as ${esc(authUser&&authUser.email||'')} (${canWrite?'Contributor':'Viewer'}). Only admins can add, edit or remove people.</p>`}
    <label for="pp_find" class="pplfind">Search<input id="pp_find" placeholder="Name, SAP no. or trade" autocomplete="off"></label>
    <div id="pplList"></div>
    <button class="ghost" type="button" id="ppClose">Close</button>`);
  $('#ppClose').onclick=closeSheet;
  $('#pp_find').oninput=renderPeopleList;
  if(isAdmin){
    const prev=()=>{const n=parsePeopleLines($('#pp_bulk').value,$('#pp_def').value.trim());$('#pp_prev').textContent=n.length?n.length+' people ready to add'+(n.some(x=>!x.trade)?' · some have no trade':''):''};
    $('#pp_bulk').oninput=prev;$('#pp_def').oninput=prev;
    $('#pplOne').onsubmit=async e=>{e.preventDefault();const name=$('#pp_name').value.trim();if(!name){toast('Enter a name');return}
      await addPeople([{name,sap:$('#pp_sap').value.trim(),trade:$('#pp_trade').value.trim()}]);$('#pp_name').value='';$('#pp_sap').value='';$('#pp_name').focus()};
    $('#pplBulk').onsubmit=async e=>{e.preventDefault();const list=parsePeopleLines($('#pp_bulk').value,$('#pp_def').value.trim());if(!list.length){toast('Paste at least one name');return}
      await addPeople(list);$('#pp_bulk').value='';prev()};
  }
  $('#panel').onclick=e=>{
    const ed=e.target.closest('[data-ped]');if(ed){const p=peopleList.find(x=>x.id===ed.dataset.ped);if(p)editPerson(p);return}
    const b=e.target.closest('[data-prm]');if(!b)return;const p=peopleList.find(x=>x.id===b.dataset.prm);if(!p)return;
    const row=$('#pr-'+p.id);row.innerHTML=`<div class="nl"><strong>Remove ${esc(p.name)}?</strong><span class="nf">${esc(p.trade||'No trade')}${p.sap?' · '+esc(p.sap):''}. They will no longer appear on the sign-in sheets.</span></div>
      <div class="nb"><button class="btn alt" type="button" data-pkeep="1">Keep</button><button class="btn danger" type="button" data-pyes="${p.id}">Remove</button></div>`;
    row.querySelector('[data-pkeep]').onclick=renderPeopleList;
    row.querySelector('[data-pyes]').onclick=async()=>{try{await db.doc(P('people')+'/'+p.id).delete();toast(p.name+' removed')}catch(err){toast('Not removed ('+(err.code||'error')+')')}}};
  renderPeopleList();
}
function editPerson(p){
  const row=$('#pr-'+p.id);if(!row)return;
  row.innerHTML=`<form class="pplform pedit" novalidate style="width:100%">
    <label for="pe_name_${p.id}">Name<input id="pe_name_${p.id}" value="${esc(p.name)}" autocomplete="off"></label>
    <div class="grid2"><label for="pe_sap_${p.id}">SAP no. (badge)<input id="pe_sap_${p.id}" value="${esc(p.sap||'')}" inputmode="numeric" autocomplete="off"></label>
    <label for="pe_trade_${p.id}">Trade / role<input id="pe_trade_${p.id}" value="${esc(p.trade||'')}" list="craftList" autocomplete="off"></label></div>
    <div class="grid2"><button class="btn alt" type="button" data-pcancel="1">Cancel</button><button class="btn" type="submit">Save</button></div></form>`;
  row.querySelector('[data-pcancel]').onclick=renderPeopleList;
  row.querySelector('form').onsubmit=async e=>{e.preventDefault();
    const name=$('#pe_name_'+p.id).value.trim();if(!name){toast('Name cannot be empty');return}
    let trade=$('#pe_trade_'+p.id).value.trim();if(normTrade(trade)==='generaloperative')trade='General Operative';
    const upd={name,sap:$('#pe_sap_'+p.id).value.trim(),trade,editedAt:Date.now(),editedByName:myName()};
    try{await db.doc(P('people')+'/'+p.id).update(upd);Object.assign(p,upd);toast('Saved');renderPeopleList()}
    catch(err){toast('Not saved ('+(err.code||'error')+')')}};
  $('#pe_trade_'+p.id).focus();
}
async function addPeople(list){
  let addedN=0,skipped=0;
  for(const p of list){
    const dup=peopleList.find(x=>(p.sap&&x.sap===p.sap)||(!p.sap&&x.name.toLowerCase()===p.name.toLowerCase()&&normTrade(x.trade)===normTrade(p.trade)));
    if(dup){skipped++;continue}
    const id='p'+(p.sap||hashStr(p.name+p.trade))+Math.random().toString(36).slice(2,5);
    try{await db.doc(P('people')+'/'+id).set({name:p.name,sap:p.sap||'',trade:p.trade||'',active:true,addedAt:Date.now(),addedByName:myName()});addedN++;
      peopleList.push(Object.assign({id},p))}catch(e){toast('Stopped ('+(e.code||'error')+')');break}
  }
  toast(`Added ${addedN}`+(skipped?` · ${skipped} already on the list`:''));renderPeopleList();
}
function renderPeopleList(){
  const el=$('#pplList');if(!el)return;
  const q=($('#pp_find')&&$('#pp_find').value||'').toLowerCase();
  const groups=new Map(),labels=new Map();
  for(const p of peopleList){const k=normTrade(p.trade)||'~';if(!groups.has(k)){groups.set(k,[]);labels.set(k,p.trade||'No trade')}
    if(p.trade&&p.trade!==p.trade.toUpperCase())labels.set(k,p.trade);groups.get(k).push(p)}
  for(const [k,v] of [...groups])if(labels.get(k)!==k){groups.delete(k);groups.set(labels.get(k),v)}
  const sorted=[...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
  $('#pplSum').innerHTML=`<div class="pplTotal"><b>${peopleList.length}</b> people</div>`+sorted.map(([t,ps])=>`<span class="pchip"><b>${ps.length}</b> ${esc(t)}</span>`).join('');
  el.innerHTML=sorted.map(([t,ps])=>{const vis=ps.filter(p=>!q||(p.name+' '+p.sap+' '+t).toLowerCase().includes(q)).sort((a,b)=>a.name.localeCompare(b.name));if(!vis.length)return '';
    return `<section class="pgrp"><h4>${esc(t)} <span class="nc">${ps.length}</span></h4><div class="nights">${vis.map(p=>`<div class="night" id="pr-${p.id}"><div class="nl"><strong>${esc(p.name)}</strong><span class="nf">${p.sap?'SAP '+esc(p.sap):'No SAP no.'}</span></div>${isAdmin?`<div class="nb"><button class="link" type="button" data-ped="${p.id}">Edit</button><button class="link danger" type="button" data-prm="${p.id}">Remove</button></div>`:''}</div>`).join('')}</div></section>`}).join('')
    ||'<p>No people yet.'+(isAdmin?' Use Add people to paste the lists for each trade.':'')+'</p>';
}

/* ---------- concern list ---------- */
let concerns=[],unsubConcerns=null;
const CTYPES=['Concern','Blocker','Fix needed'];
function subscribeConcerns(){
  if(unsubConcerns)return;
  unsubConcerns=db.collection(P('concerns')).orderBy('at','desc').limit(300).onSnapshot(s=>{
    concerns=s.docs.map(d=>Object.assign({id:d.id},d.data()));updateConcernBadge();
    if(sheetView==='concerns')renderConcernList()},()=>{});
}
const openConcerns=()=>concerns.filter(c=>!c.resolved);
function updateConcernBadge(){const b=$('#btnConcerns');if(!b)return;const n=openConcerns().length;b.innerHTML='Concern list'+(n?` <span class="cnt">${n}</span>`:'')}
function concernTrades(){return ['All',...dabsTrades()]}
function openConcernsView(){
  openView('concerns',`<h3>Concern list</h3>
    <p>Concerns, blockers and things that need fixing. Open items carry over to the next night and print in red under the DABS in the Excel export.</p>
    ${canWrite?`<form id="cnForm" novalidate class="pplform">
      <div class="grid2"><label for="cn_type">Type<select id="cn_type">${CTYPES.map(t=>`<option>${t}</option>`).join('')}</select></label>
      <label for="cn_trade">Affected trade<select id="cn_trade">${concernTrades().map(t=>`<option>${esc(t)}</option>`).join('')}</select></label></div>
      <label for="cn_desc">Description<textarea id="cn_desc" rows="3" placeholder="e.g. Scaffold on 1GST1930TY-GL04 not tagged, plating could not start"></textarea></label>
      <p class="err" id="cn_err" hidden></p>
      <button class="btn" type="submit">Add to list</button></form>`:''}
    <div id="cnList"></div>
    <button class="ghost" type="button" id="cnClose">Close</button>`);
  $('#cnClose').onclick=closeSheet;
  const fm=$('#cnForm');if(fm)fm.onsubmit=async e=>{e.preventDefault();const desc=$('#cn_desc').value.trim();
    if(!desc){const p=$('#cn_err');p.textContent='Write what the concern is.';p.hidden=false;return}
    const id='c'+Date.now().toString(36)+Math.random().toString(36).slice(2,5);
    const body={type:$('#cn_type').value,trade:$('#cn_trade').value,desc,by:authUser.uid,byName:myName(),at:Date.now(),nightLabel:cur?cur.label:'',resolved:false,hist:[]};
    concerns.unshift(Object.assign({id},body));renderConcernList();updateConcernBadge();$('#cn_desc').value='';$('#cn_err').hidden=true;
    db.doc(P('concerns')+'/'+id).set(body).catch(err=>toast('Not saved ('+(err.code||'error')+')'));toast('Added to the concern list')};
  $('#panel').onclick=e=>{const b=e.target.closest('[data-cn]');if(!b)return;const c=concerns.find(x=>x.id===b.dataset.c);if(!c)return;concernAction(b.dataset.cn,c)};
  renderConcernList();
}
function concernRow(c){
  const carried=cur&&c.at<cur.loadedAt&&!c.resolved;
  return `<div class="cn ${c.resolved?'done':''}" id="cn-${c.id}">
    <div class="cnh"><span class="ctype t-${normTrade(c.type)}">${esc(c.type)}</span><span class="ctrade">${esc(c.trade||'All')}</span>${carried?`<span class="nc">since ${esc(new Date(c.at).toLocaleDateString('en-GB',{day:'numeric',month:'short'}))}</span>`:''}</div>
    <div class="cnd">${esc(c.desc)}</div>
    <div class="nc">${esc(c.byName||'Someone')} · ${esc(fmtDT(c.at))}${c.editedAt?` · edited by ${esc(c.editedByName||'someone')} ${esc(fmtDT(c.editedAt))}`:''}${c.resolved?` · resolved by ${esc(c.resolvedByName||'someone')} ${esc(fmtDT(c.resolvedAt))}`:''}</div>
    ${canWrite?`<div class="cnb">${c.resolved?`<button class="link" type="button" data-cn="reopen" data-c="${c.id}">Reopen</button>`
      :`<button class="link" type="button" data-cn="resolve" data-c="${c.id}">Resolved</button><button class="link" type="button" data-cn="edit" data-c="${c.id}">Edit</button>`}<button class="link danger" type="button" data-cn="del" data-c="${c.id}">Delete</button></div>`:''}
  </div>`;
}
function renderConcernList(){
  const el=$('#cnList');if(!el)return;
  const open=openConcerns();const since=Date.now()-7*864e5;
  const done=concerns.filter(c=>c.resolved&&(c.resolvedAt||0)>=since);
  el.innerHTML=`<h4 class="cnhd">Open <span class="nc">${open.length}</span></h4>${open.map(concernRow).join('')||'<p class="nc">Nothing open.</p>'}
    ${done.length?`<details class="cndone"><summary>Resolved in the last 7 days (${done.length})</summary>${done.map(concernRow).join('')}</details>`:''}`;
}
function concernWarn(c,verb,onYes){
  const mine=authUser&&c.by===authUser.uid;
  openView('concernWarn',`<h3>${verb} this ${esc(c.type.toLowerCase())}?</h3>
    <div class="cn"><div class="cnd">${esc(c.desc)}</div></div>
    <p>Entered by <strong>${esc(mine?'you':(c.byName||'a teammate'))}</strong> at ${esc(fmtDT(c.at))}.${mine?'':' Make sure they know about the change.'}</p>
    <div class="grid2"><button class="btn alt" type="button" id="cwNo">Cancel</button><button class="btn${verb==='Delete'?' danger':''}" type="button" id="cwYes">${verb}</button></div>`);
  $('#cwNo').onclick=openConcernsView;$('#cwYes').onclick=onYes;
}
function concernAction(act,c){
  const ref=db.doc(P('concerns')+'/'+c.id);const h=x=>[...(c.hist||[]),{at:Date.now(),byName:myName(),ch:x}].slice(-20);
  if(act==='del')concernWarn(c,'Delete',async()=>{try{await ref.delete();concerns=concerns.filter(x=>x.id!==c.id);toast('Deleted')}catch(e){toast('Not deleted ('+(e.code||'error')+')')}openConcernsView();updateConcernBadge()});
  else if(act==='resolve')concernWarn(c,'Resolve',()=>{ref.update({resolved:true,resolvedAt:Date.now(),resolvedBy:authUser.uid,resolvedByName:myName(),hist:h('Resolved')}).catch(e=>toast('Not saved ('+(e.code||'error')+')'));
    Object.assign(c,{resolved:true,resolvedAt:Date.now(),resolvedByName:myName()});openConcernsView();updateConcernBadge();toast('Marked as resolved')});
  else if(act==='reopen')concernWarn(c,'Reopen',()=>{ref.update({resolved:false,resolvedAt:null,resolvedByName:'',hist:h('Reopened')}).catch(()=>{});
    Object.assign(c,{resolved:false});openConcernsView();updateConcernBadge()});
  else if(act==='edit')concernWarn(c,'Edit',()=>{
    openView('concernEdit',`<h3>Edit ${esc(c.type.toLowerCase())}</h3><form id="ceForm" novalidate class="pplform">
      <div class="grid2"><label for="ce_type">Type<select id="ce_type">${CTYPES.map(t=>`<option${t===c.type?' selected':''}>${t}</option>`).join('')}</select></label>
      <label for="ce_trade">Affected trade<select id="ce_trade">${[...new Set([...concernTrades(),c.trade||'All'])].map(t=>`<option${t===(c.trade||'All')?' selected':''}>${esc(t)}</option>`).join('')}</select></label></div>
      <label for="ce_desc">Description<textarea id="ce_desc" rows="4">${esc(c.desc)}</textarea></label>
      <div class="grid2"><button class="btn alt" type="button" id="ceNo">Cancel</button><button class="btn" type="submit">Save</button></div></form>`);
    $('#ceNo').onclick=openConcernsView;
    $('#ceForm').onsubmit=e=>{e.preventDefault();const nd=$('#ce_desc').value.trim();if(!nd){toast('Description cannot be empty');return}
      const upd={type:$('#ce_type').value,trade:$('#ce_trade').value,desc:nd,editedAt:Date.now(),editedByName:myName(),hist:h('Edited (was: '+c.desc.slice(0,80)+')')};
      ref.update(upd).catch(err=>toast('Not saved ('+(err.code||'error')+')'));Object.assign(c,upd);openConcernsView();toast('Saved')};
  });
}
function exportConcerns(){
  const from=cur?cur.loadedAt:0;
  return concerns.filter(c=>!c.resolved||(c.resolvedAt||0)>=from).sort((a,b)=>a.at-b.at);
}

/* ---------- menu ---------- */
function openMenu(){
  const role=isAdmin?'Admin':canWrite?'Contributor':'Viewer';
  const it=(id,label,sub,show=true)=>show?`<button class="mi" type="button" id="${id}"><strong>${label}</strong><span>${sub}</span></button>`:'';
  openView('menu',`<h3>Menu</h3><div class="menu">
    ${it('mSearch','Search history','Find a tag or activity: every night it was on, status and blockers')}
    ${it('mArchive','Archive','Earlier nights, open or delete')}
    ${it('mConcerns','Concern list',openConcerns().length+' open')}
    ${it('mBrief','Daily Activity Briefing','Word documents for each trade',!!doc)}
    ${it('mPeople','Personel',peopleList.length+' people · edit names and roles')}
    ${it('mReplace','Replace tonight\'s DABS','Load a corrected file for '+(cur?cur.label:''),isAdmin&&!!cur)}
    ${it('mSettings','Settings','Briefing details and sign-in crafts',isAdmin)}
    ${it('mTeam','Team & approvals','Who can use the app',isAdmin)}
    ${it('mCopy','Copy live nights into test','Test version only: practise with real DABS',ENV==='test'&&isAdmin)}
    </div>
    <div class="bk"><h4>Account</h4><p>Signed in as <strong>${esc(myName())}</strong> (${esc(authUser&&authUser.email||'')}) · ${role}</p><p class="nc">App version ${APP_VERSION}</p>
      <button class="btn alt" type="button" id="signOut">Sign out</button></div>
    ${standalone()?'':`<div class="bk"><h4>Install on your phone</h4>${installPrompt?'<button class="btn" type="button" id="doInstall">Install app</button>':''}
      <p><strong>Android (Chrome):</strong> menu ⋮ → Install app.<br><strong>iPhone (Safari):</strong> Share button → Add to Home Screen.</p></div>`}
    <button class="ghost" type="button" id="mClose">Close</button>`);
  const on=(id,f)=>{const b=$('#'+id);if(b)b.onclick=f};
  on('mClose',closeSheet);on('mSearch',openSearch);on('mArchive',openArchive);on('mConcerns',openConcernsView);on('mBrief',openBriefing);on('mPeople',openPeople);
  on('mSettings',openSettings);on('mTeam',openTeam);on('mCopy',copyLiveToTest);on('mReplace',()=>{closeSheet();$('#replaceInput').click()});
  on('signOut',async()=>{closeSheet();await fbSignOut()});
  on('doInstall',async()=>{try{installPrompt.prompt();await installPrompt.userChoice}catch(e){}installPrompt=null;openMenu()});
}

/* ---------- test only: copy real nights so the new version can be tried on them ---------- */
async function copyLiveToTest(){
  if(ENV!=='test'||!isAdmin)return;
  openView('copy',`<h3>Copy live nights into test</h3><p>The last ${NIGHTS_ON_OPEN} nights from the live app are copied here with their statuses and notes. The live app is only read, never changed.</p>
    <div class="grid2"><button class="btn alt" type="button" id="cpNo">Cancel</button><button class="btn" type="button" id="cpYes">Copy</button></div><p class="nc" id="cpMsg"></p>`);
  $('#cpNo').onclick=closeSheet;
  $('#cpYes').onclick=async()=>{const b=$('#cpYes');b.disabled=true;const msg=t=>{const m=$('#cpMsg');if(m)m.textContent=t};
    try{const ln=await db.collection(LIVE('nights')).orderBy('loadedAt','desc').limit(NIGHTS_ON_OPEN).get();let n=0;
      for(const d of ln.docs){const id=d.id,data=d.data();if(allNights().some(x=>x.id===id))continue;msg('Copying '+(data.label||id)+'…');
        const fid=data.fileId||id;const f=await db.doc(LIVE('files')+'/'+fid).get();if(!f.exists)continue;
        await db.doc(P('files')+'/'+fid).set(f.data());
        const its=await db.collection(LIVE('nights')+'/'+id+'/items').get();
        for(const it of its.docs)await db.doc(P('nights')+'/'+id+'/items/'+it.id).set(it.data());
        const key=data.dateKey||dateKeyFrom(data.label,data.loadedAt);
        await db.doc(P('nights')+'/'+id).set(Object.assign({},data,{dateKey:key||''}));
        if(key)try{await db.doc(P('dates')+'/'+key).create({nightId:id,by:authUser.uid,byName:myName(),at:Date.now()})}catch(e){}
        n++}
      toast(n?'Copied '+n+' nights':'Nothing new to copy');closeSheet()}
    catch(e){msg('Copy stopped ('+(e.code||e.message||'error')+')');b.disabled=false}};
}

/* ---------- search history: every night a tag or activity was on the DABS ---------- */
const sNorm=x=>String(x||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
function histEntries(){
  const src=new Map(tagHist);
  for(const n of allNights())if(!src.has(n.id)&&n.items)src.set(n.id,{label:n.label,at:n.loadedAt,dateKey:nightDateKey(n),items:n.items});
  if(cur&&doc)src.set(cur.id,{label:cur.label,at:cur.loadedAt,dateKey:nightDateKey(cur),items:compactItems()});
  return src;
}
function dateText(e){if(e.dateKey){const [y,m,d]=e.dateKey.split('-');return new Date(+y,+m-1,+d).toLocaleDateString('en-GB',{weekday:'short',day:'numeric',month:'short',year:'numeric'})}return e.label||''}
function runSearch(q){
  q=String(q||'').trim();if(q.length<2)return null;
  const qn=sNorm(q),words=q.toLowerCase().split(/\s+/).filter(Boolean);
  const groups=new Map();
  for(const [nid,e] of histEntries())for(const it of e.items||[]){
    const tagHit=qn.length>=2&&tokens(it.t).concat([sNorm(it.t)]).some(tk=>sNorm(tk).includes(qn));
    const txt=(it.a+' '+(it.n||'')+' '+(it.dr||'')).toLowerCase();
    const actHit=words.every(w=>txt.includes(w));
    if(!tagHit&&!actHit)continue;
    const key=it.t||it.a;if(!groups.has(key))groups.set(key,{tag:it.t,rows:[]});
    groups.get(key).rows.push({nid,e,it});
  }
  const out=[...groups.values()];
  for(const g of out)g.rows.sort((a,b)=>b.e.at-a.e.at);
  out.sort((a,b)=>b.rows[0].e.at-a.rows[0].e.at);
  return out;
}
function openSearch(){
  openView('search',`<h3>Search history</h3>
    <p>Type a tag number (part of it is enough, e.g. <strong>GL04</strong> or <strong>1CEX2478</strong>) or words from the activity (e.g. <strong>lisega</strong>). Every night the job was on the DABS is listed with its status and blockers.</p>
    <label for="sh_q">Tag or activity<input id="sh_q" autocomplete="off" inputmode="search" placeholder="e.g. 1GST1930TY-GL04"></label>
    <div id="sh_res"></div>
    <button class="ghost" type="button" id="shClose">Close</button>`);
  $('#shClose').onclick=closeSheet;
  let t=null;$('#sh_q').oninput=e=>{clearTimeout(t);t=setTimeout(()=>renderSearch(e.target.value),200)};
  $('#sh_res').onclick=async e=>{const b=e.target.closest('[data-shopen]');if(!b)return;
    let n=allNights().find(x=>x.id===b.dataset.shopen);
    if(!n){try{const d=await db.doc(P('nights')+'/'+b.dataset.shopen).get();if(d.exists)n=Object.assign({id:d.id},d.data())}catch(err){}}
    if(!n){toast('That night is no longer in the archive');return}
    closeSheet();try{await openNight(n);window.scrollTo(0,0);toast('Opened '+n.label)}catch(err){toast(err.message)}};
  setTimeout(()=>$('#sh_q').focus(),50);
}
function renderSearch(q){
  const el=$('#sh_res');if(!el)return;const res=runSearch(q);
  if(!res){el.innerHTML='';return}
  if(!res.length){el.innerHTML=`<p class="nc">Nothing found for “${esc(q)}”. ${tagHistReady?'':'History is still loading, try again in a moment.'}</p>`;return}
  const tone=s=>s==='DONE'?'ok':s==='ONGOING'?'go':s==='NOT DONE'?'no':'un';
  el.innerHTML=`<p class="nc">${res.length} ${res.length===1?'job':'jobs'} found</p>`+res.slice(0,60).map(g=>{
    const r=g.rows,first=r[r.length-1],last=r[0];
    const doneAt=r.slice().reverse().find(x=>x.it.s==='DONE');
    const notDone=new Set(r.filter(x=>x.it.s!=='DONE').map(x=>x.nid)).size;
    return `<section class="shg"><h4 class="tag">${esc(g.tag||'—')}</h4>
      <p class="nc">On ${r.length} ${r.length===1?'night':'nights'} · first ${esc(dateText(first.e))} · last ${esc(dateText(last.e))}${doneAt?` · <strong class="ok">DONE on ${esc(dateText(doneAt.e))}</strong>`:''}${notDone?` · not done on ${notDone}`:''}</p>
      <ul class="shl">${r.map(x=>`<li><button class="link" type="button" data-shopen="${esc(x.nid)}">${esc(dateText(x.e))}</button>
        <span class="hs ${tone(x.it.s)}">${esc(x.it.s||'NOT UPDATED')}</span>
        <span class="sha">${esc(x.it.a)}${x.it.l?' · Lvl '+esc(x.it.l):''}${(x.it.tr||[]).length?' · '+esc(x.it.tr.join(', ')):''}${x.it.x?' · added on night':''}</span>
        ${(x.it.b||[]).length?`<span class="shb">Blocker: ${esc(x.it.b.join(', '))}</span>`:''}
        ${x.it.n?`<span class="shn">${esc(x.it.n)}</span>`:''}</li>`).join('')}</ul></section>`}).join('');
}
/* ---------- Daily Activity Briefing: per-trade Word documents in the site DAB format ---------- */
const lvlShort=l=>String(l||'').replace(/\.0+$/,'').replace(/^([+-]?\d+)\.(\d*?)0+$/,'$1.$2').replace(/\.$/,'');
function previousNight(){if(!cur)return null;return byDate().filter(n=>n.loadedAt<cur.loadedAt)[0]||null}
function tradeItems(T){return allItems().filter(it=>tradeNames(it).includes(T))}
/* Turn the DABS activity into what this trade actually does on that job. */
const TRADE_WORDS=[
  ['Lifting',[[/lifting aid/i,'Lifting aid for fit-up'],[/fit.?up/i,'Lifting aid for fit-up'],[/erect/i,'Lifting activity for erection'],[/deliver/i,'Delivery (offload and lift)'],[/skat|movement|move/i,'Skating / load movement'],[/weld/i,'Lifting support for welding'],[/lift/i,'Lifting activity']]],
  ['Piping',[[/fit.?up/i,'Fit-up activity'],[/weld/i,'Spool welding'],[/erect/i,'Spool erection'],[/lift/i,'Rigging support for lifting'],[/scaffold/i,'Access check for scaffold']]],
  ['Plating',[[/erect/i,'Support erection'],[/weld/i,'Support welding'],[/fit.?up/i,'Support fit-up'],[/lift/i,'Support lifting preparation']]],
  ['Scaffolding',[[/scaffold/i,'Scaffold erection for access'],[/strip|dismantl/i,'Scaffold strip'],[/.*/,c=>'Scaffold access for '+c.toLowerCase()]]],
  ['QC',[[/weld/i,'Weld inspection'],[/fit.?up/i,'Fit-up inspection'],[/.*/,c=>'QC inspection: '+c.toLowerCase()]]]
];
function tradeActivity(act,T){
  const raw=String(act||'').replace(/\s+/g,' ').trim();
  const pri=(/^(P\d)\s*:\s*/i.exec(raw)||[])[1]||'';
  const quals=[...raw.matchAll(/\(([^)]+)\)/g)].map(m=>m[1].trim());
  const core=raw.replace(/^(P\d)\s*:\s*/i,'').replace(/\([^)]*\)/g,'').trim();
  const rules=(TRADE_WORDS.find(([k])=>normTrade(k)===normTrade(T))||[])[1]||[];
  let phrase=core;for(const [re,txt] of rules)if(re.test(core)){phrase=typeof txt==='function'?txt(core):txt;break}
  const q=quals.filter(x=>!phrase.toLowerCase().includes(x.toLowerCase()));
  return (pri?pri.toUpperCase()+' ':'')+phrase+(q.length?' ('+q.join(', ')+')':'');
}
function suggestToday(T){
  const lines=[],seen=new Set();
  for(const it of tradeItems(T)){const k=(it.tag||it.act)+'|'+it.level+'|'+it.act;if(seen.has(k))continue;seen.add(k);
    const act=tradeActivity(it.act,T);
    lines.push({p:/^P\d/.test(act)?act.slice(0,2):'P9',t:'*'+(it.tag||it.act)+(it.level&&it.level!=='—'?' ('+lvlShort(it.level)+')':'')+' – '+act})}
  return lines.map((x,i)=>Object.assign(x,{i})).sort((a,b)=>a.p.localeCompare(b.p)||a.i-b.i).map(x=>x.t).join('\n');
}
function suggestPrev(T){
  const lines=[];const pn=previousNight();const from=pn?pn.loadedAt:0;
  for(const c of concerns)if((c.trade||'All')==='All'||c.trade===T){
    if(!c.resolved)lines.push(c.desc);
    else if((c.resolvedAt||0)>=from)lines.push('Resolved: '+c.desc)}
  return lines.join('\n');
}
function suggestNotWell(T){
  const pn=previousNight();if(!pn)return '';
  const src=(tagHist.get(pn.id)||{}).items||pn.items||[];const by=new Map();
  for(const it of src){if(it.tr&&it.tr.length&&!it.tr.includes(T))continue;
    for(const b of it.b||[]){if(!by.has(b))by.set(b,[]);by.get(b).push(it.t||it.a)}}
  return [...by.entries()].map(([b,ts])=>b+': '+[...new Set(ts)].join(', ')).join('\n');
}
function briefingData(){
  const b=(cur&&cur.briefing)?JSON.parse(JSON.stringify(cur.briefing)):{};
  b.trades=b.trades||{};
  for(const T of dabsTrades().filter(t=>tradeItems(t).length||settings.mapping[t])){
    const t=b.trades[T]=b.trades[T]||{};
    // suggestions stay up to date until someone types in that box
    t.edited=t.edited||{};const ed=f=>t.edited[f]&&t[f]!=null;
    if(!ed('today'))t.today=suggestToday(T);
    if(!ed('prev'))t.prev=suggestPrev(T);
    if(!ed('notWell'))t.notWell='';   // left for the Black Hat to fill in
    if(t.well==null)t.well='';
    if(!ed('other'))t.other=settings.otherTeams||'';
    if(!ed('site'))t.site=settings.siteNews||'';
    if(t.bh==null)t.bh='';
  }
  return b;
}
let brief=null,briefTrade=null,briefT=null;
function saveBriefing(){
  clearTimeout(briefT);briefT=setTimeout(()=>{if(!cur||!canWrite||!brief)return;cur.briefing=brief;
    db.doc(P('nights')+'/'+cur.id).update({briefing:brief}).catch(()=>{})},800);
}
function signers(T){
  const want=(settings.mapping[T]||[]).map(normTrade);
  return peopleList.filter(p=>want.includes(normTrade(p.trade))).sort((a,b)=>want.indexOf(normTrade(a.trade))-want.indexOf(normTrade(b.trade))||a.name.localeCompare(b.name));
}
function openBriefing(){
  if(!doc){toast('Load tonight\'s DABS first');return}
  brief=briefingData();const trades=Object.keys(brief.trades);
  if(!briefTrade||!trades.includes(briefTrade))briefTrade=trades[0];
  const ta=(id,label,v,rows)=>`<label for="${id}">${label}<textarea id="${id}" rows="${rows||3}"${canWrite?'':' readonly'}>${esc(v||'')}</textarea></label>`;
  const t=brief.trades[briefTrade]||{};
  openView('briefing',`<h3>Daily Activity Briefing · ${esc(cur.label)}</h3>
    <p>Fill the names once for the night; the rest is suggested from the DABS and the concern list. Everyone on the team sees the same briefing.</p>
    <div class="grid2"><label for="bf_ic">Incident Controller<input id="bf_ic" value="${esc(brief.ic||'')}"${canWrite?'':' readonly'}></label>
      <label for="bf_tmic">TMIC / FPOS person<input id="bf_tmic" value="${esc(brief.tmic||'')}"${canWrite?'':' readonly'}></label></div>
    <label for="bf_fa">First Aider<input id="bf_fa" value="${esc(brief.firstAider||'')}"${canWrite?'':' readonly'}></label>
    <div class="filters btabs" role="tablist">${trades.map(T=>`<button class="chip" type="button" role="tab" data-bt="${esc(T)}" aria-pressed="${T===briefTrade}">${esc(T)}</button>`).join('')}</div>
    <div class="bform">
      <label for="bf_bh">Black Hat supervisor facilitating the brief<input id="bf_bh" value="${esc(t.bh)}"${canWrite?'':' readonly'}></label>
      ${ta('bf_prev','Feedback from any issues raised at the previous DAB',t.prev,3)}
      ${ta('bf_well','What worked well? (previous workday)',t.well,2)}
      ${ta('bf_notWell','What didn\'t work well and how can we make it better?',t.notWell,3)}
      ${ta('bf_today','Today: who is doing what and where, big new risks',t.today,4)}
      ${ta('bf_other','Will any other teams be working in our areas?',t.other,3)}
      ${ta('bf_site','What is happening on site that the team need to know?',t.site,3)}
      ${canWrite?'<button class="link" type="button" id="bf_reset">Refill suggestions for this trade</button>':''}
      <div class="bsign" id="bf_sign"></div>
    </div>
    <div class="grid2"><button class="btn alt" type="button" id="bf_all">All trades (one file)</button><button class="btn" type="button" id="bf_one">Word: ${esc(briefTrade)}</button></div>
    <button class="ghost" type="button" id="bfClose">Close</button>`);
  refreshBriefingPeople();
  $('#bfClose').onclick=closeSheet;
  const bindCommon=(id,k)=>{$(id).oninput=e=>{brief[k]=e.target.value;saveBriefing()}};
  bindCommon('#bf_ic','ic');bindCommon('#bf_tmic','tmic');bindCommon('#bf_fa','firstAider');
  for(const [id,k] of [['#bf_bh','bh'],['#bf_prev','prev'],['#bf_well','well'],['#bf_notWell','notWell'],['#bf_today','today'],['#bf_other','other'],['#bf_site','site']])
    $(id).oninput=e=>{const t=brief.trades[briefTrade];t[k]=e.target.value;t.edited=t.edited||{};t.edited[k]=true;saveBriefing()};
  for(const b of document.querySelectorAll('[data-bt]'))b.onclick=()=>{briefTrade=b.dataset.bt;openBriefing()};
  const rs=$('#bf_reset');if(rs)rs.onclick=()=>{const T=briefTrade;const k=brief.trades[T];
    Object.assign(k,{edited:{},today:suggestToday(T),prev:suggestPrev(T),notWell:'',other:settings.otherTeams||'',site:settings.siteNews||''});saveBriefing();openBriefing()};
  $('#bf_one').onclick=()=>downloadBriefing([briefTrade]);
  $('#bf_all').onclick=()=>downloadBriefing(trades);
}
function refreshBriefingPeople(){
  const el=$('#bf_sign');if(!el||!briefTrade)return;const ps=signers(briefTrade);
  el.innerHTML=`<strong>Sign-in sheet: ${ps.length} ${ps.length===1?'person':'people'}</strong> <span class="nc">(${esc((settings.mapping[briefTrade]||[]).join(', ')||'no crafts set: Menu → Settings')})</span>`;
}
async function downloadBriefing(trades){
  if(!brief)brief=briefingData();
  try{
    const u8=await buildDabDocx(trades);
    const date=(nightDateKey(cur)||'').split('-').reverse().join('-');
    const name='DAB '+(trades.length===1?trades[0]+' ':'')+(date||cur.label)+'.docx';
    deliverFile(name,u8,'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Briefing ready',openBriefing);
  }catch(e){toast('Could not create the document ('+(e.message||'error')+')')}
}

/* --- WordprocessingML, written directly so the app needs no extra library --- */
const xe=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function wRun(t,o={}){
  const rp=(o.b?'<w:b/>':'')+(o.i?'<w:i/>':'')+(o.color?`<w:color w:val="${o.color}"/>`:'')+(o.sz?`<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>`:'');
  return `<w:r>${rp?`<w:rPr>${rp}</w:rPr>`:''}<w:t xml:space="preserve">${xe(t)}</w:t></w:r>`;
}
function wP(runs,o={}){
  const pp=`<w:pPr><w:spacing w:before="${o.before||0}" w:after="${o.after||0}" w:line="240" w:lineRule="auto"/>${o.jc?`<w:jc w:val="${o.jc}"/>`:''}</w:pPr>`;
  return `<w:p>${pp}${Array.isArray(runs)?runs.join(''):runs}</w:p>`;
}
const wText=(t,o={})=>String(t||'').split('\n').map(line=>wP(wRun(line,o),{jc:o.jc})).join('')||wP('',{});
function wCell(content,o={}){
  const pr=`<w:tcPr><w:tcW w:w="${o.w}" w:type="dxa"/>${o.span>1?`<w:gridSpan w:val="${o.span}"/>`:''}${o.fill?`<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>`:''}<w:tcMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar><w:vAlign w:val="${o.v||'center'}"/></w:tcPr>`;
  return `<w:tc>${pr}${content||wP('')}</w:tc>`;
}
function wRow(cells,o={}){return `<w:tr><w:trPr>${o.h?`<w:trHeight w:val="${o.h}" w:hRule="${o.rule||'atLeast'}"/>`:''}<w:cantSplit/></w:trPr>${cells.join('')}</w:tr>`}
function wTable(grid,rows,o={}){
  const b=o.noBorder?'<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
    :'<w:tblBorders><w:top w:val="single" w:sz="8" w:color="404040"/><w:left w:val="single" w:sz="8" w:color="404040"/><w:bottom w:val="single" w:sz="8" w:color="404040"/><w:right w:val="single" w:sz="8" w:color="404040"/><w:insideH w:val="single" w:sz="4" w:color="404040"/><w:insideV w:val="single" w:sz="4" w:color="404040"/></w:tblBorders>';
  return `<w:tbl><w:tblPr><w:tblW w:w="${grid.reduce((a,b)=>a+b,0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>${b}</w:tblPr><w:tblGrid>${grid.map(g=>`<w:gridCol w:w="${g}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
}
const PAGE_BREAK='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const C={lbl:'FAD7BF',red:'E04B3F',green:'2FA352',amber:'F5B041',sec:'F8CBAD',band:'D8E4BC',sign:'E9925F'};
function dabPage1(T){
  const S=settings,b=brief,t=b.trades[T]||{};
  const G=[2150,1650,1500,1800,1450,1916],W=G.reduce((a,b)=>a+b,0);
  const L=(txt,fill,o={})=>wCell(wText(txt,{b:true,sz:o.sz||19,color:o.color}),{w:o.w,span:o.span,fill:fill||C.lbl});
  const V=(txt,o={})=>wCell(wText(txt,{sz:o.sz||20,jc:o.jc}),{w:o.w,span:o.span});
  const date=(nightDateKey(cur)||'').split('-').reverse().join('/');
  const head=(txt,fill)=>wRow([wCell(wText(txt,{b:true,sz:19}),{w:W,span:6,fill:fill||C.sec})]);
  const band=txt=>wRow([wCell(wText(txt,{b:true,sz:19,jc:'center'}),{w:W,span:6,fill:C.band})]);
  const body=(txt,h)=>wRow([wCell(wText(txt,{sz:20}),{w:W,span:6,v:'top'})],{h:h||1000});
  const rows=[
    wRow([L('Date',null,{w:G[0]}),V(date,{w:G[1]}),L('Start Time',null,{w:G[2]}),V(S.startTime,{w:G[3]}),L('Finish Time',null,{w:G[4]}),V(S.finishTime,{w:G[5]})],{h:360}),
    wRow([L('Project Name',null,{w:G[0]}),V(S.projectName,{w:G[1]+G[2],span:2}),L('Company',null,{w:G[3]}),V(S.company,{w:G[4]+G[5],span:2})],{h:340}),
    wRow([L('Work Location',null,{w:G[0]}),V(S.workLocation,{w:W-G[0],span:5})],{h:340}),
    wRow([L('Black Hat Supervisor Facilitating the Brief',null,{w:G[0]+G[1]+G[2]+G[3],span:4}),V(t.bh,{w:G[4]+G[5],span:2})],{h:340}),
    wRow([L('Incident Controller',C.red,{w:G[0],color:'FFFFFF'}),V(b.ic,{w:G[1]+G[2],span:2}),L('TMIC / FPOS Person',C.green,{w:G[3],color:'FFFFFF',sz:17}),V(b.tmic,{w:G[4]+G[5],span:2})],{h:340}),
    wRow([L('IC Radio Channel',C.red,{w:G[0],color:'FFFFFF'}),V(S.icChannel,{w:G[1]+G[2],span:2}),L('TMIC / FPOS Channel',C.green,{w:G[3],color:'FFFFFF',sz:17}),V(S.tmicChannel,{w:G[4]+G[5],span:2})],{h:340}),
    wRow([L('Nearest Internal Muster',C.red,{w:G[0],color:'FFFFFF'}),V(S.muster,{w:G[1]+G[2],span:2}),L('First Aider',C.green,{w:G[3],color:'FFFFFF'}),V(b.firstAider,{w:G[4]+G[5],span:2})],{h:340}),
    wRow([L('Fire Assembly Point (FAP)',C.amber,{w:G[0]}),V(S.fap,{w:G[1]+G[2],span:2,jc:'center'}),L('RVP Point',C.amber,{w:G[3]}),V(S.rvp,{w:G[4]+G[5],span:2,jc:'center'})],{h:520}),
    head('Feedback from any issues raised at the previous DAB'),body(t.prev,1000),
    band('Previous Workday'),
    head('What do the team believe worked well? Feedback from the team.'),body(t.well,1000),
    head('What do the team believe didn’t work well and how can we make it better? Feedback from the team.'),body(t.notWell,1000),
    band('Today'),
    head('Who is doing what and where are they doing it? Discuss any big risks which are new.'),body(t.today,1700),
    head('Will any other teams be working in our areas or working beside us? How will we manage this?'),body(t.other,1100),
    head('What is happening on site that the team need to know? E.g. changes to the roads, shift pattern changes.'),body(t.site,1100),
  ];
  return wP(wRun('Daily Activity Briefing',{b:true,sz:36}),{jc:'center',after:160})+wTable(G,rows);
}
function dabPage2(T){
  const G=[3300,2000,2500,2666],W=G.reduce((a,b)=>a+b,0);
  const top=wTable([W/2,W/2],[wRow([wCell(wP([wRun('ENKA ',{b:true,sz:30,color:'1F4E9A'}),wRun('UK',{b:true,sz:30,color:'1F4E9A'})]),{w:W/2}),
    wCell(wP(wRun('Daily Activity Briefing',{b:true,sz:30}),{jc:'right'}),{w:W/2})])],{noBorder:true});
  const hd=['Name (print)','SAP No. (Badge)','Trade','Signature'];
  const ps=signers(T);const n=Math.max(ps.length+4,22);
  const rows=[wRow(hd.map((h,i)=>wCell(wText(h,{b:true,sz:20,color:'FFFFFF',jc:'center'}),{w:G[i],fill:C.sign})),{h:420})];
  for(let i=0;i<n;i++){const p=ps[i]||{};rows.push(wRow([wCell(wText(p.name||'',{sz:21}),{w:G[0]}),wCell(wText(p.sap||'',{sz:21}),{w:G[1]}),wCell(wText(p.trade||'',{sz:21}),{w:G[2]}),wCell('',{w:G[3]})],{h:500}))}
  return top+wP('',{after:120})+wTable(G,rows);
}
async function buildDabDocx(trades){
  const body=trades.map((T,i)=>(i?PAGE_BREAK:'')+dabPage1(T)+PAGE_BREAK+dabPage2(T)).join('');
  const NSW='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const hf=(tag)=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} ${NSW}>${wP(wRun('NOT PROTECTIVELY MARKED',{b:true,sz:24,color:'7F7F7F'}),{jc:'center'})}</w:${tag}>`;
  const document=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NSW}><w:body>${body}
    <w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>
    <w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="900" w:right="720" w:bottom="800" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NSW}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
  const z=new JSZip();
  z.file('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
  z.file('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  z.file('word/_rels/document.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`);
  z.file('docProps/core.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Daily Activity Briefing</dc:title><dc:creator>${xe(myName())}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/,'Z')}</dcterms:created></cp:coreProperties>`);
  z.file('word/document.xml',document);z.file('word/styles.xml',styles);z.file('word/header1.xml',hf('hdr'));z.file('word/footer1.xml',hf('ftr'));
  return z.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
window.__buildDab=t=>buildDabDocx(t);window.__briefTrades=()=>Object.keys((brief||{}).trades||{});
/* ---------- events, with a check before changing someone else's entry or a DONE job ---------- */
const okUntil={};   // after confirming once, further edits to the same job go through for a short while
function guard(id,it,next,newVal,col){
  const s=st[id];const now=Date.now();
  if(okUntil[id]&&okUntil[id]>now){next();return}
  const other=s.by&&authUser&&s.by!==authUser.uid;
  const wasDone=col!=null&&s.status[col]==='DONE'&&newVal!=='DONE';
  if(!other&&!wasDone){next();return}
  const who=other?(s.byName||'a teammate'):'you';
  const what=wasDone?`<p><strong>${esc(tradeLabel(it,col))}</strong> is marked <strong>DONE</strong>${other?` by <strong>${esc(who)}</strong>`:''}${s.at?' at '+esc(fmtTime(s.at)):''}. Change it to <strong>${esc(newVal||'not set')}</strong>?</p>`
    :`<p>This job was last updated by <strong>${esc(who)}</strong>${s.at?' at '+esc(fmtTime(s.at)):''}. Change their entry?</p>`;
  openView('guard',`<h3>${wasDone?'Change a finished job?':'Change a teammate\'s entry?'}</h3>
    <p class="nf"><strong>${esc(it.tag||it.act)}</strong> · ${esc(it.act)}</p>${what}
    <p class="nc">The change is saved in the job's history with your name.</p>
    <div class="grid2"><button class="btn alt" type="button" id="gNo">Cancel</button><button class="btn" type="button" id="gYes">Change</button></div>`);
  $('#gNo').onclick=()=>{closeSheet();rerenderCard(id)};
  $('#gYes').onclick=()=>{okUntil[id]=Date.now()+120000;closeSheet();next()};
}
document.addEventListener('click',e=>{
  const b=e.target.closest('[data-act],[data-view],[data-level]');if(!b)return;
  if(b.dataset.view){filter.view=b.dataset.view;renderFilters();renderMain();return}
  if(b.dataset.level!==undefined&&!b.dataset.act){filter.level=filter.level===b.dataset.level?null:b.dataset.level;renderFilters();renderMain();return}
  if(b.closest('#panel'))return;
  const id=b.dataset.id;const s=st[id];
  if(b.dataset.act==='addnew'){if(canWrite)openAddForm(+b.dataset.sec);return}
  if(!canWrite||!s)return;const it=findItem(id);if(!it)return;
  if(b.dataset.act==='st'){const col=+b.dataset.col;const nv=s.status[col]===b.dataset.v?null:b.dataset.v;
    guard(id,it,()=>{s.status[col]=nv;rerenderCard(id);writeItem(id)},nv,col)}
  else if(b.dataset.act==='blk'){const x=b.dataset.b;
    guard(id,it,()=>{s.blockers=s.blockers.includes(x)?s.blockers.filter(y=>y!==x):[...s.blockers,x];rerenderCard(id);writeItem(id)})}
  else if(b.dataset.act==='openblk'){s.open=true;rerenderCard(id);writeItem(id);const t=$('#note-'+id);if(t)t.focus()}
  else if(b.dataset.act==='remove'){confirmRemove(id)}
});
document.addEventListener('focusin',e=>{const t=e.target;if(!(t.dataset&&t.dataset.act==='note')||!canWrite)return;
  const id=t.dataset.id;const s=st[id];const it=findItem(id);if(!s||!it||!s.note)return;
  const other=s.by&&authUser&&s.by!==authUser.uid;if(!other||(okUntil[id]&&okUntil[id]>Date.now()))return;
  t.blur();guard(id,it,()=>{const n=$('#note-'+id);if(n){n.focus();n.setSelectionRange(n.value.length,n.value.length)}})});
document.addEventListener('input',e=>{const t=e.target;if(t.dataset&&t.dataset.act==='note'&&canWrite){st[t.dataset.id].note=t.value;writeItem(t.dataset.id,900)}});
document.addEventListener('change',e=>{const t=e.target;if(t.dataset&&t.dataset.act==='addtrade'&&t.value&&canWrite){const id=t.dataset.id,it=findItem(id),v=+t.value;
  guard(id,it,()=>{st[id].extra.push(v);rerenderCard(id);writeItem(id)});if(!(okUntil[id]>Date.now()))t.value=''}});
$('#btnNew').addEventListener('click',()=>{if(!canWrite){toast('View only: ask the night shift lead for Contributor access');return}$('#fileInput').click()});
$('#fileInput').addEventListener('change',async e=>{
  const f=e.target.files[0];e.target.value='';if(!f)return;
  toast('Checking '+f.name+'…');
  try{await uploadDabs(f)}
  catch(err){openSheet(`<h3>Could not load this file</h3><p class="err">${esc(err.message||err.code||'Unknown error')}</p><button class="btn" type="button" id="cOk">OK</button>`);$('#cOk').onclick=closeSheet}
});
$('#replaceInput').addEventListener('change',async e=>{
  const f=e.target.files[0];e.target.value='';if(!f||!cur||!isAdmin)return;
  try{const u8=new Uint8Array(await f.arrayBuffer());const d=await parseDabs(u8,f.name);const b64=b64enc(u8);
    confirmReplace(cur,u8,d,f.name,b64,'f'+hashStr(b64)+u8.length.toString(36))}
  catch(err){openSheet(`<h3>Could not read this file</h3><p class="err">${esc(err.message||'Unknown error')}</p><button class="btn" type="button" id="cOk">OK</button>`);$('#cOk').onclick=closeSheet}
});
$('#btnMenu').addEventListener('click',openMenu);
$('#btnConcerns').addEventListener('click',openConcernsView);

/* ---------- add activity ---------- */
function openAddForm(si){
  const levels=[...new Set(allItems().map(i=>i.level))].sort((a,b)=>parseFloat(a)-parseFloat(b));
  const acts=[...new Set(allItems().map(i=>i.act.replace(/\s+/g,' ').trim()))];
  const secOpts=doc.sections.map((s,i)=>`<option value="${i}"${i===si?' selected':''}>${esc(s.short)}</option>`).join('');
  openView('add',`<h3>Add activity</h3><p>For work that is not on tonight's DABS. The whole team sees it, and it is exported in its own block under the sheet.</p>
    <form id="addForm" novalidate style="display:flex;flex-direction:column;gap:10px">
      <div class="grid2">
        <label for="aSec">Section<select id="aSec">${secOpts}</select></label>
        <label for="aLvl">Level<select id="aLvl">${levels.map(l=>`<option>${esc(l)}</option>`).join('')}<option value="__other">Other…</option></select></label>
      </div>
      <label for="aLvlOther" id="aLvlOtherWrap" hidden>Level (m)<input id="aLvlOther" inputmode="decimal" placeholder="+8.50"></label>
      <label for="aTag"><span id="aTagLbl"></span><input id="aTag" autocomplete="off" placeholder="e.g. 1CEX2478TY--PS005"></label>
      <label for="aAct">Activity description<input id="aAct" list="actList" autocomplete="off" placeholder="e.g. Support erection"></label>
      <datalist id="actList">${acts.map(a=>`<option value="${esc(a)}"></option>`).join('')}</datalist>
      <fieldset style="border:0;padding:0;margin:0;display:flex;flex-direction:column;gap:6px"><legend class="lbl" style="font-size:12px;color:var(--muted);margin-bottom:4px">Trades</legend><div class="bchips" id="aTrades"></div></fieldset>
      <label for="aWhy">Why added (optional)<input id="aWhy" placeholder="e.g. Fallback, scaffold on GL04 not handed over"></label>
      <p class="err" id="aErr" hidden></p>
      <div class="grid2"><button class="btn alt" type="button" id="aCancel">Cancel</button><button class="btn" type="submit">Add</button></div>
    </form>`);
  const picked=new Set();
  const fillSec=()=>{const sec=doc.sections[+$('#aSec').value];$('#aTagLbl').textContent=sec.tagLabel;picked.clear();
    $('#aTrades').innerHTML=sec.trades.map(t=>`<button type="button" class="bchip pick" data-col="${t.col}" aria-pressed="false">${esc(t.name)}</button>`).join('')};
  fillSec();
  $('#aSec').onchange=fillSec;
  $('#aLvl').onchange=()=>{$('#aLvlOtherWrap').hidden=$('#aLvl').value!=='__other'};
  $('#aTrades').onclick=e=>{const b=e.target.closest('.pick');if(!b)return;const c=+b.dataset.col;
    if(picked.has(c))picked.delete(c);else picked.add(c);b.setAttribute('aria-pressed',String(picked.has(c)))};
  $('#aCancel').onclick=closeSheet;
  $('#addForm').onsubmit=e=>{e.preventDefault();
    const sec=+$('#aSec').value;const tag=$('#aTag').value.trim(),act=$('#aAct').value.trim();
    let lvl=$('#aLvl').value;if(lvl==='__other'){const v=$('#aLvlOther').value.trim().replace(',','.');const n=parseFloat(v);lvl=isNaN(n)?v:fmtLevel(n)}
    const err=$('#aErr');
    if(!act){err.textContent='Enter the activity description.';err.hidden=false;return}
    if(!lvl){err.textContent='Enter the level.';err.hidden=false;return}
    const trades=doc.sections[sec].trades.filter(t=>picked.has(t.col)).map(t=>({col:t.col,name:t.name}));
    const id='a'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
    added.push({id,sec,added:true,level:lvl,shift:'Night',bh:'',tag,act,dayRemark:'',trades,pre:{},created:Date.now(),createdByName:myName()});
    st[id]={status:{},blockers:[],note:$('#aWhy').value.trim()?'Added on night: '+$('#aWhy').value.trim():'',extra:[],hist:[]};
    base[id]={status:{},blockers:[],note:''};
    writeItem(id);closeSheet();filter.level=null;renderAll();
    const el=$('#card-'+id);if(el)el.scrollIntoView({block:'center'});toast('Activity added for the team');};
  setTimeout(()=>$('#aTag').focus(),50);
}
function confirmRemove(id){
  const it=findItem(id);if(!it)return;const s=st[id]||{};
  const other=s.by&&authUser&&s.by!==authUser.uid;
  openView('remove',`<h3>Remove this activity?</h3><p><strong>${esc(it.tag||it.act)}</strong> was added on the night${it.createdByName?' by <strong>'+esc(it.createdByName)+'</strong>':''}. It is removed for the whole team, with its status and notes.</p>
    ${other?`<p>Last updated by <strong>${esc(s.byName||'a teammate')}</strong>${s.at?' at '+esc(fmtTime(s.at)):''}.</p>`:''}
    <div class="grid2"><button class="btn alt" type="button" id="cNo">Keep</button><button class="btn danger" type="button" id="cYes">Remove</button></div>`);
  $('#cNo').onclick=closeSheet;
  $('#cYes').onclick=()=>{
    clearTimeout(timers[id]);dirty.delete(id);added=added.filter(a=>a.id!==id);delete st[id];delete base[id];closeSheet();renderAll();
    db.doc(P('nights')+'/'+cur.id+'/items/'+id).delete().catch(e=>toast('Not removed ('+(e.code||'error')+')'));scheduleSummary();toast('Activity removed')};
}
/* ---------- sheets ---------- */
function openSheet(html){$('#panel').innerHTML=html;$('#sheet').hidden=false}
function closeSheet(){$('#sheet').hidden=true;sheetView=null}
$('#sheet').addEventListener('click',e=>{if(e.target.id==='sheet')closeSheet()});

function remarkText(s){
  const parts=[];if(s.blockers.length)parts.push('Blocker: '+s.blockers.join(', '));if(s.note.trim())parts.push(s.note.trim());return parts.join(' - ');
}
function insertOrdered(root,el){
  const ORDER=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','drawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'];
  const i=ORDER.indexOf(el.localName);
  const after=Array.from(root.children).find(c=>ORDER.indexOf(c.localName)>i);
  root.insertBefore(el,after||null);
}
function insertStyleOrdered(root,el){
  const ORDER=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst'];
  const i=ORDER.indexOf(el.localName);
  const after=Array.from(root.children).find(c=>ORDER.indexOf(c.localName)>i);
  root.insertBefore(el,after||null);
}
async function buildXlsx(){
  const zip=await JSZip.loadAsync(bytes);
  const {path}=await sheetPaths(zip);
  const sx=parseXml(await zip.file(path).async('string'));
  const sy=parseXml(await zip.file('xl/styles.xml').async('string'));
  const ss=await readShared(zip);
  const sroot=sy.documentElement,wroot=sx.documentElement;
  const E=(d,n,attrs)=>{const e=d.createElementNS(NS,n);for(const k in attrs||{})e.setAttribute(k,attrs[k]);return e};
  const getOrMake=(name)=>{let e=kid(sroot,name);if(!e){e=E(sy,name,{count:'0'});insertStyleOrdered(sroot,e)}return e};
  const fonts=getOrMake('fonts'),fills=getOrMake('fills'),xfs=getOrMake('cellXfs');
  const recount=e=>e.setAttribute('count',String(e.children.length));
  const COL={DONE:['FFC6EFCE','FF006100'],ONGOING:['FFC6EFCE','FF006100'],'NOT DONE':['FFFFC7CE','FF9C0006']};
  const fillIdx={};
  for(const k of ['FFC6EFCE','FFFFC7CE']){const f=E(sy,'fill'),p=E(sy,'patternFill',{patternType:'solid'});p.appendChild(E(sy,'fgColor',{rgb:k}));p.appendChild(E(sy,'bgColor',{indexed:'64'}));f.appendChild(p);fills.appendChild(f);fillIdx[k]=fills.children.length-1}
  recount(fills);
  const xfCache={};
  function styled(sIdx,status){
    const key=sIdx+'|'+status;if(key in xfCache)return xfCache[key];
    const base=xfs.children[sIdx]||xfs.children[0];const xf=base.cloneNode(true);
    const fBase=fonts.children[+(base.getAttribute('fontId')||0)]||fonts.children[0];
    const fnt=fBase?fBase.cloneNode(true):E(sy,'font');
    if(!kid(fnt,'b'))fnt.insertBefore(E(sy,'b'),fnt.firstChild);
    const oc=kid(fnt,'color');if(oc)fnt.removeChild(oc);
    const col=E(sy,'color',{rgb:COL[status][1]});
    const ref=Array.from(fnt.children).find(c=>['name','family','charset','scheme'].includes(c.localName));
    fnt.insertBefore(col,ref||null);fonts.appendChild(fnt);recount(fonts);
    xf.setAttribute('fontId',String(fonts.children.length-1));xf.setAttribute('applyFont','1');
    xf.setAttribute('fillId',String(fillIdx[COL[status][0]]));xf.setAttribute('applyFill','1');
    let al=kid(xf,'alignment');if(!al){al=E(sy,'alignment');xf.insertBefore(al,xf.firstChild)}
    al.setAttribute('horizontal','center');al.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
    xfs.appendChild(xf);recount(xfs);return xfCache[key]=xfs.children.length-1;
  }
  function wrapStyle(sIdx){
    const key=sIdx+'|wrap';if(key in xfCache)return xfCache[key];
    const base=xfs.children[sIdx]||xfs.children[0];const xf=base.cloneNode(true);
    let al=kid(xf,'alignment');if(!al){al=E(sy,'alignment');xf.insertBefore(al,xf.firstChild)}
    al.setAttribute('wrapText','1');al.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
    xfs.appendChild(xf);recount(xfs);return xfCache[key]=xfs.children.length-1;
  }
  // dxfs for conditional formatting
  const dxfs=getOrMake('dxfs');
  const mkDxf=(fill,font)=>{const d=E(sy,'dxf'),f=E(sy,'font');f.appendChild(E(sy,'b'));f.appendChild(E(sy,'color',{rgb:font}));d.appendChild(f);
    const fl=E(sy,'fill'),p=E(sy,'patternFill',{patternType:'solid'});p.appendChild(E(sy,'bgColor',{rgb:fill}));fl.appendChild(p);d.appendChild(fl);dxfs.appendChild(d);recount(dxfs);return dxfs.children.length-1};
  const dxGreen=mkDxf('FFC6EFCE','FF006100'),dxRed=mkDxf('FFFFC7CE','FF9C0006');

  const sheetData=kid(wroot,'sheetData');
  const origS=new Map();for(const ce of sheetData.getElementsByTagNameNS(NS,'c'))origS.set(ce.getAttribute('r'),+(ce.getAttribute('s')||0));
  const getS0=(r,c)=>origS.has(colName(c)+r)?origS.get(colName(c)+r):null;
  const rowEl=r=>{let found=null,after=null;for(const re of kids(sheetData,'row')){const n=+re.getAttribute('r');if(n===r){found=re;break}if(n>r){after=re;break}}
    if(!found){found=E(sx,'row',{r:String(r)});sheetData.insertBefore(found,after)}return found};
  const cellEl=(r,c)=>{const re=rowEl(r);re.removeAttribute('spans');let after=null;
    for(const ce of kids(re,'c')){const [cc]=splitRef(ce.getAttribute('r'));if(cc===c)return ce;if(cc>c){after=ce;break}}
    const ce=E(sx,'c',{r:colName(c)+r});re.insertBefore(ce,after);return ce};
  const getS=(r,c)=>{const re=rowEl(r);for(const ce of kids(re,'c'))if(ce.getAttribute('r')===colName(c)+r)return +(ce.getAttribute('s')||0);return null};
  const setText=(ce,text,s)=>{while(ce.firstChild)ce.removeChild(ce.firstChild);ce.removeAttribute('t');
    ce.setAttribute('t','inlineStr');const is=E(sx,'is'),t=E(sx,'t');t.setAttributeNS(XMLNS_XML,'xml:space','preserve');t.textContent=text;is.appendChild(t);ce.appendChild(is);
    if(s!=null)ce.setAttribute('s',String(s))};

  const ranges=[];
  for(const sec of doc.sections){
    if(!sec.items.length||!sec.trades.length)continue;
    const r1=sec.items[0].row,r2=sec.items[sec.items.length-1].row;
    const c1=Math.min(...sec.trades.map(t=>t.col)),c2=Math.max(...sec.trades.map(t=>t.col));
    ranges.push({ref:colName(c1)+r1+':'+colName(c2)+r2,tl:colName(c1)+r1});
    // remark header if it was missing
    {const h=cellValue(cellEl(sec.header,sec.cols.remark),ss);
      if(h==null||h===''){setText(cellEl(sec.header,sec.cols.remark),'Remark/Blocker',getS(sec.header,sec.cols.act))}}
    for(const it of sec.items){
      const s=st[it.id];
      for(const col of allCols(it)){
        const v=s.status[col]||'NOT DONE';
        const base=getS(it.row,col)??getS(it.row,sec.cols.act)??0;
        setText(cellEl(it.row,col),v,styled(base,v));
      }
      const add=remarkText(s);
      if(add){
        const text=it.dayRemark?it.dayRemark+' | '+add:add;
        const base=getS(it.row,sec.cols.remark)??getS(it.row,sec.cols.act)??0;
        setText(cellEl(it.row,sec.cols.remark),text,wrapStyle(base));
      }
    }
  }
  // activities added on the night: own block under the sheet, same columns as the section
  let maxRow=0;for(const re of kids(sheetData,'row'))maxRow=Math.max(maxRow,+re.getAttribute('r'));
  let maxCol=0;
  doc.sections.forEach((sec,si)=>{
    const extra=added.filter(a=>a.sec===si);if(!extra.length)return;
    const first=sec.items[0]?sec.items[0].row:sec.header;
    const hdrCols=sec.hdr.map(([c])=>c);if(!hdrCols.includes(sec.cols.remark))hdrCols.push(sec.cols.remark);
    const c1=Math.min(...hdrCols),c2=Math.max(...hdrCols);maxCol=Math.max(maxCol,c2);
    let r=maxRow+3;
    setText(cellEl(r,sec.titleCol),'NIGHT ADDED - '+sec.title,getS0(sec.titleRow,sec.titleCol));
    r+=2;
    for(let c=c1;c<=c2;c++){const h=sec.hdr.find(x=>x[0]===c);setText(cellEl(r,c),h?h[1]:(c===sec.cols.remark?'Remark/Blocker':''),getS0(sec.header,c))}
    const r1=r+1;
    for(const it of extra){
      r++;const s=st[it.id];
      for(let c=c1;c<=c2;c++){const base=getS0(first,c)??getS0(first,sec.cols.act)??0;
        let v='';
        if(c===sec.cols.level)v=it.level;else if(c===sec.cols.shift)v=it.shift;else if(c===sec.cols.tag)v=it.tag;
        else if(c===sec.cols.act)v=it.act;else if(c===sec.cols.bh)v='';
        else if(c===sec.cols.remark){setText(cellEl(r,c),remarkText(s),wrapStyle(base));continue}
        if(allCols(it).includes(c)){const sv=s.status[c]||'NOT DONE';setText(cellEl(r,c),sv,styled(base,sv));continue}
        setText(cellEl(r,c),v,base);
      }
    }
    if(sec.trades.length){const t1=Math.min(...sec.trades.map(t=>t.col)),t2=Math.max(...sec.trades.map(t=>t.col));
      ranges.push({ref:colName(t1)+r1+':'+colName(t2)+r,tl:colName(t1)+r1})}
    maxRow=r;
  });
  // report stamp at the top right
  const sec0=doc.sections[0];
  const addFont=(o)=>{const f=E(sy,'font');if(o.b)f.appendChild(E(sy,'b'));if(o.i)f.appendChild(E(sy,'i'));f.appendChild(E(sy,'sz',{val:String(o.sz)}));f.appendChild(E(sy,'color',{rgb:o.color}));f.appendChild(E(sy,'name',{val:'Arial'}));fonts.appendChild(f);recount(fonts);return fonts.children.length-1};
  const addXf=(fontId,o={})=>{const xf=E(sy,'xf',{numFmtId:'0',fontId:String(fontId),fillId:String(o.fill||0),borderId:String(o.border||0),xfId:'0',applyFont:'1',applyAlignment:'1',applyBorder:o.border?'1':'0'});
    xf.appendChild(E(sy,'alignment',Object.assign({vertical:o.v||'center'},o.h?{horizontal:o.h}:{},o.wrap?{wrapText:'1'}:{})));xfs.appendChild(xf);recount(xfs);return xfs.children.length-1};
  const stamp='Report exported '+new Date().toLocaleString('en-GB',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})+' by '+(myName()||'night shift');
  setText(cellEl(sec0.titleRow,sec0.cols.remark),stamp,addXf(addFont({i:true,sz:9,color:'FF7F7F7F'}),{h:'right',v:'top'}));
  maxCol=Math.max(maxCol,sec0.cols.remark);
  const cl=exportConcerns();
  if(cl.length){
    const hdrXf=xfs.children[getS0(sec0.header,sec0.cols.act)??0];const border=hdrXf?+(hdrXf.getAttribute('borderId')||0):0;
    const fT=addFont({b:true,sz:12,color:'FFC00000'}),fH=addFont({b:true,sz:10,color:'FFC00000'}),fR=addFont({sz:10,color:'FFC00000'});
    const xT=addXf(fT),xH=addXf(fH,{border,wrap:true}),xR=addXf(fR,{border,wrap:true,v:'top'});
    let r=maxRow+3;
    setText(cellEl(r,2),'CONCERN LIST · open items carried over and tonight\'s concerns ('+cl.filter(c=>!c.resolved).length+' open)',xT);
    r+=2;
    const cols=[[2,'Type'],[3,'Trade'],[4,'Entered by'],[5,'Date / time'],[6,'Description'],[11,'Status']];
    let mc=kid(wroot,'mergeCells');if(!mc){mc=E(sx,'mergeCells',{count:'0'});insertOrdered(wroot,mc)}
    const merge=ref=>{mc.appendChild(E(sx,'mergeCell',{ref}));mc.setAttribute('count',String(mc.children.length))};
    for(const [c,h] of cols)setText(cellEl(r,c),h,xH);for(let c=7;c<=10;c++)setText(cellEl(r,c),'',xH);merge('F'+r+':J'+r);
    for(const c of cl){r++;
      const vals={2:c.type,3:c.trade||'All',4:c.byName||'',5:fmtDT(c.at),6:c.desc,11:c.resolved?'Resolved '+fmtDT(c.resolvedAt)+(c.resolvedByName?' by '+c.resolvedByName:''):(cur&&c.at<cur.loadedAt?'Open (carried over)':'Open')};
      for(const k in vals)setText(cellEl(r,+k),vals[k],xR);for(let k=7;k<=10;k++)setText(cellEl(r,k),'',xR);merge('F'+r+':J'+r);
      const lines=Math.max(1,Math.ceil(String(c.desc).length/80),String(c.desc).split('\n').length,Math.ceil(String(vals[11]).length/22));
      const re=rowEl(r);re.setAttribute('ht',String(Math.max(15,14*lines)));re.setAttribute('customHeight','1');
    }
    maxRow=r;maxCol=Math.max(maxCol,11);
  }
  const dim=kid(wroot,'dimension');
  if(dim&&maxRow){const ref=dim.getAttribute('ref')||'A1';const [a,b]=ref.split(':');const end=b||a;const [ec,er]=splitRef(end);
    dim.setAttribute('ref',(b?a:'A1')+':'+colName(Math.max(ec,maxCol))+Math.max(er,maxRow))}
  // conditional formatting + dropdowns so the file keeps working in Excel
  let prio=0;for(const r of wroot.getElementsByTagNameNS(NS,'cfRule'))prio=Math.max(prio,+(r.getAttribute('priority')||0));
  for(const rg of ranges){
    const cf=E(sx,'conditionalFormatting',{sqref:rg.ref});
    const add=(dx,f)=>{const r=E(sx,'cfRule',{type:'expression',dxfId:String(dx),priority:String(++prio)});const fe=E(sx,'formula');fe.textContent=f;r.appendChild(fe);cf.appendChild(r)};
    add(dxGreen,`OR(UPPER(TRIM(${rg.tl}))="DONE",UPPER(TRIM(${rg.tl}))="ONGOING")`);
    add(dxRed,`UPPER(TRIM(${rg.tl}))="NOT DONE"`);
    insertOrdered(wroot,cf);
  }
  if(ranges.length){
    let dvs=kid(wroot,'dataValidations');if(!dvs){dvs=E(sx,'dataValidations',{count:'0'});insertOrdered(wroot,dvs)}
    const dv=E(sx,'dataValidation',{type:'list',allowBlank:'1',showErrorMessage:'1',errorTitle:'Status',error:'Use DONE, ONGOING or NOT DONE',sqref:ranges.map(r=>r.ref).join(' ')});
    const f1=E(sx,'formula1');f1.textContent='"DONE,ONGOING,NOT DONE"';dv.appendChild(f1);dvs.appendChild(dv);dvs.setAttribute('count',String(dvs.children.length));
  }
  zip.file(path,serXml(sx));zip.file('xl/styles.xml',serXml(sy));
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
function outName(){const base=doc.fileName.replace(/\.xlsx$/i,'').replace(/\s*-\s*Status$/i,'');return base+' - Status.xlsx'}
/* ---------- export: reminder first, then the file ---------- */
$('#btnExport').addEventListener('click',()=>{
  if(!doc)return;
  const c=counts();const open=openConcerns().length;
  openView('exportCheck',`<h3>Before you export</h3>
    <p><strong>Did you update the concern list?</strong> ${open?`There ${open===1?'is':'are'} <strong>${open}</strong> open ${open===1?'item':'items'}; ${open===1?'it prints':'they print'} in red under the DABS.`:'It is empty, so nothing will print under the DABS.'}</p>
    ${c.unset?`<p><strong>${c.unset}</strong> ${c.unset===1?'activity has':'activities have'} not been updated and will be exported as NOT DONE.</p>`:''}
    <div class="grid2"><button class="btn alt" type="button" id="exCn">Open concern list</button><button class="btn" type="button" id="exGo">Export Excel</button></div>
    <button class="ghost" type="button" id="exNo">Cancel</button>`);
  $('#exNo').onclick=closeSheet;$('#exCn').onclick=openConcernsView;$('#exGo').onclick=()=>{closeSheet();doExport()};
});
async function doExport(){
  let out;try{out=await buildXlsx()}catch(err){openSheet(`<h3>Export failed</h3><p class="err">${esc(err.message)}</p><button class="btn" type="button" id="cOk">OK</button>`);$('#cOk').onclick=closeSheet;return}
  window.__lastExport=out;
  deliverFile(outName(),out,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Excel ready');
}
function saveToPhone(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},4000)}
function deliverFile(name,u8,mime,heading,back){
  const blob=new Blob([u8],{type:mime});const file=new File([blob],name,{type:mime});
  const canShareFile=!!(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]}));
  window.__lastFile={name,size:u8.length};
  openView('file',`<h3>${esc(heading)}</h3><p>${esc(name)}</p>
    <div class="grid2"><button class="btn alt" type="button" id="xShare"${canShareFile?'':' hidden'}>Share file</button><button class="btn" type="button" id="xSave">Save to phone</button></div>
    <p class="nc" id="xMsg"></p>
    <button class="ghost" type="button" id="xClose">${back?'Back':'Close'}</button>`);
  $('#xClose').onclick=()=>back?back():closeSheet();
  $('#xSave').onclick=()=>{saveToPhone(blob,name);toast('Saved to Downloads')};
  $('#xShare').onclick=async()=>{
    try{await navigator.share({files:[file],title:name})}
    catch(e){if(e&&e.name==='AbortError')return;   // the person closed the share menu
      saveToPhone(blob,name);
      $('#xMsg').textContent='This phone would not share the file directly, so it was saved to Downloads. Attach it from there in WhatsApp or Outlook.';
      toast('Saved to Downloads')}};
}

/* ---------- archive ---------- */
let installPrompt=null;
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();installPrompt=e});
const standalone=()=>matchMedia('(display-mode: standalone)').matches||navigator.standalone===true;
function miniBar(sm){if(!sm||!sm.total)return '';const p=n=>(n/sm.total*100)+'%';
  return `<div class="bar sm"><span style="width:${p(sm.DONE)};background:var(--ok)"></span><span style="width:${p(sm.ONGOING)};background:var(--go)"></span><span style="width:${p(sm['NOT DONE'])};background:var(--no)"></span></div>`}
function openArchive(){
  const list=byDate();
  openView('archive',`<h3>Archive</h3><p>${list.length} ${list.length===1?'night':'nights'} shown${olderNights.length?'':' (the last '+NIGHTS_ON_OPEN+')'}.</p>
    <div class="nights">${list.map(n=>{const sm=n.summary;const isCur=cur&&n.id===cur.id;
      return `<div class="night${isCur?' cur':''}" id="nt-${n.id}">
        <div class="nl"><div class="nt"><strong>${esc(n.label||'—')}</strong>${isCur?'<span class="badge">On screen</span>':''}</div>
          <span class="nf">${esc(n.fileName)}</span>${miniBar(sm)}
          <span class="nc">${sm?`DONE ${sm.DONE} · ONGOING ${sm.ONGOING} · NOT DONE ${sm['NOT DONE']}${sm.unset?` · Not updated ${sm.unset}`:''}`:'Not started'}</span>
          ${n.createdByName?`<span class="nc">Loaded by ${esc(n.createdByName)} · ${esc(fmtDT(n.loadedAt))}</span>`:''}
          ${n.replacedByName?`<span class="nc">Replaced by ${esc(n.replacedByName)} · ${esc(fmtDT(n.replacedAt))}</span>`:''}</div>
        <div class="nb">${isCur?'':`<button class="btn alt" type="button" data-arch="open" data-n="${n.id}">Open</button>`}
          ${isAdmin?`<button class="link danger" type="button" data-arch="del" data-n="${n.id}">Delete</button>`:''}</div></div>`}).join('')}</div>
    ${olderNights.length?'':`<button class="btn alt" type="button" id="arOlder">Show older nights</button>`}
    <button class="ghost" type="button" id="arClose">Close</button>`);
  $('#arClose').onclick=closeSheet;
  const ol=$('#arOlder');if(ol)ol.onclick=async()=>{ol.disabled=true;ol.textContent='Loading…';try{await loadOlderNights();openArchive()}catch(e){toast('Could not load ('+(e.code||'error')+')');ol.disabled=false}};
  $('#panel').onclick=async e=>{const b=e.target.closest('[data-arch]');if(!b)return;const n=allNights().find(x=>x.id===b.dataset.n);if(!n)return;
    if(b.dataset.arch==='open'){try{closeSheet();await openNight(n);window.scrollTo(0,0);toast('Opened '+n.label)}catch(err){toast(err.message)}}
    else if(b.dataset.arch==='del'){
      const row=$('#nt-'+n.id);row.innerHTML=`<div class="nl"><strong>Delete ${esc(n.label)} for everyone?</strong><span class="nf">Its statuses, notes and added activities are removed for the whole team.</span></div>
        <div class="nb"><button class="btn alt" type="button" data-arch="keep" data-n="${n.id}">Keep</button><button class="btn danger" type="button" data-arch="really" data-n="${n.id}">Delete</button></div>`}
    else if(b.dataset.arch==='keep'){openArchive()}
    else if(b.dataset.arch==='really'){
      if(!navigator.onLine){toast('Deleting a night needs signal');return}
      b.disabled=true;b.textContent='Deleting…';
      try{const its=await db.collection(P('nights')+'/'+n.id+'/items').get();
        for(const d of its.docs)await db.doc(P('nights')+'/'+n.id+'/items/'+d.id).delete();
        await db.doc(P('files')+'/'+(n.fileId||n.id)).delete();
        const key=nightDateKey(n);if(key){const l=await db.doc(P('dates')+'/'+key).get();if(l.exists&&l.data().nightId===n.id)await db.doc(P('dates')+'/'+key).delete()}
        const hm=db.doc(P('tagHist')+'/'+monthKey(n.loadedAt));const h=await hm.get();
        if(h.exists){const ns=Object.assign({},(h.data()||{}).nights||{});if(ns[n.id]){delete ns[n.id];await hm.set({nights:ns})}}
        tagHist.delete(n.id);olderNights=olderNights.filter(x=>x.id!==n.id);
        await db.doc(P('nights')+'/'+n.id).delete();toast('Night deleted');openArchive()}
      catch(err){toast('Delete failed ('+(err.code||'error')+')');openArchive()}}
  };
}

/* ---------- team approvals (admins) ---------- */
let unsubTeam=null;
const ROLES=[['viewer','Viewer'],['contributor','Contributor'],['admin','Admin']];
function openTeam(){
  openView('team',`<h3>Team & approvals</h3>
    <p>New people appear here after they create an account in the app. <strong>Contributors</strong> mark jobs, add activities and concerns. <strong>Viewers</strong> only read. <strong>Admins</strong> have all your rights: approve people, delete nights, replace a DABS, change settings and the people list.</p>
    <div id="teamList" class="nights"><p class="loading">Loading…</p></div>
    <button class="ghost" type="button" id="tmClose">Close</button>`);
  const stop=()=>{if(unsubTeam)unsubTeam();unsubTeam=null};
  $('#tmClose').onclick=()=>{stop();closeSheet()};
  stop();
  unsubTeam=db.collection(P('members')).onSnapshot(snap=>{
    const el=$('#teamList');if(!el||sheetView!=='team'){stop();return}
    const order={pending:0,contributor:1,admin:1,viewer:1,removed:2};
    const ms=snap.docs.map(d=>Object.assign({uid:d.id},d.data())).sort((a,b)=>(order[a.role]??1)-(order[b.role]??1)||String(a.name||a.email).localeCompare(String(b.name||b.email)));
    el.innerHTML=ms.length?ms.map(m=>`<div class="night" id="tm-${m.uid}"><div class="nl"><div class="nt"><strong>${esc(m.name||m.email)}</strong>${m.role==='pending'?'<span class="badge">Waiting</span>':''}${m.role==='removed'?'<span class="nc">Removed</span>':''}</div><span class="nf">${esc(m.email)}</span></div>
      <div class="nb">${m.role==='pending'?`<button class="btn" type="button" data-tm="approve" data-u="${m.uid}">Approve</button><button class="link danger" type="button" data-tm="remove" data-u="${m.uid}">Reject</button>`
        :m.role==='removed'?`<button class="btn alt" type="button" data-tm="restore" data-u="${m.uid}">Restore</button>`
        :`<select data-tm="role" data-u="${m.uid}" aria-label="Role for ${esc(m.name||m.email)}">${ROLES.map(([v,l])=>`<option value="${v}"${m.role===v?' selected':''}>${l}</option>`).join('')}</select><button class="link danger" type="button" data-tm="remove" data-u="${m.uid}">Remove</button>`}</div></div>`).join('')
      :'<p>No one has asked to join yet. Send your team the app link; when they create an account they appear here.</p>';
  },e=>{const el=$('#teamList');if(el)el.innerHTML=`<p class="err">Could not load the team (${esc(e.code||'error')}).</p>`});
  $('#panel').onclick=async e=>{const b=e.target.closest('button[data-tm]');if(!b)return;const u=b.dataset.u;
    const role={approve:'contributor',remove:'removed',restore:'viewer'}[b.dataset.tm];
    try{await db.doc(P('members')+'/'+u).update({role});toast({approve:'Approved as Contributor',remove:'Access removed',restore:'Restored as Viewer'}[b.dataset.tm])}
    catch(err){toast('Not changed ('+(err.code||'error')+')')}};
  $('#panel').onchange=async e=>{const s=e.target.closest('select[data-tm="role"]');if(!s)return;
    try{await db.doc(P('members')+'/'+s.dataset.u).update({role:s.value});toast('Role updated')}catch(err){toast('Not changed ('+(err.code||'error')+')')}};
}

/* ---------- sign in ---------- */
let fsDb=null,auth=null,unsubMember=null;
const AUTH_ERR={'auth/invalid-credential':'Email or password is wrong.','auth/wrong-password':'Email or password is wrong.','auth/user-not-found':'No account with this email. Create one first.',
  'auth/email-already-in-use':'An account with this email already exists. Sign in instead.','auth/weak-password':'Use a password of at least 6 characters.',
  'auth/invalid-email':'Check the email address.','auth/network-request-failed':'No signal. Signing in needs a connection the first time.',
  'auth/popup-closed-by-user':'Google sign-in was closed.','auth/popup-blocked':'The Google window was blocked. Use email and password instead.',
  'auth/too-many-requests':'Too many tries. Wait a few minutes and try again.','auth/operation-not-supported-in-this-environment':'Google sign-in does not work here. Use email and password.'};
const authMsg=e=>AUTH_ERR[e&&e.code]||('Something went wrong ('+((e&&e.code)||'error')+').');
function screen(html){doc=null;cur=null;renderTop();$('#banner').innerHTML='';$('#main').innerHTML=html}
function showAuth(mode){
  const create=mode==='create';
  screen(`<section class="empty auth"><h2>${create?'Create your account':'Sign in'}</h2>
    <p>${create?'After you create it, the night shift lead approves you. Then you can mark jobs from your phone.':'Night shift DABS for the Turbine Building piping team.'}</p>
    <form id="authForm" class="authform" novalidate>
      ${create?'<label for="auName">Your name<input id="auName" autocomplete="name" placeholder="e.g. Lee Shepherd"></label>':''}
      <label for="auEmail">Email<input id="auEmail" type="email" autocomplete="email" inputmode="email"></label>
      <label for="auPass">Password<input id="auPass" type="password" autocomplete="${create?'new-password':'current-password'}" minlength="6"></label>
      <p class="err" id="auErr" hidden></p>
      <button class="btn" type="submit">${create?'Create account':'Sign in'}</button>
      <button class="link" type="button" id="auSwitch">${create?'I already have an account':'New here? Create an account'}</button>
      ${create?'':'<button class="link" type="button" id="auForgot">Forgot password</button>'}
      <div class="or"><span>or</span></div>
      <button class="btn alt" type="button" id="auGoogle">Continue with Google</button>
    </form></section>`);
  const err=m=>{const p=$('#auErr');p.textContent=m;p.hidden=false};
  $('#auSwitch').onclick=()=>showAuth(create?'signin':'create');
  const fg=$('#auForgot');if(fg)fg.onclick=async()=>{const em=$('#auEmail').value.trim();if(!em){err('Enter your email first, then tap Forgot password.');return}
    try{await sendPasswordResetEmail(auth,em);toast('Password reset email sent to '+em)}catch(e){err(authMsg(e))}};
  $('#auGoogle').onclick=async()=>{try{await signInWithPopup(auth,new GoogleAuthProvider())}catch(e){err(authMsg(e))}};
  $('#authForm').onsubmit=async e=>{e.preventDefault();
    const em=$('#auEmail').value.trim(),pw=$('#auPass').value;const nm=create?$('#auName').value.trim():'';
    if(create&&!nm){err('Enter your name so the team knows who updated a job.');return}
    if(!em||!pw){err('Enter your email and password.');return}
    const btn=e.submitter||$('#authForm .btn');btn.disabled=true;
    try{
      if(create){const c=await createUserWithEmailAndPassword(auth,em,pw);await updateProfile(c.user,{displayName:nm});
        try{await sendEmailVerification(c.user)}catch(_){}
        authUser=c.user;startMember(c.user)}
      else await signInWithEmailAndPassword(auth,em,pw);
    }catch(x){err(authMsg(x));btn.disabled=false}};
}
function showWaiting(){
  screen(`<section class="empty"><h2>Waiting for approval</h2>
    <p>Your account <strong>${esc(authUser&&authUser.email||'')}</strong> has been sent to the night shift lead. The DABS opens here as soon as you are approved; you can keep this screen open.</p>
    <button class="btn alt" type="button" id="wOut">Sign out</button></section>`);
  $('#wOut').onclick=fbSignOut;
}
function showVerify(u){
  screen(`<section class="empty"><h2>Verify your email</h2>
    <p>This account manages the team, so the email must be confirmed first. Open the link sent to <strong>${esc(u.email)}</strong>, then come back.</p>
    <div class="grid2"><button class="btn alt" type="button" id="vSend">Send the email again</button><button class="btn" type="button" id="vDone">I've verified</button></div>
    <button class="link" type="button" id="vOut">Sign out</button></section>`);
  $('#vSend').onclick=async()=>{try{await sendEmailVerification(u);toast('Email sent')}catch(e){toast(authMsg(e))}};
  $('#vDone').onclick=async()=>{try{await u.reload();await u.getIdToken(true)}catch(e){}if(auth.currentUser&&auth.currentUser.emailVerified)startMember(auth.currentUser);else toast('Not verified yet. Open the link in the email first.')};
  $('#vOut').onclick=fbSignOut;
}
async function fbSignOut(){if(unsubMember){unsubMember();unsubMember=null}if(unsubTeam){unsubTeam();unsubTeam=null}stopListening();closeSheet();try{await signOut(auth)}catch(e){}}
function subscribeAll(){subscribeSettings();subscribePeople();subscribeConcerns();subscribeTagHist();subscribeNights()}
function startMember(u){
  if(unsubMember){unsubMember();unsubMember=null}
  const owner=[].concat(OWNER_EMAIL||[]).map(x=>String(x).toLowerCase()).includes((u.email||'').toLowerCase());
  if(owner&&!u.emailVerified){showVerify(u);return}
  const r=fdoc(fsDb,P('members')+'/'+u.uid);
  unsubMember=fsnap(r,async s=>{
    let role=owner?'admin':(s.exists()?s.data().role:null);
    if(!s.exists()&&!owner){role='pending';
      setDoc(r,{email:u.email||'',name:myName(),role:'pending',requestedAt:Date.now()}).catch(()=>{})}
    else if(s.exists()&&!owner&&myName()&&s.data().name!==myName())updateDoc(r,{name:myName()}).catch(()=>{});
    if(!role||role==='pending'||role==='removed'){canWrite=false;isAdmin=false;stopListening();showWaiting();return}
    canWrite=role==='contributor'||role==='admin';isAdmin=role==='admin';
    document.documentElement.classList.toggle('ro',!canWrite);
    if(!prepBy)prepBy=myName();
    updateSync();
    if(!unsubNights){setSync('load');renderAll();subscribeAll()}else{renderAll();updateSync()}
  },()=>{showWaiting()});
}

/* ---------- boot ---------- */
(()=>{
  loadLocal();prepBy=meta.prepBy||'';
  if(ENV!=='live'){document.documentElement.classList.add('env-test');const tb=$('#testBar');if(tb)tb.hidden=false}
  renderAll();
  if(!firebaseConfig||!firebaseConfig.apiKey||/PASTE/.test(firebaseConfig.apiKey)){
    screen(`<section class="empty"><h2>Almost ready</h2><p>Paste your Firebase settings into <strong>firebase-config.js</strong> and upload it again. The setup guide shows where to find them.</p></section>`);return}
  const app=initializeApp(firebaseConfig);
  try{fsDb=initializeFirestore(app,{localCache:persistentLocalCache({tabManager:persistentMultipleTabManager()})})}
  catch(e){fsDb=initializeFirestore(app,{})}
  auth=getAuth(app);
  const wrap=s=>({id:s.id,exists:s.exists(),data:()=>s.data(),metadata:s.metadata});
  const mkDoc=p=>{const r=fdoc(fsDb,p);return {id:r.id,path:p,get:async()=>wrap(await getDoc(r)),set:d=>setDoc(r,d),merge:d=>setDoc(r,d,{merge:true}),
    create:d=>setDoc(r,d),   // the database rules refuse it when the document already exists
    update:d=>updateDoc(r,d),delete:()=>deleteDoc(r),
    onSnapshot:(n,e,m)=>fsnap(r,{includeMetadataChanges:!!m},s=>n(wrap(s)),e)}};
  const mkQ=(p,cons)=>{const q=()=>cons.length?fquery(fcol(fsDb,p),...cons):fcol(fsDb,p);return {
    orderBy:(f,d)=>mkQ(p,[...cons,forder(f,d||'asc')]),limit:n=>mkQ(p,[...cons,flimit(n)]),
    get:async()=>{const s=await getDocs(q());return {docs:s.docs.map(wrap),size:s.size,empty:s.empty}},
    onSnapshot:(n,e,m)=>fsnap(q(),{includeMetadataChanges:!!m},s=>n({docs:s.docs.map(wrap),size:s.size,empty:s.empty,metadata:s.metadata,
      docChanges:()=>s.docChanges().map(c=>({type:c.type,doc:wrap(c.doc)}))}),e)}};
  db={doc:mkDoc,collection:p=>mkQ(p,[])};
  setSync('load');
  onAuthStateChanged(auth,u=>{
    authUser=u;
    if(!u){if(unsubMember){unsubMember();unsubMember=null}stopListening();canWrite=false;isAdmin=false;showAuth('signin');return}
    startMember(u);
  });
})();
