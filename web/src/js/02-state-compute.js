/* ---------- trạng thái ---------- */
const S={thang:'2026-11',step:1,tab:'td',prdot:{},po:{},pn:{},dh:{},ds:{},prbs:{},migo:{},ck:null,hd:{},dn:{},pb:null,nhan:null,td:{},dm:{},khach:null,dutru:null,tt:{},log:[],names:{}};
const ED={tdKey:null,td:null,tdDirty:false,nhom:'NV',week:null,khach:null,khDirty:false,dmMon:null,dmRows:null,dmDirty:false,lines:null,lnDirty:false,filt:{q:'',ncc:'',n2:''},paste:false,ai:null};
let unsubMonth=[];

function toast(m){const t=$('#toast');t.textContent=m;t.hidden=false;clearTimeout(toast.h);toast.h=setTimeout(()=>t.hidden=true,3200);}
async function log(buoc,viec){ try{ await store.add('nhatky',{ts:Date.now(),thang:S.thang,buoc,viec,ai:uid}); }catch(e){} }
async function setTT(key,ok,extra){ const tt={...S.tt}; tt[key]={ok,luc:Date.now(),ai:uid,...(extra||{})}; await store.set('trangthai/'+S.thang,tt); }
const locked=()=>!!(S.tt.chot&&S.tt.chot.ok);
const W=()=>canWrite&&!locked();

/* ---------- đăng ký dữ liệu ---------- */
function subscribeMonth(){
  unsubMonth.forEach(f=>f()); unsubMonth=[];
  S.khach=null;S.dutru=null;S.tt={};ED.khach=null;ED.lines=null;ED.khDirty=ED.lnDirty=false;
  unsubMonth.push(store.watchDoc('khach/'+S.thang,d=>{S.khach=d;if(!ED.khDirty)ED.khach=null;render();}));
  unsubMonth.push(store.watchDoc('dutru/'+S.thang,d=>{S.dutru=d;if(!ED.lnDirty)ED.lines=null;if(!(ED.pb&&ED.pb.dirty))ED.pb=null;render();}));
  S.pb=null;S.nhan=null;ED.pb=null;ED.sel=null;ED.bs={};
  unsubMonth.push(store.watchDoc('pb/'+S.thang,d=>{S.pb=d;if(!(ED.pb&&ED.pb.dirty))ED.pb=null;render();}));
  unsubMonth.push(store.watchDoc('nhanhang/'+S.thang,d=>{S.nhan=d;render();}));
  unsubMonth.push(store.watchDoc('trangthai/'+S.thang,d=>{S.tt=d||{};render();}));
}
function subscribeAll(){
  store.watchCol('thucdon',a=>{S.td={};a.forEach(d=>S.td[d.id]=d);if(!ED.tdDirty)ED.td=null;render();});
  store.watchCol('dinhmuc',a=>{S.dm={};a.forEach(d=>S.dm[d.id]=d);if(!ED.dmDirty)ED.dmRows=null;render();});
  store.watchDoc('kho/ton',d=>{S.kho=d;render();});
  store.watchDoc('kho/map',d=>{S.khomap=d;if(!ED.mapDirty)ED.map=null;render();});
  store.watchCol('hoadon',a=>{S.hd={};a.forEach(d=>S.hd[d.id]={...d});render();});
  store.watchCol('dntt',a=>{S.dn={};a.forEach(d=>S.dn[d.id]={...d});render();});
  store.watchCol('doisoat',a=>{S.ds={};a.forEach(d=>S.ds[d.id]={...d});render();});
  store.watchCol('prbs',a=>{S.prbs={};a.forEach(d=>S.prbs[d.id]={...d});render();});
  store.watchCol('migo',a=>{S.migo={};a.forEach(d=>S.migo[d.id]={...d});render();});
  store.watchDoc('cauhinh/chuky',d=>{S.ck=d;render();});
  store.watchCol('phieunhan',a=>{S.pn={};a.forEach(d=>S.pn[d.id]={...d});render();});
  store.watchCol('donhang',a=>{S.dh={};a.forEach(d=>S.dh[d.id]={...d});render();});
  store.watchCol('po',a=>{S.po={};a.forEach(d=>S.po[d.id]={...d});render();});
  store.watchCol('prdot',a=>{S.prdot={};a.forEach(d=>S.prdot[d.id]={...d});render();});
  store.watchCol('lenh',a=>{S.lenh={};a.forEach(d=>S.lenh[d.id]={...d});renderLenh();applyLenh();render();});
  store.watchCol('nhatky',a=>{S.log=a;resolveNames();renderLog();},{order:'ts',dir:'desc',limit:80});
  subscribeMonth();
}
async function resolveNames(){
  if(!user||!user.profiles) return;
  const ids=[...new Set(S.log.map(l=>l.ai).concat(Object.values(S.tt).map(x=>x&&x.ai)).filter(Boolean))].filter(i=>!(i in S.names));
  if(!ids.length) return;
  try{const ps=await user.profiles(ids);ids.forEach(i=>S.names[i]=(ps[i]&&ps[i].name)||'');renderLog();}catch(e){}
}
const nm=id=>id?(S.names[id]||'Người dùng'):'Claude';

/* ---------- thực đơn tháng ---------- */
function monthMenu(){
  const ds=new Set(monthDays(S.thang)),out=[];
  for(const doc of Object.values(S.td)) for(const it of (doc.items||[])) if(ds.has(it.d)) out.push({...it,nhom:doc.nhom});
  return out;
}
const dishesOfMonth=()=>{const m=new Map();monthMenu().forEach(it=>{const k=slug(it.mon);if(!m.has(k))m.set(k,{mon:it.mon,n:0});m.get(k).n++;});return [...m.entries()].map(([k,v])=>({k,...v})).sort((a,b)=>a.mon.localeCompare(b.mon,'vi'));};
function khachDays(){const src=ED.khach||(S.khach&&S.khach.days)||{};return src;}

/* ---------- TÍNH DỰ TRÙ ---------- */
function khoMapAll(){ const m={}; Object.entries(KHOMAP0).forEach(([k,v])=>m[k]=v[0]); Object.assign(m,(S.khomap&&S.khomap.map)||{},ED.map||{}); return m; }
function khoByMa(){ if(!S.kho) return null; const map=khoMapAll(), out=new Map();
  Object.entries(S.kho.items||{}).forEach(([id,v])=>{ const ma=map[id]; if(!ma||v.act===false) return; const o=out.get(ma)||{ton:0,min:0}; o.ton+=v.ton||0; o.min+=v.min||0; out.set(ma,o); });
  return out; }
function compute(dpT,dpK,withHist,useKho,histNgoai){ if(histNgoai==null) histNgoai=true;
  const days=monthDays(S.thang),K=(S.khach&&S.khach.days)||{};
  const slot={}; for(const it of monthMenu()){ const k=it.d+'|'+it.nhom+'|'+it.b; (slot[k]=slot[k]||[]).push(it); }
  const covered=new Set(),missing=new Set(),qM={},inRecipe=new Set(); let menuN=0;
  for(const d of days){ const k1=+d.slice(8)<=15, r=K[d]||[];
    for(const [g,N] of Object.entries(NHOM)){ if(N.chung) continue;
      for(const b of N.bua){ const idx=N.idx[b]; const its=slot[d+'|'+g+'|'+b]||(g!=='NV'?slot[d+'|NTC|'+b]:null); if(!its||!its.length) continue;
        covered.add(d+'|'+idx); const s=+r[idx]||0;
        const c1=its.filter(i=>/^chọn 1/i.test(i.loai||'')).length;
        for(const it of its){ menuN++; const f=(it.tl!=null&&it.tl!=='')?(+it.tl)/100:(c1&&/^chọn 1/i.test(it.loai||'')?1/c1:1); const rec=(it.nl&&it.nl.length)?it:S.dm[slug(it.mon)]; if(!rec||!(rec.nl||[]).length){missing.add(it.mon);continue;}
          for(const n of rec.nl){ if(!IT.has(n.ma)) continue; inRecipe.add(n.ma); const o=qM[n.ma]||(qM[n.ma]={t:0,k1:0}); o.t+=n.q*s*f; if(k1) o.k1+=n.q*s*f; } } } } }
  let tot=0,totK1=0,unc=0,uncK1=0;
  for(const d of days){const r=K[d]||[],k1=+d.slice(8)<=15; for(let i=0;i<NK;i++){const s=+r[i]||0;tot+=s;if(k1)totK1+=s;if(!covered.has(d+'|'+i)){unc+=s;if(k1)uncK1+=s;}}}
  const kho=useKho?khoByMa():null;
  const prev=new Map(((S.dutru&&S.dutru.lines)||[]).map(l=>[l.ma,l]));
  const lines=[];
  for(const x of ITEMS){
    const m=qM[x.ma]; let hb=0,hbK1=0;
    if(withHist&&x.h){const rate=x.h[0]/SUAT_T9; const ngoai=!inRecipe.has(x.ma)&&histNgoai, base=ngoai?tot:unc, bK1=ngoai?totK1:uncK1; hb=rate*base; hbK1=rate*bK1;}
    const t=(m?m.t:0)+hb, kh=kho&&x.n2==='Đồ khô'?kho.get(x.ma):null;
    if(t<=0.0001&&!(kh&&kh.min>kh.ton)) continue;
    const k1=(m?m.k1:0)+hbK1, dp=x.n2==='Đồ khô'?dpK:dpT;
    let need=t*(1+dp/100); if(kh) need=Math.max(0,need-kh.ton+kh.min);
    const dx=roundQ(need,x.dvt), p=prev.get(x.ma);
    const keep=p&&p.e; const ncc=(p&&p.ncc!=null&&x.ncc.some(n=>n.ncc===p.ncc))?p.ncc:(x.ncc[0]?x.ncc[0].ncc:'');
    lines.push({ma:x.ma,tm:r3(m?m.t:0),ls:r3(hb),dx,k1r:t?+(k1/t).toFixed(4):0.5,chot:keep?p.chot:dx,e:!!keep,ncc,note:p?(p.note||''):'',ton:kh?r3(kh.ton):null,mn:kh?r3(kh.min):null});
  }
  return {lines,warn:{missing:[...missing],tot,unc,menuN},params:{dpT,dpK,withHist,histNgoai,useKho:!!kho}};
}
const r3=n=>Math.round(n*1000)/1000;
function roundQ(q,dvt){ if(/^(KG|L|LIT|LÍT)$/.test(dvt)) return Math.ceil(q*10-1e-9)/10; return Math.ceil(q-1e-9); }
const priceOf=(x,ncc)=>{const n=x.ncc.find(n=>n.ncc===ncc)||x.ncc[0];return n?n.g:0;};

