/* ---------- cấu hình: build.py chèn từ web/config/app.config.json ---------- */
const CFG=__CONFIG__;
"use strict";
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const VND=n=>Math.round(n||0).toLocaleString('vi-VN');
const TR=n=>(Math.round((n||0)/1e5)/10).toLocaleString('vi-VN')+' tr';
const Q=n=>(+n||0).toLocaleString('vi-VN',{maximumFractionDigits:2});
const QI=n=>(+n||0).toLocaleString('vi-VN',{maximumFractionDigits:3,useGrouping:false});
const QG=n=>(+n||0).toLocaleString('vi-VN',{maximumFractionDigits:3});
const pnumVN=v=>{const s=String(v??'').replace(/\s/g,''); if(/,/.test(s)||/^[1-9]\d{0,2}(\.\d{3})+$/.test(s)) return pnum(s.replace(/\./g,'')); return pnum(s);};
const pnum=v=>{if(v===''||v==null)return 0;const n=parseFloat(String(v).replace(/\s/g,'').replace(',','.'));return isFinite(n)?n:0;};

/* ---------- dữ liệu nhúng ---------- */
const RAW=JSON.parse($('#data').textContent);
const mkItem=a=>({ma:a[0],ten:a[1],n2:a[2],n3:a[3],dvt:a[4],ncc:a[5].map(x=>({ncc:x[0],m:x[1],g:x[2]})),h:a[6]||null});
const ITEMS=RAW.items.map(mkItem), IT=new Map(ITEMS.map(x=>[x.ma,x])); let GIA_V=RAW.v||'T10/2026';
function setItems(arr,v){ ITEMS.splice(0,ITEMS.length,...arr.map(mkItem)); IT.clear(); ITEMS.forEach(x=>IT.set(x.ma,x)); GIA_V=v||GIA_V; const dl=document.getElementById('dl-item'); if(dl) dl.innerHTML=ITEMS.map(x=>`<option value="${x.ma}">${esc(x.ten)} (${esc(x.dvt)})</option>`).join(''); }
const SUAT_T9=10111;                         // T9: NV 5.240 + nội trú 4.871 suất (gồm bữa phụ)
const CHI_T9=439711957;                       // thực chi T9 theo Công nợ.xlsx
const DM_SUAT=CHI_T9/SUAT_T9;                 // ≈ 43.489 đ/suất: ngân sách tham chiếu theo khách
const NV_TB={1:[181,30],2:[160,27],3:[156,30],4:[189,27],5:[185,30],6:[112,26],0:[38,25]}; // TB suất NV T9 theo thứ (trưa, tối)
const NT_TB={TH:[26,29,24,22],VIP:[1,2,2,3],AN:[10,11,11,22]};  // TB suất/ngày T9: sáng, trưa, tối, phụ · An Nhiên = VVIP khoa Sản + VVIP Nội trú Ngoại; VIP = VVIP các khoa còn lại
const B4=['Sáng','Trưa','Tối','Phụ'];
const NHOM={NV:{ten:'Nhân viên',bua:['Trưa','Tối'],idx:{'Trưa':0,'Tối':1}},
  NTC:{ten:'Nội trú chung',bua:B4,chung:true},
  TH:{ten:'Nội trú Thường',bua:B4,idx:{'Sáng':2,'Trưa':3,'Tối':4,'Phụ':5}},
  VIP:{ten:'Nội trú VIP',bua:B4,idx:{'Sáng':6,'Trưa':7,'Tối':8,'Phụ':9}},
  AN:{ten:'Gói An Nhiên',bua:B4,idx:{'Sáng':10,'Trưa':11,'Tối':12,'Phụ':13}}};
const NK=14, Z=()=>Array(NK).fill(0);
const CG=[['Nhân viên',['Trưa','Tối']],['Nội trú Thường',B4],['Nội trú VIP',B4],['Gói An Nhiên',B4]];
const KHO={id:CFG.khoSheetId,server:'Google Sheets'};
const KHOMAP0=__KHOMAP__;                     // gợi ý ghép mã kho → mã hàng theo tên [mã, độ khớp]
const STEPS=[['Dự trù','Thực đơn · khách · ngân sách'],['Lập PR','Theo tháng, duyệt trong ngày'],['Tạo PO','Mỗi NCC một PO/tháng'],['Đặt & nhận hàng','Số nhận là dữ liệu chuẩn'],['Đối soát NCC','Chu kỳ 15 ngày / tháng'],['Chỉnh PO & PR bổ sung','1 lần cuối chu kỳ'],['MIGO cuối kỳ','1 MIGO mỗi PO, = PO'],['ĐNTT theo PO','PO = MIGO = hóa đơn']];
const SUB=[['1.1','Thực đơn tuần','td'],['1.2','Lượng khách','khach'],['1.3','Định mức món','dm'],['1.4','Dự trù & ngân sách','tinh'],['1.5','Chốt dự trù','chot']];
const DOW=['CN','T2','T3','T4','T5','T6','T7'];

/* ---------- ngày tháng ---------- */
const P=s=>{const[y,m,d]=s.split('-').map(Number);return new Date(Date.UTC(y,m-1,d));};
const iso=d=>d.toISOString().slice(0,10);
const addD=(s,n)=>{const d=P(s);d.setUTCDate(d.getUTCDate()+n);return iso(d);};
const dm=s=>s.slice(8)+'/'+s.slice(5,7);
const monday=s=>{const w=P(s).getUTCDay();return addD(s,w===0?-6:1-w);};
function monthDays(t){const[y,m]=t.split('-').map(Number);const n=new Date(Date.UTC(y,m,0)).getUTCDate();return Array.from({length:n},(_,i)=>`${t}-${String(i+1).padStart(2,'0')}`);}
function weeksOf(t){const ds=monthDays(t);const ws=[];let w=monday(ds[0]);while(w<=ds[ds.length-1]){ws.push(w);w=addD(w,7);}return ws;}
const slug=s=>String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/đ/g,'d').replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'').slice(0,120)||'mon';
const tstr=ts=>{const d=new Date(+ts+7*3600e3),z=n=>String(n).padStart(2,'0');return `${z(d.getUTCDate())}/${z(d.getUTCMonth()+1)} ${z(d.getUTCHours())}:${z(d.getUTCMinutes())}`;};

/* ---------- kho dữ liệu: file Google Sheets PRTruoc_DuLieu trong thư mục dự án (G:\My Drive\30102026.QuyTrinhPRTruoc) ---------- */
const FILE={sheet:CFG.dataSheetId,seed:CFG.seedFolderId,xuat:CFG.xuatFolderId};
const COLS=['lenh','dexuat','agentchay','cauhinh','thucdon','dinhmuc','khach','dutru','trangthai','kho','pb','prdot','nhanhang','po','donhang','phieunhan','doisoat','prbs','migo','hoadon','dntt'];
let db=null,user=null,sample=null,dlcap=null,mcp=null,gmail=null,uid='',uname='',canWrite=true,SH=false;
const MEM=new Map(),MEML=new Set(),ROW=new Map(),CMD=[]; let cmdSeq=0, WQ=Promise.resolve();
const tstrFull=ts=>{const d=new Date(+ts+7*3600e3),z=n=>String(n).padStart(2,'0');return `${z(d.getUTCDate())}/${z(d.getUTCMonth()+1)}/${d.getUTCFullYear()} ${z(d.getUTCHours())}:${z(d.getUTCMinutes())}:${z(d.getUTCSeconds())}`;};
function renderCmd(){ const el=document.getElementById('cmd'); if(!el) return;
  const act=CMD.filter(c=>c.st!=='ok').slice(0,20), ok=CMD.filter(c=>c.st==='ok'), last=ok[0];
  const li=c=>`<li class="c">${c.st==='run'?'<span class="spin" aria-label="đang chạy"></span>':'<span class="pill bad">Lỗi</span>'}<b>${esc(c.label)}${c.err?`<div class="dirty" style="font-weight:500">${esc(c.err)}</div>`:''}</b><span class="num">${c.st==='run'?'…':((c.ms/1000).toFixed(1).replace('.',',')+' s')}</span></li>`;
  el.innerHTML=(act.map(li).join('')+(ok.length?`<li class="c"><span class="pill ok">✓</span><b style="font-weight:500">${ok.length} lượt đọc, ghi thành công<div class="meta">Gần nhất: ${esc(last.label)}</div></b><span class="num">${tstr(last.t0).slice(6)}</span></li>`:''))||'<li class="muted">Chưa đọc, ghi gì.</li>';
}
async function cmd(label,fn){ const c={label,t0:Date.now(),st:'run'}; CMD.unshift(c); if(CMD.length>80) CMD.pop(); renderCmd();
  try{ const r=await fn(); c.st='ok'; c.ms=Date.now()-c.t0; renderCmd(); return r; }catch(e){ c.st='err'; c.err=(e&&(e.message||e.code))||String(e); c.ms=Date.now()-c.t0; renderCmd(); throw e; } }
const call=(server,tool,input,label)=>cmd(label||`${server} · ${tool}`,()=>mcp.callTool(server,tool,input,{cache:false}));
const SJ=(p,o)=>cmd(o&&o.images?'Claude đọc ảnh biên bản':'Claude soạn dữ liệu',()=>sample.json(p,o));
const pay=r=>{ let p=r&&r.payload; if(!p&&r&&r.content){ try{ p=JSON.parse((r.content.find(c=>c.type==='text')||{}).text); }catch(e){} } return p||{}; };
const colL=n=>{let s='';n++;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};
const CH=45000;
function notify(){ MEML.forEach(f=>f()); }
function rowVals(id,data){ const json=data==null?'':JSON.stringify(data), parts=[]; for(let i=0;i<json.length;i+=CH) parts.push("'"+json.slice(i,i+CH)); return ["'"+id,data==null?'đã xóa':tstrFull(Date.now()),...parts]; }
/* Connector Google Sheets không có append_values: đọc cột A để tìm dòng trống kế tiếp rồi ghi bằng update_values. */
async function appendRows(col,rows,label){ if(!rows.length) return 0; const w=Math.max(...rows.map(r=>r.length)); rows.forEach(r=>{while(r.length<w)r.push('');});
  const a=(pay(await call('Google Sheets','get_values',{spreadsheetId:FILE.sheet,range:`${col}!A1:A`},`Tìm dòng trống ${col}`)).values)||[]; let n=a.length+1; if(n<2) n=2;
  await call('Google Sheets','update_values',{spreadsheetId:FILE.sheet,range:`${col}!A${n}:${colL(w-1)}${n+rows.length-1}`,values:rows},label||`Thêm ${rows.length} dòng vào ${col}`); return n; }
async function shWrite(path,data){
  const i=path.indexOf('/'), col=path.slice(0,i), id=path.slice(i+1), vals=rowVals(id,data), prev=ROW.get(path), real=vals.length-2;
  if(prev){ while(vals.length-2<prev.n) vals.push(''); await call('Google Sheets','update_values',{spreadsheetId:FILE.sheet,range:`${col}!A${prev.row}:${colL(vals.length-1)}${prev.row}`,values:[vals]},(data==null?'Xóa ':'Ghi ')+`${col} · ${id}`); ROW.set(path,{row:prev.row,n:real}); }
  else if(data!=null){ const n=await appendRows(col,[vals],`Thêm ${col} · ${id}`); if(n) ROW.set(path,{row:n,n:real}); }
}
let PENDING=0;
function enqueue(fn){ PENDING++; WQ=WQ.then(fn).finally(()=>{PENDING--;}).catch(e=>{ toast('Không ghi được vào file dữ liệu: '+((e&&(e.message||e.code))||e)+'. Bấm “Tải lại dữ liệu”.'); }); return WQ; }
const store={
  async set(path,data){ const d=JSON.parse(JSON.stringify(data)); MEM.set(path,d); notify(); if(SH) return enqueue(()=>shWrite(path,d)); },
  async del(path){ MEM.delete(path); notify(); if(SH) return enqueue(()=>shWrite(path,null)); },
  async add(col,data){
    if(col==='nhatky'){ const id='m'+Date.now()+Math.random().toString(36).slice(2,5), d={...data,nguoi:uname}; MEM.set('nhatky/'+id,d); notify();
      if(SH) return enqueue(()=>appendRows('nhatky',[[d.ts,"'"+tstrFull(d.ts),"'"+(d.thang||''),"'"+(d.buoc||''),d.viec,uname||'']],'Ghi nhật ký')); return; }
    return this.set(col+'/a'+Date.now().toString(36)+Math.random().toString(36).slice(2,5),data); },
  watchDoc(path,cb){ const f=()=>cb(MEM.get(path)||null); MEML.add(f); f(); return ()=>MEML.delete(f); },
  watchCol(col,cb,q){ const f=()=>{let a=[...MEM.entries()].filter(([k])=>k.startsWith(col+'/')&&k.split('/').length===2).map(([k,v])=>({id:k.slice(col.length+1),...v}));
      if(q&&q.order) a.sort((x,y)=>(x[q.order]>y[q.order]?1:-1)*(q.dir==='desc'?-1:1)); if(q&&q.limit) a=a.slice(0,q.limit); cb(a);};
    MEML.add(f); f(); return ()=>MEML.delete(f); }
};
window.__store=store;
function applyConfig(){ const g=MEM.get('cauhinh/gia'), bs=MEM.get('cauhinh/giabs'); if((g&&g.items)||(bs&&bs.items)){ const base=(g&&g.items)||RAW.items, have=new Set(base.map(a=>a[0])); setItems([...base,...((bs&&bs.items)||[]).filter(a=>!have.has(a[0]))],g&&g.v); }
  const n=MEM.get('cauhinh/nccinfo'); if(n) Object.assign(NCCINFO,n); const l=MEM.get('cauhinh/lifnr'); if(l) Object.assign(LIFNR,l); const k=MEM.get('cauhinh/khomap'); if(k) Object.assign(KHOMAP0,k); }
async function loadAll(){
  $('#conn').textContent='Đang đọc file dữ liệu…'; const bad=[];
  const ALL=[...COLS,'nhatky'], RES=await Promise.all(ALL.map(col=>call('Google Sheets','get_values',{spreadsheetId:FILE.sheet,range:`${col}!A1:Z5000`},`Đọc sheet ${col}`)));
  for(const [ci,col] of ALL.entries()){
    const V=(pay(RES[ci]).values)||[];
    [...MEM.keys()].filter(k=>k.startsWith(col+'/')).forEach(k=>MEM.delete(k)); [...ROW.keys()].filter(k=>k.startsWith(col+'/')).forEach(k=>ROW.delete(k));
    if(col==='nhatky'){ V.slice(1).forEach((r,i)=>{ if(r[0]) MEM.set('nhatky/r'+(i+2),{ts:+r[0],thang:r[2]||'',buoc:r[3]||'',viec:r[4]||'',nguoi:r[5]||''}); }); continue; }
    const st=V[0]&&V[0][0]==='id'?1:0;
    if(!V.length) await call('Google Sheets','update_values',{spreadsheetId:FILE.sheet,range:`${col}!A1:C1`,values:[['id','cap_nhat','du_lieu']]},`Tạo tiêu đề sheet ${col}`);
    V.slice(st).forEach((r,i)=>{ const rn=i+st+1, id=r[0]; if(!id) return; const parts=r.slice(2).filter(x=>x!==''), path=col+'/'+id; ROW.set(path,{row:rn,n:parts.length}); if(!parts.length) return;
      try{ MEM.set(path,JSON.parse(parts.join(''))); }catch(e){ bad.push(path); } });
  }
  applyConfig(); notify();
  const empty=!MEM.has('cauhinh/gia')&&![...MEM.keys()].some(k=>k.startsWith('thucdon/'));
  $('#seedNote').hidden=!empty; $('#conn').textContent=`File dữ liệu · ${[...MEM.keys()].length} bản ghi · đọc ${tstr(Date.now())}`; $('#conn').className='pill ok';
  if(bad.length) toast('Có '+bad.length+' bản ghi hỏng JSON, bỏ qua: '+bad.slice(0,3).join(', '));
}
async function seedAll(){
  const r=pay(await call('Google Drive','download_file_content',{fileId:FILE.seed},'Tải DuLieuApp/khoitao.json'));
  const txt=new TextDecoder().decode(Uint8Array.from(atob(r.content||''),c=>c.charCodeAt(0))); const S0=JSON.parse(txt);
  for(const [col,docs] of Object.entries(S0)){ if(col==='nhatky') continue; const rows=Object.entries(docs).map(([id,d])=>rowVals(id,d)); const w=Math.max(...rows.map(x=>x.length)); rows.forEach(x=>{while(x.length<w)x.push('');});
    await appendRows(col,rows,`Nạp ${rows.length} bản ghi vào sheet ${col}`); }
  const L=Object.values(S0.nhatky||{}).sort((a,b)=>a.ts-b.ts).map(d=>[d.ts,"'"+tstrFull(d.ts),"'"+(d.thang||''),"'"+(d.buoc||''),d.viec,'Claude (khởi tạo)']);
  if(L.length) await appendRows('nhatky',L,`Nạp ${L.length} dòng nhật ký`);
  await loadAll();
}
async function saveOut(filename,blob){
  if(mcp){ const b64=await new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(String(fr.result).split(',')[1]);fr.onerror=rej;fr.readAsDataURL(blob);});
    await call('Google Drive','create_file',{title:filename,parentId:FILE.xuat,base64Content:b64,contentMimeType:blob.type||'application/octet-stream',disableConversionToGoogleType:true},`Lưu ${filename} vào thư mục XuatSAP`);
    toast(`Đã lưu ${filename} vào thư mục XuatSAP của dự án.`); return; }
  if(dlcap) return dlcap.save({filename,data:blob}); throw new Error('Không có nơi lưu file');
}

