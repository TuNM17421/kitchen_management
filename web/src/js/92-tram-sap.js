/* ---------- TRẠM SAP (máy công ty chạy TramSAP.ps1, nhận lệnh qua sheet lenh) ---------- */
S.lenh={}; S.tram=null; const APPLY=new Set();
const SAPKEY={PR:'khongLuu',PO:'khongLuu',POSUA:'khongLuuSua',MIGO:'khongLuuMigo'}, SAPDEF={khongLuu:false,khongLuuSua:true,khongLuuMigo:true};
const SAPTEN={PR:'CreatePR (ME51N)',PO:'CreatePO (ME21N)',POSUA:'sửa PO (ME22N)',MIGO:'nhập kho MIGO',DNTT:'tạo file ĐNTT (TaoDNTT)'};
function sapLuuThat(loai){ const c=MEM.get('cauhinh/sap')||{}, k=SAPKEY[loai||'PR']; return !(k in c?c[k]:SAPDEF[k]); }
const TTL={cho:['Chờ trạm nhận','run'],dang:['Trạm đang chạy','run'],xong:['Xong','ok'],loi:['Lỗi / chưa đủ','bad']};
const lenhBusy=id=>!!(id&&S.lenh&&S.lenh[id]&&['cho','dang'].includes(S.lenh[id].tt));
function tramOnline(){ return !!(S.tram&&S.tram.ts&&(Date.now()-S.tram.ts)<180000); }
function renderTram(){ const el=$('#tramPill'); if(!el) return;
  if(!SH){ el.textContent='Trạm SAP: chưa kết nối'; el.className='pill stt'; return; }
  if(!S.tram||!S.tram.ts){ el.textContent='Trạm SAP: chưa cài'; el.className='pill stt bad'; return; }
  const ago=Math.max(0,Math.round((Date.now()-S.tram.ts)/1000));
  el.textContent=`Trạm SAP ${tramOnline()?'trực tuyến':'ngoại tuyến'} · ${S.tram.may||''} · ${ago<120?ago+' giây':Math.round(ago/60)+' phút'} trước`; el.className='pill stt '+(tramOnline()?'ok':'bad'); }
function lenhPill(id){ if(!lenhBusy(id)) return ''; const L=S.lenh[id], t=TTL[L.tt]; return ` <span class="pill ${t[1]}">${t[0]}${L.tiendo?' · '+esc(L.tiendo):''}</span>`; }
function tramBox(loai){ const lt=sapLuuThat(loai);
  if(loai==='DNTT') return `<div class="box"><div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center"><b>Trạm SAP</b><span class="pill ${tramOnline()?'ok':'bad'}">${tramOnline()?'trực tuyến':'ngoại tuyến'}</span></div><div class="meta" style="margin-top:6px">Bấm <b>Tạo file ĐNTT</b>: máy công ty chạy TaoDNTT (không đụng SAP), lưu file ĐNTT + PDF ở C:\\TramSAP\\DNTT_xuat và chép lên thư mục XuatSAP của dự án; link hiện ngay trong bảng.${tramOnline()?'':' Trạm đang ngoại tuyến: lệnh chờ tới khi trạm chạy.'}</div></div>`;
  return `<div class="box"><div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center"><b>Trạm SAP</b><span class="meta">· ${SAPTEN[loai]}</span><span class="pill ${tramOnline()?'ok':'bad'}">${tramOnline()?'trực tuyến':'ngoại tuyến'}</span>
    <span class="pill ${lt?'bad':''}">${lt?'LƯU THẬT':'Không lưu (chạy thử)'}</span>${canWrite?`<button class="btn sm" data-tram="mode" data-loai="${loai}" style="width:auto">${lt?'Chuyển sang không lưu':'Chuyển sang lưu thật'}</button>`:''}</div>
    <details class="tram-more"><summary>Chế độ này làm gì?</summary><div class="meta">${lt?(loai==='MIGO'?'Lưu thật: macro tự Post khi đã tích đủ mọi dòng của chứng từ; thiếu dòng thì <b>không Post</b> và báo về đây.':'Lưu thật: macro tự bấm Save khi số lượng và giá trên SAP khớp tuyệt đối với dữ liệu gửi đi; nếu lệch thì <b>không lưu</b> và báo về đây.'):'Không lưu (chạy thử): macro điền màn hình'+(loai==='MIGO'?', bấm Check':'')+' rồi dừng, không có số chứng từ. Chạy thử thành công thì chuyển sang lưu thật.'} Trạm chạy ${SAPTEN[loai]} trên cửa sổ SAP đầu tiên của máy công ty.${tramOnline()?'':' Trạm đang ngoại tuyến: lệnh sẽ chờ tới khi máy công ty chạy ChayTramSAP.bat.'}</div></details></div>`; }
function prSapBar(sel){
  const ok=sel.filter(b=>!lenhBusy(b.lenh)), nccs=[...new Set(sel.flatMap(b=>b.lines.map(l=>l.ncc)))].sort();
  return canWrite?`<div class="toolbar"><button class="btn primary" data-tram="pr" ${ok.length?'':'disabled'}>Tạo PR trên SAP (${ok.length} đợt đã chọn)</button></div>`:''; }
function tachBox(sel){ const nccs=[...new Set(sel.flatMap(b=>b.lines.map(l=>l.ncc)))].sort();
  return canWrite&&nccs.length?`<details class="more"><summary>Thao tác khác</summary><div class="toolbar"><label class="meta" for="tach-ncc">Tách 1 NCC thành đợt riêng</label><select id="tach-ncc"><option value="">Chọn NCC…</option>${nccs.map(n=>`<option>${esc(n)}</option>`).join('')}</select><button class="btn sm" data-tram="tach">Tách</button></div></details>`:''; }
async function tachNCC(ncc){
  const bs=batchesOf(S.thang).filter(b=>(b.tt==='nhap'||b.tt==='xuat')&&!lenhBusy(b.lenh)&&ED.sel.has(b.id)&&b.lines.some(l=>l.ncc===ncc)); let n=0;
  for(const b of bs){ const mv=b.lines.filter(l=>l.ncc===ncc), keep=b.lines.filter(l=>l.ncc!==ncc); if(!keep.length) continue;
    const tg=L=>Math.round(L.reduce((a,l)=>a+l.sl*l.g,0)), sfx=slug(ncc).replace(/-/g,'').slice(0,8).toUpperCase(), nid=b.id+'-'+sfx;
    const d={...b,nhom:b.nhom+'-'+sfx,lines:mv,tong:tg(mv),tt:'nhap',soPR:'',lenh:'',luc:Date.now(),ai:uid,ghichu:'Tách từ '+b.nhom,hist:[{ts:Date.now(),tt:'nhap',ai:uid,note:'Tách '+ncc+' từ '+b.nhom}]}; delete d.id;
    await store.set('prdot/'+nid,d); await setBatch(b.id,{lines:keep,tong:tg(keep)},'Tách '+ncc+' sang '+d.nhom); ED.sel.add(nid); ED.sel.delete(b.id); n++; }
  if(n){ await log('2.2',`Tách ${ncc} thành ${n} đợt riêng`); toast(`Đã tách ${ncc} thành ${n} đợt riêng (đã chọn sẵn các đợt mới).`); render(); }
  else toast('Không có đợt nào cần tách (NCC này đã là đợt riêng, hoặc chưa chọn đợt).'); }
// Hủy lệnh còn chờ (trạm chưa nhận) khi người dùng đã tự làm bằng tay: tránh trạm chạy lại sau này.
async function boLenhCho(lid,why){ const L=lid&&S.lenh&&S.lenh[lid]; if(!L||L.tt!=='cho') return false;
  const x={...L,tt:'loi',log:why,ketqua:null,apDung:true,apLuc:Date.now()}; delete x.id; await store.set('lenh/'+lid,x); return true; }
async function guiLenh(loai,rows,ref,moTa,lifnr){
  if(!SH){ toast('Chưa kết nối file dữ liệu, không gửi được lệnh.'); return null; }
  const id='L'+Date.now().toString(36), lt=sapLuuThat(loai);
  await store.set('lenh/'+id,{loai,thang:S.thang,luuThat:lt,taoLuc:Date.now(),nguoi:uname,moTa,rows,ref,lifnr:lifnr||{},tt:'cho'});
  await log({PR:'2.2',PO:'3.2',POSUA:'3.4',MIGO:'7.1',DNTT:'8.2'}[loai],`Gửi trạm SAP: ${SAPTEN[loai]} ${moTa} (${lt?'lưu thật':'không lưu'})`);
  toast(tramOnline()?'Đã gửi lệnh. Trạm SAP sẽ nhận trong khoảng 15 giây.':'Đã gửi lệnh. Trạm SAP đang ngoại tuyến: lệnh chờ tới khi trạm chạy.');
  setTimeout(pollTram,4000); return id; }
async function lenhPR(){
  const bs=batchesOf(S.thang).filter(b=>(b.tt==='nhap'||b.tt==='xuat')&&ED.sel.has(b.id)&&!lenhBusy(b.lenh));
  if(!bs.length){ toast('Chọn ít nhất 1 đợt PR chưa gửi.'); return; }
  const rows=[]; let i=0;
  bs.forEach(b=>b.lines.forEach((l,li)=>rows.push({c:[++i,b.nhom,l.ncc,l.ma,l.ten,l.sl,l.g,b.dt,b.nhom,'',`PR trước ${mmY(b.thang)} · ${b.loai}`],b:b.id,li})));
  const id=await guiLenh('PR',rows,{batches:bs.map(b=>b.id)},bs.map(b=>b.nhom).join(', ')+` (${rows.length} dòng)`); if(!id) return;
  for(const b of bs) await setBatch(b.id,{lenh:id,tt:b.tt==='nhap'?'xuat':b.tt},'Gửi trạm SAP tạo PR'); }
function prItemOf(soPR,ma,ncc){ const b=Object.values(S.prdot||{}).find(x=>x.soPR===soPR); const l=b&&b.lines.find(x=>x.ma===ma&&x.ncc===ncc); return l&&l.item?String(l.item):''; }
async function lenhPO(id){
  const p=S.po[id]; if(!p) return; const lif=p.lifnr||LIFNR[String(p.ncc).toUpperCase()]||'';
  if(!lif){ toast('NCC '+p.ncc+' chưa có mã SAP (LIFNR).'); return; }
  const prs=[...new Set(p.lines.map(l=>l.soPR).filter(Boolean))], nd=prs.filter(x=>!Object.values(S.prdot||{}).some(b=>b.soPR===x&&b.tt==='duyet'));
  if(!prs.length||nd.length||p.lines.some(l=>!l.soPR)){ toast(nd.length?'PR chưa duyệt: '+nd.join(', ')+'. Bấm “Đã duyệt” ở 2.3 trước.':'Có dòng PO chưa có số PR.'); return; }
  const nhomOf=x=>{const b=Object.values(S.prdot||{}).find(y=>y.soPR===x);return b?b.nhom:'';};
  const rows=p.lines.map((l,i)=>({c:[i+1,nhomOf(l.soPR),p.ncc,l.ma,l.ten,l.sl,l.g,l.dt,nhomOf(l.soPR),'',`PO ${mmY(p.thang)} · ${p.ncc}`],soPR:l.soPR,item:prItemOf(l.soPR,l.ma,p.ncc)}));
  const lid=await guiLenh('PO',rows,{po:id},`${p.ncc} (${rows.length} dòng)`,{[p.ncc]:lif}); if(!lid) return;
  await setPO(id,{lenh:lid},'Gửi trạm SAP tạo PO'); }
async function applyLenh(){
  for(const L of Object.values(S.lenh||{})){ if(!['xong','loi'].includes(L.tt)||L.apDung||APPLY.has(L.id)) continue; APPLY.add(L.id);
    try{ const K=(L.ketqua&&L.ketqua.rows)||[], note=[];
      if(L.loai==='PR'){ for(const bid of (L.ref&&L.ref.batches)||[]){ const b=S.prdot[bid]; if(!b) continue;
          const idx=(L.rows||[]).map((r,i)=>r.b===bid?i:-1).filter(i=>i>=0), kq=idx.map(i=>K.find(k=>+k.i===i)||{}), so=[...new Set(kq.map(k=>k.soPR).filter(Boolean))];
          if(so.length===1&&kq.length&&kq.every(k=>k.soPR)){
            const lines=b.lines.map((l,li)=>{const j=L.rows.findIndex(r=>r.b===bid&&r.li===li); const k=K.find(x=>+x.i===j)||{}; return {...l,item:k.item||l.item||''};});
            await setBatch(bid,{soPR:so[0],tt:'sap',sapLuc:Date.now(),lenh:'',lines},'Trạm SAP tạo PR '+so[0]); note.push(`${b.nhom} → PR ${so[0]}`); }
          else { await setBatch(bid,{lenh:''},'Trạm SAP chưa tạo được PR: '+((kq.find(k=>k.tt)||{}).tt||'xem nhật ký lệnh')); note.push(`${b.nhom}: ${L.luuThat?'chưa tạo được PR':'chạy thử xong, chưa lưu'}`); } } }
      else if(L.loai==='POSUA'||L.loai==='MIGO'){ await applyLenh2(L,K,note); }
      else if(L.loai==='DNTT'){ await applyLenh3(L,K,note); }
      else if(L.loai==='PO'){ const pid=L.ref&&L.ref.po, p=pid&&S.po[pid];
        if(p){ const P=((L.ketqua&&L.ketqua.po)||[]).find(x=>String(x.ncc).toUpperCase()===String(p.ncc).toUpperCase())||{};
          const so=P.soPO||(K.find(k=>k.soPO)||{}).soPO||'', ts=P.tongSAP!==undefined&&P.tongSAP!==''&&!isNaN(+String(P.tongSAP).replace(',','.'))?+String(P.tongSAP).replace(',','.'):null;
          if(so){ await setPO(pid,{soPO:so,tongSAP:ts,tt:'sap',sapLuc:Date.now(),lenh:''},'Trạm SAP tạo PO '+so); note.push(`${p.ncc} → PO ${so}`); }
          else { await setPO(pid,{lenh:'',tongSAP:ts},'Trạm SAP chưa tạo được PO: '+(P.tt||'xem nhật ký lệnh')); note.push(`${p.ncc}: ${L.luuThat?'chưa tạo được PO':'chạy thử xong, chưa lưu'}`); } } }
      const x={...L,apDung:true,apLuc:Date.now()}; delete x.id; await store.set('lenh/'+L.id,x);
      await log({PR:'2.3',PO:'3.2',POSUA:'3.4',MIGO:'7.1',DNTT:'8.2'}[L.loai]||'2.2',`Kết quả trạm SAP (${L.luuThat?'lưu thật':'không lưu'}): ${note.join('; ')||L.tt}`);
      toast(`Trạm SAP: ${note.join('; ')||L.tt}`);
    }catch(e){ APPLY.delete(L.id); }
  } }
function renderLenh(){ const el=$('#lenhList'); if(!el) return;
  const A=Object.values(S.lenh||{}).sort((a,b)=>(b.taoLuc||0)-(a.taoLuc||0)).slice(0,12);
  const nCho=Object.values(S.lenh||{}).filter(L=>L.tt==='cho').length;
  el.innerHTML=(nCho>1&&canWrite?`<li><button class="btn sm warn" data-tram="huy-all">Hủy tất cả ${nCho} lệnh đang chờ</button></li>`:'')+(A.length?A.map(L=>{const t=TTL[L.tt]||[L.tt,''];
    return `<li><div class="h"><span class="pill ${t[1]}">${t[0]}</span><span class="pill">${esc(L.loai)}</span><span class="pill ${L.luuThat?'bad':''}">${L.luuThat?'lưu thật':'không lưu'}</span><time>${tstr(L.taoLuc)}</time></div>
    <div>${esc(L.moTa||'')}${L.tiendo?` · <i>${esc(L.tiendo)}</i>`:''}</div>
    ${L.log?`<details><summary class="meta">Nhật ký macro</summary><pre style="white-space:pre-wrap;font-size:11px;max-height:200px;overflow:auto;margin:4px 0">${esc(L.log)}</pre></details>`:''}
    ${L.tt==='cho'&&canWrite?`<button class="btn sm warn" data-tram="huy" data-id="${L.id}">Hủy lệnh</button>`:''}</li>`;}).join('')
    :'<li class="muted">Chưa gửi lệnh nào. Bấm “Tạo PR trên SAP” ở 2.2 hoặc “Tạo trên SAP” ở 3.2.</li>'); }
async function reloadCol(col){
  const V=(pay(await mcp.callTool('Google Sheets','get_values',{spreadsheetId:FILE.sheet,range:`${col}!A1:Z2000`},{cache:false})).values)||[];
  if(PENDING>0) return; const st=V[0]&&V[0][0]==='id'?1:0;
  [...MEM.keys()].filter(k=>k.startsWith(col+'/')).forEach(k=>MEM.delete(k)); [...ROW.keys()].filter(k=>k.startsWith(col+'/')).forEach(k=>ROW.delete(k));
  V.slice(st).forEach((r,i)=>{ const rn=i+st+1,id=r[0]; if(!id) return; const parts=r.slice(2).filter(x=>x!==''),path=col+'/'+id; ROW.set(path,{row:rn,n:parts.length}); if(!parts.length) return; try{ MEM.set(path,JSON.parse(parts.join(''))); }catch(e){} });
  notify(); }
let pollBusy=false, pollN=0;
async function pollTram(){ if(!SH||!mcp||pollBusy) return; pollBusy=true;
  try{ const r=pay(await mcp.callTool('Google Sheets','get_values',{spreadsheetId:FILE.sheet,range:'tram!A1:E2'},{cache:false})).values||[];
    if(r[1]) S.tram={may:r[1][0],ts:+r[1][1]||0,ping:r[1][2],ver:r[1][3]};
    const act=Object.values(S.lenh||{}).some(L=>lenhBusy(L.id)||(['xong','loi'].includes(L.tt)&&!L.apDung));
    if(act&&PENDING===0) await reloadCol('lenh'); pollN++;
  }catch(e){} finally{ pollBusy=false; renderTram(); if(S.step===2||S.step===3) render(); } }
setInterval(()=>{ const act=Object.values(S.lenh||{}).some(L=>lenhBusy(L.id)); if(act||pollN%4===0) pollTram(); else pollN++; },15000);
var TRAM_RUN=new Set(); // lệnh đang gửi: chặn bấm lại khi trang vẽ lại nút (tránh gửi trùng)
function tramRunMark(){ const v=$('#view'); if(!v||!TRAM_RUN||!TRAM_RUN.size) return; v.querySelectorAll('[data-tram]').forEach(b=>{ if(TRAM_RUN.has(b.dataset.tram+'|'+(b.dataset.id||''))){ b.disabled=true; b.setAttribute('aria-busy','true'); } }); }
document.addEventListener('click',async e=>{ const t=e.target.closest('[data-tram]'); if(!t) return; const a=t.dataset.tram;
  const rk=a+'|'+(t.dataset.id||''), guard=!['mode','tach'].includes(a); if(guard){ if(TRAM_RUN.has(rk)) return; TRAM_RUN.add(rk); }
  try{
    if(a==='mode'){ const lo=t.dataset.loai||'PR', lt=sapLuuThat(lo), c={...(MEM.get('cauhinh/sap')||{})}; c[SAPKEY[lo]]=lt; c.luc=Date.now(); c.nguoi=uname; await store.set('cauhinh/sap',c);
      await log('2.2',`Chế độ trạm SAP – ${SAPTEN[lo]}${lo==='PR'||lo==='PO'?' và '+SAPTEN[lo==='PR'?'PO':'PR']:''}: ${lt?'không lưu (chạy thử)':'LƯU THẬT'}`); render(); }
    else if(a==='posua'){ t.disabled=true; await lenhPOSua(t.dataset.id); }
    else if(a==='dntt'){ t.disabled=true; await lenhDNTT([...(ED.dnsel||[])]); }
    else if(a==='migo'){ t.disabled=true; await lenhMIGO(mgSelIds()); }
    else if(a==='pr'){ t.disabled=true; await lenhPR(); }
    else if(a==='po'){ t.disabled=true; await lenhPO(t.dataset.id); }
    else if(a==='po-many'){ t.disabled=true; const bo=ED.poBo||new Set(), ids=poOf(S.thang).filter(p=>p.tt==='cho'&&!lenhBusy(p.lenh)&&!bo.has(p.id)).map(p=>p.id); for(const id of ids) await lenhPO(id); ED.poBo=null; render(); }
    else if(a==='tach'){ const n=$('#tach-ncc').value; if(!n){toast('Chọn NCC cần tách.');return;} await tachNCC(n); }
    else if(a==='huy'){ await WQ; await reloadCol('lenh'); const L=S.lenh[t.dataset.id]; if(!L||L.tt!=='cho'){toast('Trạm đã nhận lệnh, không hủy được.');return;}
      await boLenhCho(L.id,'Người dùng hủy trước khi trạm nhận.'); await log('2.2','Hủy lệnh trạm SAP: '+(L.moTa||L.id)); }
    else if(a==='huy-all'){ await WQ; await reloadCol('lenh'); let n=0; for(const L of Object.values(S.lenh||{})) if(L.tt==='cho'&&await boLenhCho(L.id,'Người dùng hủy tất cả lệnh đang chờ.')) n++;
      if(n) await log('2.2',`Hủy ${n} lệnh trạm SAP đang chờ`); toast(n?`Đã hủy ${n} lệnh đang chờ.`:'Không còn lệnh nào đang chờ.'); }
  }catch(err){ toast('Lỗi: '+((err&&(err.message||err.code))||err)); }
  finally{ if(guard){ TRAM_RUN.delete(rk); render(); } }
});

