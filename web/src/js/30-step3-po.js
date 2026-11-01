/* ================= BƯỚC 3 · PO ================= */
const SUB3=[['3.1','PO cần tạo','pocan'],['3.2','Tạo PO trên SAP','posap'],['3.3','PO so với nhận hàng','podc'],['3.4','Lịch sử PO','pols']];
const TTPO={cho:['Chờ tạo trên SAP','run'],sap:['Đã có số PO','ok'],huy:['Đã hủy','bad']};
const poAll=()=>Object.values(S.po||{});
const poOf=t=>poAll().filter(p=>p.thang===t&&p.tt!=='huy').sort((a,b)=>a.ncc.localeCompare(b.ncc,'vi')||a.luc-b.luc);
const poPill=p=>{const t=TTPO[p.tt]||['?',''];return `<span class="pill ${t[1]}">${t[0]}</span>${p.sua?' <span class="pill bad">Đã sửa SL – cần sửa ME22N</span>':''}`;};
function poPending(){
  const used=new Set(poAll().filter(p=>p.tt!=='huy').flatMap(p=>p.src||[]));
  const g=new Map();
  batchesOf(S.thang).filter(b=>b.tt==='duyet'&&b.loai==='Chính').forEach(b=>b.lines.forEach(l=>{const key=b.id+'|'+l.ncc; if(used.has(key)) return;
    const o=g.get(l.ncc)||{ncc:l.ncc,src:new Set(),lines:[],loai:new Set(),pr:new Set()}; o.src.add(key); o.loai.add(b.loai); o.pr.add(b.soPR);
    o.lines.push({ma:l.ma,ten:l.ten,dvt:l.dvt,sl:l.sl,g:l.g,soPR:b.soPR,dt:b.dt}); g.set(l.ncc,o);}));
  return [...g.values()].map(o=>({...o,src:[...o.src],loai:[...o.loai],pr:[...o.pr],tong:Math.round(o.lines.reduce((a,l)=>a+l.sl*l.g,0))})).sort((a,b)=>b.tong-a.tong);
}
async function setPO(id,patch,note){ const p=S.po[id];
  if(patch.soPO&&!p.soPO&&!patch.sapLines){ const L=(patch.lines||p.lines).map((l,i)=>({...l,ebelp:+l.ebelp||(i+1)*10})); patch={...patch,lines:L,sapLines:L.map(l=>({ebelp:l.ebelp,ma:l.ma,sl:+l.sl,g:+l.g}))}; }
  const d={...p,...patch,hist:[...(p.hist||[]),{ts:Date.now(),tt:patch.tt||p.tt,ai:uid,note:note||''}]}; delete d.id; await store.set('po/'+id,d); }
/* ---- 3.1 PO cần tạo ---- */
function vPOcan(){
  const pend=poPending(), made=poOf(S.thang), noL=pend.filter(o=>!hasLifnr(o.ncc));
  return head2('3.1 PO cần tạo','Gom các dòng của PR đã duyệt theo nhà cung cấp: mỗi NCC một PO cho cả tháng (gồm dòng của PR Nhân viên và PR Bệnh nhân). PO là dữ liệu nội bộ để đặt hàng hằng ngày và lập ĐNTT, không gửi NCC. PO theo tháng nên không vượt kỳ giá. Sau khi lập danh sách, sang 3.2 bấm <b>Tạo trên SAP</b> cho từng NCC: trạm SAP tự chạy CreatePO và ghi số PO về.')+
  `<div class="kpis">${pend.length?`<div class="kv big"><b>Chờ lập PO</b><div class="num">${pend.length} NCC</div><small>${TR(pend.reduce((a,o)=>a+o.tong,0))}</small></div>`:''}
   <div class="kv"><b>PO đã lập tháng ${mmY(S.thang)}</b><div class="num">${made.length}</div><small>${made.filter(p=>p.soPO).length} đã có số PO</small></div>
   <div class="kv"><b>PR đã duyệt</b><div class="num">${batchesOf(S.thang).filter(b=>b.tt==='duyet').length} đợt</div><small>${batchesOf(S.thang).filter(b=>b.tt==='sap'||b.tt==='xuat'||b.tt==='nhap').length} đợt chưa duyệt (chưa lên PO)</small></div></div>
  ${noL.length?`<div class="warn">Chưa có mã NCC SAP: <b>${noL.map(o=>esc(o.ncc)).join(', ')}</b>. Thêm vào sheet Mapping (cột E:F) trước khi chạy CreatePO.</div>`:''}
  ${pend.length?`<div class="toolbar">${canWrite?`<button class="btn primary" data-act3="po-make">Lập danh sách ${pend.length} PO</button>`:''}<span class="meta">Tạo bản ghi PO “chờ tạo trên SAP” để theo dõi số PO, email và nhận hàng.</span></div>
  <div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>Nhà cung cấp</th><th>Mã NCC SAP</th><th>PR nguồn</th><th>Loại</th><th class="r">Dòng PO</th><th class="r">Tổng tiền</th></tr></thead><tbody>
  ${pend.map(o=>`<tr><td><b>${esc(o.ncc)}</b></td><td class="num">${hasLifnr(o.ncc)?LIFNR[o.ncc.toUpperCase()]:'<span class="pill bad">thiếu</span>'}</td><td class="num" title="${o.pr.map(esc).join(', ')}">${o.pr.length} PR</td><td>${o.loai.join(', ')}</td><td class="r num">${o.lines.length}</td><td class="r num">${VND(o.tong)}</td></tr>`).join('')}
  </tbody><tfoot><tr><td colspan="5">Tổng</td><td class="r num">${VND(pend.reduce((a,o)=>a+o.tong,0))}</td></tr></tfoot></table></div>`
  :`<div class="empty"><b>Không còn dòng PR nào chờ lên PO</b>${made.length?'Tất cả PR đã duyệt đã có trong danh sách PO.':'Duyệt PR ở bước 2.3 trước.'}<div class="toolbar" style="justify-content:center;margin-top:12px">${made.length?'<button class="btn sm primary" data-act="go-next" data-tab="posap">Mở 3.2 để tạo trên SAP</button>':'<button class="btn sm" data-act="go-next" data-step="2">Mở bước 2</button>'}</div></div>`}`;
}
/* ---- 3.2 Ghi số PO ---- */
function vPOsap(){
  const L=poOf(S.thang), can=p=>p.tt==='cho'&&canWrite&&!lenhBusy(p.lenh)&&!!(p.lifnr||LIFNR[String(p.ncc).toUpperCase()]);
  const elig=L.filter(can), bo=ED.poBo||(ED.poBo=new Set());
  const sel=elig.filter(p=>!bo.has(p.id));
  return head2('3.2 Tạo PO trên SAP','Chọn các PO rồi bấm <b>Tạo trên SAP</b>: trạm SAP chạy CreatePO cho từng NCC (chỉ khi mọi PR nguồn đã duyệt), tự ghi số PO và tổng tiền SAP đọc lại. Chênh lệch khác 0 nghĩa là giá hoặc số lượng trên PO khác PR. Vẫn ghi tay được nếu tạo PO ngoài trạm.')+
  (L.length?`${tramBox('PO')}
  ${canWrite&&elig.length?`<div class="toolbar"><button class="btn primary" data-tram="po-many" ${sel.length?'':'disabled'}>Tạo ${sel.length} PO đã chọn trên SAP</button><span class="meta">${L.filter(p=>p.soPO).length}/${L.length} PO đã có số</span></div>`:`<p class="meta">${L.filter(p=>p.soPO).length}/${L.length} PO đã có số.</p>`}
  <div class="tbl-wrap" style="max-height:none"><table class="t po-tbl"><colgroup><col class="c-chk"><col><col class="c-tien"><col class="c-so"><col class="c-tt"><col class="c-acts"></colgroup><thead><tr><th>${elig.length?`<input type="checkbox" id="posel-all" aria-label="Chọn tất cả PO chờ tạo" ${sel.length===elig.length?'checked':''}>`:'<span class="sr">Chọn</span>'}</th><th>Nhà cung cấp</th><th class="r">Tổng theo PR</th><th>Số PO (45… / 42…) · tổng trên SAP</th><th>Trạng thái</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>
  ${L.map(p=>{const ed=p.tt==='cho'||p.tt==='sap', d=p.tongSAP!=null&&p.tongSAP!==''?(+p.tongSAP)-p.tong:null;
    return `<tr><td class="chk">${can(p)?`<input type="checkbox" data-posel="${p.id}" ${bo.has(p.id)?'':'checked'} aria-label="Chọn PO ${esc(p.ncc)}">`:''}</td>
    <td class="ncc"><b>${esc(p.ncc)}</b>${p.loai==='Bổ sung'?' <span class="pill run">bổ sung</span>':''}<small>Mã NCC ${esc(p.lifnr||'–')} · ${p.lines.length} dòng · <span title="${(p.pr||[]).map(esc).join(', ')}">${(p.pr||[]).length} PR</span></small></td>
    <td class="r num tien">${VND(p.tong)}</td>
    <td class="so">${ed&&canWrite?`<div class="so-in"><input type="text" id="sopo-${p.id}" value="${esc(p.soPO||'')}" placeholder="Số PO" aria-label="Số PO của ${esc(p.ncc)}"><input type="text" inputmode="numeric" id="tsap-${p.id}" value="${p.tongSAP!=null?QI(p.tongSAP):''}" placeholder="Tổng trên SAP" aria-label="Tổng tiền trên SAP của ${esc(p.ncc)}"></div>`:`<span class="num">${esc(p.soPO||'')}</span>${p.tongSAP!=null?` · <span class="num">${VND(p.tongSAP)}</span>`:''}`}${d==null?'':`<small>Chênh lệch <span class="pill ${Math.abs(d)<1?'ok':'bad'}">${d>0?'+':''}${VND(d)}</span></small>`}</td>
    <td class="tt">${poPill(p)}${lenhPill(p.lenh)}</td>
    <td class="acts"><div class="acts-in">${ed&&canWrite?`<button class="btn sm" data-act3="po-save" data-id="${p.id}">Lưu số PO</button>`:''}${can(p)?`<button class="btn sm" data-tram="po" data-id="${p.id}">Tạo trên SAP</button>`:''}</div></td></tr>`;}).join('')}</tbody></table></div>
  ${canWrite?`<details class="more"${ED.poPaste?' open':''}><summary>Thao tác khác</summary><div class="toolbar"><button class="btn sm" data-act3="po-paste">Dán số PO từ sheet PO</button><span class="meta">Mỗi dòng: Nhà cung cấp ⇥ Số PO ⇥ Tổng tiền trên SAP (tùy chọn).</span></div>
  ${ED.poPaste?`<div class="box"><textarea class="paste" id="po-txt" aria-label="Dán số PO"></textarea><div class="toolbar"><button class="btn sm primary" data-act3="po-paste-ok">Ghi số PO</button><button class="btn sm" data-act3="po-paste">Đóng</button></div></div>`:''}</details>`:''}`
  :'<div class="empty"><b>Chưa có PO nào</b>Lập danh sách PO ở 3.1.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-tab="pocan">Mở 3.1</button></div></div>');
}
/* ---- 3.3 PO so với nhận hàng ---- */
function poRecvCompute(){
  const P=poOf(S.thang).filter(p=>p.soPO||ED.podcAll), rv=recvRows();
  const days=monthDays(S.thang), dim=days.length; let last=''; rv.forEach(r=>{if(r.d>last)last=r.d;});
  let asOf=last||todayVN(); if(asOf<days[0]) asOf=''; if(asOf>days[dim-1]) asOf=days[dim-1];
  const frac=asOf?(+asOf.slice(8))/dim:0;
  const got=new Map(); rv.forEach(r=>{const k=String(r.ncc||'').toUpperCase()+'|'+r.ma; const o=got.get(k)||{sl:0,n:0,migo:new Set()}; o.sl+=+r.sl||0; o.n++; if(r.migo) o.migo.add(r.migo); got.set(k,o);});
  const rows=[], byKey=new Map();
  P.slice().sort((a,b)=>a.luc-b.luc).forEach(p=>{ const m=new Map(); p.lines.forEach((l,i)=>{const o=m.get(l.ma)||{ma:l.ma,ten:l.ten,dvt:l.dvt,g:l.g,sl:0,items:[]};o.sl+=l.sl;o.items.push((i+1)*10);m.set(l.ma,o);});
    m.forEach(o=>{const k=p.ncc.toUpperCase()+'|'+o.ma; const r={po:p,...o,key:k,sR:0}; rows.push(r); (byKey.get(k)||byKey.set(k,[]).get(k)).push(r);}); });
  byKey.forEach((list,k)=>{ let rem=(got.get(k)||{sl:0}).sl; list.forEach((r,i)=>{ const take=i===list.length-1?rem:Math.min(rem,r.sl); r.sR=Math.max(0,take); rem-=r.sR; }); });
  const out=[];
  got.forEach((g,k)=>{ if(byKey.has(k)) return; const [ncc,ma]=k.split('|'); const x=IT.get(ma)||{ten:ma,dvt:'',ncc:[]}; out.push({po:null,ncc,ma,ten:x.ten,dvt:x.dvt,g:priceOf(x,ncc)||0,sl:0,items:[],sR:g.sl,key:k}); });
  [...rows,...out].forEach(r=>{ const kh=r.sl*frac, db=frac>0?r.sR/frac:r.sR; r.kh=kh; r.db=db; r.lech=r.sR-kh;
    if(!r.po) r.st='ngoai'; else if(r.sR>r.sl+1e-9) r.st='vuot'; else if(frac>0&&frac<1&&r.sR>0&&db>r.sl*1.0001) r.st='sap_vuot';
    else if(r.sR>=r.sl-1e-9&&r.sl>0) r.st='du'; else if(r.sR===0) r.st=frac>=0.5?'thieu':'chua_nhan'; else if(frac>=0.5&&db<r.sl*0.8) r.st='thieu'; else r.st='dang';
    r.need=(r.st==='vuot'||r.st==='sap_vuot'||r.st==='ngoai')?roundQ(Math.max(0,Math.max(r.sR,frac<1?db:r.sR)-r.sl),r.dvt):0; });
  return {rows:[...rows,...out],asOf,frac,nRecv:rv.length};
}
const PODC={ngoai:['Nhận ngoài PO','bad',1],vuot:['Vượt PO','bad',2],sap_vuot:['Dự kiến vượt','run',3],thieu:['Nhận chậm / dự kiến dư','run',4],dang:['Đúng kế hoạch','ok',5],chua_nhan:['Chưa nhận','',6],du:['Đã nhận đủ','ok',7]};
function vPOdc(){
  const C=poRecvCompute(), f=ED.podcF||'', nf=ED.podcN||'';
  const nccs=[...new Set(C.rows.map(r=>r.po?r.po.ncc:r.ncc))].sort((a,b)=>a.localeCompare(b,'vi'));
  const cnt={}; C.rows.forEach(r=>cnt[r.st]=(cnt[r.st]||0)+1);
  const shown=C.rows.filter(r=>(!f||r.st===f)&&(!nf||(r.po?r.po.ncc:r.ncc)===nf)).sort((a,b)=>PODC[a.st][2]-PODC[b.st][2]||Math.abs(b.lech*b.g)-Math.abs(a.lech*a.g));
  const tPO=C.rows.reduce((a,r)=>a+r.sl*r.g,0), tR=C.rows.reduce((a,r)=>a+r.sR*r.g,0), tKH=C.rows.reduce((a,r)=>a+r.kh*r.g,0);
  const thua=C.rows.filter(r=>r.st==='vuot'||r.st==='ngoai').reduce((a,r)=>a+(r.sR-r.sl)*r.g,0), du=C.rows.filter(r=>r.st==='thieu').reduce((a,r)=>a+Math.max(0,r.sl-r.db)*r.g,0);
  ED.posl=ED.posl||{};
  const body=shown.slice(0,400).map(r=>{const s=PODC[r.st],pct=r.sl?r.sR/r.sl*100:0,k=r.po?r.po.id+'|'+r.ma:'';
    return `<tr class="${ED.posl[k]!=null?'edited':''}"><td>${esc(r.po?r.po.ncc:r.ncc)}<div class="meta num">${r.po?esc(r.po.soPO||'chưa có số'):'—'}${r.items.length?' · dòng '+r.items.join(','):''}</div></td><td class="num">${r.ma}</td><td class="wrap">${esc(r.ten)}</td><td>${esc(r.dvt)}</td>
    <td class="r num">${Q(r.sl)}</td><td class="r num">${C.frac>0?Q(Math.round(r.kh*10)/10):'–'}</td><td class="r num">${Q(r.sR)}</td>
    <td class="r num"><span class="${r.lech>0.05?'dirty':r.lech<-0.05?'muted':''}">${C.frac>0?(r.lech>0?'+':'')+Q(Math.round(r.lech*10)/10):'–'}</span></td>
    <td class="r num">${r.sl?Math.round(pct)+'%':'–'}${pct>0?`<div class="bar" style="width:${Math.min(100,pct)}%;background:var(${pct>100?'--bad':'--accent'})"></div>`:''}</td>
    <td class="r num">${C.frac>0&&C.frac<1?Q(Math.round(r.db*10)/10):'–'}</td><td><span class="pill ${s[1]}">${s[0]}</span></td>
    <td class="r num">${r.need>0?Q(r.need):''}</td></tr>`;}).join('');
  const nEdit=Object.keys(ED.posl).length, nBS=C.rows.filter(r=>r.need>0).length;
  return head2('3.3 PO so với nhận hàng','So từng dòng PO với số thực nhận (phiếu MIGO) theo NCC và mã hàng. Kế hoạch đến ngày = SL PO × phần tháng đã qua; cột Chênh cho biết đang nhập nhanh (+) hay chậm (−) so với kế hoạch. Báo cáo để theo dõi trong kỳ; việc chỉnh PO và PR bổ sung làm một lần cuối chu kỳ ở bước 6.1, sau khi đối soát NCC.',
    '<span class="pill">Theo dõi, không bắt buộc</span>')+
  (!C.rows.length?`<div class="empty"><b>Chưa có PO nào có số</b>Ghi số PO ở 3.2 rồi quay lại đây để so với số nhận hàng.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-tab="posap">Mở 3.2</button><label class="meta"><input type="checkbox" id="podc-all" ${ED.podcAll?'checked':''}> Tính cả PO chưa có số</label></div></div>`:

  `<div class="kpis"><div class="kv big"><b>Đã nhận / tổng PO</b><div class="num">${TR(tR)} / ${TR(tPO)}</div><small>${tPO?Math.round(tR/tPO*100):0}% · tới ${C.asOf?dm(C.asOf):'—'} · ${C.nRecv} dòng nhận</small></div>
   <div class="kv"><b>So với kế hoạch đến ngày</b><div class="num">${C.frac>0?(tR-tKH>=0?'+':'')+TR(tR-tKH):'—'}</div><small>KH đến ngày ${TR(tKH)}</small></div>
   <div class="kv"><b>Nhập thừa (vượt + ngoài PO)</b><div class="num">${TR(thua)}</div><small>${(cnt.vuot||0)+(cnt.ngoai||0)} dòng · cần PR bổ sung</small></div>
   <div class="kv"><b>Dự kiến dư cuối tháng</b><div class="num">${TR(du)}</div><small>${cnt.thieu||0} dòng nhận chậm</small></div></div>
  <div class="toolbar"><select id="podc-ncc" aria-label="Lọc theo NCC"><option value="">Tất cả NCC</option>${nccs.map(n=>`<option ${n===nf?'selected':''}>${esc(n)}</option>`).join('')}</select>
   <div class="seg"><button data-act3="podc-f" data-f="" aria-pressed="${!f}">Tất cả ${C.rows.length}</button>${Object.entries(PODC).filter(([k])=>cnt[k]).map(([k,v])=>`<button data-act3="podc-f" data-f="${k}" aria-pressed="${f===k}">${v[0]} ${cnt[k]}</button>`).join('')}</div><span class="sp"></span>
   <label class="meta"><input type="checkbox" id="podc-all" ${ED.podcAll?'checked':''}> Tính cả PO chưa có số</label></div>
  <div class="meta">${S.nhan?`Số nhận: ${(S.nhan.rows||[]).length} dòng · ${esc(S.nhan.nguon||'')} · cập nhật ${tstr(S.nhan.luc)}. Dán thêm số nhận ở bước 2.4.`:'Chưa có số nhận hàng tháng này. Bước 4 (MIGO) sẽ ghi vào; trước đó dán tay ở bước 2.4.'}</div>
  <div class="tbl-wrap"><table class="t"><thead><tr><th>NCC · PO</th><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th class="r">SL PO</th><th class="r">KH đến ngày</th><th class="r">Đã nhận</th><th class="r">Chênh</th><th class="r">% PO</th><th class="r">Dự báo cuối tháng</th><th>Tình trạng</th><th class="r">Dự kiến thiếu</th></tr></thead><tbody>${body}</tbody></table></div>`);
}
async function makeBSFrom(items,note,buoc){
  const p=S.pb?pbRatio({...S.pb,lines:(S.pb.lines||[])}):pbDefaultPct()/100;
  const nv=[],bn=[]; items.forEach(o=>{const a=splitNV(o.q,p,o.dvt);if(a>0)nv.push({ma:o.ma,ncc:o.ncc,sl:a,g:o.g});if(o.q-a>0)bn.push({ma:o.ma,ncc:o.ncc,sl:+(o.q-a).toFixed(3),g:o.g});});
  let n=0; for(const [dt,L] of [['NhanVien',nv],['BenhNhan',bn]]) if(L.length){const b=makeBatch('Bổ sung',dt,L,note); await store.set('prdot/'+b.id,b.doc); n++;}
  ED.sel=null; await log(buoc,`Tạo ${n} đợt PR bổ sung cho ${items.length} mã (${TR(items.reduce((a,o)=>a+o.q*o.g,0))})`); return n;
}
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act3]'); if(!t) return; const a=t.dataset.act3;
  try{
  if(a==='podc-f'){ ED.podcF=t.dataset.f; render(); }
  else if(a==='posl-save'){ const ch=ED.posl||{}, byPO={};
    Object.entries(ch).forEach(([k,v])=>{const [id,ma]=k.split('|');(byPO[id]=byPO[id]||[]).push([ma,v]);});
    for(const [id,list] of Object.entries(byPO)){ const p=S.po[id]; const lines=p.lines.map(l=>({...l})); const notes=[];
      for(const [ma,nv] of list){ const L=lines.filter(l=>l.ma===ma), old=L.reduce((a,l)=>a+l.sl,0); if(Math.abs(nv-old)<1e-9) continue; const x=IT.get(ma);
        L.forEach((l,i)=>{ l.sl=i===L.length-1?+(nv-L.slice(0,-1).reduce((a,y)=>a+y.sl,0)).toFixed(3):(old?(isKG(l.dvt)?Math.round(l.sl/old*nv*10)/10:Math.round(l.sl/old*nv)):0); if(l.sl<0) l.sl=0; });
        notes.push(`${ma} ${Q(old)} → ${Q(nv)}`); }
      if(!notes.length) continue; const tong=Math.round(lines.reduce((a,l)=>a+l.sl*l.g,0));
      await setPO(id,{lines,tong,sua:true},'Sửa SL: '+notes.join('; ')); await log('3.3',`Sửa SL PO ${p.soPO||p.ncc}: ${notes.join('; ')} (cần sửa ME22N)`); }
    ED.posl={}; toast('Đã lưu. Sang 3.4 bấm “Cập nhật trên SAP” để trạm sửa PO (ME22N).'); }
  else if(a==='po-bs'){ const C=poRecvCompute(); const items=C.rows.filter(r=>r.need>0).map(r=>({ma:r.ma,ncc:r.po?r.po.ncc:r.ncc,g:r.g,dvt:r.dvt,q:ED.pobs&&ED.pobs[r.key]!=null?ED.pobs[r.key]:r.need})).filter(o=>o.q>0);
    if(!items.length){toast('Không có dòng nào cần bổ sung.');return;}
    await makeBSFrom(items,`Bổ sung từ đối chiếu PO – nhận hàng tới ${C.asOf?dm(C.asOf):''}`,'3.3'); ED.pobs={}; toast('Đã tạo PR bổ sung. Sang bước 2.2 để xuất lên SAP.'); }
  else if(a==='po-sapok'){ const p=S.po[t.dataset.id]; await setPO(p.id,{sua:false},'Đã cập nhật SL trên SAP (ME22N)'); await log('3.4',`PO ${p.soPO}: đã cập nhật SAP sau khi sửa SL`); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target;
  if(t.dataset.posl){ ED.posl=ED.posl||{}; ED.posl[t.dataset.posl]=Math.max(0,pnum(t.value)); clearTimeout(vPOdc.h); vPOdc.h=setTimeout(render,700); }
  else if(t.dataset.pobs){ ED.pobs=ED.pobs||{}; ED.pobs[t.dataset.pobs]=Math.max(0,pnum(t.value)); }
});
document.addEventListener('change',e=>{ const t=e.target; if(t.id==='podc-all'){ED.podcAll=t.checked;render();} else if(t.id==='podc-ncc'){ED.podcN=t.value;render();} });

/* ---- 3.4 Lịch sử PO ---- */
function vPOls(){
  const all=poAll().filter(p=>ED.poAll||p.thang===S.thang).sort((a,b)=>b.luc-a.luc);
  const body=all.map(p=>{const open=ED.poOpen===p.id;
    const acts=canWrite?[p.sua&&p.soPO?(lenhBusy(p.lenh)?lenhPill(p.lenh):`<button class="btn sm primary" data-tram="posua" data-id="${p.id}">Cập nhật trên SAP</button><button class="btn sm" data-act3="po-sapok" data-id="${p.id}">Đã sửa tay</button>`):'',p.tt!=='huy'?`<button class="btn sm danger-text" data-act3="po-huy" data-id="${p.id}">${ED.poHuy===p.id?'Bấm lần nữa để hủy':'Hủy'}</button>`:''].join(''):'';
    return `<tr><td class="exp"><button class="x" data-act3="po-open" data-id="${p.id}" aria-expanded="${open}" aria-label="${open?'Ẩn':'Xem'} chi tiết PO ${esc(p.ncc)}">${open?'▾':'▸'}</button></td>
    <td class="dg"><b>${esc(p.ncc)}</b><small><span class="num">${p.soPO?'PO '+esc(p.soPO):'chưa có số PO'}</span> · ${esc(p.loai)}${ED.poAll?' · '+mmY(p.thang):''}</small></td>
    <td class="r tien"><b class="num">${VND(p.tong)}</b><small>${p.lines.length} dòng</small></td>
    <td class="tt">${poPill(p)}${p.sua&&p.soPO?'<small>Đã sửa SL, chờ cập nhật SAP</small>':''}</td>
    <td class="meta when">Lập ${tstr(p.luc)}${p.sapLuc&&p.soPO?`<br>Có số ${tstr(p.sapLuc)}`:''}</td>
    <td class="acts"><div class="acts-in">${acts}</div></td></tr>
    ${open?`<tr class="detail"><td colspan="6"><div class="two" style="grid-template-columns:minmax(0,2fr) minmax(0,1fr)"><div class="tbl-wrap" style="max-height:320px"><table class="t"><thead><tr><th>Dòng PO</th><th>Mã</th><th>Tên</th><th class="r">SL</th><th class="r">Đơn giá</th><th>PR</th><th>Đối tượng</th></tr></thead><tbody>
      ${p.lines.map((l,i)=>`<tr><td class="num">${l.ebelp||((i+1)*10+' ?')}</td><td class="num">${l.ma}</td><td class="wrap">${esc(l.ten)}</td><td class="r num">${Q(l.sl)} ${esc(l.dvt)}</td><td class="r num">${VND(l.g)}</td><td class="num">${esc(l.soPR)}</td><td>${DTG[l.dt]?DTG[l.dt].k:''}</td></tr>`).join('')}</tbody></table></div>
      <div><b class="meta">Lịch sử PO</b><ol style="margin:6px 0 0;padding-left:18px;font-size:13px">${(p.hist||[]).map(h=>`<li><span class="num meta">${tstr(h.ts)}</span> · ${(TTPO[h.tt]||[h.tt])[0]} · ${esc(nm(h.ai))}${h.note?' · '+esc(h.note):''}</li>`).join('')}</ol>
      <p class="meta">Dòng PO lấy theo thứ tự tạo trên SAP; dòng có “?” là dòng mới chưa cập nhật lên SAP.</p></div></div></td></tr>`:''}`;}).join('');
  const m=poOf(S.thang).filter(p=>p.tt!=='huy'), nSua=m.filter(p=>p.sua&&p.soPO).length;
  return head2('3.4 Lịch sử PO','Mọi PO theo nhà cung cấp: chờ tạo trên SAP → có số PO. PO đã sửa số lượng hoặc thêm dòng PR bổ sung: bấm <b>Cập nhật trên SAP</b>, trạm SAP chạy ME22N (sửa SL các dòng có sẵn, thêm dòng mới) và lưu khi khớp tuyệt đối.')+
  (m.length?`<div class="kpis"><div class="kv big"><b>PO tháng ${mmY(S.thang)}</b><div class="num">${TR(m.reduce((a,p)=>a+p.tong,0))}</div><small>${m.length} NCC</small></div>
   <div class="kv"><b>Có số PO</b><div class="num">${m.filter(p=>p.soPO).length} / ${m.length}</div><small>${m.filter(p=>p.tt==='cho').length} PO chờ tạo trên SAP</small></div>
   <div class="kv"><b>Chờ cập nhật SAP</b><div class="num">${nSua}</div><small>PO đã sửa số lượng</small></div></div>`:'')+
  (nSua?tramBox('POSUA'):'')+
  `<div class="toolbar"><label class="meta"><input type="checkbox" id="po-all" ${ED.poAll?'checked':''}> Hiện cả các tháng khác</label></div>
  ${all.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl"><colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"><col class="c-when"><col class="c-acts"></colgroup><thead><tr><th><span class="sr">Chi tiết</span></th><th>Nhà cung cấp</th><th class="r">Tiền</th><th>Trạng thái</th><th>Lập / có số</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>${body}</tbody></table></div>`
  :'<div class="empty"><b>Chưa có PO nào</b>Lập danh sách PO ở 3.1.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-tab="pocan">Mở 3.1</button></div></div>'}`;
}
/* ---- sự kiện bước 3 ---- */
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act3]'); if(!t) return; const a=t.dataset.act3, id=t.dataset.id;
  try{
  if(a==='po-make'){ const pend=poPending(); let n=0;
    for(const o of pend){ const pid=S.thang+'-'+slug(o.ncc)+'-'+Date.now().toString(36)+n;
      await store.set('po/'+pid,{thang:S.thang,ncc:o.ncc,lifnr:LIFNR[o.ncc.toUpperCase()]||'',loai:o.loai.includes('Bổ sung')&&!o.loai.includes('Chính')?'Bổ sung':o.loai.join('+'),src:o.src,pr:o.pr,lines:o.lines,tong:o.tong,tt:'cho',soPO:'',tongSAP:null,email:[],luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'cho',ai:uid,note:'Lập từ PR '+o.pr.join(', ')}]}); n++; }
    await log('3.1',`Lập danh sách ${n} PO tháng ${mmY(S.thang)} (${TR(pend.reduce((s,o)=>s+o.tong,0))})`); toast(`Đã lập ${n} PO. Sau khi tạo trên SAP, ghi số PO ở 3.2.`); }
  else if(a==='po-save'){ const v=($('#sopo-'+id).value||'').trim(), ts=$('#tsap-'+id).value.trim(); if(!/^\d{6,12}$/.test(v)){toast('Số PO là dãy 6–12 chữ số.');return;}
    const p=S.po[id], tsv=ts===''?null:pnum(ts.replace(/\./g,'')); await setPO(id,{soPO:v,tongSAP:tsv,tt:p.tt==='cho'?'sap':p.tt,sapLuc:p.sapLuc||Date.now()},'Ghi số PO '+v+(tsv!=null?' · SAP '+VND(tsv):''));
    await log('3.2',`Ghi số PO ${v} cho ${p.ncc}${tsv!=null&&Math.abs(tsv-p.tong)>=1?` – lệch ${VND(tsv-p.tong)} so với PR`:''}`); }
  else if(a==='po-paste'){ ED.poPaste=!ED.poPaste; render(); }
  else if(a==='po-paste-ok'){ const L=poOf(S.thang), bad=[]; let n=0;
    for(const line of $('#po-txt').value.split(/\r?\n/)){ if(!line.trim()) continue; const c=line.split('\t').map(x=>x.trim()); const p=L.find(x=>x.ncc.toUpperCase()===(c[0]||'').toUpperCase()&&(x.tt==='cho'||x.tt==='sap'));
      if(!p||!/^\d{6,12}$/.test(c[1]||'')){bad.push(line);continue;} const tsv=c[2]?pnum(c[2].replace(/\./g,'')):null; await setPO(p.id,{soPO:c[1],tongSAP:tsv,tt:'sap',sapLuc:p.sapLuc||Date.now()},'Dán từ sheet PO'); n++; }
    if(n) await log('3.2',`Ghi ${n} số PO (dán từ sheet PO)`); ED.poPaste=bad.length>0; render(); if(ED.poPaste) $('#po-txt').value=bad.join('\n'); toast(`Đã ghi ${n} số PO${bad.length?`, ${bad.length} dòng không khớp NCC hoặc số PO`:''}.`); }
  else if(a==='po-open'){ ED.poOpen=ED.poOpen===id?null:id; render(); }
  else if(a==='po-huy'){ const p=S.po[id]; if(ED.poHuy!==id){ED.poHuy=id;toast('Bấm Hủy lần nữa để hủy PO '+(p.soPO||p.ncc)+'. Các dòng PR sẽ quay lại 3.1.');render();return;} ED.poHuy=null;
    await setPO(id,{tt:'huy'},'Hủy PO'); await log('3.4',`Hủy PO ${p.soPO||'(chưa có số)'} – ${p.ncc}`); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('change',e=>{ const t=e.target; if(t.id==='po-all'){ ED.poAll=t.checked; render(); }
  else if(t.dataset&&t.dataset.posel){ ED.poBo=ED.poBo||new Set(); t.checked?ED.poBo.delete(t.dataset.posel):ED.poBo.add(t.dataset.posel); render(); }
  else if(t.id==='posel-all'){ const L=poOf(S.thang).filter(p=>p.tt==='cho'&&!lenhBusy(p.lenh)); ED.poBo=new Set(t.checked?[]:L.map(p=>p.id)); render(); } });

