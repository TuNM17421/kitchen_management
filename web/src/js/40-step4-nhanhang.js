/* ================= BƯỚC 4 · ĐẶT HÀNG, NHẬN HÀNG & MIGO TỪNG PHẦN ================= */
const _x=0;
const SUB4=[['4.1','Đặt hàng theo PO','dat'],['4.2','Nhận hàng','nhan'],['4.3','Lịch sử nhận hàng','nls']];
const TTPN={nhan:['Đã nhận, chờ MIGO','run'],xuat:['Đã xuất cho macro','run'],migo:['Đã MIGO','ok'],huy:['Đã hủy','bad']};
const pnAll=()=>Object.values(S.pn||{}).filter(p=>p.tt!=='huy');
const pnOf=t=>pnAll().filter(p=>p.thang===t).sort((a,b)=>a.d.localeCompare(b.d)||a.luc-b.luc);
const dhOf=t=>Object.values(S.dh||{}).filter(d=>d.thang===t&&d.tt!=='huy').sort((a,b)=>a.d.localeCompare(b.d)||a.luc-b.luc);
const poById=id=>S.po&&S.po[id];
const poItems=p=>p.lines.map((l,i)=>({...l,prItem:l.item,item:+l.ebelp||(i+1)*10}));
function poMerge(p){ const m=new Map(); poItems(p).forEach(l=>{const o=m.get(l.ma)||{ma:l.ma,ten:l.ten,dvt:l.dvt,g:l.g,sl:0,items:[]};o.sl+=l.sl;o.items.push(l.item);m.set(l.ma,o);}); return [...m.values()]; }
function poUsed(poId,exceptId){ const u={}; pnAll().filter(p=>p.poId===poId&&p.id!==exceptId).forEach(p=>p.lines.forEach(l=>(l.alloc||[]).forEach(a=>u[a.item]=(u[a.item]||0)+a.sl))); return u; }
function poRecvByMa(poId){ const u={}; pnAll().filter(p=>p.poId===poId).forEach(p=>p.lines.forEach(l=>u[l.ma]=(u[l.ma]||0)+(+l.sl||0))); return u; }
function poOrdByMa(poId){ const u={}; dhOf(S.thang).filter(d=>d.poId===poId).forEach(d=>d.lines.forEach(l=>u[l.ma]=(u[l.ma]||0)+(+l.sl||0))); return u; }
function allocItems(p,ma,qty,used){ let rem=qty; const out=[];
  poItems(p).filter(x=>x.ma===ma).forEach(x=>{const cap=Math.max(0,x.sl-(used[x.item]||0)), t=Math.min(cap,rem); if(t>1e-9){out.push({item:x.item,sl:+t.toFixed(3)});used[x.item]=(used[x.item]||0)+t;rem-=t;}});
  return {alloc:out,over:+Math.max(0,rem).toFixed(3)}; }
/* số nhận dùng chung cho báo cáo 2.4 và 3.3: phiếu nhận (bước 4) + số dán tay / nạp từ Công nợ */
function recvRows(){ const a=((S.nhan&&S.nhan.rows)||[]).slice();
  pnOf(S.thang).forEach(p=>p.lines.forEach(l=>{ if(+l.sl>0) a.push({d:p.d,ma:l.ma,ncc:p.ncc,sl:+l.sl,migo:p.migo||'',po:p.soPO,nguon:'Phiếu '+p.ma}); })); return a; }
const poWithNo=()=>poOf(S.thang).filter(p=>p.soPO&&p.tt!=='huy');
const tomorrowVN=()=>{const d=new Date(Date.now()+7*3600e3+864e5);return d.toISOString().slice(0,10);};

/* ---- 4.1 Đặt hàng theo PO ---- */
const emptyPO4=t=>head2(t,'')+'<div class="empty"><b>Chưa có PO nào có số</b>Ghi số PO ở bước 3.2 rồi quay lại đây.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-step="3">Mở bước 3</button></div></div>';
const poOpt=(P,cur)=>P.map(x=>`<option value="${x.id}" ${x.id===cur?'selected':''}>${esc(x.ncc)} · PO ${esc(x.soPO)}</option>`).join('');
function vDat(){
  const P=poWithNo(); if(!P.length) return emptyPO4('4.1 Đặt hàng theo PO');
  if(!ED.datPO||!P.find(p=>p.id===ED.datPO)) ED.datPO=P[0].id; if(!ED.datD) ED.datD=tomorrowVN();
  const p=poById(ED.datPO), M=poMerge(p), rec=poRecvByMa(p.id), ord=poOrdByMa(p.id); ED.datQ=ED.datQ||{};
  const fq=fold(ED.datF||''); let tong=0, n=0; const over=[];
  const rows=M.map(m=>{const q=ED.datQ[m.ma]; const con=Math.max(0,m.sl-(rec[m.ma]||0)); if(q>0){n++;tong+=q*m.g; if(q>con+1e-9) over.push(m.ma);}
    const hide=(ED.datOnly&&!(q>0))||(fq&&!fold(m.ma+' '+m.ten).includes(fq));
    return `<tr class="${q>0?'edited':''}" data-k="${esc(m.ma+' '+m.ten)}" ${hide?'hidden':''}><td class="nm"><b>${esc(m.ten)}</b><small><span class="num">${m.ma}</span> · dòng PO ${m.items.join(', ')}</small></td><td class="dvt">${esc(m.dvt)}</td><td class="r num slpo">${Q(m.sl)}</td>
    <td class="r num con"><b>${Q(con)}</b><small>đặt ${Q(ord[m.ma]||0)} · nhận ${Q(rec[m.ma]||0)}</small></td>
    <td class="r q"><input class="q" type="text" inputmode="decimal" id="dq-${m.ma}" data-dq="${m.ma}" value="${q?QI(q):''}" aria-label="SL đặt ${esc(m.ten)}" placeholder="0" ${canWrite?'':'disabled'}>${q>con+1e-9?'<small class="bad-t">vượt PO</small>':''}</td></tr>`;}).join('');
  const done=dhOf(S.thang).filter(d=>d.poId===p.id).slice().reverse();
  return head2('4.1 Đặt hàng theo PO','Chọn ngày giao và nhà cung cấp, nhập số lượng đặt cho từng mã. Đơn hàng mang số PO và dòng PO để bên giao ghi lên biên bản; số đặt được dùng làm số nhận mặc định ở bước 4.2.')+
  `<div class="toolbar"><span class="fld"><label class="meta" for="dat-d">Ngày giao</label><input type="date" id="dat-d" value="${ED.datD}"></span>
   <span class="fld"><label class="meta" for="dat-po">PO</label><select id="dat-po">${poOpt(P,p.id)}</select></span></div>
  <div class="toolbar"><input type="search" id="dat-find" value="${esc(ED.datF||'')}" placeholder="Tìm mã hoặc tên hàng" aria-label="Tìm mã hoặc tên hàng">
   <div class="seg"><button data-act4="dat-only" data-v="" aria-pressed="${!ED.datOnly}">Tất cả ${M.length}</button><button data-act4="dat-only" data-v="1" aria-pressed="${!!ED.datOnly}">Đã nhập ${n}</button></div></div>
  ${over.length?`<div class="warn">${over.length} mã đặt vượt phần còn lại của PO. SAP không cho MIGO vượt PO: làm PR bổ sung (bước 3.3 hoặc 2.4) trước khi nhận.</div>`:''}
  <div class="tbl-wrap"><table class="t dat-tbl"><colgroup><col><col class="c-dvt"><col class="c-n"><col class="c-con"><col class="c-q"></colgroup><thead><tr><th>Mặt hàng</th><th>ĐVT</th><th class="r">SL PO</th><th class="r">Còn lại</th><th class="r">SL đặt</th></tr></thead><tbody>${rows}</tbody></table></div>
  ${canWrite?`<div class="toolbar savebar"><span class="meta">${n?`<b>${n} mã</b> · <b class="num">${VND(tong)} đ</b>`:'Chưa nhập số lượng'}</span><span class="sp"></span><button class="btn" data-act4="dat-copy" ${tong?'':'disabled'}>Copy đơn hàng</button><button class="btn primary" data-act4="dat-save" ${tong?'':'disabled'}>Lưu đơn hàng</button></div>`:''}
  ${done.length?`<details class="more"${ED.datDone?' open':''} id="dat-done"><summary>Đơn đã đặt của PO này (${done.length})</summary><ul class="dh-list">${done.map(d=>`<li><span><b>${dm(d.d)}</b> · ${d.lines.length} mã · <span class="num">${VND(d.tong)} đ</span> ${d.pn?'<span class="pill ok">đã nhận</span>':'<span class="pill run">chưa nhận</span>'}</span><button class="btn sm" data-act4="dh-copy" data-id="${d.id}">Copy lại</button></li>`).join('')}</ul></details>`:''}`;
}
function datText(p,d,lines){ return `ĐƠN ĐẶT HÀNG – Bếp Vinmec Ocean Park 2\nNgày giao: ${dm(d)}/${d.slice(0,4)}\nNCC: ${p.ncc} · Số PO: ${p.soPO}\n`+lines.map((l,i)=>`${i+1}. [PO ${p.soPO} dòng ${l.items.join('/')}] ${l.ma} ${l.ten}: ${Q(l.sl)} ${l.dvt}`).join('\n')+`\nĐề nghị ghi số PO và mã hàng trên biên bản giao hàng.`; }

/* ---- 4.2 Nhận hàng ---- */
function nhanDraft(){
  if(ED.nd) return ED.nd;
  const P=poWithNo(); const dhs=dhOf(S.thang).filter(d=>!d.pn);
  if(ED.ndDH==null) ED.ndDH=dhs[0]?dhs[0].id:'';
  const dh=dhs.find(d=>d.id===ED.ndDH)||null; const po=dh?poById(dh.poId):(P.find(p=>p.id===ED.ndPO)||P[0]);
  if(!po) return null;
  const M=new Map(poMerge(po).map(m=>[m.ma,m]));
  const lines=dh?dh.lines.map(l=>{const m=M.get(l.ma)||{};return {ma:l.ma,ten:m.ten||l.ten,dvt:m.dvt||l.dvt,g:m.g||l.g,slDat:l.sl,sl:l.sl,gach:false,note:''};}):[];
  ED.nd={dh:dh?dh.id:'',poId:po.id,d:dh?dh.d:todayVN(),lines,nguon:dh?'Đơn hàng '+dm(dh.d):'Nhập tay'}; return ED.nd;
}
function vNhan(){
  const P=poWithNo(); if(!P.length) return emptyPO4('4.2 Nhận hàng');
  const N=nhanDraft(), po=poById(N.poId), used=poUsed(po.id), dhs=dhOf(S.thang).filter(d=>!d.pn);
  const M=poMerge(po), sim={...used}; let tong=0, tDat=0;
  const rows=N.lines.map((l,i)=>{const q=l.gach?0:+l.sl||0; const a=allocItems(po,l.ma,q,sim); tong+=q*l.g; tDat+=(l.slDat||0)*l.g; const d=l.slDat!=null&&!l.gach?q-l.slDat:0;
    return `<tr class="${l.gach?'gach':(Math.abs(d)>1e-9?'edited':'')}"><td class="nm"><b>${esc(l.ten)}</b><small class="num">${l.ma}</small></td><td class="dvt">${esc(l.dvt)}</td><td class="r num dat"><span class="lbl">Đặt </span>${l.slDat!=null?Q(l.slDat):'–'}</td>
    <td class="r q"><input class="q" type="text" inputmode="decimal" id="nq-${i}" data-nq="${i}" value="${QI(l.sl)}" aria-label="Thực nhận ${esc(l.ten)}" ${l.gach||!canWrite?'disabled':''}>${Math.abs(d)>1e-9?`<small class="${d>0?'bad-t':''}">${d>0?'+':''}${Q(d)} so với đặt</small>`:''}${a.alloc.length>1?`<small>dòng PO ${a.alloc.map(x=>x.item+': '+Q(x.sl)).join(', ')}</small>`:''}${a.over>0?`<small class="bad-t">vượt PO ${Q(a.over)}</small>`:''}</td>
    <td class="gch"><label><input type="checkbox" data-ngach="${i}" ${l.gach?'checked':''}> Gạch bỏ</label></td>
    <td class="note">${l.note||ED.ndNote===i?`<input type="text" id="nn-${i}" data-nn="${i}" value="${esc(l.note||'')}" placeholder="Ghi chú" aria-label="Ghi chú ${esc(l.ten)}">`:`<button class="btn sm ghost" data-act4="nn-open" data-i="${i}">✎ Ghi chú</button>`}</td>
    <td class="del"><button class="x" data-act4="nl-del" data-i="${i}" aria-label="Xóa dòng ${esc(l.ten)}">×</button></td></tr>`;}).join('');
  const addable=M.filter(m=>!N.lines.some(l=>l.ma===m.ma)), ok=ED.ndOk;
  return head2('4.2 Nhận hàng','Mỗi lần giao là một phiếu nhận: ngày giao, PO, số thực nhận từng mã. Mặc định thực nhận = số đặt; dòng bị gạch bỏ trên biên bản = 0. Đây là dữ liệu chuẩn để đối soát NCC (bước 5); phần vượt PO được xử lý một lần cuối chu kỳ (bước 6). Có thể đọc ảnh biên bản để điền tự động.')+
  (ok?`<div class="okbox" role="status"><span>Đã lưu phiếu <b class="num">${esc(ok)}</b>.${dhs.length?` Còn ${dhs.length} đơn chưa nhận, đã mở đơn tiếp theo.`:''}</span><button class="btn sm" data-act="go-next" data-tab="nls">Xem ở 4.3</button><button class="x" data-act4="ok-x" aria-label="Đóng thông báo">×</button></div>`:'')+
  `<div class="toolbar"><span class="fld"><label class="meta" for="nd-dh">Đơn hàng</label><select id="nd-dh"><option value="">Nhận không theo đơn</option>${dhs.map(d=>`<option value="${d.id}" ${d.id===N.dh?'selected':''}>${dm(d.d)} · ${esc(d.ncc)} · ${d.lines.length} mã</option>`).join('')}</select></span>
   ${N.dh?'':`<span class="fld"><label class="meta" for="nd-po">PO</label><select id="nd-po">${poOpt(P,po.id)}</select></span>`}
   <span class="fld"><label class="meta" for="nd-d">Ngày giao</label><input type="date" id="nd-d" value="${N.d}" title="Cũng là ngày hạch toán MIGO"></span>
   <span class="sp"></span>${canWrite?(sample?`<label class="btn sm" for="nd-img">${ED.ndAI?esc(ED.ndAI):'Đọc ảnh biên bản'}</label><input type="file" id="nd-img" accept="image/*" multiple hidden ${ED.ndAI?'disabled':''}>`:`<button class="btn sm" disabled title="Cần cho phép trang hỏi Claude để đọc ảnh">Đọc ảnh biên bản</button>`):''}</div>
  <div class="meta">PO <span class="num">${esc(po.soPO)}</span> · ${esc(po.ncc)} · nguồn: ${esc(N.nguon)}</div>
  ${N.lines.length?`<div class="tbl-wrap"><table class="t nh-tbl"><colgroup><col><col class="c-dvt"><col class="c-n"><col class="c-q"><col class="c-g"><col class="c-note"><col class="c-x"></colgroup><thead><tr><th>Mặt hàng</th><th>ĐVT</th><th class="r">SL đặt</th><th class="r">Thực nhận</th><th><span class="sr">Gạch bỏ</span></th><th><span class="sr">Ghi chú</span></th><th><span class="sr">Xóa</span></th></tr></thead><tbody>${rows}</tbody></table></div>`:'<div class="empty"><b>Phiếu chưa có dòng</b>Chọn đơn hàng, đọc ảnh biên bản, hoặc thêm mã từ PO bên dưới.</div>'}
  ${canWrite?`<div class="toolbar savebar"><select id="nd-add" aria-label="Thêm mã từ PO"><option value="">+ Thêm mã từ PO…</option>${addable.map(m=>`<option value="${m.ma}">${m.ma} · ${esc(m.ten)}</option>`).join('')}</select>
   <span class="meta">${N.lines.length} mã · <b class="num">${VND(tong)} đ</b>${N.dh&&Math.abs(tong-tDat)>=1?` · <span class="${tong>tDat?'bad-t':''}">${tong>tDat?'+':''}${VND(Math.round(tong-tDat))} so với đơn</span>`:''}</span><span class="sp"></span>
   <button class="btn" data-act4="nd-reset">Làm lại</button><button class="btn primary" data-act4="nd-save" ${N.lines.length?'':'disabled'}>Lưu phiếu nhận</button></div>`:''}`;
}
async function readImages(files){
  if(!sample){toast('Trang chưa được phép hỏi Claude.');return;}
  const N=nhanDraft(), po=poById(N.poId), M=poMerge(po);
  let lim=null; try{ lim=await sample.limits(); }catch(e){}
  if(lim&&!lim.images){ toast('Phiên xem này chưa cho phép gửi ảnh cho Claude.'); return; }
  ED.ndAI='Claude đang đọc ảnh…'; render();
  const prompt=`Đây là ảnh biên bản giao hàng của nhà cung cấp ${po.ncc} cho Bếp Vinmec Ocean Park 2 (PO ${po.soPO}).
Đọc từng dòng hàng và ghép với DANH SÁCH DÒNG PO bên dưới (theo mã hàng nếu có, không thì theo tên). Quy tắc:
- "sl" = số ở ô Thực nhận. Nếu ô Thực nhận để trống thì để null (sẽ lấy số đặt/số PO trên biên bản).
- "sl_dat" = số đặt / SL PO in trên biên bản nếu có.
- Dòng bị GẠCH BỎ (không giao) thì "gach_bo": true.
- Nếu không chắc mã thì "can_kiem_tra": true và ghi lý do ngắn.
Trả JSON đúng dạng: {"ngay":"dd/mm/yyyy hoặc null","lines":[{"ma":"THP…","sl":12.5,"sl_dat":12,"gach_bo":false,"can_kiem_tra":false,"ghi_chu":""}]}
DANH SÁCH DÒNG PO (mã|tên|ĐVT):
${M.map(m=>`${m.ma}|${m.ten}|${m.dvt}`).join('\n')}`;
  try{
    const res=await SJ(prompt,{images:[...files],modelTier:'default'});
    let n=0,chk=0; const ok=new Set(M.map(m=>m.ma));
    (res&&res.lines||[]).forEach(r=>{ const ma=String(r.ma||'').toUpperCase(); if(!ok.has(ma)) {chk++;return;} const m=M.find(x=>x.ma===ma);
      const slDat=r.sl_dat!=null?+r.sl_dat:null, sl=r.gach_bo?0:(r.sl!=null?+r.sl:(slDat!=null?slDat:0));
      const ex=N.lines.find(l=>l.ma===ma); const note=[r.can_kiem_tra?'⚠ cần kiểm tra':'',r.ghi_chu||''].filter(Boolean).join(' · ');
      if(ex){ ex.sl=sl; ex.gach=!!r.gach_bo; ex.note=note; if(ex.slDat==null) ex.slDat=slDat; } else N.lines.push({ma,ten:m.ten,dvt:m.dvt,g:m.g,slDat,sl,gach:!!r.gach_bo,note}); n++; if(r.can_kiem_tra) chk++; });
    const dd=res&&res.ngay&&String(res.ngay).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if(dd) N.d=`${dd[3]}-${dd[2].padStart(2,'0')}-${dd[1].padStart(2,'0')}`;
    N.nguon=(N.nguon||'')+' + ảnh ('+files.length+')'; N.anh=files.length; ED.ndAI=null; render(); toast(`Claude đọc được ${n} dòng${chk?`, ${chk} dòng cần kiểm tra`:''}. Soát lại rồi Lưu phiếu.`);
  }catch(e){ ED.ndAI=null; render(); toast(e&&e.code==='rate_limited'?'Claude đang bận, thử lại sau.':'Không đọc được ảnh: '+((e&&e.message)||e)); }
}

/* ---- 4.3 Lịch sử nhận hàng ---- */
function vNls(){
  const L0=pnOf(S.thang).slice().reverse(), D=dhOf(S.thang).slice().reverse(), v=ED.nlsV||'pn', nf=ED.nlsN||'', f=ED.nlsF||'';
  const stOf=p=>lockedDay(p.ncc,p.d)?'khop':p.tt;
  const nccs=[...new Set([...L0,...D].map(x=>x.ncc))].sort((a,b)=>a.localeCompare(b,'vi'));
  const cnt={}; L0.filter(p=>!nf||p.ncc===nf).forEach(p=>{const s=stOf(p); cnt[s]=(cnt[s]||0)+1;});
  const L=L0.filter(p=>(!nf||p.ncc===nf)&&(!f||stOf(p)===f)), DD=D.filter(d=>!nf||d.ncc===nf);
  const FL={nhan:'Chờ MIGO',xuat:'Đã xuất macro',migo:'Đã MIGO',khop:'Đã khớp NCC'};
  const body=L.map(p=>{const open=ED.pnOpen===p.id, lk=lockedDay(p.ncc,p.d);
    return `<tr><td class="exp"><button class="x" data-act4="pn-open" data-id="${p.id}" aria-expanded="${open}" aria-label="${open?'Ẩn':'Xem'} chi tiết phiếu ${esc(p.ma)}">${open?'▾':'▸'}</button></td>
    <td class="dg"><b class="num">${esc(p.ma)}</b><small>${dm(p.d)} · ${esc(p.ncc)} · PO <span class="num">${esc(p.soPO)}</span></small></td>
    <td class="r tien"><b class="num">${VND(p.tong)}</b><small>${p.lines.length} dòng</small></td>
    <td class="tt"><span class="pill ${TTPN[p.tt][1]}">${TTPN[p.tt][0]}</span>${lk?' <span class="pill ok">đã khớp NCC</span>':''}${p.migo?`<small class="num">MIGO ${esc(p.migo)}</small>`:''}${p.lines.some(l=>l.over>0)?'<small class="bad-t">có phần vượt PO</small>':''}</td>
    <td class="meta when">Lập ${tstr(p.luc)} · ${esc(nm(p.ai))}<br>${esc(p.nguon||'')}</td>
    <td class="acts"><div class="acts-in">${!lk&&canWrite?`<button class="btn sm danger-text" data-act4="pn-huy" data-id="${p.id}">${ED.pnHuy===p.id?'Bấm lần nữa để hủy':'Hủy'}</button>`:''}</div></td></tr>
    ${open?`<tr class="detail"><td colspan="6"><div class="tbl-wrap" style="max-height:260px" aria-label="Chi tiết phiếu ${esc(p.ma)}"><table class="t"><thead><tr><th>Mã</th><th>Tên</th><th class="r">Đặt</th><th class="r">Thực nhận</th><th>Dòng PO</th><th class="r">Vượt PO</th><th>Ghi chú</th></tr></thead><tbody>${p.lines.map(l=>`<tr><td class="num">${l.ma}</td><td class="wrap">${esc(l.ten)}</td><td class="r num">${l.slDat!=null?Q(l.slDat):''}</td><td class="r num">${l.gach?'<span class="pill">gạch bỏ</span>':Q(l.sl)}</td><td class="num">${(l.alloc||[]).map(a=>a.item+': '+Q(a.sl)).join(', ')}</td><td class="r num">${l.over>0?Q(l.over):''}</td><td>${esc(l.note||'')}</td></tr>`).join('')}</tbody></table></div>
      <ol class="meta" style="margin:6px 0 0;padding-left:18px">${(p.hist||[]).map(h=>`<li>${tstr(h.ts)} · ${(TTPN[h.tt]||[h.tt])[0]} · ${esc(nm(h.ai))}${h.note?' · '+esc(h.note):''}</li>`).join('')}</ol></td></tr>`:''}`;}).join('');
  const dbody=DD.map(d=>`<tr><td class="exp"></td><td class="dg"><b>${dm(d.d)} · ${esc(d.ncc)}</b><small>PO <span class="num">${esc(d.soPO)}</span></small></td><td class="r tien"><b class="num">${VND(d.tong)}</b><small>${d.lines.length} mã</small></td>
    <td class="tt">${d.pn?'<span class="pill ok">đã có phiếu</span>':'<span class="pill run">chưa nhận</span>'}</td><td class="meta when">Đặt ${tstr(d.luc)}</td>
    <td class="acts"><div class="acts-in">${!d.pn?`<button class="btn sm" data-act="go-next" data-tab="nhan">Nhận ở 4.2</button>`:''}${canWrite&&!d.pn?`<button class="btn sm danger-text" data-act4="dh-huy" data-id="${d.id}">${ED.dhHuy===d.id?'Bấm lần nữa để hủy':'Hủy'}</button>`:''}</div></td></tr>`).join('');
  const lk=L0.filter(p=>lockedDay(p.ncc,p.d)).length, cho=D.filter(d=>!d.pn).length;
  const cols='<colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"><col class="c-when"><col class="c-acts"></colgroup>';
  return head2('4.3 Lịch sử nhận hàng','Mọi đơn đặt hàng và phiếu nhận trong tháng, kèm phân bổ dự kiến vào dòng PO. Phiếu thuộc kỳ đã chốt khớp với NCC bị khóa, không hủy được.','<span class="pill">Theo dõi</span>')+
  (L0.length||D.length?`<div class="kpis"><div class="kv big"><b>Đã nhận tháng ${mmY(S.thang)}</b><div class="num">${TR(L0.reduce((a,p)=>a+p.tong,0))}</div><small>${L0.length} phiếu nhận</small></div>
   <div class="kv"><b>Đã chốt khớp NCC</b><div class="num">${lk} / ${L0.length} phiếu</div><small>${L0.length-lk} phiếu chưa đối soát</small></div>
   <div class="kv"><b>Đơn chưa nhận</b><div class="num">${cho} đơn</div><small>trên ${D.length} đơn đặt hàng</small></div></div>
  <div class="toolbar"><div class="seg"><button data-act4="nls-v" data-v="pn" aria-pressed="${v==='pn'}">Phiếu nhận ${L0.length}</button><button data-act4="nls-v" data-v="dh" aria-pressed="${v==='dh'}">Đơn đặt hàng ${D.length}</button></div>
   <select id="nls-ncc" aria-label="Lọc theo NCC"><option value="">Tất cả NCC</option>${nccs.map(n=>`<option ${n===nf?'selected':''}>${esc(n)}</option>`).join('')}</select>
   ${v==='pn'&&Object.keys(cnt).length>1?`<div class="seg"><button data-act4="nls-f" data-f="" aria-pressed="${!f}">Tất cả</button>${Object.keys(FL).filter(k=>cnt[k]).map(k=>`<button data-act4="nls-f" data-f="${k}" aria-pressed="${f===k}">${FL[k]} ${cnt[k]}</button>`).join('')}</div>`:''}</div>
  ${v==='pn'?(L.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl">${cols}<thead><tr><th><span class="sr">Chi tiết</span></th><th>Phiếu nhận</th><th class="r">Tiền</th><th>Trạng thái</th><th>Lập / nguồn</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>${body}</tbody></table></div>`:'<div class="empty"><b>Không có phiếu nhận nào khớp bộ lọc</b></div>')
   :(DD.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl">${cols}<thead><tr><th><span class="sr">Chi tiết</span></th><th>Đơn đặt hàng</th><th class="r">Tiền</th><th>Trạng thái</th><th>Đặt lúc</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>${dbody}</tbody></table></div>`:'<div class="empty"><b>Không có đơn nào khớp bộ lọc</b></div>')}`
  :'<div class="empty"><b>Chưa có đơn hay phiếu nhận nào</b>Đặt hàng ở 4.1, khi hàng về thì ghi nhận ở 4.2.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-tab="dat">Mở 4.1</button></div></div>');
}

/* ---- sự kiện bước 4 ---- */
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act4]'); if(!t) return; const a=t.dataset.act4, id=t.dataset.id;
  try{
  if(a==='dat-save'||a==='dat-copy'){ const p=poById(ED.datPO), M=poMerge(p); const L=M.filter(m=>ED.datQ[m.ma]>0).map(m=>({ma:m.ma,ten:m.ten,dvt:m.dvt,g:m.g,sl:ED.datQ[m.ma],items:m.items}));
    if(a==='dat-copy'){ try{ await navigator.clipboard.writeText(datText(p,ED.datD,L)); toast('Đã copy đơn hàng.'); }catch(err){ toast('Không copy được.'); } return; }
    if(ED.datD.slice(0,7)!==p.thang){ toast('Ngày giao phải nằm trong tháng của PO ('+mmY(p.thang)+').'); return; }
    const did=p.thang+'-'+ED.datD.slice(8)+'-'+slug(p.ncc)+'-'+Date.now().toString(36), tong=Math.round(L.reduce((s,l)=>s+l.sl*l.g,0));
    await store.set('donhang/'+did,{thang:p.thang,d:ED.datD,ncc:p.ncc,poId:p.id,soPO:p.soPO,lines:L.map(({items,...l})=>l),tong,tt:'dat',pn:'',luc:Date.now(),ai:uid});
    await log('4.1',`Đơn hàng ${dm(ED.datD)} – ${p.ncc} (PO ${p.soPO}): ${L.length} mã, ${TR(tong)}`); ED.datQ={}; ED.datOnly=false; ED.ndDH=null; ED.nd=null; render(); toast('Đã lưu đơn hàng. Khi hàng về, ghi nhận ở 4.2.'); }
  else if(a==='dh-copy'){ const d=S.dh[id], p=poById(d.poId), M=new Map(poMerge(p).map(m=>[m.ma,m])); try{ await navigator.clipboard.writeText(datText(p,d.d,d.lines.map(l=>({...l,items:(M.get(l.ma)||{items:[]}).items})))); toast('Đã copy đơn '+dm(d.d)+'.'); }catch(err){ toast('Không copy được.'); } }
  else if(a==='dat-only'){ ED.datOnly=!!t.dataset.v; render(); }
  else if(a==='nn-open'){ ED.ndNote=+t.dataset.i; render(); const el=$('#nn-'+t.dataset.i); if(el) el.focus(); }
  else if(a==='ok-x'){ ED.ndOk=null; render(); }
  else if(a==='nls-v'){ ED.nlsV=t.dataset.v; render(); }
  else if(a==='nls-f'){ ED.nlsF=t.dataset.f; render(); }
  else if(a==='nl-del'){ ED.nd.lines.splice(+t.dataset.i,1); ED.ndNote=null; render(); }
  else if(a==='nd-reset'){ ED.nd=null; ED.ndOk=null; render(); }
  else if(a==='nd-save'){ const N=ED.nd; if(!N||ED.ndBusy) return; const po=poById(N.poId); if(N.d.slice(0,7)!==po.thang){toast('Ngày giao phải nằm trong tháng của PO.');return;} if(lockedDay(po.ncc,N.d)){toast('Kỳ này của '+po.ncc+' đã chốt khớp; mở lại ở 5.3 trước.');return;}
    const used=poUsed(po.id); const lines=N.lines.map(l=>{const q=l.gach?0:Math.max(0,+l.sl||0); const r=allocItems(po,l.ma,q,used); return {...l,sl:q,alloc:r.alloc,over:r.over};});
    const tong=Math.round(lines.reduce((s,l)=>s+l.sl*l.g,0)), seq=pnOf(po.thang).filter(p=>p.d===N.d).length+1;
    ED.ndBusy=true; t.disabled=true; const ma=`NH${N.d.slice(8)}${N.d.slice(5,7)}-${slug(po.ncc).replace(/-/g,'').slice(0,10).toUpperCase()}-${seq}`, pid=po.thang+'-'+Date.now().toString(36);
    await store.set('phieunhan/'+pid,{thang:po.thang,ma,d:N.d,ncc:po.ncc,poId:po.id,soPO:po.soPO,dh:N.dh,lines,tong,nguon:N.nguon,tt:'nhan',migo:'',luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'nhan',ai:uid,note:N.nguon}]});
    if(N.dh){ const d=S.dh[N.dh]; const x={...d,pn:pid,tt:'nhan'}; delete x.id; await store.set('donhang/'+N.dh,x); }
    const ov=lines.filter(l=>l.over>0); await log('4.2',`Phiếu nhận ${ma}: ${lines.length} mã, ${TR(tong)}${ov.length?`; vượt PO ${ov.length} mã (${ov.map(l=>l.ma).join(', ')})`:''}`);
    ED.nd=null; ED.ndDH=null; ED.ndNote=null; ED.ndOk=ma; ED.ndBusy=false; render(); toast(ov.length?'Đã lưu phiếu. Phần vượt PO sẽ xử lý cuối chu kỳ (bước 6).':'Đã lưu phiếu nhận.'); }
  else if(a==='pn-open'){ ED.pnOpen=ED.pnOpen===id?null:id; render(); }
  else if(a==='pn-huy'){ const p=S.pn[id]; if(ED.pnHuy!==id){ED.pnHuy=id;toast('Bấm Hủy lần nữa để hủy phiếu '+p.ma+'.');render();return;} ED.pnHuy=null;
    const x={...p,tt:'huy',hist:[...(p.hist||[]),{ts:Date.now(),tt:'huy',ai:uid,note:'Hủy phiếu'}]}; delete x.id; await store.set('phieunhan/'+id,x);
    if(p.dh&&S.dh[p.dh]){ const d={...S.dh[p.dh],pn:'',tt:'dat'}; delete d.id; await store.set('donhang/'+p.dh,d); } await log('4.3',`Hủy phiếu nhận ${p.ma}`); }
  else if(a==='dh-huy'){ if(ED.dhHuy!==id){ED.dhHuy=id;toast('Bấm Hủy lần nữa để hủy đơn hàng.');render();return;} ED.dhHuy=null; const d={...S.dh[id],tt:'huy'}; delete d.id; await store.set('donhang/'+id,d); await log('4.3',`Hủy đơn hàng ${dm(d.d)} – ${d.ncc}`); }
  }catch(err){ ED.ndBusy=false; toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target;
  if(t.dataset.dq){ ED.datQ[t.dataset.dq]=Math.max(0,pnum(t.value)); clearTimeout(vDat.h); vDat.h=setTimeout(render,700); }
  else if(t.dataset.nq!=null){ ED.nd.lines[+t.dataset.nq].sl=Math.max(0,pnum(t.value)); clearTimeout(vNhan.h); vNhan.h=setTimeout(render,700); }
  else if(t.dataset.nn!=null){ ED.nd.lines[+t.dataset.nn].note=t.value; }
  else if(t.id==='dat-find'){ ED.datF=t.value; const q=fold(t.value); document.querySelectorAll('.dat-tbl tbody tr').forEach(r=>{ const inp=r.querySelector('[data-dq]'); r.hidden=(ED.datOnly&&!(pnum(inp.value)>0))||(!!q&&!fold(r.dataset.k).includes(q)); }); }
});
document.addEventListener('change',e=>{ const t=e.target;
  if(t.id==='dat-po'){ ED.datPO=t.value; ED.datQ={}; render(); }
  else if(t.id==='dat-d'){ ED.datD=t.value; render(); }
  else if(t.id==='nd-dh'){ ED.ndDH=t.value; ED.nd=null; ED.ndOk=null; render(); }
  else if(t.id==='nd-po'){ ED.ndPO=t.value; ED.nd=null; render(); }
  else if(t.id==='nd-d'){ ED.nd.d=t.value; }
  else if(t.dataset.ngach!=null){ const l=ED.nd.lines[+t.dataset.ngach]; l.gach=t.checked; render(); }
  else if(t.id==='nd-add'&&t.value){ const po=poById(ED.nd.poId), m=poMerge(po).find(x=>x.ma===t.value); ED.nd.lines.push({ma:m.ma,ten:m.ten,dvt:m.dvt,g:m.g,slDat:null,sl:0,gach:false,note:''}); render(); }
  else if(t.id==='nd-img'&&t.files&&t.files.length){ readImages([...t.files]); }
  else if(t.id==='nls-ncc'){ ED.nlsN=t.value; ED.nlsF=''; render(); }
});

