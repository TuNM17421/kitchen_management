/* ---------- BẢNG ĐIỀU PHỐI AGENT ----------
   Mỗi bước có 1 agent (model Claude riêng). Agent chỉ ĐỌC dữ liệu và ĐỀ XUẤT; mọi ghi file / lệnh SAP / email
   chỉ chạy khi người dùng bấm Duyệt, bằng chính các hàm của giao diện. Agent chạy nền theo lịch ghi đề xuất vào sheet dexuat. */
const TIERS={quick:'Nhanh',default:'Tiêu chuẩn',complex:'Mạnh'};
const AGENTS=[
  {id:'dieuphoi',ten:'Điều phối',buoc:0,tier:'complex',vai:'Xem toàn bộ quy trình tháng, chỉ ra việc nào đang tắc, agent nào nên chạy tiếp, lệnh trạm SAP nào lỗi.'},
  {id:'dutru',ten:'Dự trù',buoc:1,tier:'complex',vai:'Soát dự trù tháng: mặt hàng tăng/giảm bất thường so với tháng trước, ngân sách so tham chiếu, thiếu giá, thiếu NCC.'},
  {id:'pr',ten:'Lập PR',buoc:2,tier:'default',vai:'Soát đợt PR: NCC thiếu mã SAP, đợt chưa gửi trạm, PR chờ duyệt quá ngày, đề xuất gửi lệnh tạo PR.'},
  {id:'po',ten:'Tạo PO',buoc:3,tier:'default',vai:'Soát PO: PR đã duyệt chưa lên PO, PO chưa có số, PO đã sửa chờ cập nhật SAP, dòng nhận vượt/ngoài PO.'},
  {id:'nhan',ten:'Nhận hàng',buoc:4,tier:'quick',vai:'Soát số nhận gần đây: ngày thiếu phiếu nhận, đơn đặt chưa nhận, số lượng lạ.'},
  {id:'doisoat',ten:'Đối soát NCC',buoc:5,tier:'complex',vai:'Theo chu kỳ từng NCC: kỳ đã hết hạn chưa gửi bảng kê / chưa chốt khớp, đề xuất soạn bảng kê.'},
  {id:'chinhpo',ten:'Chỉnh PO',buoc:6,tier:'default',vai:'Ở mốc ngày 15 / cuối tháng: PO cần tăng/giảm, mã cần PR bổ sung (1 lần duy nhất mỗi mốc).'},
  {id:'migo',ten:'MIGO',buoc:7,tier:'quick',vai:'Kỳ đã khớp chưa lập MIGO, chứng từ chưa nhập kho, PO còn sửa chặn MIGO.'},
  {id:'dntt',ten:'ĐNTT',buoc:8,tier:'quick',vai:'Kỳ đủ điều kiện chưa lập ĐNTT, thiếu hóa đơn, ĐNTT chưa tạo file / chưa trình ký.'}];
const HANHDONG={
  mo:'Mở bước', lenh_pr:'Gửi trạm: tạo PR', lenh_po:'Gửi trạm: tạo PO', lenh_posua:'Gửi trạm: sửa PO (ME22N)', lap_migo:'Lập chứng từ MIGO',
  lenh_migo:'Gửi trạm: MIGO', lap_dntt:'Lập ĐNTT', lenh_dntt:'Gửi trạm: tạo file ĐNTT', bangke:'Mở bảng kê để tạo nháp email', chinh_po:'Áp dụng tăng/giảm PO tại mốc', pr_bosung:'Tạo PR bổ sung tại mốc'};
S.dx={}; S.ar={}; const AG_RUN=new Set();
const agCfg=()=>MEM.get('cauhinh/agent')||{};
const agTier=a=>(agCfg()[a.id]&&agCfg()[a.id].tier)||a.tier;
const cut=(a,n)=>a.length>n?[...a.slice(0,n),{conLai:a.length-n}]:a;
const td0=()=>todayVN();
/* ---- ngữ cảnh từng agent: số liệu gọn, có id để đề xuất hành động ---- */
function agCtx(id){
  const t=S.thang, B=batchesOf(t), P=poOf(t), M=migoAll().filter(m=>m.thang===t), L=Object.values(S.lenh||{}).sort((a,b)=>b.taoLuc-a.taoLuc);
  const base={thang:t,homNay:td0(),tram:{trucTuyen:tramOnline(),may:S.tram&&S.tram.may||''},cheDoLuu:{PR:sapLuuThat('PR'),PO:sapLuuThat('PO'),POSUA:sapLuuThat('POSUA'),MIGO:sapLuuThat('MIGO')}};
  const bS=b=>({id:b.id,nhom:b.nhom,loai:b.loai,doiTuong:b.dt,dong:b.lines.length,tien:b.tong,tt:b.tt,soPR:b.soPR||'',dangGuiTram:lenhBusy(b.lenh),ngayTao:tstr(b.luc),thieuMaNCC:[...new Set(b.lines.map(l=>l.ncc).filter(n=>!hasLifnr(n)))]});
  const pS=p=>({id:p.id,ncc:p.ncc,soPO:p.soPO||'',tt:p.tt,dong:p.lines.length,tien:p.tong,tienSAP:p.tongSAP,canSuaSAP:!!p.sua,dangGuiTram:lenhBusy(p.lenh),coAnhChupSAP:!!p.sapLines});
  if(id==='dieuphoi'){ const R=dnRows();
    return {...base,pr:{dot:B.length,choGui:B.filter(b=>b.tt==='nhap'||b.tt==='xuat').length,choDuyet:B.filter(b=>b.tt==='sap').length,daDuyet:B.filter(b=>b.tt==='duyet').length},
      po:{tong:P.length,chuaSo:P.filter(p=>!p.soPO).length,canSuaSAP:P.filter(p=>p.sua).length,prChuaLenPO:poPending().length},
      doiSoat:dsNccs().map(n=>({ncc:n,ky:kyList(t,n).map(k=>({k:k.k,den:k.to,tt:(dsDoc(n,k.k)||{}).tt||'chua'}))})),
      migo:{cho:M.filter(m=>m.tt==='cho'||m.tt==='xuat').length,xong:M.filter(m=>m.tt==='migo').length,kyChoLap:mgCandidates().length},
      dntt:{duDieuKien:R.filter(r=>r.ready&&!r.dn).length,daLap:R.filter(r=>r.dn&&r.dn.tt!=='huy').length},
      lenhGanDay:L.slice(0,10).map(x=>({loai:x.loai,moTa:x.moTa,tt:x.tt,luc:tstr(x.taoLuc),loi:x.tt==='loi'?String(x.log||'').slice(0,300):''}))}; }
  if(id==='dutru'){ const dt=S.dutru; if(!dt) return {...base,chuaCoDuTru:true};
    const prev=new Map(); Object.values(S.prdot||{}).filter(b=>b.thang<t).forEach(b=>b.lines.forEach(l=>prev.set(l.ma,(prev.get(l.ma)||0)+l.sl)));
    const lines=(dt.lines||[]).map(l=>{const x=IT.get(l.ma)||{}; return {ma:l.ma,ten:x.ten||l.ten||'',ncc:l.ncc||'',dvt:x.dvt||'',sl:+l.sl||0,gia:+l.g||(x.ncc?priceOf(x,l.ncc):0)||0,thangTruocPR:prev.get(l.ma)||null};});
    lines.forEach(l=>l.tien=Math.round(l.sl*l.gia)); lines.sort((a,b)=>b.tien-a.tien);
    return {...base,tongNganSach:dt.tong,daChot:locked(),soDong:lines.length,thieuGia:lines.filter(l=>!l.gia).map(l=>l.ma).slice(0,30),canhBaoCu:dt.warn||[],dongLonNhat:cut(lines,60)}; }
  if(id==='pr') return {...base,daLuuPhanBo:!!S.pb,dot:B.map(bS)};
  if(id==='po'){ const C=poRecvCompute(); return {...base,poChoLap:poPending().map(o=>({ncc:o.ncc,dong:o.lines.length,tien:o.tong,coMaNCC:hasLifnr(o.ncc)})),po:P.map(pS),
    dongBatThuong:cut(C.rows.filter(r=>['ngoai','vuot','sap_vuot'].includes(r.st)).map(r=>({ncc:r.po?r.po.ncc:r.ncc,soPO:r.po?r.po.soPO:'',ma:r.ma,ten:r.ten,slPO:r.sl,daNhan:+r.sR.toFixed(3),tt:r.st,canThem:r.need})),40)}; }
  if(id==='nhan'){ const rv=recvRows(); const from=new Date(Date.now()-10*864e5).toISOString().slice(0,10);
    const g={}; rv.filter(r=>r.d>=from).forEach(r=>{const k=r.d+'|'+r.ncc; g[k]=(g[k]||0)+1;});
    const dh=Object.values(S.dh||{}).filter(d=>d.thang===t&&d.tt==='dat'&&d.d<td0()).map(d=>({ngay:d.d,ncc:d.ncc,soPO:d.soPO,dong:d.lines.length}));
    return {...base,phieuTheoNgayNCC:Object.entries(g).map(([k,n])=>({ngay:k.split('|')[0],ncc:k.split('|')[1],dong:n})).sort((a,b)=>a.ngay<b.ngay?1:-1).slice(0,80),donDatChuaNhan:cut(dh,40)}; }
  if(id==='doisoat') return {...base,ncc:dsNccs().map(n=>({ncc:n,chuKy:ckOf(n),email:!!((NCCINFO[n]||{}).email||[]).length,ky:kyList(t,n).map(k=>{const d=dsDoc(n,k.k)||{};return {k:k.k,tu:k.from,den:k.to,daHet:k.to<td0(),soDongNhan:kyRows(n,k).length,tt:d.tt||'chua',guiLuc:d.guiLuc?tstr(d.guiLuc):''};})}))};
  if(id==='chinhpo') return {...base,moc:mocList().map(m=>{const R=cpoCompute(m); return {moc:m.m,ten:m.ten,ngay:m.d,daDen:m.d<=td0(),daTaoPRBS:!!(S.prbs||{})[t+'_m'+m.m],
      dong:cut(R.filter(r=>r.act!=='khop').map(r=>({ncc:r.ncc,ma:r.ma,ten:r.ten,daNhan:r.r,slPO:r.q,prDuyet:r.pr,de:r.act,tang:r.tang,bs:r.bs,giam:r.giam,kyDaKhop:r.ok})),40)};})};
  if(id==='migo') return {...base,kyChoLap:mgCandidates().map(c=>({ncc:c.po.ncc,soPO:c.po.soPO,ky:c.ky.k,dong:c.lines.length,tien:c.tong,poThieu:c.short.length,poCanSua:!!c.po.sua})),
    chungTu:M.filter(m=>m.tt!=='huy').map(m=>({id:m.id,ma:m.ma,ncc:m.ncc,soPO:m.soPO,ngay:m.ngay,tt:m.tt,soMIGO:m.soMIGO||'',dangGuiTram:lenhBusy(m.lenh),poCanSua:!!(S.po[m.poId]||{}).sua}))};
  if(id==='dntt') return {...base,ky:dnRows().map(r=>({ncc:r.po.ncc,soPO:r.po.soPO,ky:r.ky.k,tien:r.tien,dieuKien:r.ck,duDieuKien:r.ready,dn:r.dn?{id:r.dn.id,tt:r.dn.tt,coFile:!!r.dn.fileDNTT,dangGuiTram:lenhBusy(r.dn.lenh)}:null,hoaDon:r.hd?r.hd.soHD:''}))};
  return base; }
function agPrompt(a,ctx){
  return `Bạn là agent "${a.ten}" trong hệ thống mua hàng Bộ phận Bếp (Vinmec Ocean Park 2). Quy trình: dự trù → PR theo tháng → PO mỗi NCC/tháng → đặt & nhận hàng (số nhận là chuẩn) → đối soát NCC theo chu kỳ 15 ngày hoặc tháng → chỉnh PO + 1 PR bổ sung duy nhất mỗi mốc → MIGO cuối kỳ (PO = MIGO) → ĐNTT theo PO. PR phải được duyệt trong ngày; hàng tươi giá theo tháng, đồ khô theo quý.
Nhiệm vụ: ${a.vai}
Bạn CHỈ đọc và đề xuất; người dùng duyệt từng đề xuất. Không bịa id: chỉ dùng id có trong dữ liệu. Viết tiếng Việt ngắn gọn, số có dấu chấm ngăn nghìn.
Các loại hành động được phép (field "hanhDong"):
- {"loai":"mo","buoc":1-8} – chỉ mở bước để người dùng tự xem
- {"loai":"lenh_pr","batches":[id đợt PR tt nhap/xuat]}
- {"loai":"lenh_po","po":"id PO tt cho"}
- {"loai":"lenh_posua","po":"id PO canSuaSAP"}
- {"loai":"lap_migo"} – lập chứng từ cho các kỳ đã khớp
- {"loai":"lenh_migo","migo":[id chứng từ tt cho/xuat]}
- {"loai":"lap_dntt"} – lập ĐNTT cho các kỳ đủ điều kiện
- {"loai":"lenh_dntt","dn":[id ĐNTT]}
- {"loai":"bangke","ncc":"tên NCC","ky":số kỳ}
- {"loai":"chinh_po","moc":"15"|"cuoi"}
- {"loai":"pr_bosung","moc":"15"|"cuoi"}
Không đề xuất hành động khi không cần. Mức độ: "cao" (chặn quy trình / sai tiền), "vua", "thap".
Trả về JSON đúng dạng: {"tomTat":"1-3 câu","canhBao":[{"mucDo":"cao|vua|thap","noiDung":"..."}],"deXuat":[{"tieuDe":"...","lyDo":"...","mucDo":"cao|vua|thap","hanhDong":{...}}]}
Dữ liệu hiện tại (JSON):
${JSON.stringify(ctx)}`; }
async function chayAgent(a){
  if(!sample){ toast('Trang chưa được phép dùng Claude.'); return; } if(AG_RUN.has(a.id)) return; AG_RUN.add(a.id); renderAgents();
  try{ const ctx=agCtx(a.id), tier=agTier(a);
    const r=await cmd(`Agent ${a.ten} (${TIERS[tier]})`,()=>sample.json(agPrompt(a,ctx),{modelTier:tier,cache:false}));
    const luc=Date.now(), dx=(r&&Array.isArray(r.deXuat)?r.deXuat:[]).filter(d=>d&&d.tieuDe).slice(0,12);
    await store.set('agentchay/'+a.id+'_'+S.thang,{agent:a.id,thang:S.thang,luc,tier,nguon:'trang',tomTat:String(r&&r.tomTat||''),canhBao:(r&&r.canhBao||[]).slice(0,12),nDeXuat:dx.length});
    for(const [i,d] of dx.entries()){ const hd=d.hanhDong&&HANHDONG[d.hanhDong.loai]?d.hanhDong:null;
      await store.set('dexuat/'+a.id+'-'+luc.toString(36)+i,{agent:a.id,thang:S.thang,luc,nguon:'trang',tieuDe:String(d.tieuDe).slice(0,200),lyDo:String(d.lyDo||'').slice(0,600),mucDo:d.mucDo||'vua',hanhDong:hd,tt:'cho'}); }
    await log(String(a.buoc||'0'),`Agent ${a.ten}: ${dx.length} đề xuất – ${String(r&&r.tomTat||'').slice(0,120)}`);
  }catch(e){ toast(`Agent ${a.ten} lỗi: `+((e&&(e.message||e.code))||e)); }
  finally{ AG_RUN.delete(a.id); renderAgents(); } }
function moBuoc(n,tab){ ED.mode='quytrinh'; S.step=n; S.tab=tab||curSub()[0][2]; render(); window.scrollTo({top:0,behavior:'smooth'}); }
async function clickSau(n,tab,sel,prep){ moBuoc(n,tab); if(prep) prep(); render(); await new Promise(r=>setTimeout(r,60)); const b=document.querySelector(sel); if(!b||b.disabled) throw new Error('Nút không khả dụng (dữ liệu đã thay đổi?)'); b.click(); await new Promise(r=>setTimeout(r,400)); }
async function thucHien(d){
  const h=d.hanhDong||{}; const has=(col,id)=>!!(S[col]||{})[id];
  switch(h.loai){
    case 'mo': moBuoc(+h.buoc||1); return 'Đã mở bước '+h.buoc;
    case 'lenh_pr': { const ids=(h.batches||[]).filter(id=>has('prdot',id)); if(!ids.length) throw new Error('Không còn đợt PR hợp lệ'); ED.sel=new Set(ids); await lenhPR(); return 'Đã gửi lệnh tạo PR'; }
    case 'lenh_po': if(!has('po',h.po)) throw new Error('Không thấy PO'); await lenhPO(h.po); return 'Đã gửi lệnh tạo PO';
    case 'lenh_posua': if(!has('po',h.po)) throw new Error('Không thấy PO'); await lenhPOSua(h.po); return 'Đã gửi lệnh sửa PO';
    case 'lap_migo': await clickSau(7,'mg','[data-act7="mg-make"]'); return 'Đã lập chứng từ MIGO';
    case 'lenh_migo': { const ids=(h.migo||[]).filter(id=>has('migo',id)); if(!ids.length) throw new Error('Không còn chứng từ hợp lệ'); await lenhMIGO(ids); return 'Đã gửi lệnh MIGO'; }
    case 'lap_dntt': await clickSau(8,'dn','[data-act8="dn-make"]'); return 'Đã lập ĐNTT';
    case 'lenh_dntt': { const ids=(h.dn||[]).filter(id=>has('dn',id)); if(!ids.length) throw new Error('Không còn ĐNTT hợp lệ'); await lenhDNTT(ids); return 'Đã gửi lệnh tạo file ĐNTT'; }
    case 'bangke': moBuoc(5,'bk'); ED.dsN=h.ncc; ED.dsK=+h.ky||0; render(); return 'Đã mở bảng kê – kiểm tra rồi bấm tạo nháp email';
    case 'chinh_po': await clickSau(6,'cpo','[data-act6="cpo-apply"]',()=>{ED.moc=h.moc;}); return 'Đã áp dụng chỉnh PO';
    case 'pr_bosung': await clickSau(6,'cpo','[data-act6="cpo-bs"]',()=>{ED.moc=h.moc;}); return 'Đã tạo PR bổ sung';
    default: throw new Error('Loại hành động không hỗ trợ'); } }
async function duyetDX(id){ const d=S.dx[id]; if(!d||d.tt!=='cho') return; let kq='', tt='duyet';
  try{ kq=d.hanhDong?await thucHien(d):'Đã ghi nhận'; }catch(e){ kq='Lỗi: '+((e&&e.message)||e); tt='loi'; }
  await store.set('dexuat/'+id,{...stripId(d),tt,ketQua:kq,xuLyLuc:Date.now(),nguoiXuLy:uname}); await log('AG',`Duyệt đề xuất “${d.tieuDe}”: ${kq}`); toast(kq); }
const MD={cao:['Cao','bad'],vua:['Vừa','run'],thap:['Thấp','']};
function renderAgents(){ const el=$('#agentPanel'); if(!el) return; try{applyMode();}catch(e){} if(ED.mode!=='agent') return;
  const dx=Object.values(S.dx||{}).filter(d=>d.thang===S.thang&&d.tt==='cho').sort((a,b)=>({cao:0,vua:1,thap:2}[a.mucDo]??1)-({cao:0,vua:1,thap:2}[b.mucDo]??1)||b.luc-a.luc);
  const done=Object.values(S.dx||{}).filter(d=>d.thang===S.thang&&d.tt!=='cho').sort((a,b)=>(b.xuLyLuc||0)-(a.xuLyLuc||0)).slice(0,6);
  const dem=S.ar['dem_'+S.thang];
  el.innerHTML=`<div class="sh" style="margin:0"><div><h2>Bảng điều phối agent</h2><p>Mỗi bước có một agent Claude riêng (model chọn theo độ khó). Agent đọc dữ liệu và <b>đề xuất</b>; ghi dữ liệu, lệnh trạm SAP, email chỉ chạy khi anh bấm <b>Duyệt</b>. Agent đêm chạy theo lịch mỗi sáng 5:47 và ghi đề xuất vào đây.${dem?` Lần chạy nền gần nhất: ${tstr(dem.luc)}.`:''}</p></div>
    <div class="toolbar"><button class="btn sm primary" data-ag="all" ${sample?'':'disabled'}>Chạy tất cả agent</button></div></div>
  <div class="agents">${AGENTS.map(a=>{const r=S.ar[a.id+'_'+S.thang], run=AG_RUN.has(a.id), n=dx.filter(d=>d.agent===a.id).length;
    return `<div class="ag"><div class="ag-h"><b>${a.buoc?a.buoc+'. ':''}${a.ten}</b>${n?`<span class="pill bad">${n} đề xuất</span>`:''}</div>
      <div class="meta">${esc(a.vai)}</div>
      <div class="ag-f"><select data-agtier="${a.id}" aria-label="Model agent ${a.ten}">${Object.entries(TIERS).map(([k,v])=>`<option value="${k}" ${agTier(a)===k?'selected':''}>${v}</option>`).join('')}</select>
      <button class="btn sm" data-ag="run" data-id="${a.id}" ${run||!sample?'disabled':''}>${run?'<span class="spin"></span> Đang chạy':'Chạy'}</button></div>
      ${r?`<div class="ag-r"><span class="meta num">${tstr(r.luc)} · ${TIERS[r.tier]||''}</span><div>${esc(r.tomTat)}</div>${(r.canhBao||[]).slice(0,3).map(c=>`<div class="meta"><span class="pill ${(MD[c.mucDo]||MD.vua)[1]}">${(MD[c.mucDo]||MD.vua)[0]}</span> ${esc(c.noiDung)}</div>`).join('')}</div>`:'<div class="meta">Chưa chạy tháng này.</div>'}</div>`;}).join('')}</div>
  <h3 class="eyebrow" style="margin:4px 0 0">Đề xuất chờ duyệt (${dx.length})</h3>
  ${dx.length?`<div class="tbl-wrap" style="max-height:340px"><table class="t"><thead><tr><th>Agent</th><th>Mức</th><th>Đề xuất</th><th>Hành động khi duyệt</th><th></th></tr></thead><tbody>
   ${dx.map(d=>{const a=AGENTS.find(x=>x.id===d.agent)||{ten:d.agent==='dem'?'Agent đêm':d.agent};const m=MD[d.mucDo]||MD.vua;
     return `<tr><td>${esc(a.ten)}<div class="meta">${d.nguon==='lich'?'chạy nền':'trên trang'} · ${tstr(d.luc)}</div></td><td><span class="pill ${m[1]}">${m[0]}</span></td><td class="wrap"><b>${esc(d.tieuDe)}</b><div class="meta">${esc(d.lyDo||'')}</div></td>
     <td>${d.hanhDong?esc(HANHDONG[d.hanhDong.loai]||d.hanhDong.loai):'<span class="meta">Chỉ ghi nhận</span>'}</td>
     <td style="white-space:nowrap">${canWrite?`<button class="btn sm primary" data-ag="ok" data-id="${d.id}">Duyệt</button> <button class="btn sm" data-ag="bo" data-id="${d.id}">Bỏ qua</button>`:''}</td></tr>`;}).join('')}</tbody></table></div>`:'<div class="meta">Không có đề xuất nào đang chờ.</div>'}
  ${done.length?`<details><summary class="meta">Đã xử lý gần đây</summary><ul style="margin:6px 0;padding-left:18px;font-size:13px">${done.map(d=>`<li>${esc(d.tieuDe)} – <span class="pill ${d.tt==='loi'?'bad':d.tt==='bo'?'':'ok'}">${d.tt==='duyet'?'Đã duyệt':d.tt==='bo'?'Bỏ qua':'Lỗi'}</span> <span class="meta">${esc(d.ketQua||'')}</span></li>`).join('')}</ul></details>`:''}`; }
store.watchCol('dexuat',a=>{S.dx={};a.forEach(d=>S.dx[d.id]={...d});renderAgents();});
store.watchCol('agentchay',a=>{S.ar={};a.forEach(d=>S.ar[d.id]={...d});renderAgents();});
document.addEventListener('change',async e=>{ const s=e.target.closest('[data-agtier]'); if(!s) return; const c={...agCfg()}; c[s.dataset.agtier]={tier:s.value}; await store.set('cauhinh/agent',c); renderAgents(); });
document.addEventListener('click',async e=>{ const t=e.target.closest('[data-ag]'); if(!t) return; const a=t.dataset.ag;
  if(a==='thu'){ ED.agOff=true; renderAgents(); } else if(a==='mo'){ ED.agOff=false; renderAgents(); }
  else if(a==='mo2'){ moBuoc(+t.dataset.b); }
  else if(a==='run'){ await chayAgent(AGENTS.find(x=>x.id===t.dataset.id)); }
  else if(a==='all'){ for(const x of AGENTS.slice(1)) await chayAgent(x); await chayAgent(AGENTS[0]); }
  else if(a==='ok'){ t.disabled=true; await duyetDX(t.dataset.id); }
  else if(a==='bo'){ const d=S.dx[t.dataset.id]; if(d) await store.set('dexuat/'+d.id,{...stripId(d),tt:'bo',xuLyLuc:Date.now(),nguoiXuLy:uname}); } });

window.__ag={AGENTS,agCtx,chayAgent,duyetDX,S:()=>S};
