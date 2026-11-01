/* ================= BƯỚC 8 · ĐNTT THEO PO ================= */
const SUB8=[['8.1','Hóa đơn & điều kiện','hd'],['8.2','Lập ĐNTT','dn'],['8.3','Lịch sử & thanh toán','dnls']];
const SHEETMAU={"NCC MẪU 23": "NCC MẪU 23", "NCC MẪU 26": "NCC MẪU 26", "NCC MẪU 11": "ĐNTT NCC MẪU 11", "NCC MẪU 06": "ĐNTT NCC MẪU 06", "NCC MẪU 30": "ĐNTT NCC MẪU 30", "NCC MẪU 05": "NCC MẪU 05", "NCC MẪU 02": "ĐNTT NCC MẪU 02", "NCC MẪU 35": "ĐNTT NCC MẪU 35", "NCC MẪU 36": "NCC MẪU 36", "NCC MẪU 18": "ĐNTT NCC MẪU 18", "NCC MẪU 21": "NCC MẪU 21", "NCC MẪU 22": "NCC MẪU 22", "NCC MẪU 24": "NCC MẪU 24", "NCC MẪU 07": "ĐNTT NCC MẪU 07", "NCC MẪU 33": "ĐNTT NCC MẪU 33", "NCC MẪU 12": "ĐNTT NCC MẪU 12", "NCC MẪU 14": "ĐNTT NCC MẪU 14", "NCC MẪU 17": "NCC MẪU 17"};
const TTDN={lap:['Đã lập','run'],xuat:['Đã tạo file ĐNTT','run'],trinh:['Đang trình ký','run'],ketoan:['Đã chuyển kế toán','run'],tt:['Đã thanh toán','ok'],huy:['Đã hủy','bad']};
const dnId=(t,ncc,k)=>`${t}_${slug(ncc)}_k${k}`;
function dnRows(){
  const out=[];
  poOf(S.thang).forEach(po=>{ const kys=kyList(S.thang,po.ncc); const items=new Map(poItems(po).map(l=>[l.item,l]));
    const allMg=migoAll().filter(m=>m.poId===po.id), poQ=new Map(); poItems(po).forEach(l=>poQ.set(l.item,l.sl));
    kys.forEach((ky,ki)=>{ const mg=allMg.filter(m=>m.k===ky.k); const ds=dsDoc(po.ncc,ky.k); if(!mg.length&&!(ds&&ds.tt==='khop')) return;
      let tien=0; const grp={}; mg.forEach(m=>m.lines.forEach(l=>{const it=items.get(l.item)||{g:priceFor(po.ncc,l.ma)}; const v=l.sl*it.g; tien+=v; const x=IT.get(l.ma); const n2=x?x.n2:''; grp[n2]=(grp[n2]||0)+v;}));
      const last=ki===kys.length-1; let poEqMigo=true; if(last){ const used={}; allMg.forEach(m=>m.lines.forEach(l=>used[l.item]=(used[l.item]||0)+l.sl)); poQ.forEach((q,it)=>{ if(Math.abs(q-(used[it]||0))>1e-6) poEqMigo=false; }); }
      const loai=(grp['Đồ khô']||0)>(grp['Thực phẩm tươi sống']||0)?'đồ khô':'thực phẩm tươi sống';
      const bb=new Set(recvFull().filter(r=>r.ncc===po.ncc&&r.d>=ky.from&&r.d<=ky.to&&r.sl>0).map(r=>r.d+'|'+(r.src.t==='pn'?r.src.id:'nh'))).size;
      const hd=(S.hd||{})[dnId(S.thang,po.ncc,ky.k)]||null, dn=(S.dn||{})[dnId(S.thang,po.ncc,ky.k)]||null;
      const ck={khop:!!(ds&&ds.tt==='khop'),migo:mg.length>0&&mg.every(m=>m.tt==='migo'),migoLap:mg.length>0,po:!po.sua,poEq:!last||poEqMigo,hd:!!(hd&&hd.soHD&&Math.abs((+hd.tien||0)-Math.round(tien))<=1)};
      out.push({po,ky,last,mg,tien:Math.round(tien),loai,bb,hd,dn,ck,ready:ck.khop&&ck.migoLap&&ck.po&&ck.poEq&&ck.hd}); }); });
  return out;
}
const ckPill=(ok,txt,warn)=>`<span class="pill ${ok?'ok':(warn?'run':'bad')}" title="${esc(txt)}">${ok?'✓':'✗'} ${txt}</span>`;
function noiDung(r){ return `Thanh toán chi phí ${r.loai} tháng ${+S.thang.slice(5)}${r.ky.k?` kỳ ${r.ky.k}`:''}`; }
function vHD(){
  const R=dnRows(); ED.hd=ED.hd||{};
  return head2('8.1 Hóa đơn & điều kiện lập ĐNTT','Mỗi NCC mỗi kỳ: tiền ĐNTT = giá trị MIGO kỳ theo đơn giá PO (PO = MIGO). Nhập số hóa đơn NCC xuất theo số PO; ĐNTT chỉ lập khi đủ điều kiện: đã khớp NCC, đã lập MIGO, PO không còn chờ sửa ME22N, kỳ cuối tháng thì PO = tổng MIGO, và tiền hóa đơn khớp.')+
  (R.length?`<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>NCC · PO</th><th>Kỳ</th><th class="r">Tiền theo PO/MIGO</th><th>Số hóa đơn</th><th>Ngày HĐ</th><th class="r">Tiền HĐ</th><th>Điều kiện</th><th></th></tr></thead><tbody>
  ${R.map((r,i)=>{const k=dnId(S.thang,r.po.ncc,r.ky.k), h=ED.hd[k]||r.hd||{}; const lockDn=r.dn&&r.dn.tt!=='huy';
    return `<tr><td><b>${esc(r.po.ncc)}</b><div class="meta num">${esc(r.po.soPO)}</div></td><td>${kyName(r.ky)}</td><td class="r num">${VND(r.tien)}</td>
    <td><input type="text" id="hd-so-${i}" data-hd="${k}" data-f="soHD" value="${esc(h.soHD||'')}" style="width:100px" ${canWrite&&!lockDn?'':'disabled'}></td>
    <td><input type="date" id="hd-ng-${i}" data-hd="${k}" data-f="ngay" value="${esc(h.ngay||'')}" ${canWrite&&!lockDn?'':'disabled'}></td>
    <td class="r"><input type="text" inputmode="numeric" id="hd-t-${i}" data-hd="${k}" data-f="tien" value="${h.tien!=null&&h.tien!==''?QI(h.tien):''}" style="width:120px;text-align:right" ${canWrite&&!lockDn?'':'disabled'}>${h.tien!=null&&h.tien!==''&&Math.abs(+h.tien-r.tien)>1?`<div class="pill bad">lệch ${VND(+h.tien-r.tien)}</div>`:''}</td>
    <td style="white-space:normal;min-width:260px">${ckPill(r.ck.khop,'Khớp NCC')} ${ckPill(r.ck.migoLap,'Lập MIGO')} ${ckPill(r.ck.migo,'Có số MIGO',true)} ${ckPill(r.ck.po,'PO đã cập nhật SAP')} ${r.last?ckPill(r.ck.poEq,'PO = MIGO'):''} ${ckPill(r.ck.hd,'Hóa đơn khớp')}</td>
    <td>${canWrite&&!lockDn?`<button class="btn sm" data-act8="hd-save" data-k="${k}" data-i="${i}">Lưu</button>`:(r.dn?`<span class="pill ${TTDN[r.dn.tt][1]}">${TTDN[r.dn.tt][0]}</span>`:'')}</td></tr>`;}).join('')}
  </tbody></table></div><p class="meta">“Có số MIGO” là cảnh báo vàng: ở chế độ không lưu MIGO chưa có số chứng từ, vẫn cho lập ĐNTT để chạy thử.</p>`
  :'<div class="empty"><b>Chưa có kỳ nào đến bước thanh toán</b>Cần kỳ đã chốt khớp NCC (bước 5) và đã lập MIGO (bước 7).</div>');
}
function dnSheetRows(list){ return list.map((r,i)=>[i+1,r.po.ncc,r.po.soPO,r.tien,+S.thang.slice(5),r.ky.k,noiDung(r),'','',r.bb,'KHOP MIGO KY',SHEETMAU[r.po.ncc.toUpperCase()]||'','DNTT.xlsx','','','']); }
function vDN(){
  const R=dnRows(), ready=R.filter(r=>r.ready&&(!r.dn||r.dn.tt==='huy')), made=R.filter(r=>r.dn&&r.dn.tt!=='huy');
  if(!ED.dnsel) ED.dnsel=new Set(made.filter(r=>r.dn.tt==='lap').map(r=>r.dn.id));
  const sel=made.filter(r=>ED.dnsel.has(r.dn.id)), rows=dnSheetRows(sel);
  return head2('8.2 Lập ĐNTT theo PO','Lập ĐNTT cho các NCC × kỳ đủ điều kiện. Nội dung theo mẫu cũ: “Thanh toán chi phí {thực phẩm tươi sống / đồ khô} tháng {tháng của số liệu} kỳ {kỳ}”, số tiền theo PO, số biên bản = số lần giao trong kỳ, sheet mẫu theo bảng NCC của DNTT_CauHinh. Chọn ĐNTT đã lập rồi bấm <b>Tạo file ĐNTT</b>: trạm trên máy công ty chạy TaoDNTT, file ĐNTT + PDF tự chép về thư mục XuatSAP.')+tramBox('DNTT')+
  `<div class="kpis"><div class="kv big"><b>Đủ điều kiện chưa lập</b><div class="num">${ready.length}</div><small>${TR(ready.reduce((a,r)=>a+r.tien,0))}</small></div>
   <div class="kv"><b>Đã lập tháng ${mmY(S.thang)}</b><div class="num">${made.length}</div><small>${TR(made.reduce((a,r)=>a+r.dn.tien,0))}</small></div>
   <div class="kv"><b>Chưa đủ điều kiện</b><div class="num">${R.filter(r=>!r.ready&&!r.dn).length}</div><small>xem 8.1</small></div></div>
  ${ready.length?`<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>NCC</th><th>PO</th><th>Kỳ</th><th>Nội dung</th><th class="r">Số tiền</th><th class="r">Biên bản</th><th>Hóa đơn</th><th>Sheet mẫu</th></tr></thead><tbody>
   ${ready.map(r=>`<tr><td><b>${esc(r.po.ncc)}</b></td><td class="num">${esc(r.po.soPO)}</td><td>${kyName(r.ky)}</td><td class="wrap">${esc(noiDung(r))}</td><td class="r num">${VND(r.tien)}</td><td class="r num">${r.bb}</td><td class="num">${esc(r.hd.soHD)}</td><td>${SHEETMAU[r.po.ncc.toUpperCase()]?esc(SHEETMAU[r.po.ncc.toUpperCase()]):'<span class="pill bad">chưa có sheet mẫu</span>'}</td></tr>`).join('')}</tbody></table></div>
   ${canWrite?`<div class="toolbar"><button class="btn primary" data-act8="dn-make">Lập ${ready.length} ĐNTT</button></div>`:''}`:''}
  ${made.length?`<h3 class="eyebrow" style="margin:6px 0 0">ĐNTT đã lập</h3><div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th></th><th>NCC</th><th>PO</th><th>Kỳ</th><th>Nội dung</th><th class="r">Số tiền</th><th>Trạng thái</th></tr></thead><tbody>
   ${made.map(r=>`<tr><td><input type="checkbox" data-dnsel="${r.dn.id}" ${ED.dnsel.has(r.dn.id)?'checked':''} aria-label="Chọn"></td><td>${esc(r.po.ncc)}</td><td class="num">${esc(r.po.soPO)}</td><td>${kyName(r.ky)}</td><td class="wrap">${esc(r.dn.noiDung)}</td><td class="r num">${VND(r.dn.tien)}</td><td><span class="pill ${TTDN[r.dn.tt][1]}">${TTDN[r.dn.tt][0]}</span>${lenhPill(r.dn.lenh)}${dnLinks(r.dn)}</td></tr>`).join('')}</tbody></table></div>
   ${canWrite?`<div class="toolbar"><button class="btn primary" data-tram="dntt" ${sel.filter(r=>!lenhBusy(r.dn.lenh)&&(r.dn.tt==='lap'||r.dn.tt==='xuat')).length?'':'disabled'}>Tạo file ĐNTT (${sel.filter(r=>!lenhBusy(r.dn.lenh)&&(r.dn.tt==='lap'||r.dn.tt==='xuat')).length} đã chọn)</button></div>`:''}
   <div class="toolbar"><button class="btn" data-act8="dn-xlsx" ${rows.length?'':'disabled'}>Tải file cho sheet DNTT</button><button class="btn" data-act8="dn-copy" ${rows.length?'':'disabled'}>Copy để dán</button><span class="meta">${rows.length} dòng · dán vào sheet DNTT ô A3, rồi Alt+F8 → TaoDNTT</span></div>`:''}
  ${!ready.length&&!made.length?'<div class="empty"><b>Chưa có ĐNTT nào đủ điều kiện</b>Hoàn tất điều kiện ở 8.1.</div>':''}`;
}
function vDNls(){
  const L=Object.values(S.dn||{}).filter(d=>(ED.dnAll||d.thang===S.thang)&&d.tt!=='huy').sort((a,b)=>b.luc-a.luc);
  const nxt={lap:'xuat',xuat:'trinh',trinh:'ketoan',ketoan:'tt'};
  return head2('8.3 Lịch sử ĐNTT & thanh toán','Theo dõi từng ĐNTT: đã lập → tạo file ĐNTT → trình ký → chuyển kế toán → đã thanh toán.',`<label class="meta"><input type="checkbox" id="dn-all" ${ED.dnAll?'checked':''}> Tất cả các tháng</label>`)+
  `<div class="kpis"><div class="kv big"><b>Tổng ĐNTT</b><div class="num">${TR(L.reduce((a,d)=>a+d.tien,0))}</div><small>${L.length} đề nghị</small></div>
   <div class="kv"><b>Đã thanh toán</b><div class="num">${TR(L.filter(d=>d.tt==='tt').reduce((a,d)=>a+d.tien,0))}</div><small>${L.filter(d=>d.tt==='tt').length} đề nghị</small></div>
   <div class="kv"><b>Còn phải trả</b><div class="num">${TR(L.filter(d=>d.tt!=='tt').reduce((a,d)=>a+d.tien,0))}</div></div></div>
  ${L.length?`<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>NCC</th><th>Tháng · kỳ</th><th>PO</th><th>Hóa đơn</th><th class="r">Số tiền</th><th>Trạng thái</th><th>Cập nhật</th><th></th></tr></thead><tbody>
   ${L.map(d=>`<tr><td><b>${esc(d.ncc)}</b><div class="meta">${esc(d.noiDung)}</div></td><td>${mmY(d.thang)}${d.k?' · kỳ '+d.k:''}</td><td class="num">${esc(d.soPO)}</td><td class="num">${esc(d.soHD||'')}</td><td class="r num">${VND(d.tien)}</td><td><span class="pill ${TTDN[d.tt][1]}">${TTDN[d.tt][0]}</span>${dnLinks(d)}</td><td class="meta">${tstr((d.hist||[]).slice(-1)[0]?.ts||d.luc)}</td>
     <td>${canWrite?`${nxt[d.tt]?`<button class="btn sm" data-act8="dn-next" data-id="${d.id}">→ ${TTDN[nxt[d.tt]][0]}</button>`:''} ${d.tt!=='tt'?`<button class="btn sm warn" data-act8="dn-huy" data-id="${d.id}">Hủy</button>`:''}`:''}</td></tr>`).join('')}</tbody></table></div>`:'<div class="empty"><b>Chưa có ĐNTT</b>Lập ở 8.2.</div>'}`;
}
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act8]'); if(!t) return; const a=t.dataset.act8;
  try{
  if(a==='hd-save'){ const k=t.dataset.k, h={...((S.hd||{})[k]||{}),...((ED.hd||{})[k]||{})}; const [th,,kk]=[k.slice(0,7),0,k.split('_k').pop()];
    const r=dnRows().find(x=>dnId(S.thang,x.po.ncc,x.ky.k)===k); if(!r) return;
    await store.set('hoadon/'+k,{thang:S.thang,ncc:r.po.ncc,k:r.ky.k,poId:r.po.id,soPO:r.po.soPO,soHD:String(h.soHD||'').trim(),ngay:h.ngay||'',tien:h.tien===''||h.tien==null?null:+h.tien,luc:Date.now(),ai:uid});
    delete ED.hd[k]; await log('8.1',`Hóa đơn ${h.soHD||''} ${r.po.ncc} ${kyName(r.ky)}: ${VND(h.tien)} (PO/MIGO ${VND(r.tien)})`); }
  else if(a==='dn-make'){ const R=dnRows().filter(r=>r.ready&&(!r.dn||r.dn.tt==='huy')); let n=0;
    for(const r of R){ const id=dnId(S.thang,r.po.ncc,r.ky.k); await store.set('dntt/'+id,{thang:S.thang,ncc:r.po.ncc,k:r.ky.k,poId:r.po.id,soPO:r.po.soPO,tien:r.tien,noiDung:noiDung(r),loai:r.loai,soBB:r.bb,soHD:r.hd.soHD,ngayHD:r.hd.ngay,sheetMau:SHEETMAU[r.po.ncc.toUpperCase()]||'',migo:r.mg.map(m=>m.soMIGO||m.ma),tt:'lap',luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'lap',ai:uid,note:'Lập từ PO/MIGO'}]}); n++; }
    ED.dnsel=null; await log('8.2',`Lập ${n} ĐNTT (${TR(R.reduce((s,r)=>s+r.tien,0))})`); }
  else if(a==='dn-copy'||a==='dn-xlsx'){ const R=dnRows().filter(r=>r.dn&&ED.dnsel.has(r.dn.id)); const rows=dnSheetRows(R);
    if(a==='dn-copy'){ try{ await navigator.clipboard.writeText(rows.map(r=>r.join('\t')).join('\r\n')); toast('Đã copy '+rows.length+' dòng. Dán vào sheet DNTT ô A3.'); }catch(err){ toast('Không copy được.'); return; } }
    else { if(!window.XLSX){toast('Không tạo được file; dùng nút Copy.');return;} const H=['STT','Nha_cung_cap','So_PO','So_tien_PO','Thang','Ky','Noi_dung_thanh_toan','Ngay_lap','Han_thanh_toan','So_bien_ban','Kiem_tra_ky','Sheet_mau','File_mau','Nguoi_thu_huong','File_DNTT_da_tao','Status'];
      const ws=XLSX.utils.aoa_to_sheet([H,...rows]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'DNTT');
      try{ await saveOut(`DNTT_${S.thang}_${todayVN()}.xlsx`,new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); }catch(err){ if(err&&err.code!=='cancelled') toast('Không lưu được file.'); return; } }
    for(const r of R) if(r.dn.tt==='lap') await store.set('dntt/'+r.dn.id,{...stripId(r.dn),tt:'xuat',hist:[...(r.dn.hist||[]),{ts:Date.now(),tt:'xuat',ai:uid,note:'Xuất cho sheet DNTT'}]});
    await log('8.2',`Xuất ${rows.length} ĐNTT cho sheet DNTT`); }
  else if(a==='dn-next'){ const d=S.dn[t.dataset.id], nx={lap:'xuat',xuat:'trinh',trinh:'ketoan',ketoan:'tt'}[d.tt];
    await store.set('dntt/'+d.id,{...stripId(d),tt:nx,[nx+'Luc']:Date.now(),hist:[...(d.hist||[]),{ts:Date.now(),tt:nx,ai:uid,note:''}]}); await log('8.3',`ĐNTT ${d.ncc} ${mmY(d.thang)}${d.k?' kỳ '+d.k:''}: ${TTDN[nx][0]}`); }
  else if(a==='dn-huy'){ const d=S.dn[t.dataset.id]; if(ED.dnHuy!==d.id){ED.dnHuy=d.id;toast('Bấm Hủy lần nữa để hủy ĐNTT '+d.ncc+'.');return;} ED.dnHuy=null;
    await store.set('dntt/'+d.id,{...stripId(d),tt:'huy',hist:[...(d.hist||[]),{ts:Date.now(),tt:'huy',ai:uid,note:'Hủy'}]}); await log('8.3',`Hủy ĐNTT ${d.ncc}`); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target; if(t.dataset.hd){ ED.hd=ED.hd||{}; const k=t.dataset.hd; ED.hd[k]=ED.hd[k]||{...((S.hd||{})[k]||{})}; ED.hd[k][t.dataset.f]=t.dataset.f==='tien'?(t.value.trim()===''?'':pnum(t.value.replace(/\./g,''))):t.value; } });
document.addEventListener('change',e=>{ const t=e.target; if(t.dataset.dnsel){ ED.dnsel=ED.dnsel||new Set(); t.checked?ED.dnsel.add(t.dataset.dnsel):ED.dnsel.delete(t.dataset.dnsel); render(); } else if(t.id==='dn-all'){ ED.dnAll=t.checked; render(); } });

