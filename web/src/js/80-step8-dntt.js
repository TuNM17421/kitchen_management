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
function noiDung(r){ return `Thanh toán chi phí ${r.loai} tháng ${+S.thang.slice(5)}${r.ky.k?` kỳ ${r.ky.k}`:''}`; }
function dnSheetRows(list){ return list.map((r,i)=>[i+1,r.po.ncc,r.po.soPO,r.tien,+S.thang.slice(5),r.ky.k,noiDung(r),'','',r.bb,'KHOP MIGO KY',SHEETMAU[r.po.ncc.toUpperCase()]||'','DNTT.xlsx','','','']); }
const CK8=[['khop','Khớp NCC',5,'sofile','Mở 5.2'],['migoLap','Lập MIGO',7,'mg','Mở 7.1'],['po','PO đã cập nhật SAP',3,'pols','Mở 3.4'],['poEq','PO = tổng MIGO',6,'cpo','Mở 6.1'],['hd','Hóa đơn khớp',0,'','']];
const hdKey=r=>dnId(S.thang,r.po.ncc,r.ky.k);
const dnLive=r=>!!(r.dn&&r.dn.tt!=='huy');
const ckMiss=r=>CK8.filter(c=>!r.ck[c[0]]);
const kyNg=r=>r.ky.k?'kỳ '+r.ky.k:'cả tháng';
const migoWarn=r=>r.mg.length&&!r.ck.migo?`<span class="pill run">${r.mg.some(m=>!mgReady(m))?'MIGO chờ hết kỳ':'MIGO chưa có số'}</span>`:'';
const okBox=(m,x)=>m?`<div class="okbox" role="status"><span>${m}</span><button class="x" data-act8="${x}" aria-label="Đóng thông báo">×</button></div>`:'';
const fli=(d,c,n,t,b)=>`<li class="${d?'done':c?'cur':''}"><div class="fh"><span class="fn" aria-hidden="true">${d?'✓':n}</span><b>${t}</b></div>${b}</li>`;
const DAY=864e5, QUAHAN=15;
const dnQuaHan=d=>d.tt==='ketoan'&&d.ketoanLuc&&Date.now()-d.ketoanLuc>QUAHAN*DAY;

/* ---- 8.1 Hóa đơn & điều kiện ---- */
function vHD(){
  const R=dnRows(); ED.hd=ED.hd||{};
  const H=head2('8.1 Hóa đơn & điều kiện lập ĐNTT','Mỗi NCC mỗi kỳ: tiền ĐNTT = giá trị MIGO kỳ theo đơn giá PO. Nhập số hóa đơn NCC xuất theo số PO, ngày và tiền hóa đơn rồi bấm Lưu. ĐNTT chỉ lập khi đủ điều kiện: đã khớp NCC, đã lập MIGO, PO không còn chờ sửa ME22N, kỳ cuối tháng thì PO = tổng MIGO, và tiền hóa đơn khớp. Điều kiện còn thiếu có nút mở đúng bước để xử lý.');
  if(!R.length) return H+'<div class="empty"><b>Chưa có kỳ nào đến bước thanh toán</b>Cần kỳ đã chốt khớp NCC (bước 5) và đã lập MIGO (bước 7).<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-step="5" data-tab="sofile">Mở 5.2</button><button class="btn sm" data-act="go-next" data-step="7" data-tab="mg">Mở 7.1</button></div></div>';
  const f=ED.hdF||'can', nf=ED.hdN||'', nccs=[...new Set(R.map(r=>r.po.ncc))].sort((a,b)=>a.localeCompare(b,'vi'));
  const Rn=R.filter(r=>!nf||r.po.ncc===nf), G={can:Rn.filter(r=>!dnLive(r)),lap:Rn.filter(dnLive),all:Rn}, L=G[f]||Rn;
  const open=R.filter(r=>!dnLive(r)), okN=open.filter(r=>r.ready), needHd=open.filter(r=>!r.ck.hd), other=open.filter(r=>ckMiss(r).some(c=>c[0]!=='hd'));
  const body=L.map((r,i)=>{ const k=hdKey(r), dirty=!!ED.hd[k], h=ED.hd[k]||r.hd||{}, lock=dnLive(r)||!canWrite, dis=lock?'disabled':'', lb=esc(`${r.po.ncc} ${kyNg(r)}`);
    const tv=h.tien!=null&&h.tien!==''?+h.tien:null, lech=tv!=null&&Math.abs(tv-r.tien)>1, miss=ckMiss(r);
    const cond=dnLive(r)?`<span class="pill ${TTDN[r.dn.tt][1]}">ĐNTT: ${TTDN[r.dn.tt][0]}</span><small>Hủy ĐNTT ở 8.3 nếu cần sửa hóa đơn</small>`
      :(r.ready?'<span class="pill ok">✓ Đủ điều kiện</span> ':`<span class="pill bad">Thiếu ${miss.length} điều kiện</span> `)+migoWarn(r)+(miss.length?`<small>Thiếu: ${miss.map(c=>c[1]).join(', ')}</small>`:'<small>Lập ở 8.2</small>')+
       (miss.some(c=>c[2])?`<div class="acts-in hd-go">${miss.filter(c=>c[2]).map(c=>`<button class="btn sm" data-act="go-next" data-step="${c[2]}" data-tab="${c[3]}">${c[4]}</button>`).join('')}</div>`:'');
    return `<tr class="${dirty?'dirty':''}"><td class="exp meta num">${i+1}</td>
    <td class="dg"><b>${esc(r.po.ncc)}</b><small>PO <span class="num">${esc(r.po.soPO)}</span> · ${kyName(r.ky)}</small></td>
    <td class="r tien"><b class="num">${VND(r.tien)}</b><small>theo PO/MIGO</small></td>
    <td class="tt"><div class="hd-f"><label><span>Số hóa đơn</span><input type="text" id="hd-so-${i}" data-hd="${k}" data-f="soHD" value="${esc(h.soHD||'')}" aria-label="Số hóa đơn ${lb}" ${dis}></label>
      <label><span>Ngày hóa đơn</span><input type="date" id="hd-ng-${i}" data-hd="${k}" data-f="ngay" min="${r.ky.from}" value="${esc(h.ngay||'')}" aria-label="Ngày hóa đơn ${lb}" ${dis}></label>
      <label class="w"><span>Tiền hóa đơn</span><span class="hd-t"><input type="text" inputmode="numeric" id="hd-t-${i}" data-hd="${k}" data-f="tien" value="${tv!=null?VND(tv):''}" aria-label="Tiền hóa đơn ${lb}" ${dis}>${!lock&&tv!==r.tien?`<button class="btn sm ghost" data-act8="hd-fill" data-k="${k}" data-t="${r.tien}">= Tiền PO</button>`:''}</span></label></div>
      ${lech?`<small class="bad-t">Lệch ${VND(tv-r.tien)} đ so với PO/MIGO</small>`:''}</td>
    <td class="when">${cond}</td>
    <td class="acts"><div class="acts-in">${lock?'':`<button class="btn sm ${dirty?'primary':''}" data-act8="hd-save" data-k="${k}" ${ED.hdBusy?'disabled':''}>Lưu</button>`}</div></td></tr>`;}).join('');
  return H+`<div class="kpis"><div class="kv big"><b>Đủ điều kiện, chưa lập ĐNTT</b><div class="num">${okN.length}</div><small>${TR(okN.reduce((a,r)=>a+r.tien,0))}</small></div>
   <div class="kv"><b>Cần nhập hóa đơn</b><div class="num">${needHd.length}</div><small>${needHd.length?'nhập ở bảng dưới':'không có'}</small></div>
   <div class="kv"><b>Thiếu điều kiện khác</b><div class="num">${other.length}</div><small>${other.length?'bấm nút Mở ở từng dòng':'không có'}</small></div>
   <div class="kv"><b>Đã lập ĐNTT</b><div class="num">${R.length-open.length}</div><small>/ ${R.length} NCC × kỳ</small></div></div>`+
  okBox(ED.hdOk&&(esc(ED.hdOk.m)+(ED.hdOk.ready?' <button class="btn sm" data-act="go-next" data-tab="dn">Mở 8.2</button>':'')),'hd-okx')+
  `<div class="toolbar">${nccs.length>1?`<select id="hd-ncc" aria-label="Lọc theo NCC"><option value="">Tất cả NCC</option>${nccs.map(n=>`<option ${n===nf?'selected':''}>${esc(n)}</option>`).join('')}</select>`:''}
   <div class="seg wrapseg">${[['can','Chưa lập ĐNTT'],['lap','Đã lập ĐNTT'],['all','Tất cả']].map(([v,t])=>`<button data-act8="hd-f" data-f="${v}" aria-pressed="${f===v}">${t} ${G[v].length}</button>`).join('')}</div></div>`+
  (L.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl hd-tbl"><colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"><col class="c-when"><col class="c-acts"></colgroup><thead><tr><th><span class="sr">STT</span></th><th>NCC · PO · kỳ</th><th class="r">Tiền</th><th>Hóa đơn NCC</th><th>Điều kiện</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>${body}</tbody></table></div>
   <p class="meta">“MIGO chưa có số” là cảnh báo vàng: ở chế độ không lưu MIGO chưa có số chứng từ, vẫn cho lập ĐNTT để chạy thử.</p>`
  :`<div class="empty"><b>${f==='can'?'Mọi kỳ đã lập ĐNTT':'Không có dòng nào khớp bộ lọc'}</b></div>`);
}

/* ---- 8.2 Lập ĐNTT ---- */
function vDN(){
  const R=dnRows(), ready=R.filter(r=>r.ready&&!dnLive(r)), made=R.filter(dnLive), bo=ED.dnBo||(ED.dnBo=new Set()), notReady=R.filter(r=>!r.ready&&!dnLive(r));
  const rsel=ready.filter(r=>!bo.has(hdKey(r))), F=made.filter(r=>r.dn.tt==='lap'||r.dn.tt==='xuat'), elig=F.filter(r=>!lenhBusy(r.dn.lenh));
  if(!ED.dnsel) ED.dnsel=new Set(F.filter(r=>r.dn.tt==='lap').map(r=>r.dn.id));
  const fsel=F.filter(r=>ED.dnsel.has(r.dn.id)), tsel=elig.filter(r=>ED.dnsel.has(r.dn.id)), rows=dnSheetRows(fsel), on=tramOnline();
  const nrBtn=notReady.length?`<span class="meta">${notReady.length} NCC × kỳ chưa đủ điều kiện</span><button class="btn sm" data-act="go-next" data-tab="hd">Mở 8.1</button>`:'';
  const s1=ready.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl dnr-tbl"><colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"><col class="c-when"></colgroup><thead><tr><th><input type="checkbox" id="dnr-all" aria-label="Chọn tất cả kỳ đủ điều kiện" ${rsel.length===ready.length?'checked':''}></th><th>NCC · PO · kỳ</th><th class="r">Số tiền</th><th>Hóa đơn · biên bản</th><th>Lưu ý</th></tr></thead><tbody>
    ${ready.map(r=>{const k=hdKey(r), sm=SHEETMAU[r.po.ncc.toUpperCase()];return `<tr><td class="exp"><input type="checkbox" data-dnr="${k}" ${bo.has(k)?'':'checked'} aria-label="Chọn ${esc(r.po.ncc)} ${kyNg(r)}"></td>
     <td class="dg"><b>${esc(r.po.ncc)}</b><small>PO <span class="num">${esc(r.po.soPO)}</span> · ${kyName(r.ky)}</small><small>${esc(noiDung(r))}</small></td>
     <td class="r tien"><b class="num">${VND(r.tien)}</b></td>
     <td class="tt">HĐ <span class="num">${esc(r.hd.soHD)}</span>${r.hd.ngay?` · ${dm(r.hd.ngay)}`:''}<small>${r.bb} biên bản giao hàng</small></td>
     <td class="when">${sm?`<small>Sheet ${esc(sm)}</small>`:'<span class="pill bad">Chưa có sheet mẫu</span>'}${migoWarn(r)}</td></tr>`;}).join('')}</tbody></table></div>
    ${canWrite?`<div class="toolbar"><button class="btn primary" data-act8="dn-make" ${rsel.length&&!ED.dnBusy?'':'disabled'}>Lập ${rsel.length} ĐNTT · ${VND(rsel.reduce((a,r)=>a+r.tien,0))} đ</button>${nrBtn}</div>`:''}`
    :`<div class="toolbar"><span class="meta">${made.length?'Đã lập ĐNTT cho mọi kỳ đủ điều kiện.':'Chưa có kỳ nào đủ điều kiện.'}</span>${nrBtn}</div>`;
  const s2=F.length?`${tramBox('DNTT')}<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl dnf-tbl"><colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"></colgroup><thead><tr><th>${elig.length?`<input type="checkbox" id="dnf-all" aria-label="Chọn tất cả ĐNTT" ${elig.every(r=>ED.dnsel.has(r.dn.id))?'checked':''}>`:'<span class="sr">Chọn</span>'}</th><th>ĐNTT</th><th class="r">Số tiền</th><th>Trạng thái</th></tr></thead><tbody>
    ${F.map(r=>`<tr><td class="exp">${lenhBusy(r.dn.lenh)?'':`<input type="checkbox" data-dnsel="${r.dn.id}" ${ED.dnsel.has(r.dn.id)?'checked':''} aria-label="Chọn ĐNTT ${esc(r.po.ncc)} ${kyNg(r)}">`}</td>
     <td class="dg"><b>${esc(r.po.ncc)}</b><small>PO <span class="num">${esc(r.po.soPO)}</span> · ${kyName(r.ky)}</small><small>${esc(r.dn.noiDung)}</small></td>
     <td class="r tien"><b class="num">${VND(r.dn.tien)}</b></td>
     <td class="tt"><span class="pill ${TTDN[r.dn.tt][1]}">${TTDN[r.dn.tt][0]}</span>${lenhPill(r.dn.lenh)}${dnLinks(r.dn)}</td></tr>`).join('')}</tbody></table></div>
    ${canWrite?`<div class="toolbar"><button class="btn primary" data-tram="dntt" ${tsel.length?'':'disabled'}>${on?'Tạo file ĐNTT':'Gửi lệnh tạo file ĐNTT'} · ${tsel.length} ĐNTT</button><span class="meta">${on?'Trạm tạo file ĐNTT + PDF, link hiện ngay trong bảng':'Trạm đang ngoại tuyến, lệnh chạy khi trạm mở'}</span></div>
    <details class="more"${ED.dnMacro?' open':''} id="dn-macro"><summary>Chạy bằng macro (khi không dùng trạm)</summary>
     <div class="toolbar"><button class="btn sm" data-act8="dn-xlsx" ${rows.length?'':'disabled'}>Tải file cho sheet DNTT</button><button class="btn sm" data-act8="dn-copy" ${rows.length?'':'disabled'}>Copy để dán</button></div>
     <p class="meta">${rows.length} dòng của các ĐNTT đã chọn · dán vào sheet DNTT ô A3, rồi Alt+F8 → TaoDNTT. Sau khi tải hoặc copy, ĐNTT chuyển sang “Đã tạo file ĐNTT”.</p></details>`:''}`
    :`<p class="meta">${made.length?'Mọi ĐNTT đã có file, tiếp tục trình ký ở 8.3.':'Lập ĐNTT ở bước 1 trước.'}</p>`;
  const paid=made.filter(r=>r.dn.tt==='tt').length, st1=!ready.length&&made.length>0, st2=made.length>0&&!F.some(r=>r.dn.tt==='lap'), st3=made.length>0&&paid===made.length;
  return head2('8.2 Lập ĐNTT theo PO','Lập ĐNTT cho các NCC × kỳ đủ điều kiện. Nội dung theo mẫu cũ: “Thanh toán chi phí {thực phẩm tươi sống / đồ khô} tháng {tháng} kỳ {kỳ}”, số tiền theo PO, số biên bản = số lần giao trong kỳ, sheet mẫu theo bảng NCC của DNTT_CauHinh. Sau đó chọn ĐNTT và bấm Tạo file ĐNTT: trạm trên máy công ty chạy TaoDNTT, file ĐNTT + PDF tự chép về thư mục XuatSAP.')+
  `<div class="kpis"><div class="kv big"><b>Đủ điều kiện chưa lập</b><div class="num">${ready.length}</div><small>${TR(ready.reduce((a,r)=>a+r.tien,0))}</small></div>
   <div class="kv"><b>Chờ tạo file</b><div class="num">${F.filter(r=>r.dn.tt==='lap').length}</div><small>ĐNTT đã lập</small></div>
   <div class="kv"><b>Đã lập tháng ${mmY(S.thang)}</b><div class="num">${made.length}</div><small>${TR(made.reduce((a,r)=>a+r.dn.tien,0))}</small></div>
   <div class="kv"><b>Chưa đủ điều kiện</b><div class="num">${notReady.length}</div><small>${notReady.length?'xem 8.1':'không có'}</small></div></div>`+
  okBox(ED.dnOk&&esc(ED.dnOk),'dn-okx')+
  `<ol class="flow">${fli(st1,!st1,1,'Lập ĐNTT cho các kỳ đủ điều kiện',s1)}${fli(st2,st1&&!st2,2,'Tạo file ĐNTT',s2)}${fli(st3,st2&&!st3,3,'Trình ký & thanh toán',`<div class="toolbar"><span class="meta">${paid} / ${made.length} ĐNTT đã thanh toán</span><button class="btn sm" data-act="go-next" data-tab="dnls">Xem ở 8.3</button></div>`)}</ol>`;
}

/* ---- 8.3 Lịch sử & thanh toán ---- */
const NXT8={xuat:'trinh',trinh:'ketoan',ketoan:'tt'};
function vDNls(){
  const A=Object.values(S.dn||{}).filter(d=>(ED.dnAll||d.thang===S.thang)&&d.tt!=='huy').sort((a,b)=>b.luc-a.luc), f=ED.dnlsF||'', nf=ED.dnlsN||'';
  const nccs=[...new Set(A.map(d=>d.ncc))].sort((a,b)=>a.localeCompare(b,'vi')), An=A.filter(d=>!nf||d.ncc===nf), cnt={}; An.forEach(d=>cnt[d.tt]=(cnt[d.tt]||0)+1);
  const L=An.filter(d=>!f||d.tt===f), qh=A.filter(dnQuaHan), td=todayVN(), ask=ED.dnAsk||{};
  const H=head2('8.3 Lịch sử ĐNTT & thanh toán','Theo dõi từng ĐNTT: đã lập → tạo file ĐNTT (ở 8.2) → trình ký → chuyển kế toán → đã thanh toán. Mỗi lần chuyển trạng thái cần xác nhận; khi thanh toán ghi ngày và số UNC. Lọc theo một trạng thái để chuyển cả nhóm.',`<label class="meta"><input type="checkbox" id="dn-all" ${ED.dnAll?'checked':''}> Tất cả các tháng</label>`);
  if(!A.length) return H+'<div class="empty"><b>Chưa có ĐNTT</b>Lập ở 8.2.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-tab="dn">Mở 8.2</button></div></div>';
  const askBox=d=>{ if(ask.id!==d.id) return ''; const lb=esc(`${d.ncc} ${d.k?'kỳ '+d.k:'cả tháng'}`);
    if(ask.k==='huy') return `<div class="dn-ask" role="group" aria-label="Hủy ĐNTT ${lb}"><b>Hủy ĐNTT ${lb}?</b> <span class="meta">Hóa đơn ở 8.1 sẽ mở lại để sửa, rồi lập ĐNTT mới.</span><label class="fld"><span>Lý do${d.tt==='ketoan'?' (bắt buộc, đã chuyển kế toán)':''}</span><input type="text" id="dn-lydo" aria-label="Lý do hủy ${lb}"></label><div class="acts-in"><button class="btn sm danger-text" data-act8="dn-huy-ok" data-id="${d.id}">Hủy ĐNTT</button><button class="btn sm ghost" data-act8="dn-ask">Thôi</button></div></div>`;
    const nx=NXT8[d.tt]; return `<div class="dn-ask" role="group" aria-label="Chuyển trạng thái ${lb}"><b>Chuyển ${lb} sang “${TTDN[nx][0]}”?</b>${nx==='tt'?`<label class="fld"><span>Ngày thanh toán</span><input type="date" id="dn-ngtt" value="${td}" aria-label="Ngày thanh toán ${lb}"></label><label class="fld"><span>Số UNC (không bắt buộc)</span><input type="text" id="dn-unc" aria-label="Số UNC ${lb}"></label>`:''}<div class="acts-in"><button class="btn sm primary" data-act8="dn-next-ok" data-id="${d.id}">Xác nhận</button><button class="btn sm ghost" data-act8="dn-ask">Thôi</button></div></div>`; };
  const body=L.map(d=>{ const open=ED.dnOpen===d.id, qhd=dnQuaHan(d), lb=esc(`${d.ncc} ${d.k?'kỳ '+d.k:'cả tháng'}`), last=(d.hist||[]).slice(-1)[0], ab=askBox(d);
    const acts=canWrite?`${d.tt==='lap'?'<button class="btn sm" data-act="go-next" data-tab="dn">Tạo file ở 8.2</button>':''}${NXT8[d.tt]?`<button class="btn sm" data-act8="dn-ask" data-id="${d.id}" data-k="next" aria-expanded="${ask.id===d.id&&ask.k==='next'}">→ ${TTDN[NXT8[d.tt]][0]}</button>`:''}${d.tt!=='tt'?`<button class="btn sm danger-text" data-act8="dn-ask" data-id="${d.id}" data-k="huy" aria-expanded="${ask.id===d.id&&ask.k==='huy'}">Hủy</button>`:''}`:'';
    const det=open?`<dl class="dn-det"><dt>Hóa đơn</dt><dd class="num">${esc(d.soHD||'–')}${d.ngayHD?' · '+dm(d.ngayHD)+'/'+d.ngayHD.slice(0,4):''}</dd><dt>MIGO</dt><dd class="num">${esc((d.migo||[]).join(', ')||'–')}</dd><dt>Biên bản</dt><dd>${d.soBB??'–'} · sheet ${esc(d.sheetMau||'chưa có')}</dd>${d.tt==='tt'?`<dt>Thanh toán</dt><dd>${d.ngayTT?dm(d.ngayTT)+'/'+d.ngayTT.slice(0,4):tstr(d.ttLuc)}${d.soUNC?' · UNC <span class="num">'+esc(d.soUNC)+'</span>':''}</dd>`:''}${d.fileDNTT?`<dt>File</dt><dd>${esc(String(d.fileDNTT).split('\\').pop())}${dnLinks(d)}</dd>`:''}</dl>
      <ol class="dn-hist" aria-label="Lịch sử trạng thái">${(d.hist||[]).map(h=>`<li><span class="meta num">${tstr(h.ts)}</span> ${TTDN[h.tt]?TTDN[h.tt][0]:esc(h.tt)}${h.note?` · <span class="meta">${esc(h.note)}</span>`:''}</li>`).join('')}</ol>`:'';
    return `<tr class="${qhd?'qh':''}"><td class="exp"><button class="x" data-act8="dn-open" data-id="${d.id}" aria-expanded="${open}" aria-label="${open?'Ẩn':'Xem'} chi tiết ĐNTT ${lb}">${open?'▾':'▸'}</button></td>
    <td class="dg"><b>${esc(d.ncc)}</b><small>${esc(d.noiDung)}</small><small>PO <span class="num">${esc(d.soPO)}</span>${d.soHD?` · HĐ <span class="num">${esc(d.soHD)}</span>`:''}</small></td>
    <td class="r tien"><b class="num">${VND(d.tien)}</b></td>
    <td class="tt"><span class="pill ${TTDN[d.tt][1]}">${TTDN[d.tt][0]}</span>${lenhPill(d.lenh)}${dnLinks(d)}${qhd?`<small class="bad-t">Chuyển kế toán ${Math.floor((Date.now()-d.ketoanLuc)/DAY)} ngày, chưa thanh toán</small>`:''}${d.tt==='tt'&&d.ngayTT?`<small>Trả ngày ${dm(d.ngayTT)}${d.soUNC?' · UNC '+esc(d.soUNC):''}</small>`:''}</td>
    <td class="meta when">${mmY(d.thang)}${d.k?' · kỳ '+d.k:''}<br>Cập nhật ${tstr(last?.ts||d.luc)}</td>
    <td class="acts"><div class="acts-in">${acts}</div></td></tr>${open||ab?`<tr class="detail"><td colspan="6">${ab}${det}</td></tr>`:''}`;}).join('');
  const bulk=canWrite&&NXT8[f]&&L.length>1?`<div class="toolbar dn-bulk">${f==='ketoan'?`<label class="fld"><span>Ngày thanh toán</span><input type="date" id="dn-bulk-ng" value="${ED.dnBulkNg||td}"></label>`:''}<button class="btn ${ED.dnBulk?'primary':''}" data-act8="dn-bulk">${ED.dnBulk?`Bấm lần nữa để chuyển ${L.length} ĐNTT`:`→ ${TTDN[NXT8[f]][0]} cho cả ${L.length} ĐNTT`}</button>${ED.dnBulk?'<button class="btn ghost" data-act8="dn-bulk-x">Thôi</button>':''}</div>`:'';
  const thangA=A.filter(d=>d.thang===S.thang), allPaid=thangA.length&&thangA.every(d=>d.tt==='tt');
  return H+`<div class="kpis"><div class="kv big"><b>Tổng ĐNTT</b><div class="num">${TR(A.reduce((a,d)=>a+d.tien,0))}</div><small>${A.length} đề nghị</small></div>
   <div class="kv"><b>Đã thanh toán</b><div class="num">${TR(A.filter(d=>d.tt==='tt').reduce((a,d)=>a+d.tien,0))}</div><small>${A.filter(d=>d.tt==='tt').length} đề nghị</small></div>
   <div class="kv"><b>Còn phải trả</b><div class="num">${TR(A.filter(d=>d.tt!=='tt').reduce((a,d)=>a+d.tien,0))}</div><small>${A.filter(d=>d.tt!=='tt').length} đề nghị</small></div>
   <div class="kv"><b>Quá ${QUAHAN} ngày chưa trả</b><div class="num">${qh.length}</div><small>${qh.length?TR(qh.reduce((a,d)=>a+d.tien,0)):'tính từ lúc chuyển kế toán'}</small></div></div>`+
  (allPaid&&!ED.dnAll?`<div class="okbox" role="status"><span>Tháng ${mmY(S.thang)} đã thanh toán đủ ${thangA.length} ĐNTT · ${VND(thangA.reduce((a,d)=>a+d.tien,0))} đ.</span></div>`:'')+
  `<div class="toolbar">${nccs.length>1?`<select id="dnls-ncc" aria-label="Lọc theo NCC"><option value="">Tất cả NCC</option>${nccs.map(n=>`<option ${n===nf?'selected':''}>${esc(n)}</option>`).join('')}</select>`:''}
   <div class="seg wrapseg"><button data-act8="dnls-f" data-f="" aria-pressed="${!f}">Tất cả ${An.length}</button>${Object.keys(TTDN).filter(k=>cnt[k]).map(k=>`<button data-act8="dnls-f" data-f="${k}" aria-pressed="${f===k}">${TTDN[k][0]} ${cnt[k]}</button>`).join('')}</div></div>`+bulk+
  (L.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl dnls-tbl"><colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"><col class="c-when"><col class="c-acts"></colgroup><thead><tr><th><span class="sr">Chi tiết</span></th><th>NCC · nội dung</th><th class="r">Số tiền</th><th>Trạng thái</th><th>Tháng · cập nhật</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>${body}</tbody></table></div>`:'<div class="empty"><b>Không có ĐNTT nào khớp bộ lọc</b></div>');
}

async function dnSetTT(d,nx,extra,note){ await store.set('dntt/'+d.id,{...stripId(d),...extra,tt:nx,[nx+'Luc']:Date.now(),hist:[...(d.hist||[]),{ts:Date.now(),tt:nx,ai:uid,note:note||''}]}); }
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act8]'); if(!t) return; const a=t.dataset.act8;
  try{
  if(a==='hd-okx'){ ED.hdOk=null; render(); }
  else if(a==='dn-okx'){ ED.dnOk=null; render(); }
  else if(a==='hd-f'){ ED.hdF=t.dataset.f; render(); }
  else if(a==='dnls-f'){ ED.dnlsF=t.dataset.f; ED.dnBulk=false; render(); }
  else if(a==='dn-open'){ ED.dnOpen=ED.dnOpen===t.dataset.id?null:t.dataset.id; render(); }
  else if(a==='dn-ask'){ const id=t.dataset.id; ED.dnAsk=id&&!(ED.dnAsk&&ED.dnAsk.id===id&&ED.dnAsk.k===t.dataset.k)?{id,k:t.dataset.k}:null; render(); if(ED.dnAsk){ const el=document.querySelector('#dn-lydo,#dn-ngtt,.dn-ask .btn.primary'); if(el) el.focus(); } }
  else if(a==='hd-fill'){ const k=t.dataset.k; ED.hd=ED.hd||{}; ED.hd[k]={...((S.hd||{})[k]||{}),...(ED.hd[k]||{}),tien:+t.dataset.t}; render(); }
  else if(a==='hd-save'){ if(ED.hdBusy) return; const k=t.dataset.k, h={...((S.hd||{})[k]||{}),...((ED.hd||{})[k]||{})};
    const r=dnRows().find(x=>hdKey(x)===k); if(!r) return;
    const errs=[]; if(!String(h.soHD||'').trim()) errs.push('nhập số hóa đơn'); if(!h.ngay) errs.push('chọn ngày hóa đơn'); else if(h.ngay<r.ky.from) errs.push(`ngày hóa đơn không được trước ngày đầu kỳ ${dm(r.ky.from)}`); if(h.tien===''||h.tien==null) errs.push('nhập tiền hóa đơn');
    if(errs.length){ toast('Chưa lưu: cần '+errs.join(', ')+'.'); return; }
    ED.hdBusy=true; render();
    try{ await store.set('hoadon/'+k,{thang:S.thang,ncc:r.po.ncc,k:r.ky.k,poId:r.po.id,soPO:r.po.soPO,soHD:String(h.soHD).trim(),ngay:h.ngay,tien:+h.tien,luc:Date.now(),ai:uid});
      delete ED.hd[k]; await log('8.1',`Hóa đơn ${h.soHD} ${r.po.ncc} ${kyName(r.ky)}: ${VND(h.tien)} (PO/MIGO ${VND(r.tien)})`);
      const r2=dnRows().find(x=>hdKey(x)===k)||r, lech=Math.abs(+h.tien-r.tien)>1, miss=ckMiss(r2);
      ED.hdOk={m:`Đã lưu hóa đơn ${String(h.soHD).trim()} cho ${r.po.ncc} ${kyNg(r)}. `+(lech?`Tiền hóa đơn lệch ${VND(+h.tien-r.tien)} đ so với PO/MIGO, chưa lập ĐNTT được.`:r2.ready?'Đủ điều kiện, lập ĐNTT ở 8.2.':`Còn thiếu: ${miss.map(c=>c[1]).join(', ')}.`),ready:!lech&&r2.ready}; }
    finally{ ED.hdBusy=false; render(); } }
  else if(a==='dn-make'){ if(ED.dnBusy) return; const bo=ED.dnBo||new Set(); const R=dnRows().filter(r=>r.ready&&!dnLive(r)&&!bo.has(hdKey(r))); if(!R.length) return; let n=0;
    ED.dnBusy=true; render();
    try{ for(const r of R){ const id=hdKey(r); await store.set('dntt/'+id,{thang:S.thang,ncc:r.po.ncc,k:r.ky.k,poId:r.po.id,soPO:r.po.soPO,tien:r.tien,noiDung:noiDung(r),loai:r.loai,soBB:r.bb,soHD:r.hd.soHD,ngayHD:r.hd.ngay,sheetMau:SHEETMAU[r.po.ncc.toUpperCase()]||'',migo:r.mg.map(m=>m.soMIGO||m.ma),tt:'lap',luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'lap',ai:uid,note:'Lập từ PO/MIGO'}]}); n++; }
      ED.dnsel=null; ED.dnBo=new Set(); ED.hdOk=null; ED.dnOk=`Đã lập ${n} ĐNTT · ${VND(R.reduce((s,r)=>s+r.tien,0))} đ. Bước tiếp: chọn ĐNTT và bấm Tạo file ĐNTT.`; await log('8.2',`Lập ${n} ĐNTT (${TR(R.reduce((s,r)=>s+r.tien,0))})`); }
    finally{ ED.dnBusy=false; render(); } }
  else if(a==='dn-copy'||a==='dn-xlsx'){ const R=dnRows().filter(r=>r.dn&&(r.dn.tt==='lap'||r.dn.tt==='xuat')&&(ED.dnsel||new Set()).has(r.dn.id)); const rows=dnSheetRows(R); if(!rows.length) return;
    if(a==='dn-copy'){ try{ await navigator.clipboard.writeText(rows.map(r=>r.join('\t')).join('\r\n')); }catch(err){ toast('Không copy được.'); return; } }
    else { if(!window.XLSX){toast('Không tạo được file; dùng nút Copy.');return;} const H=['STT','Nha_cung_cap','So_PO','So_tien_PO','Thang','Ky','Noi_dung_thanh_toan','Ngay_lap','Han_thanh_toan','So_bien_ban','Kiem_tra_ky','Sheet_mau','File_mau','Nguoi_thu_huong','File_DNTT_da_tao','Status'];
      const ws=XLSX.utils.aoa_to_sheet([H,...rows]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'DNTT');
      try{ await saveOut(`DNTT_${S.thang}_${todayVN()}.xlsx`,new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); }catch(err){ if(err&&err.code!=='cancelled') toast('Không lưu được file.'); return; } }
    for(const r of R) if(r.dn.tt==='lap') await dnSetTT(r.dn,'xuat',{},a==='dn-copy'?'Copy cho sheet DNTT':'Tải file cho sheet DNTT');
    ED.dnMacro=true; ED.dnOk=`${a==='dn-copy'?'Đã copy':'Đã tải file'} ${rows.length} dòng cho sheet DNTT (dán ô A3, Alt+F8 → TaoDNTT). Các ĐNTT chuyển sang “Đã tạo file ĐNTT”, trình ký ở 8.3.`; ED.dnsel=null;
    await log('8.2',`Xuất ${rows.length} ĐNTT cho sheet DNTT`); render(); }
  else if(a==='dn-next-ok'){ const d=S.dn[t.dataset.id], nx=d&&NXT8[d.tt]; if(!nx){ toast('ĐNTT chưa có file: tạo file ở 8.2 trước.'); return; }
    let extra={}, note=''; if(nx==='tt'){ const ng=($('#dn-ngtt')||{}).value||todayVN(), unc=(($('#dn-unc')||{}).value||'').trim(); if(ng<(d.ngayHD||'')){ toast('Ngày thanh toán trước ngày hóa đơn.'); return; } extra={ngayTT:ng,soUNC:unc}; note=`Thanh toán ${dm(ng)}${unc?' · UNC '+unc:''}`; }
    ED.dnAsk=null; await dnSetTT(d,nx,extra,note); await log('8.3',`ĐNTT ${d.ncc} ${mmY(d.thang)}${d.k?' kỳ '+d.k:''}: ${TTDN[nx][0]}`); toast(`${d.ncc}: ${TTDN[nx][0]}.`); render(); }
  else if(a==='dn-huy-ok'){ const d=S.dn[t.dataset.id]; if(!d||d.tt==='tt') return; const ly=(($('#dn-lydo')||{}).value||'').trim(); if(d.tt==='ketoan'&&!ly){ toast('ĐNTT đã chuyển kế toán: nhập lý do hủy.'); $('#dn-lydo')&&$('#dn-lydo').focus(); return; }
    ED.dnAsk=null; await dnSetTT(d,'huy',{lyDoHuy:ly},'Hủy'+(ly?': '+ly:'')); await log('8.3',`Hủy ĐNTT ${d.ncc}${ly?' – '+ly:''}`); toast(`Đã hủy ĐNTT ${d.ncc}. Sửa hóa đơn ở 8.1 rồi lập lại.`); render(); }
  else if(a==='dn-bulk-x'){ ED.dnBulk=false; render(); }
  else if(a==='dn-bulk'){ const f=ED.dnlsF, nx=NXT8[f]; if(!nx) return; if(!ED.dnBulk){ ED.dnBulk=true; ED.dnBulkNg=($('#dn-bulk-ng')||{}).value||''; render(); return; }
    const nf=ED.dnlsN||'', L=Object.values(S.dn||{}).filter(d=>(ED.dnAll||d.thang===S.thang)&&d.tt===f&&(!nf||d.ncc===nf)); const ng=($('#dn-bulk-ng')||{}).value||ED.dnBulkNg||todayVN();
    ED.dnBulk=false; for(const d of L) await dnSetTT(d,nx,nx==='tt'?{ngayTT:ng,soUNC:''}:{},nx==='tt'?`Thanh toán ${dm(ng)} (chuyển cả nhóm)`:'Chuyển cả nhóm');
    await log('8.3',`Chuyển ${L.length} ĐNTT sang ${TTDN[nx][0]}`); toast(`Đã chuyển ${L.length} ĐNTT sang “${TTDN[nx][0]}”.`); render(); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target; if(t.dataset.hd){ ED.hd=ED.hd||{}; const k=t.dataset.hd; ED.hd[k]=ED.hd[k]||{...((S.hd||{})[k]||{})}; ED.hd[k][t.dataset.f]=t.dataset.f==='tien'?(t.value.trim()===''?'':pnum(t.value.replace(/\./g,''))):t.value; const tr=t.closest('tr'); if(tr) tr.classList.add('dirty'); const b=tr&&tr.querySelector('[data-act8="hd-save"]'); if(b) b.classList.add('primary'); } });
document.addEventListener('change',e=>{ const t=e.target;
  if(t.dataset.hd&&t.dataset.f==='tien'&&t.value.trim()!==''){ t.value=VND(pnum(t.value.replace(/\./g,''))); }
  else if(t.dataset.dnsel){ ED.dnsel=ED.dnsel||new Set(); t.checked?ED.dnsel.add(t.dataset.dnsel):ED.dnsel.delete(t.dataset.dnsel); render(); }
  else if(t.id==='dnf-all'){ const F=dnRows().filter(r=>dnLive(r)&&(r.dn.tt==='lap'||r.dn.tt==='xuat')&&!lenhBusy(r.dn.lenh)); ED.dnsel=t.checked?new Set(F.map(r=>r.dn.id)):new Set(); render(); }
  else if(t.dataset.dnr){ ED.dnBo=ED.dnBo||new Set(); t.checked?ED.dnBo.delete(t.dataset.dnr):ED.dnBo.add(t.dataset.dnr); render(); }
  else if(t.id==='dnr-all'){ ED.dnBo=t.checked?new Set():new Set(dnRows().filter(r=>r.ready&&!dnLive(r)).map(hdKey)); render(); }
  else if(t.id==='hd-ncc'){ ED.hdN=t.value; render(); }
  else if(t.id==='dnls-ncc'){ ED.dnlsN=t.value; ED.dnBulk=false; render(); }
  else if(t.id==='dn-bulk-ng'){ ED.dnBulkNg=t.value; }
  else if(t.id==='dn-all'){ ED.dnAll=t.checked; ED.dnBulk=false; render(); } });
document.addEventListener('toggle',e=>{ if(e.target.id==='dn-macro') ED.dnMacro=e.target.open; },true);
