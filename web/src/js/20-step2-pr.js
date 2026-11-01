/* ================= BƯỚC 2 · LẬP PR ================= */
const SUB2=[['2.1','Phân bổ ngân sách','pb'],['2.2','Xuất PR lên SAP','xuat'],['2.3','Lịch sử PR','ls'],['2.4','PR so với nhận hàng','dc']];
const LIFNR={"NCC MẪU 23": "2900005", "NCC MẪU 26": "2900014", "NCC MẪU 11": "2900015", "NCC MẪU 05": "2900032", "NCC MẪU 02": "2900008", "NCC MẪU 03": "2900021", "NCC MẪU 04": "2900024", "NCC MẪU 35": "2900028", "NCC MẪU 06": "2900016", "NCC MẪU 36": "2900011", "NCC MẪU 12": "2900030", "NCC MẪU 09": "2900020", "NCC MẪU 10": "2900019", "NCC MẪU 13": "2900004", "NCC MẪU 14": "2900002", "NCC MẪU 16": "2900031", "NCC MẪU 15": "2900026", "NCC MẪU 18": "2900013", "NCC MẪU 19": "2900027", "NCC MẪU 21": "2900018", "NCC MẪU 22": "2900022", "NCC MẪU 24": "2900023", "NCC MẪU 27": "2900025", "NCC MẪU 28": "2900029", "NCC MẪU 17": "2900035", "NCC MẪU 32": "2900001", "NCC MẪU 07": "2900003", "NCC MẪU 01": "2900006", "NCC MẪU 30": "2900007", "NCC MẪU 31": "2900009", "NCC MẪU 33": "2900010", "NCC MẪU 34": "2900012", "NCC MẪU 29": "2900017", "NCC MẪU 08": "2900034", "NCC MẪU 20": "2900036", "NCC MẪU 25": "2900033"};
const SAPCFG={name1:'2200',ekorg:'1701',kostl:'2200115000',geber:'xanh',ekgrp:'105'};
const DTG={NhanVien:{ten:'Nhân viên',dispo:'be1',k:'NV'},BenhNhan:{ten:'Bệnh nhân',dispo:'be2',k:'BN'}};
const TTPR={nhap:['Nháp',''],xuat:['Đã xuất file','run'],sap:['Đã tạo trên SAP','run'],duyet:['Đã duyệt','ok'],huy:['Đã hủy','bad']};
const mmY=t=>t.slice(5)+'/'+t.slice(0,4), tK=t=>'T'+(+t.slice(5));
const isKG=dvt=>/^(KG|L)$/.test(dvt);
const todayVN=()=>{const d=new Date(Date.now()+7*3600e3);return d.toISOString().slice(0,10);};
const hasLifnr=n=>!!LIFNR[String(n||'').toUpperCase()];

function pbDefaultPct(){ const K=(S.khach&&S.khach.days)||{}; let nv=0,all=0; Object.values(K).forEach(r=>r.forEach((v,i)=>{all+=+v||0;if(i<2)nv+=+v||0;})); return all?Math.round(nv/all*1000)/10:50; }
function splitNV(sl,p,dvt){ let nv=isKG(dvt)?Math.round(sl*p*10)/10:Math.round(sl*p); return Math.min(sl,Math.max(0,nv)); }
function pbState(){
  if(ED.pb) return ED.pb;
  const base=(S.dutru&&S.dutru.lines)||[], sv=S.pb, old=new Map(((sv&&sv.lines)||[]).map(l=>[l.ma,l]));
  const P={mode:sv?sv.mode:'pct',pct:sv?sv.pct:pbDefaultPct(),amt:sv?sv.amt:0,dirty:false,
    lines:base.filter(l=>l.chot>0&&IT.has(l.ma)).map(l=>{const x=IT.get(l.ma),o=old.get(l.ma),g=priceOf(x,l.ncc);
      return (o&&o.sl===l.chot&&o.ncc===l.ncc)?{ma:l.ma,ncc:l.ncc,g,sl:l.chot,nv:o.nv,e:!!o.e}:{ma:l.ma,ncc:l.ncc,g,sl:l.chot,nv:null,e:false};})};
  ED.pb=P; pbApply(true); return P;
}
function pbRatio(P){ if(P.mode!=='amt') return Math.min(1,Math.max(0,(+P.pct||0)/100));
  const fx=P.lines.filter(l=>l.e), fxNV=fx.reduce((a,l)=>a+l.g*l.nv,0), rest=P.lines.filter(l=>!l.e).reduce((a,l)=>a+l.g*l.sl,0);
  return rest>0?Math.min(1,Math.max(0,((+P.amt||0)-fxNV)/rest)):0; }
function pbApply(onlyNull){ const P=ED.pb,p=pbRatio(P); P.lines.forEach(l=>{ if(l.e) return; if(onlyNull&&l.nv!=null) return; l.nv=splitNV(l.sl,p,IT.get(l.ma).dvt); }); }
function pbTotals(P){ let nv=0,bn=0; P.lines.forEach(l=>{nv+=l.g*l.nv;bn+=l.g*(l.sl-l.nv);}); return {nv,bn,tot:nv+bn}; }
const batchesOf=t=>Object.values(S.prdot||{}).filter(b=>b.thang===t).sort((a,b)=>a.luc-b.luc);
const prMain=()=>batchesOf(S.thang).filter(b=>b.loai==='Chính'&&b.tt!=='huy');
function head2(title,desc,right){ return `<div class="sh"><div><h2>${title}</h2>${desc?`<details class="help"${HELP_OPEN?' open':''}><summary>Cách làm</summary><p>${desc}</p></details>`:''}</div><div class="toolbar">${right||''}</div></div>`; }
const ttPill=b=>{const t=TTPR[b.tt]||['?',''];let late='';if(b.tt==='sap'&&b.sapLuc&&tstr(b.sapLuc).slice(0,5)!==tstr(Date.now()).slice(0,5))late=' <span class="pill bad">Quá ngày chưa duyệt</span>';return `<span class="pill ${t[1]}">${t[0]}</span>${late}`;};

/* ---- 2.1 Phân bổ ngân sách ---- */
function vPB(){
  const dt=S.dutru;
  if(!dt||!(dt.lines||[]).length) return head2('2.1 Phân bổ ngân sách','Chia dự trù tháng vào 2 mã ngân sách Nhân viên (be1) và Bệnh nhân (be2).')+'<div class="empty"><b>Chưa có dự trù tháng này</b>Làm bước 1.4 (tính dự trù) và 1.5 (chốt) trước.</div>';
  const P=pbState(),T=pbTotals(P),p=pbRatio(P), mainUsed=prMain().length>0;
  const q=(ED.pbq||'').toLowerCase(), nEd=P.lines.filter(l=>l.e).length, onlyEd=ED.pbEd&&nEd>0;
  const rows=P.lines.map((l,i)=>[l,i]).filter(([l])=>{const x=IT.get(l.ma);return (!onlyEd||l.e)&&(!q||(x.ma+' '+x.ten+' '+l.ncc).toLowerCase().includes(q));}).map(([l,i])=>{const x=IT.get(l.ma);
    return `<tr class="${l.e?'edited':''}"><td class="num ma">${x.ma}</td><td class="wrap nm">${esc(x.ten)}</td><td class="ncc">${esc(l.ncc)}${hasLifnr(l.ncc)?'':' <span class="pill bad">thiếu mã NCC SAP</span>'}</td><td class="dvt">${esc(x.dvt)}</td>
    <td class="r num slc">${Q(l.sl)}</td><td class="r num gia">${VND(l.g)}</td>
    <td class="r nv"><input class="q" type="text" inputmode="decimal" id="pb-${i}" data-pb="${i}" aria-label="SL Nhân viên · ${x.ma} · ${esc(x.ten)}" value="${QI(l.nv)}" ${canWrite&&!mainUsed?'':'disabled'}></td>
    <td class="r num bn">${Q(l.sl-l.nv)}</td><td class="r num tnv">${VND(l.g*l.nv)}</td><td class="r num tbn">${VND(l.g*(l.sl-l.nv))}</td></tr>`;}).join('');
  return head2('2.1 Phân bổ ngân sách','Chia số lượng dự trù đã chốt vào 2 mã ngân sách. Đặt tỷ lệ % hoặc số tiền cho mã Nhân viên, phần còn lại vào Bệnh nhân. Sửa từng dòng ở cột SL Nhân viên; dòng sửa tay được giữ khi chia lại.',
    S.pb?`<span class="pill ok">✓ Đã lưu ${tstr(S.pb.luc)} · ${esc(nm(S.pb.ai))}</span>`:'<span class="pill run">Chưa lưu</span>')+
  `${!locked()?'<div class="warn">Dự trù tháng chưa chốt (bước 1.5). Số ở đây sẽ đổi nếu dự trù còn sửa.</div>':''}
   ${mainUsed?'<div class="warn">Đã tạo đợt PR chính từ phân bổ này. Muốn sửa phân bổ, hủy đợt PR chính ở bước 2.3 (khi chưa tạo trên SAP).</div>':''}
   <div class="kpis"><div class="kv big"><b>Tổng PR tháng ${mmY(S.thang)}</b><div class="num">${TR(T.tot)}</div><small>${P.lines.length} mã hàng · giá gồm VAT</small></div>
    <div class="kv"><b>Mã Nhân viên (be1)</b><div class="num">${TR(T.nv)}</div><small>${T.tot?(T.nv/T.tot*100).toFixed(1).replace('.',','):0}% tổng</small></div>
    <div class="kv"><b>Mã Bệnh nhân (be2)</b><div class="num">${TR(T.bn)}</div><small>${T.tot?(T.bn/T.tot*100).toFixed(1).replace('.',','):0}% tổng</small></div>
    <div class="kv"><b>Tỷ lệ suất nhân viên</b><div class="num">${String(pbDefaultPct()).replace('.',',')}%</div><small>theo lượng khách dự kiến (1.2)</small></div></div>
   ${canWrite&&!mainUsed?`<div class="toolbar"><div class="seg"><button data-act2="pb-mode" data-m="pct" aria-pressed="${P.mode==='pct'}">Theo %</button><button data-act2="pb-mode" data-m="amt" aria-pressed="${P.mode==='amt'}">Theo số tiền</button></div>
    ${P.mode==='pct'?`<label class="meta" for="pb-pct">Nhân viên %</label><input class="n" type="text" inputmode="decimal" id="pb-pct" value="${String(P.pct).replace('.',',')}" style="width:80px">`:`<label class="meta" for="pb-amt">Tiền mã Nhân viên (đ)</label><input type="text" inputmode="numeric" id="pb-amt" value="${QI(P.amt)}" style="width:140px">`}
    <button class="btn sm" data-act2="pb-apply">Chia lại theo tỷ lệ</button><button class="btn sm" data-act2="pb-reset">Bỏ sửa tay</button>
    ${P.dirty?'<span class="meta"><span class="dirty">Có thay đổi chưa lưu</span></span>':''}<span class="sp"></span>
    <button class="btn primary" data-act2="pb-save" ${P.dirty||!S.pb?'':'disabled'}>Lưu phân bổ</button></div>`:''}
   <div class="toolbar">${nEd?`<div class="seg" role="group" aria-label="Dòng hiển thị"><button data-act2="pb-ed" data-v="1" aria-pressed="${!!onlyEd}">Đã sửa tay ${nEd}</button><button data-act2="pb-ed" data-v="0" aria-pressed="${!onlyEd}">Tất cả ${P.lines.length}</button></div>`:''}<input type="search" id="pb-q" placeholder="Tìm mã / tên / NCC" aria-label="Tìm mã, tên hoặc NCC" value="${esc(ED.pbq||'')}" style="width:220px"></div>
   <div class="tbl-wrap"><table class="t pb-tbl sheet"><colgroup><col class="c-ma"><col><col class="c-ncc"><col class="c-dvt"><col class="c-n"><col class="c-g"><col class="c-nv"><col class="c-n"><col class="c-t"><col class="c-t"></colgroup><thead><tr><th>Mã</th><th>Tên hàng</th><th>NCC</th><th>ĐVT</th><th class="r">SL chốt</th><th class="r">Đơn giá</th><th class="r">SL Nhân viên</th><th class="r">SL Bệnh nhân</th><th class="r">Tiền NV</th><th class="r">Tiền BN</th></tr></thead><tbody>${rows}</tbody>
   <tfoot><tr><td colspan="8">Tổng</td><td class="r num">${VND(T.nv)}</td><td class="r num">${VND(T.bn)}</td></tr></tfoot></table></div>`;
}
function makeBatch(loai,dt,lines,note){
  const t=S.thang, k=DTG[dt].k, id=t+'-'+(loai==='Chính'?'C':'B')+'-'+k+'-'+Date.now().toString(36);
  const nhom=loai==='Chính'?`PR-${tK(t)}-${k}`:`PRBS-${tK(t)}-${k}-${tstr(Date.now()).slice(0,5).replace('/','')}`;
  const L=lines.filter(l=>l.sl>0).map(l=>{const x=IT.get(l.ma);return {ma:l.ma,ten:x?x.ten:l.ten||'',ncc:l.ncc,dvt:x?x.dvt:'',sl:+(+l.sl).toFixed(3),g:l.g};});
  const tong=Math.round(L.reduce((a,l)=>a+l.sl*l.g,0));
  return {id,doc:{thang:t,loai,dt,nhom,lines:L,tong,tt:'nhap',soPR:'',luc:Date.now(),ai:uid,ghichu:note||'',hist:[{ts:Date.now(),tt:'nhap',ai:uid,note:'Tạo đợt'}]}};
}
async function setBatch(id,patch,note){ const b=S.prdot[id]; const d={...b,...patch,hist:[...(b.hist||[]),{ts:Date.now(),tt:patch.tt||b.tt,ai:uid,note:note||''}]}; delete d.id; await store.set('prdot/'+id,d); }

/* ---- 2.2 Xuất PR ---- */
function congNoRows(bs){ const out=[]; let i=0;
  bs.forEach(b=>b.lines.forEach(l=>out.push([++i,'',l.ncc,l.ma,l.ten,l.sl,l.g,b.dt,b.nhom,'',`PR trước ${mmY(b.thang)} · ${b.loai}`]))); return out; }
function vXuat(){
  const all=batchesOf(S.thang).filter(b=>b.tt==='nhap'||b.tt==='xuat');
  if(!ED.sel) ED.sel=new Set(all.map(b=>b.id));
  const sel=all.filter(b=>ED.sel.has(b.id)), rows=congNoRows(sel);
  const noL=[...new Set(sel.flatMap(b=>b.lines.map(l=>l.ncc)).filter(n=>!hasLifnr(n)))];
  const main=prMain();
  const cfg=`<details class="more"><summary>Cấu hình SAP (sheet Mapping)</summary><div class="meta">Mã cơ sở ${SAPCFG.name1} · EKORG ${SAPCFG.ekorg} · Cost center ${SAPCFG.kostl} · Quỹ ${SAPCFG.geber} · EKGRP ${SAPCFG.ekgrp} · NhanVien → be1, BenhNhan → be2 · Loại PR ZEX. Macro gom các dòng cùng Nhom_PR thành 1 PR.</div></details>`;
  const s1=main.length?`<p class="fdone">Đã tạo đợt chính: ${main.map(b=>'<b class="num">'+esc(b.nhom)+'</b> ('+(TTPR[b.tt]||[''])[0].toLowerCase()+')').join(', ')}</p>`
    :!S.pb?`<div class="toolbar"><span class="meta">Cần lưu phân bổ ngân sách ở 2.1 trước.</span><button class="btn sm" data-act="go-next" data-tab="pb">Mở 2.1</button></div>`
    :canWrite?`<div class="toolbar"><button class="btn primary" data-act2="mk-main" ${!(ED.pb&&ED.pb.dirty)?'':'disabled'}>Tạo đợt PR chính tháng ${mmY(S.thang)}</button>${ED.pb&&ED.pb.dirty?'<span class="meta">Lưu thay đổi phân bổ ở 2.1 trước.</span>':''}</div>`:'<p class="meta">Chưa tạo đợt chính.</p>';
  const s2=all.length?`<div class="tbl-wrap" style="max-height:none"><table class="t bt-tbl"><thead><tr><th><span class="sr">Chọn</span></th><th>Nhom_PR</th><th>Loại</th><th>Đối tượng</th><th class="r">Số dòng</th><th class="r">Tổng tiền</th><th>Trạng thái</th></tr></thead><tbody>
    ${all.map(b=>`<tr><td><input type="checkbox" data-sel="${b.id}" ${ED.sel.has(b.id)?'checked':''} aria-label="Chọn ${esc(b.nhom)}"></td><td class="num">${esc(b.nhom)}</td><td>${b.loai}</td><td>${DTG[b.dt].ten} (${DTG[b.dt].dispo})</td><td class="r num">${b.lines.length}</td><td class="r num">${VND(b.tong)}</td><td>${ttPill(b)}${lenhPill(b.lenh)}</td></tr>`).join('')}</tbody></table></div>
    ${noL.length?`<div class="warn">NCC chưa có mã SAP (LIFNR) trong sheet Mapping: <b>${noL.map(esc).join(', ')}</b>. PR vẫn tạo được, nhưng bước tạo PO sẽ báo thiếu mã NCC; bổ sung vào Mapping cột E:F.</div>`:''}${tachBox(sel)}`
    :'<p class="meta">Không có đợt nào chờ gửi. Đợt đã có số PR xem ở 2.3; PR bổ sung tạo ở bước 6.</p>';
  const s3=all.length?`${tramBox('PR')}${prSapBar(sel)}
    <div class="toolbar"><span class="meta">Hoặc chạy tay:</span><button class="btn sm" data-act2="xuat-xlsx" ${rows.length?'':'disabled'}>Tải file Excel cho sheet CongNo</button><button class="btn sm" data-act2="xuat-copy" ${rows.length?'':'disabled'}>Copy để dán vào CongNo</button><span class="meta">${rows.length} dòng · ${sel.length} đợt được chọn</span></div>
    ${rows.length?`<details class="more"><summary>Xem trước file CongNo (${rows.length} dòng)</summary><div class="tbl-wrap"><table class="t cn-tbl"><thead><tr>${['STT','So_hoa_don','Nha_cung_cap','Ma_vat_tu','Ten_hang','So_luong','Don_gia','Doi_tuong_su_dung','Nhom_PR','Khoa_nap','Can_cu'].map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>
    ${rows.slice(0,300).map(r=>`<tr>${r.map((v,j)=>`<td class="${j===5||j===6?'r num':''}">${j===6?VND(v):j===5?Q(v):esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div></details>`:''}`
    :'<p class="meta">Tạo hoặc chọn đợt trước.</p>';
  const sent=all.length&&all.every(b=>b.tt!=='nhap');
  return head2('2.2 Xuất PR lên SAP','Tạo đợt PR từ phân bổ, rồi xuất đúng 11 cột của sheet CongNo trong Template_PR_v3.xlsm. Bấm <b>Tạo PR trên SAP</b>: trạm SAP trên máy công ty tự điền CongNo và chạy CreatePR, số PR tự ghi về 2.3. Vẫn có thể tải file / copy để chạy tay.')+
  `<ol class="flow">
    <li class="${main.length?'done':'cur'}"><div class="fh"><span class="fn" aria-hidden="true">${main.length?'✓':'1'}</span><b>Tạo đợt PR</b></div>${s1}</li>
    <li class="${!main.length?'':all.length?'cur':'done'}"><div class="fh"><span class="fn" aria-hidden="true">2</span><b>Chọn đợt gửi đi</b></div>${s2}</li>
    <li class="${sent?'done':all.length?'cur':''}"><div class="fh"><span class="fn" aria-hidden="true">${sent?'✓':'3'}</span><b>Gửi lên SAP</b></div>${s3}</li>
   </ol>${cfg}`;
}
async function exportSel(kind){
  const bs=batchesOf(S.thang).filter(b=>(b.tt==='nhap'||b.tt==='xuat')&&ED.sel.has(b.id)), rows=congNoRows(bs); if(!rows.length) return;
  if(kind==='copy'){ const tsv=rows.map(r=>r.map((v,j)=>j===5||j===6?String(v).replace('.',','):String(v)).join('\t')).join('\r\n');
    try{ await navigator.clipboard.writeText(tsv); toast('Đã copy '+rows.length+' dòng (số lẻ dùng dấu phẩy). Dán vào sheet CongNo, ô A3.'); }catch(e){ toast('Không copy được, dùng nút Tải file Excel.'); return; } }
  else{ if(!window.XLSX){toast('Chưa tải được thư viện Excel, thử lại sau ít giây.');return;}
    const H=['STT','So_hoa_don','Nha_cung_cap','Ma_vat_tu','Ten_hang(tham khao)','So_luong','Don_gia','Doi_tuong_su_dung','Nhom_PR','Khoa_nap (macro NapCongNo)','Can_cu_khop_cong_no'];
    const ws=XLSX.utils.aoa_to_sheet([H,...rows]); ws['!cols']=[6,12,22,12,40,10,12,16,18,10,28].map(w=>({wch:w}));
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'CongNo');
    const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'}), fn=`PR_${S.thang}_${bs.map(b=>b.nhom).join('_').slice(0,60)}.xlsx`;
    try{ await saveOut(fn,new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); }catch(e){ if(e&&e.code!=='cancelled') toast('Không lưu được file: '+(e.message||e.code)); return; } }
  for(const b of bs) if(b.tt==='nhap') await setBatch(b.id,{tt:'xuat',xuatLuc:Date.now()},kind==='copy'?'Copy dán CongNo':'Xuất file Excel');
  await log('2.2',`Xuất ${rows.length} dòng PR (${bs.map(b=>b.nhom).join(', ')}) ${kind==='copy'?'bằng copy':'ra file Excel'} cho sheet CongNo`);
}

/* ---- 2.3 Lịch sử PR ---- */
/* ---- sửa đợt PR (chỉ khi chưa gửi trạm / chưa có số PR) ---- */
const prSuaDuoc=b=>canWrite&&(b.tt==='nhap'||b.tt==='xuat')&&!b.soPR&&!lenhBusy(b.lenh);
function prEditor(b){
  const R=ED.preRows, tong=R.reduce((a,l)=>a+(+l.sl||0)*(+l.g||0),0), goc=b.tong;
  return `<div class="box"><div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center"><b>Sửa đợt ${esc(b.nhom)}</b><span class="meta">${DTG[b.dt].ten} (${DTG[b.dt].dispo}) · ${R.length} dòng · ${VND(tong)} đ${Math.round(tong)!==Math.round(goc)?` <span class="dirty">(trước ${VND(goc)} đ)</span>`:''}</span></div>
   ${b.tt==='xuat'?'<div class="warn">Đợt này đã xuất file CongNo. Sau khi lưu, đợt quay về Nháp – nếu đã dán file cũ vào SAP thì xuất / gửi lại.</div>':''}
   <div class="tbl-wrap" style="max-height:420px"><table class="t"><thead><tr><th>Mã</th><th>Tên hàng</th><th>NCC</th><th class="r">Số lượng</th><th>ĐVT</th><th class="r">Đơn giá</th><th class="r">Tiền</th><th></th></tr></thead><tbody>
   ${R.map((l,i)=>{const x=IT.get(l.ma), ns=x?x.ncc.map(n=>n.ncc):[l.ncc];
     return `<tr${l.moi?' style="background:var(--edit)"':''}><td class="num">${esc(l.ma)}</td><td class="wrap">${esc(l.ten)}</td>
      <td>${ns.length>1?`<select data-pr="${i}" data-f="ncc">${ns.map(n=>`<option ${n===l.ncc?'selected':''}>${esc(n)}</option>`).join('')}</select>`:esc(l.ncc)}</td>
      <td class="r"><input class="q" type="text" inputmode="decimal" data-pr="${i}" data-f="sl" value="${QI(l.sl)}" style="width:90px;text-align:right"></td><td>${esc(l.dvt)}</td>
      <td class="r"><input class="q" type="text" inputmode="decimal" data-pr="${i}" data-f="g" value="${QI(l.g)}" style="width:100px;text-align:right"></td>
      <td class="r num">${VND((+l.sl||0)*(+l.g||0))}</td><td><button class="x" data-pre="xoa" data-i="${i}" aria-label="Xóa dòng ${esc(l.ma)}">×</button></td></tr>`;}).join('')}</tbody></table></div>
   <div class="toolbar"><input type="text" id="pre-ma" list="dl-item" placeholder="Thêm mặt hàng: gõ mã hoặc tên" style="flex:1;min-width:220px"><input type="text" id="pre-sl" inputmode="decimal" placeholder="Số lượng" style="width:100px"><button class="btn sm" data-pre="them">+ Thêm</button>
    <span class="sp"></span><button class="btn sm" data-pre="huy">Bỏ thay đổi</button><button class="btn sm primary" data-pre="luu">Lưu đợt PR</button></div>
   <div class="meta">Dòng số lượng 0 sẽ bị bỏ khi lưu. Đổi NCC thì đơn giá lấy theo giá khung của NCC đó (sửa lại được). Mọi thay đổi ghi vào lịch sử đợt và nhật ký.</div></div>`; }
function preFind(q){ q=String(q||'').trim(); if(!q) return null; const u=q.toUpperCase(); if(IT.has(u)) return IT.get(u);
  const m=q.match(/^([A-Z0-9 .\-]+?)\s*[–(]/i); if(m&&IT.has(m[1].trim().toUpperCase())) return IT.get(m[1].trim().toUpperCase());
  const s=slug(q); return ITEMS.find(x=>slug(x.ten)===s)||ITEMS.find(x=>slug(x.ten).includes(s))||null; }
document.addEventListener('input',e=>{ const el=e.target.closest('[data-pr]'); if(!el||!ED.preRows) return; const l=ED.preRows[+el.dataset.pr]; if(!l) return;
  if(el.dataset.f==='sl'||el.dataset.f==='g'){ l[el.dataset.f]=pnum(el.value)||0; const td=el.closest('tr').querySelector('td:nth-child(7)'); if(td) td.textContent=VND((+l.sl||0)*(+l.g||0)); } });
document.addEventListener('change',e=>{ const el=e.target.closest('select[data-pr]'); if(!el||!ED.preRows) return; const l=ED.preRows[+el.dataset.pr]; l.ncc=el.value; const x=IT.get(l.ma); if(x) l.g=priceOf(x,l.ncc)||l.g; render(); });
document.addEventListener('click',async e=>{ const t=e.target.closest('[data-pre]'); if(!t) return; const a=t.dataset.pre;
  try{
  if(a==='mo'){ const b=S.prdot[t.dataset.id]; if(!b||!prSuaDuoc(b)){ toast('Đợt này không sửa được (đã gửi trạm hoặc đã có số PR).'); return; } ED.preId=b.id; ED.preRows=b.lines.map(l=>({...l})); render(); setTimeout(()=>{const x=$('#pre-box'); if(x) x.scrollIntoView({behavior:'smooth',block:'start'});},50); return; }
  if(!ED.preId) return; const b=S.prdot[ED.preId];
  if(a==='huy'){ ED.preId=null; ED.preRows=null; render(); }
  else if(a==='xoa'){ ED.preRows.splice(+t.dataset.i,1); render(); }
  else if(a==='them'){ const x=preFind($('#pre-ma').value), sl=pnum($('#pre-sl').value)||0; if(!x){ toast('Không tìm thấy mặt hàng trong danh mục.'); return; }
    if(ED.preRows.some(l=>l.ma===x.ma)){ toast('Mặt hàng đã có trong đợt – sửa số lượng ở dòng đó.'); return; }
    const ncc=x.ncc[0]?x.ncc[0].ncc:''; ED.preRows.push({ma:x.ma,ten:x.ten,ncc,dvt:x.dvt,sl:+sl.toFixed(3),g:priceOf(x,ncc)||0,moi:true}); render(); }
  else if(a==='luu'){ if(!b||!prSuaDuoc(b)){ toast('Đợt đã đổi trạng thái, không lưu được.'); return; }
    const L=ED.preRows.filter(l=>(+l.sl||0)>0).map(({moi,...l})=>({...l,sl:+(+l.sl).toFixed(3),g:Math.round(+l.g||0)}));
    if(!L.length){ toast('Đợt không còn dòng nào – dùng nút Hủy nếu muốn bỏ đợt.'); return; }
    const old=new Map(b.lines.map(l=>[l.ma,l])), neu=new Map(L.map(l=>[l.ma,l])), notes=[];
    L.forEach(l=>{ const o=old.get(l.ma); if(!o) notes.push(`+${l.ma} ${Q(l.sl)}`); else { if(Math.abs(o.sl-l.sl)>1e-9) notes.push(`${l.ma} ${Q(o.sl)}→${Q(l.sl)}`); if(o.g!==l.g) notes.push(`${l.ma} giá ${VND(o.g)}→${VND(l.g)}`); if(o.ncc!==l.ncc) notes.push(`${l.ma} NCC ${o.ncc}→${l.ncc}`); } });
    b.lines.forEach(o=>{ if(!neu.has(o.ma)) notes.push(`−${o.ma}`); });
    if(!notes.length){ ED.preId=null; ED.preRows=null; render(); toast('Không có thay đổi.'); return; }
    const tong=Math.round(L.reduce((a,l)=>a+l.sl*l.g,0)), note='Sửa đợt: '+notes.slice(0,12).join('; ')+(notes.length>12?` … (+${notes.length-12})`:'');
    await setBatch(b.id,{lines:L,tong,tt:'nhap'},note); await log('2.3',`Sửa đợt ${b.nhom}: ${notes.length} thay đổi, ${VND(b.tong)} → ${VND(tong)} đ`);
    ED.preId=null; ED.preRows=null; toast(`Đã lưu đợt ${b.nhom}.`); }
  }catch(err){ toast('Không lưu được: '+((err&&(err.message||err.code))||err)); } });
function vLS(){
  const all=Object.values(S.prdot||{}).filter(b=>ED.lsAll||b.thang===S.thang).sort((a,b)=>b.luc-a.luc);
  const body=all.map(b=>{const open=ED.lsOpen===b.id, cho=b.tt==='nhap'||b.tt==='xuat';
    const acts=canWrite?[prSuaDuoc(b)?`<button class="btn sm" data-pre="mo" data-id="${b.id}">✎ Sửa</button>`:'',b.tt==='xuat'||b.tt==='sap'?`<button class="btn sm" data-act2="ls-sopr" data-id="${b.id}">Lưu số PR</button>`:'',b.tt==='sap'?`<button class="btn sm primary" data-act2="ls-duyet" data-id="${b.id}">Đã duyệt</button>`:'',b.tt!=='duyet'&&b.tt!=='huy'?`<button class="btn sm danger-text" data-act2="ls-huy" data-id="${b.id}">Hủy</button>`:''].join(''):'';
    return `<tr><td class="exp"><button class="x" data-act2="ls-open" data-id="${b.id}" aria-expanded="${open}" aria-label="${open?'Ẩn':'Xem'} chi tiết ${esc(b.nhom)}">${open?'▾':'▸'}</button></td>
    <td class="dg"><b class="num">${esc(b.nhom)}</b><small>${b.loai} · ${DTG[b.dt].ten}${ED.lsAll?' · '+mmY(b.thang):''}</small></td>
    <td class="r tien"><b class="num">${VND(b.tong)}</b><small>${b.lines.length} dòng</small></td>
    <td class="tt">${ttPill(b)}${lenhPill(b.lenh)}${cho&&!b.lenh?'<small>gửi ở 2.2</small>':''}${b.tt==='xuat'||b.tt==='sap'?`<input type="text" id="sopr-${b.id}" value="${esc(b.soPR||'')}" placeholder="Số PR" aria-label="Số PR của ${esc(b.nhom)}" ${canWrite?'':'disabled'}>`:b.soPR?`<small class="num">PR ${esc(b.soPR)}</small>`:''}</td>
    <td class="meta when">Tạo ${tstr(b.luc)} · ${esc(nm(b.ai))}${b.duyetLuc?`<br>Duyệt ${tstr(b.duyetLuc)}`:''}</td>
    <td class="acts"><div class="acts-in">${acts}</div></td></tr>
    ${open?`<tr class="detail"><td colspan="6"><div class="two" style="grid-template-columns:minmax(0,2fr) minmax(0,1fr)"><div class="tbl-wrap" style="max-height:300px"><table class="t"><thead><tr><th>Mã</th><th>Tên</th><th>NCC</th><th class="r">SL</th><th class="r">Đơn giá</th><th class="r">Tiền</th></tr></thead><tbody>${b.lines.map(l=>`<tr><td class="num">${l.ma}</td><td class="wrap">${esc(l.ten)}</td><td>${esc(l.ncc)}</td><td class="r num">${Q(l.sl)} ${esc(l.dvt)}</td><td class="r num">${VND(l.g)}</td><td class="r num">${VND(l.sl*l.g)}</td></tr>`).join('')}</tbody></table></div>
      <div><b class="meta">Lịch sử đợt</b><ol style="margin:6px 0 0;padding-left:18px;font-size:13px">${(b.hist||[]).map(h=>`<li><span class="num meta">${tstr(h.ts)}</span> · ${(TTPR[h.tt]||[h.tt])[0]} · ${esc(nm(h.ai))}${h.note?' · '+esc(h.note):''}</li>`).join('')}</ol>${b.ghichu?`<p class="meta">${esc(b.ghichu)}</p>`:''}</div></div></td></tr>`:''}`;}).join('');
  const m=batchesOf(S.thang).filter(b=>b.tt!=='huy'), tong=m.reduce((a,b)=>a+b.tong,0), duyet=m.filter(b=>b.tt==='duyet');
  const choGui=m.filter(b=>b.tt==='nhap'||b.tt==='xuat').length, choDuyet=m.filter(b=>b.tt==='sap').length;
  return head2('2.3 Lịch sử tạo PR','Mọi đợt PR (chính và bổ sung) với trạng thái: nháp → đã xuất file → đã tạo trên SAP (ghi số PR) → đã duyệt. Đợt chưa có số PR và chưa gửi trạm SAP thì bấm <b>Sửa</b> để sửa số lượng, giá, NCC, thêm hoặc xóa mặt hàng. PR phải được duyệt trong ngày tạo.')+
  (m.length?`<div class="kpis"><div class="kv big"><b>PR tháng ${mmY(S.thang)}</b><div class="num">${TR(tong)}</div><small>${m.length} đợt · ${m.filter(b=>b.loai==='Bổ sung').length} bổ sung</small></div>
   <div class="kv"><b>Đã duyệt</b><div class="num">${duyet.length} / ${m.length} đợt</div><small>${TR(duyet.reduce((a,b)=>a+b.tong,0))}</small></div>
   <div class="kv"><b>Chờ gửi SAP</b><div class="num">${choGui} đợt</div><small>${choDuyet} đợt đã tạo, chờ duyệt</small></div></div>`:'')+
  `${ED.preId&&S.prdot[ED.preId]?`<div id="pre-box">${prEditor(S.prdot[ED.preId])}</div>`:''}
  <div class="toolbar"><label class="meta"><input type="checkbox" id="ls-all" ${ED.lsAll?'checked':''}> Hiện cả các tháng khác</label></div>
  ${all.length?`<div class="tbl-wrap" style="max-height:none"><table class="t ls-tbl"><colgroup><col class="c-exp"><col><col class="c-tien"><col class="c-tt"><col class="c-when"><col class="c-acts"></colgroup><thead><tr><th><span class="sr">Chi tiết</span></th><th>Đợt</th><th class="r">Tiền</th><th>Trạng thái</th><th>Tạo / duyệt</th><th><span class="sr">Thao tác</span></th></tr></thead><tbody>${body}</tbody></table></div>`
  :'<div class="empty"><b>Chưa có đợt PR nào</b>Tạo đợt PR chính ở bước 2.2.<div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn sm" data-act="go-next" data-tab="xuat">Mở 2.2</button></div></div>'}`;
}

/* ---- 2.4 PR so với nhận hàng ---- */
const DCST={chua_pr:['Chưa có PR','bad',1],vuot:['Vượt PR','bad',2],sap_vuot:['Dự kiến vượt','run',3],thua:['Dự kiến dư','run',5],chua_nhan:['Chưa nhận','',6],dang:['Đang nhận','ok',4],du:['Đã nhận đủ','ok',7]};
function dcCompute(){
  const okT=ED.dcAll?['xuat','sap','duyet']:['sap','duyet'];
  const pr=new Map(); batchesOf(S.thang).filter(b=>okT.includes(b.tt)).forEach(b=>b.lines.forEach(l=>{const o=pr.get(l.ma)||{sl:0,g:l.g,ncc:l.ncc};o.sl+=l.sl;pr.set(l.ma,o);}));
  const rv=recvRows(), got=new Map(); let last='';
  rv.forEach(r=>{const o=got.get(r.ma)||{sl:0,ncc:r.ncc,n:0};o.sl+=+r.sl||0;o.n++;got.set(r.ma,o);if(r.d>last)last=r.d;});
  const days=monthDays(S.thang), dim=days.length, td=todayVN();
  let asOf=last||td; if(asOf<days[0]) asOf=''; if(asOf>days[dim-1]) asOf=days[dim-1];
  const frac=asOf?(+asOf.slice(8))/dim:0;
  const rows=[...new Set([...pr.keys(),...got.keys()])].map(ma=>{const x=IT.get(ma)||{ma,ten:ma,dvt:'',ncc:[]},p=pr.get(ma),r=got.get(ma);
    const sP=p?p.sl:0,sR=r?r.sl:0,db=frac>0?sR/frac:sR,ncc=(p&&p.ncc)||(r&&r.ncc)||'',g=(p&&p.g)||priceOf(x,ncc)||0;
    let st; if(!sP&&sR>0) st='chua_pr'; else if(sR>sP+1e-9) st='vuot'; else if(frac>0&&frac<1&&sR>0&&db>sP*1.0001) st='sap_vuot'; else if(sR>=sP-1e-9&&sP>0) st='du';
    else if(sR===0) st='chua_nhan'; else if(frac>=0.5&&db<sP*0.8) st='thua'; else st='dang';
    const need=Math.max(0,Math.max(sR,frac>0&&frac<1?db:sR)-sP);
    return {ma,x,ncc,g,sP,sR,db,st,need:need>0?roundQ(need,x.dvt):0,du:Math.max(0,sP-Math.max(sR,frac<1?db:sR))};});
  rows.sort((a,b)=>DCST[a.st][2]-DCST[b.st][2]||b.g*b.need-a.g*a.need);
  return {rows,asOf,frac,nRecv:rv.length};
}
function vDC(){
  const C=dcCompute(), f=ED.dcF||'';
  const cnt={}; C.rows.forEach(r=>cnt[r.st]=(cnt[r.st]||0)+1);
  const need=C.rows.filter(r=>r.need>0), tPR=C.rows.reduce((a,r)=>a+r.sP*r.g,0), tR=C.rows.reduce((a,r)=>a+r.sR*r.g,0);
  const shown=C.rows.filter(r=>!f||r.st===f);
  ED.bs=ED.bs||{};
  const body=shown.slice(0,400).map(r=>{const s=DCST[r.st],pct=r.sP?r.sR/r.sP*100:0,bs=ED.bs[r.ma]!=null?ED.bs[r.ma]:r.need;
    return `<tr><td class="num">${r.ma}</td><td class="wrap">${esc(r.x.ten)}</td><td>${esc(r.ncc)}</td><td>${esc(r.x.dvt)}</td><td class="r num">${Q(r.sP)}</td><td class="r num">${Q(r.sR)}</td>
    <td class="r num">${r.sP?Math.round(pct)+'%':'–'}${pct>0?`<div class="bar" style="width:${Math.min(100,pct)}%;background:var(${pct>100?'--bad':'--accent'})"></div>`:''}</td>
    <td class="r num">${C.frac>0&&C.frac<1?Q(Math.round(r.db*10)/10):'–'}</td><td><span class="pill ${s[1]}">${s[0]}</span></td>
    <td class="r num">${r.need>0?Q(r.need):''}</td>
    <td class="r num">${r.du>0&&(r.st==='thua'||C.frac>=1)?Q(Math.round(r.du*10)/10):''}</td></tr>`;}).join('');
  return head2('2.4 PR so với nhận hàng','So số lượng trên PR (đã tạo trên SAP / đã duyệt) với số thực nhận (phiếu MIGO). Dự báo cuối tháng = đã nhận ÷ phần tháng đã qua. Báo cáo theo dõi; PR bổ sung chỉ tạo một lần cuối chu kỳ ở bước 6 sau khi số nhận đã khớp với NCC.',
    '<span class="pill">Theo dõi, không bắt buộc</span>')+
  (!C.rows.length&&!ED.nhPaste?`<div class="empty"><b>Chưa có gì để so</b>Cần ít nhất một PR đã tạo trên SAP (2.3) hoặc số nhận hàng. Bước 4 (MIGO từng phần) sẽ tự ghi số nhận vào đây.<div class="toolbar" style="justify-content:center;margin-top:12px"><label class="meta"><input type="checkbox" id="dc-all" ${ED.dcAll?'checked':''}> Tính cả PR mới xuất file</label>${canWrite?'<button class="btn sm" data-act2="nh-paste">Dán thực nhận</button>':''}</div></div>`:`<div class="kpis"><div class="kv big"><b>Cần PR bổ sung</b><div class="num">${need.length} mã</div><small>${TR(need.reduce((a,r)=>a+(ED.bs[r.ma]!=null?ED.bs[r.ma]:r.need)*r.g,0))}</small></div>
   <div class="kv"><b>Tiền theo PR</b><div class="num">${TR(tPR)}</div><small>${C.rows.filter(r=>r.sP>0).length} mã</small></div>
   <div class="kv"><b>Đã nhận (theo giá PR)</b><div class="num">${TR(tR)}</div><small>${C.nRecv} dòng nhận · tới ${C.asOf?dm(C.asOf):'—'}</small></div>
   <div class="kv"><b>Dự kiến dư cuối tháng</b><div class="num">${TR(C.rows.reduce((a,r)=>a+(r.st==='thua'||C.frac>=1?r.du*r.g:0),0))}</div><small>đóng phần dư ở bước 6</small></div></div>
  <div class="toolbar"><div class="seg"><button data-act2="dc-f" data-f="" aria-pressed="${!f}">Tất cả ${C.rows.length}</button>${Object.entries(DCST).filter(([k])=>cnt[k]).map(([k,v])=>`<button data-act2="dc-f" data-f="${k}" aria-pressed="${f===k}">${v[0]} ${cnt[k]}</button>`).join('')}</div><span class="sp"></span>
   <label class="meta"><input type="checkbox" id="dc-all" ${ED.dcAll?'checked':''}> Tính cả PR mới xuất file</label>${canWrite?`<button class="btn sm" data-act2="nh-paste">Dán thực nhận</button>`:''}</div>
  ${ED.nhPaste?`<div class="box"><label class="meta" for="nh-txt">Tạm thời trước khi có bước 4 (MIGO): dán các dòng <b>Ngày</b> ⇥ <b>Mã hàng</b> ⇥ <b>SL thực nhận</b> ⇥ NCC ⇥ Số MIGO. Ngày dạng dd/mm/yyyy; mã phải có trong bảng giá.</label><textarea class="paste" id="nh-txt"></textarea>
    <div class="toolbar"><button class="btn sm primary" data-act2="nh-ok">Thêm vào nhận hàng tháng ${mmY(S.thang)}</button><button class="btn sm" data-act2="nh-paste">Đóng</button>${(S.nhan&&S.nhan.rows||[]).length?'<span class="sp"></span><button class="btn sm warn" data-act2="nh-clear">Xóa toàn bộ số nhận tháng này</button>':''}</div></div>`:''}
  <div class="meta">${S.nhan?`Số nhận: ${(S.nhan.rows||[]).length} dòng · nguồn ${esc(S.nhan.nguon||'')} · cập nhật ${tstr(S.nhan.luc)}`:'Chưa có số nhận hàng tháng này. Bước 4 (MIGO từng phần) sẽ tự ghi vào đây; trước đó có thể dán tay.'}</div>
  ${C.rows.length?`<div class="tbl-wrap"><table class="t"><thead><tr><th>Mã</th><th>Tên hàng</th><th>NCC</th><th>ĐVT</th><th class="r">SL PR</th><th class="r">Đã nhận</th><th class="r">% nhận</th><th class="r">Dự báo cuối tháng</th><th>Tình trạng</th><th class="r">Dự kiến thiếu PR</th><th class="r">Dư</th></tr></thead><tbody>${body}</tbody></table></div>`
  :'<div class="empty"><b>Chưa có gì để so</b>Cần ít nhất một PR đã tạo trên SAP (2.3) hoặc số nhận hàng.</div>'}`);
}
async function makeBS(){
  const C=dcCompute(), pick=C.rows.map(r=>({r,q:ED.bs&&ED.bs[r.ma]!=null?ED.bs[r.ma]:r.need})).filter(o=>o.q>0);
  if(!pick.length){toast('Không có dòng nào cần bổ sung.');return;}
  const p=S.pb?pbRatio({...S.pb,lines:(S.pb.lines||[])}):pbDefaultPct()/100;
  const nv=[],bn=[]; pick.forEach(({r,q})=>{const a=splitNV(q,p,r.x.dvt);if(a>0)nv.push({ma:r.ma,ncc:r.ncc,sl:a,g:r.g});if(q-a>0)bn.push({ma:r.ma,ncc:r.ncc,sl:+(q-a).toFixed(3),g:r.g});});
  const note=`Bổ sung từ đối chiếu PR–nhận hàng tới ${C.asOf?dm(C.asOf):''}: ${pick.length} mã`; let n=0;
  for(const [dt,L] of [['NhanVien',nv],['BenhNhan',bn]]) if(L.length){const b=makeBatch('Bổ sung',dt,L,note); await store.set('prdot/'+b.id,b.doc); n++;}
  ED.bs={}; ED.sel=null; await log('2.4',`Tạo ${n} đợt PR bổ sung cho ${pick.length} mã (${TR(pick.reduce((a,o)=>a+o.q*o.r.g,0))})`); toast('Đã tạo PR bổ sung. Sang 2.2 để xuất lên SAP.'); S.tab='xuat'; render();
}

/* ---- sự kiện bước 2 ---- */
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act2]'); if(!t) return; const a=t.dataset.act2;
  try{
  if(a==='pb-ed'){ ED.pbEd=t.dataset.v==='1'; render(); return; }
  if(a==='pb-mode'){ const P=pbState(); P.mode=t.dataset.m; if(P.mode==='amt'&&!P.amt) P.amt=Math.round(pbTotals(P).nv); P.dirty=true; render(); }
  else if(a==='pb-apply'){ pbApply(false); ED.pb.dirty=true; render(); }
  else if(a==='pb-reset'){ ED.pb.lines.forEach(l=>l.e=false); pbApply(false); ED.pb.dirty=true; render(); }
  else if(a==='pb-save'){ const P=ED.pb, T=pbTotals(P); await store.set('pb/'+S.thang,{mode:P.mode,pct:+P.pct||0,amt:+P.amt||0,lines:P.lines.map(l=>({ma:l.ma,ncc:l.ncc,g:l.g,sl:l.sl,nv:l.nv,e:l.e})),tongNV:Math.round(T.nv),tongBN:Math.round(T.bn),luc:Date.now(),ai:uid});
    P.dirty=false; ED.pb=null; await log('2.1',`Lưu phân bổ: Nhân viên ${TR(T.nv)} (${(T.nv/T.tot*100).toFixed(1)}%), Bệnh nhân ${TR(T.bn)}`); toast('Đã lưu phân bổ.'); }
  else if(a==='mk-main'){ const P=pbState(); if(P.dirty){toast('Lưu phân bổ trước.');return;}
    const nv=P.lines.map(l=>({ma:l.ma,ncc:l.ncc,sl:l.nv,g:l.g})), bn=P.lines.map(l=>({ma:l.ma,ncc:l.ncc,sl:+(l.sl-l.nv).toFixed(3),g:l.g}));
    for(const [dt,L] of [['NhanVien',nv],['BenhNhan',bn]]){ const b=makeBatch('Chính',dt,L,locked()?'Từ dự trù đã chốt':'Dự trù CHƯA chốt lúc tạo'); if(b.doc.lines.length) await store.set('prdot/'+b.id,b.doc); }
    ED.sel=null; await log('2.2',`Tạo đợt PR chính tháng ${mmY(S.thang)}: PR-${tK(S.thang)}-NV và PR-${tK(S.thang)}-BN`); toast('Đã tạo đợt PR chính.'); }
  else if(a==='xuat-xlsx'){ await exportSel('xlsx'); }
  else if(a==='xuat-copy'){ await exportSel('copy'); }
  else if(a==='ls-open'){ ED.lsOpen=ED.lsOpen===t.dataset.id?null:t.dataset.id; render(); }
  else if(a==='ls-sopr'){ const id=t.dataset.id, v=($('#sopr-'+id).value||'').trim(); if(!/^\d{6,12}$/.test(v)){toast('Số PR là dãy 6–12 chữ số.');return;}
    await setBatch(id,{soPR:v,tt:'sap',sapLuc:Date.now()},'Ghi số PR '+v); await log('2.3',`Ghi số PR ${v} cho ${S.prdot[id].nhom}`); }
  else if(a==='ls-duyet'){ const id=t.dataset.id,b=S.prdot[id]; const late=b.sapLuc&&tstr(b.sapLuc).slice(0,5)!==tstr(Date.now()).slice(0,5);
    await setBatch(id,{tt:'duyet',duyetLuc:Date.now()},late?'Duyệt trễ (khác ngày tạo)':'Duyệt trong ngày'); await log('2.3',`PR ${b.soPR} (${b.nhom}) đã duyệt${late?' – trễ ngày':''}`); }
  else if(a==='ls-huy'){ const id=t.dataset.id,b=S.prdot[id]; if(ED.huyAsk!==id){ED.huyAsk=id;toast('Bấm Hủy lần nữa để xác nhận hủy '+b.nhom+'.');return;} ED.huyAsk=null;
    await setBatch(id,{tt:'huy'},'Hủy đợt'); await log('2.3',`Hủy đợt ${b.nhom}${b.soPR?' (PR '+b.soPR+')':''}`); }
  else if(a==='dc-f'){ ED.dcF=t.dataset.f; render(); }
  else if(a==='nh-paste'){ ED.nhPaste=!ED.nhPaste; render(); }
  else if(a==='nh-ok'){ const rows=[],bad=[]; for(const line of $('#nh-txt').value.split(/\r?\n/)){ if(!line.trim()) continue; const c=line.split('\t').map(x=>x.trim());
      const m=(c[0]||'').match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); const ma=(c[1]||'').toUpperCase(), sl=pnum(c[2]);
      if(!m||!ma||!(sl>0)){bad.push(line);continue;} const d=`${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; if(d.slice(0,7)!==S.thang){bad.push(line);continue;}
      const x=IT.get(ma); if(!x){bad.push(line);continue;} rows.push({d,ma,sl,ncc:(c[3]||'').toUpperCase()||(x&&x.ncc[0]?x.ncc[0].ncc:''),migo:c[4]||'',nguon:'dán tay'}); }
    if(rows.length){ const old=(S.nhan&&S.nhan.rows)||[]; await store.set('nhanhang/'+S.thang,{rows:[...old,...rows],luc:Date.now(),ai:uid,nguon:(S.nhan&&S.nhan.nguon?S.nhan.nguon+' + ':'')+'dán tay'}); await log('2.4',`Thêm ${rows.length} dòng thực nhận (dán tay)`); }
    ED.nhPaste=bad.length>0; render(); if(ED.nhPaste) $('#nh-txt').value=bad.join('\n'); toast(`Đã thêm ${rows.length} dòng${bad.length?`, ${bad.length} dòng không đọc được hoặc khác tháng`:''}.`); }
  else if(a==='nh-clear'){ if(ED.clrAsk!==1){ED.clrAsk=1;toast('Bấm lần nữa để xóa toàn bộ số nhận tháng này.');return;} ED.clrAsk=0; await store.del('nhanhang/'+S.thang); await log('2.4','Xóa toàn bộ số nhận hàng tháng'); }
  else if(a==='mk-bs'){ await makeBS(); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target;
  if(t.dataset.pb!=null){ const l=ED.pb.lines[+t.dataset.pb]; l.nv=Math.min(l.sl,Math.max(0,pnum(t.value))); l.e=true; ED.pb.dirty=true; clearTimeout(vPB.h); vPB.h=setTimeout(render,700); }
  else if(t.id==='pb-pct'){ ED.pb.pct=pnum(t.value); pbApply(false); ED.pb.dirty=true; clearTimeout(vPB.h); vPB.h=setTimeout(render,500); }
  else if(t.id==='pb-amt'){ ED.pb.amt=pnum(String(t.value).replace(/\./g,'')); pbApply(false); ED.pb.dirty=true; clearTimeout(vPB.h); vPB.h=setTimeout(render,600); }
  else if(t.id==='pb-q'){ ED.pbq=t.value; clearTimeout(vPB.q); vPB.q=setTimeout(render,250); }
  else if(t.dataset.bs){ ED.bs=ED.bs||{}; ED.bs[t.dataset.bs]=Math.max(0,pnum(t.value)); clearTimeout(vDC.h); vDC.h=setTimeout(render,700); }
});
document.addEventListener('change',e=>{ const t=e.target;
  if(t.dataset.sel){ ED.sel=ED.sel||new Set(); t.checked?ED.sel.add(t.dataset.sel):ED.sel.delete(t.dataset.sel); render(); }
  else if(t.id==='ls-all'){ ED.lsAll=t.checked; render(); }
  else if(t.id==='dc-all'){ ED.dcAll=t.checked; render(); }
});

