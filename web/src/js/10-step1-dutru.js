/* ---------- 1.1 THỰC ĐƠN TUẦN ---------- */
function tdKey(){return ED.nhom+'_'+ED.week;}
function tdRows(){ if(ED.tdKey!==tdKey()||!ED.td){ED.tdKey=tdKey();const d=S.td[ED.tdKey];ED.td=d?(d.items||[]).map(x=>({...x,monGoc:x.mon})):[];ED.tdDirty=false;} return ED.td; }
function vTD(){
  const ws=weeksOf(S.thang); if(!ED.week||!ws.includes(ED.week)) ED.week=ws.find(w=>!S.td[ED.nhom+'_'+w])||ws[0];
  const rows=tdRows(), doc=S.td[tdKey()], days=Array.from({length:7},(_,i)=>addD(ED.week,i)), N=NHOM[ED.nhom];
  const known=[...new Set(Object.values(S.td).flatMap(d=>(d.items||[]).map(x=>x.mon)))].sort((a,b)=>a.localeCompare(b,'vi'));
  const ord=r=>r.d+N.bua.indexOf(r.b);
  const idx=rows.map((r,i)=>i).sort((a,b)=>ord(rows[a])<ord(rows[b])?-1:ord(rows[a])>ord(rows[b])?1:0);
  let lastD='',body='';
  for(const i of idx){const r=rows[i];
    if(r.d!==lastD){lastD=r.d;body+=`<tr class="day"><td colspan="7">${DOW[P(r.d).getUTCDay()]} ${dm(r.d)}${r.d.slice(0,7)!==S.thang?' <span class="muted">(ngoài tháng, không tính)</span>':''}</td></tr>`;}
    const has=S.dm[slug(r.mon)], hasR=r.nl&&r.nl.length;
    body+=`<tr${r.d.slice(0,7)!==S.thang?' class="out"':''}><td><select data-td="${i}" data-f="d" aria-label="Ngày" ${W()?'':'disabled'}>${days.map(d=>`<option value="${d}" ${d===r.d?'selected':''}>${DOW[P(d).getUTCDay()]} ${dm(d)}</option>`).join('')}</select></td>
      <td><select data-td="${i}" data-f="b" aria-label="Bữa" ${W()?'':'disabled'}>${N.bua.map(b=>`<option ${b===r.b?'selected':''}>${b}</option>`).join('')}</select></td>
      <td><input type="text" id="td-l-${i}" data-td="${i}" data-f="loai" aria-label="Loại món" value="${esc(r.loai)}" title="${esc(r.loai)}" ${W()?'':'disabled'}></td>
      <td class="wrap"><input type="text" id="td-m-${i}" list="dl-mon" data-td="${i}" data-f="mon" aria-label="Tên món" value="${esc(r.mon)}" title="${esc(r.mon)}" ${W()?'':'disabled'}></td>
      <td><input type="text" inputmode="decimal" data-td="${i}" data-f="tl" value="${r.tl!=null&&r.tl!==''?String(+(+r.tl).toFixed(2)).replace('.',','):''}" placeholder="100" class="num-in" aria-label="Tỷ lệ suất" title="% số suất của bữa dùng món này (trống = theo quy tắc chung)" ${W()?'':'disabled'}></td>
      <td>${r.mon?(hasR?`<span class="pill ok" title="Định lượng riêng của bữa này (${r.nl.length} nguyên liệu)">định lượng riêng</span>`:has?'<span class="pill ok">có định mức</span>':'<span class="pill run">thiếu định mức</span>'):''}</td>
      <td>${W()?`<button class="x" data-act="td-del" data-i="${i}" aria-label="Xóa dòng">×</button>`:''}</td></tr>`;}
  const wsOpts=ws.map(w=>{const d=S.td[ED.nhom+'_'+w];return `<option value="${w}" ${w===ED.week?'selected':''}>Tuần ${dm(w)}–${dm(addD(w,6))} ${d?'· '+(d.items||[]).length+' món':'· trống'}</option>`;}).join('');
  const cover=ws.filter(w=>S.td[ED.nhom+'_'+w]).length;
  return head('1.1 Thực đơn tuần','Nhập thực đơn từng tuần cho từng nhóm khách. Mỗi món ở đây sẽ được nhân với số suất của bữa đó (bước 1.2) và định mức nguyên liệu (bước 1.3). Thực đơn “Nội trú chung” dùng cho các nhóm nội trú chưa có thực đơn riêng ở bữa đó. Món có loại “Chọn 1” (bữa sáng chọn 1 trong các món) được chia đều số suất. Bữa nào chưa có thực đơn sẽ tính theo mức tiêu thụ tháng 9.','td')+
  `<div class="toolbar"><div class="seg">${Object.entries(NHOM).map(([k,v])=>`<button data-act="td-nhom" data-k="${k}" aria-pressed="${ED.nhom===k}">${v.ten}</button>`).join('')}</div>
   <select id="td-week" aria-label="Chọn tuần">${wsOpts}</select><span class="sp"></span>
   ${W()?`<button class="btn sm" data-act="td-add">+ Thêm món</button><button class="btn sm" data-act="td-copy">Chép tuần trước</button><button class="btn sm" data-act="td-paste">Dán từ Excel</button><button class="btn sm primary" data-act="pack-open" type="button">Nạp gói dữ liệu</button><select id="td-to" aria-label="Chép sang nhóm">${Object.entries(NHOM).filter(([k])=>k!==ED.nhom).map(([k,v])=>`<option value="${k}">${v.ten}</option>`).join('')}</select><button class="btn sm" data-act="td-to" ${rows.length?'':'disabled'}>Chép tuần này sang</button>
   <button class="btn primary" data-act="td-save" ${ED.tdDirty?'':'disabled'}>Lưu tuần</button>`:''}</div>
   <div class="meta">${doc?`Đã lưu ${tstr(doc.luc)} · ${esc(nm(doc.ai))}${doc.nguon?' · '+esc(doc.nguon):''}`:'Tuần này chưa có thực đơn.'} ${ED.tdDirty?'<span class="dirty">· Có thay đổi chưa lưu</span>':''} · ${N.ten}: ${cover}/${ws.length} tuần đã có thực đơn</div>
   ${ED.paste?`<div class="box"><label class="meta" for="td-paste-txt">Dán các dòng copy từ Excel, mỗi dòng: <b>Ngày hoặc Thứ</b> ⇥ <b>Bữa</b> ⇥ <b>Loại</b> ⇥ <b>Món</b>. Ví dụ: <span class="num">T2 ⇥ Trưa ⇥ Món mặn ⇥ Thịt kho trứng</span> hoặc <span class="num">03/11 ⇥ Tối ⇥ Món rau ⇥ Cải luộc</span></label>
     <textarea class="paste" id="td-paste-txt"></textarea><div class="toolbar"><button class="btn sm primary" data-act="td-paste-ok">Thêm vào tuần</button><button class="btn sm" data-act="td-paste">Đóng</button></div></div>`:''}
   ${W()?`<details class="box" ${ED.packOpen?'open':''}><summary class="meta"><b>Nạp gói dữ liệu từ Drive</b> – thực đơn + định lượng + lượng khách soạn sẵn (vd thực đơn tuần 1 từ Excel)</summary>
     <div class="toolbar"><input type="text" id="pack-id" value="${esc(ED.packId||PACK_TUAN1||'')}" placeholder="ID hoặc link file .json trên Drive" style="flex:1;min-width:220px"><button class="btn sm primary" data-act="pack-load" ${ED.packing?'disabled':''}>${ED.packing?esc(ED.packing):'Nạp gói'}</button></div>
     <div class="meta">Gói mặc định: thực đơn tháng 11 chạy thử – thực đơn tuần 1 Excel (ở cữ, nội trú thường + nội trú chung, nhân viên) từ 02/11 lặp lại mọi tuần đến 30/11 (01/11 lấy thực đơn Chủ nhật), định mức tham chiếu, lượng khách cả tháng (CBNV 260/160/100, nội trú 40, ở cữ 15), 6 mặt hàng ngoài bảng giá. Ghi đè thực đơn tháng 11 hiện có.</div></details>`:''}
   ${rows.length?`<div class="tbl-wrap"><table class="t td-tbl"><colgroup><col class="c-d"><col class="c-b"><col class="c-l"><col><col class="c-tl"><col class="c-dm"><col class="c-x"></colgroup><thead><tr><th>Ngày</th><th>Bữa</th><th>Loại món</th><th>Tên món</th><th class="r">% suất</th><th>Định mức</th><th><span class="sr">Xóa</span></th></tr></thead><tbody>${body}</tbody></table></div>`:
   `<div class="empty"><b>Tuần ${dm(ED.week)}–${dm(addD(ED.week,6))} chưa có thực đơn ${N.ten.toLowerCase()}</b>Bấm “Chép tuần trước”, “Dán từ Excel” hoặc “+ Thêm món”.</div>`}
   <datalist id="dl-mon">${known.map(m=>`<option value="${esc(m)}">`).join('')}</datalist>
   ${W()?`<div class="toolbar"><span class="sp"></span><button class="btn" data-act="tt" data-k="td">${S.tt.td&&S.tt.td.ok?'Xác nhận lại':'Xác nhận đủ thực đơn tháng'}</button></div>`:''}`;
}
function parsePaste(txt){
  const days=Array.from({length:7},(_,i)=>addD(ED.week,i)), N=NHOM[ED.nhom], out=[], bad=[];
  const th={'t2':1,'thu hai':1,'thứ hai':1,'t3':2,'thu ba':2,'thứ ba':2,'t4':3,'thu tu':3,'thứ tư':3,'t5':4,'thu nam':4,'thứ năm':4,'t6':5,'thu sau':5,'thứ sáu':5,'t7':6,'thu bay':6,'thứ bảy':6,'cn':0,'chu nhat':0,'chủ nhật':0};
  for(const line of txt.split(/\r?\n/)){ if(!line.trim()) continue; const c=line.split('\t').map(s=>s.trim());
    if(c.length<3){bad.push(line);continue;} let d=null; const a=c[0].toLowerCase();
    const m=a.match(/^(\d{1,2})[\/\-.](\d{1,2})/); if(m){const t=`${ED.week.slice(0,4)}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; d=days.find(x=>x===t)||null;}
    else if(a in th){d=days.find(x=>P(x).getUTCDay()===th[a]);}
    const b=N.bua.find(x=>x.toLowerCase()===String(c[1]).toLowerCase())||(String(c[1]).toLowerCase().startsWith('ph')?'Phụ':null);
    const mon=c.length>=4?c[3]:c[2], loai=c.length>=4?c[2]:'';
    if(!d||!b||!mon){bad.push(line);continue;} out.push({d,b,loai,mon});
  }
  return {out,bad};
}

/* ---------- 1.2 LƯỢNG KHÁCH ---------- */
function vKhach(){
  if(!ED.khach){ED.khach=JSON.parse(JSON.stringify((S.khach&&S.khach.days)||{}));}
  const K=ED.khach, days=monthDays(S.thang); const tot=Z(); const LB=CG.flatMap(([g,b])=>b.map(x=>g+' · '+x));
  const rows=days.map(d=>{const r=Array.from({length:NK},(_,i)=>+((K[d]||[])[i])||0);r.forEach((v,i)=>tot[i]+=v);const s=r.reduce((a,b)=>a+b,0);const w=P(d).getUTCDay();
    return `<tr${w===0||w===6?' class="edited"':''}><td>${DOW[w]} ${dm(d)}</td>${r.map((v,i)=>`<td class="r"><input class="n" type="number" min="0" step="1" id="kh-${d}-${i}" aria-label="${DOW[w]} ${dm(d)} · ${LB[i]||''}" data-kh="${d}" data-i="${i}" value="${v}" ${W()?'':'disabled'}></td>`).join('')}<td class="r num">${VND(s)}</td></tr>`;}).join('');
  const all=tot.reduce((a,b)=>a+b,0), sum=(a,b)=>tot.slice(a,b).reduce((x,y)=>x+y,0);
  const pct=(v,ref)=>v?((v/ref-1)*100>=0?'+':'')+((v/ref-1)*100).toFixed(1).replace('.',',')+'% so với T9':'';
  return head('1.2 Dự báo lượng khách','Số suất dự kiến theo ngày, theo nhóm khách và bữa. Mặc định điền theo trung bình thực tế tháng 9: nhân viên theo từng thứ trong tuần, nội trú theo từng nhóm (An Nhiên = VVIP khoa Sản + VVIP Nội trú Ngoại; VIP = VVIP các khoa còn lại). Sửa trực tiếp các ngày có lễ, sự kiện hoặc dự báo công suất khác.','khach')+
  `<div class="kpis"><div class="kv big"><b>Tổng suất dự kiến tháng</b><div class="num">${VND(all)}</div><small>${pct(all,SUAT_T9)||'T9: 10.111 suất'}</small></div>
    <div class="kv"><b>Nhân viên</b><div class="num">${VND(sum(0,2))}</div><small>T9: 5.240 · ${pct(sum(0,2),5240)}</small></div>
    <div class="kv"><b>Nội trú (Thường · VIP · An Nhiên)</b><div class="num" style="font-size:16px">${VND(sum(2,6))} · ${VND(sum(6,10))} · ${VND(sum(10,14))}</div><small>T9: 4.871 suất gồm bữa phụ</small></div>
    <div class="kv"><b>Ngân sách tham chiếu theo khách</b><div class="num">${TR(all*DM_SUAT)}</div><small>${VND(DM_SUAT)} đ/suất (thực chi T9 ÷ suất T9)</small></div></div>
  ${W()?`<div class="toolbar"><button class="btn sm" data-act="kh-fill">Điền theo TB tháng 9</button><span class="meta">${S.khach?`Đã lưu ${tstr(S.khach.luc)} · ${esc(nm(S.khach.ai))}`:'Chưa lưu số khách tháng này.'} ${ED.khDirty?'<span class="dirty">· Có thay đổi chưa lưu</span>':''}</span><span class="sp"></span>
    <button class="btn primary" data-act="kh-save" ${ED.khDirty?'':'disabled'}>Lưu lượng khách</button><button class="btn" data-act="tt" data-k="khach">${S.tt.khach&&S.tt.khach.ok?'Xác nhận lại':'Xác nhận số khách'}</button></div>`:''}
  <div class="tbl-wrap"><table class="t kh-tbl"><thead><tr><th rowspan="2">Ngày</th>${CG.map(([g,b])=>`<th colspan="${b.length}" style="text-align:center">${g}</th>`).join('')}<th rowspan="2" class="r">Cộng</th></tr>
  <tr>${CG.map(([g,b])=>b.map(x=>`<th class="r th2">${x}</th>`).join('')).join('')}</tr></thead><tbody>${rows}</tbody>
  <tfoot><tr><td>Tổng</td>${tot.map(v=>`<td class="r num">${VND(v)}</td>`).join('')}<td class="r num">${VND(all)}</td></tr></tfoot></table></div>`;
}

/* ---------- 1.3 ĐỊNH MỨC MÓN ---------- */
function vDM(){
  const ds=dishesOfMonth(); if(!ED.dmMon&&ds.length) ED.dmMon=ds[0].k;
  const miss=ds.filter(d=>!S.dm[d.k]);
  const list=ds.length?ds.map(d=>{const r=S.dm[d.k];return `<button type="button" class="ditem" data-act="dm-pick" data-k="${d.k}" aria-current="${ED.dmMon===d.k}"><span>${esc(d.mon)}</span>${r?`<span class="pill ${r.nguon==='AI'?'run':'ok'}">${r.nguon==='AI'?'AI':'đã duyệt'}</span>`:'<span class="pill bad">thiếu</span>'}</button>`;}).join(''):'<div class="empty"><b>Chưa có món nào</b>Nhập thực đơn ở bước 1.1 trước.</div>';
  return head('1.3 Định mức nguyên liệu theo món','Mỗi món cần biết dùng mã hàng nào, bao nhiêu cho 1 suất. Bấm “AI đề xuất” để Claude soạn định mức từ danh mục hàng có giá khung, sau đó bếp xem và sửa. Món đã duyệt được dùng lại cho mọi tuần sau.','dm')+
  `<div class="toolbar">${W()?`<button class="btn primary" data-act="dm-ai" ${miss.length&&!ED.ai?'':'disabled'}>${ED.ai?esc(ED.ai):`AI đề xuất định mức cho ${miss.length} món thiếu`}</button>`:''}
    <span class="meta">${ds.length} món trong tháng · ${ds.length-miss.length} có định mức · ${ds.filter(d=>S.dm[d.k]&&S.dm[d.k].nguon==='AI').length} do AI soạn chưa duyệt</span><span class="sp"></span>
    ${W()?`<button class="btn sm" data-act="dm-paste">Dán định mức từ Excel</button><button class="btn" data-act="tt" data-k="dm">${S.tt.dm&&S.tt.dm.ok?'Xác nhận lại':'Xác nhận định mức'}</button>`:''}</div>
  ${ED.dmPaste?`<div class="box"><label class="meta" for="dm-paste-txt">Mỗi dòng: <b>Tên món</b> ⇥ <b>Mã hàng</b> ⇥ <b>Định lượng/suất</b> (gram với hàng tính KG, số đơn vị với hàng khác). Món đã có định mức sẽ được thay bằng các dòng dán vào.</label>
    <textarea class="paste" id="dm-paste-txt"></textarea><div class="toolbar"><button class="btn sm primary" data-act="dm-paste-ok">Nạp định mức</button><button class="btn sm" data-act="dm-paste">Đóng</button></div></div>`:''}
  <div class="two"><div class="dlist" role="group" aria-label="Danh sách món (${ds.length})">${list}</div><div>${ED.dmMon?dmEditor():''}</div></div>`;
}
function dmEditor(){
  const k=ED.dmMon, rec=S.dm[k], mon=(dishesOfMonth().find(d=>d.k===k)||{}).mon||(rec&&rec.mon)||'';
  if(!ED.dmRows||ED.dmRows.k!==k){ED.dmRows={k,rows:rec?(rec.nl||[]).map(x=>({...x})):[]};ED.dmDirty=false;}
  const rows=ED.dmRows.rows; let cost=0;
  const body=rows.map((r,i)=>{const x=IT.get(r.ma);const kg=x&&/^(KG|L)$/.test(x.dvt);const g=x?priceOf(x,x.ncc[0]&&x.ncc[0].ncc):0;cost+=g*(r.q||0);
    return `<tr><td><input type="text" id="dm-ma-${i}" list="dl-item" data-dm="${i}" data-f="ma" aria-label="Mã hàng" value="${esc(r.ma)}" ${W()?'':'disabled'}></td>
    <td class="wrap nm">${x?esc(x.ten):'<span class="pill bad">mã không có trong bảng giá</span>'}</td><td class="dvt">${x?esc(x.dvt):''}</td>
    <td class="r q"><input class="q" type="text" inputmode="decimal" aria-label="Định lượng / suất${x?' · '+esc(x.ten):''}" id="dm-q-${i}" data-dm="${i}" data-f="q" value="${kg?QI((r.q||0)*1000):QI(r.q)}" ${W()?'':'disabled'}> <span class="muted u">${kg?'g':esc(x?x.dvt.toLowerCase():'')}</span></td>
    <td class="r num tien">${VND(g*(r.q||0))}</td><td class="del">${W()?`<button class="x" data-act="dm-del" data-i="${i}" aria-label="Xóa nguyên liệu">×</button>`:''}</td></tr>`;}).join('');
  return `<div class="sh"><div><h2 style="font-size:17px">${esc(mon)}</h2><p>${rec?`${rec.nguon==='AI'?'AI soạn':'Bếp duyệt'} · ${tstr(rec.luc)}${rec.ghichu?' · '+esc(rec.ghichu):''}`:'Chưa có định mức.'} ${ED.dmDirty?'<span class="dirty">· Có thay đổi chưa lưu</span>':''}</p></div>
    <div class="kv"><b>Giá vốn nguyên liệu / suất</b><div class="num">${VND(cost)} đ</div></div></div>
    <div class="tbl-wrap"><table class="t dm-tbl"><colgroup><col class="c-ma"><col><col class="c-dvt"><col class="c-q"><col class="c-tien"><col class="c-x"></colgroup><thead><tr><th>Mã hàng</th><th>Tên hàng</th><th>ĐVT</th><th class="r">Định lượng / suất</th><th class="r">Tiền / suất</th><th><span class="sr">Xóa</span></th></tr></thead><tbody>${body||'<tr><td colspan="6" class="muted">Chưa có nguyên liệu.</td></tr>'}</tbody></table></div>
    ${W()?`<div class="toolbar" style="margin-top:8px"><button class="btn sm" data-act="dm-add">+ Thêm nguyên liệu</button><span class="sp"></span>${rec?'<button class="btn sm warn" data-act="dm-clear">Xóa định mức</button>':''}<button class="btn primary" data-act="dm-save">${rec&&rec.nguon==='AI'&&!ED.dmDirty?'Duyệt định mức':'Lưu định mức'}</button></div>`:''}
`;
}
async function runAI(){
  if(!sample){toast('Trang này chưa được phép hỏi Claude. Mở trang trên claude.ai để dùng.');return;}
  const miss=dishesOfMonth().filter(d=>!S.dm[d.k]); if(!miss.length) return;
  const cat=ITEMS.filter(x=>x.n2!=='(ngoài bảng giá T10)').map(x=>`${x.ma}|${x.ten}|${x.dvt}`).join('\n');
  const B=8; let done=0,okN=0;
  for(let i=0;i<miss.length;i+=B){
    const batch=miss.slice(i,i+B); ED.ai=`Claude đang soạn ${Math.min(i+B,miss.length)}/${miss.length} món…`; render();
    const prompt=`Bạn là bếp trưởng bếp bệnh viện ở Hà Nội, nấu cho nhân viên và bệnh nhân (khẩu phần người lớn).
Với mỗi món dưới đây, liệt kê nguyên liệu CHÍNH cần mua cho 1 SUẤT, chỉ chọn mã có trong DANH MỤC (đúng mã). Định lượng "q" tính theo ĐVT của mã đó cho 1 suất (KG: ví dụ 0.08 = 80 g; QUA/HOP/CHAI...: số đơn vị, có thể lẻ, ví dụ 0.25). Bỏ qua nước, muối, mì chính, gia vị lặt vặt dưới 3 g/suất (đã tính riêng theo lịch sử). Cơm trắng: chỉ gạo. Món tráng miệng là hoa quả/đồ đóng gói: chọn đúng mặt hàng.
Trả về JSON đúng dạng: {"mon":[{"ten":"<tên món như đề bài>","nl":[{"ma":"THP...","q":0.08}],"ghichu":"<ngắn, nếu có>"}]}
MÓN:
${batch.map(d=>'- '+d.mon).join('\n')}
DANH MỤC (mã|tên|ĐVT):
${cat}`;
    try{
      const res=await SJ(prompt,{modelTier:'default'});
      for(const m of (res&&res.mon)||[]){
        const d=batch.find(b=>slug(b.mon)===slug(m.ten)); if(!d) continue;
        const nl=(m.nl||[]).filter(n=>IT.has(n.ma)&&+n.q>0&&+n.q<20).map(n=>({ma:n.ma,q:+(+n.q).toFixed(4)}));
        if(!nl.length) continue;
        await store.set('dinhmuc/'+d.k,{mon:d.mon,nl,nguon:'AI',ghichu:String(m.ghichu||'').slice(0,200),luc:Date.now(),ai:uid}); okN++;
      }
    }catch(e){ ED.ai=null; render(); toast(e&&e.code==='rate_limited'?'Claude đang bận, thử lại sau ít phút.':'Không soạn được định mức: '+(e&&e.message||e)); if(e&&(e.code==='not_granted'||e.code==='rate_limited')) break; }
    done+=batch.length;
  }
  ED.ai=null; render(); if(okN){ await log('1.3',`AI soạn định mức cho ${okN}/${miss.length} món (chờ bếp duyệt)`); toast(`Claude đã soạn ${okN} món. Kiểm tra từng món rồi bấm Duyệt.`);}
}

/* ---------- 1.4 DỰ TRÙ & NGÂN SÁCH ---------- */
function curLines(){ if(!ED.lines) ED.lines=((S.dutru&&S.dutru.lines)||[]).map(l=>({...l})); return ED.lines; }
function vTinh(){
  const dt=S.dutru, pr=(dt&&dt.params)||{dpT:10,dpK:5,withHist:true};
  const lines=curLines(); const ds=dishesOfMonth(), miss=ds.filter(d=>!S.dm[d.k]).length, K=S.khach;
  const warns=[]; if(!K) warns.push(['Chưa lưu lượng khách tháng này (bước 1.2): phần theo thực đơn và theo lịch sử sẽ bằng 0.',0]);
  if(!ds.length) warns.push(['Chưa có thực đơn trong tháng: toàn bộ dự trù sẽ tính theo lịch sử tiêu thụ.',0]);
  if(miss) warns.push([`${miss} món chưa có định mức (bước 1.3): các món này chưa đóng góp nguyên liệu.`,0]);
  if(dt&&dt.warn&&dt.warn.tot) warns.push([`Lần tính gần nhất: ${VND(dt.warn.tot)} suất, trong đó ${VND(dt.warn.tot-dt.warn.unc)} suất có thực đơn + định mức, ${VND(dt.warn.unc)} suất tính theo lịch sử.`,1]);
  let tx=0,tc=0; const byN={},byG={};
  lines.forEach(l=>{const x=IT.get(l.ma);if(!x)return;const g=priceOf(x,l.ncc);tx+=g*l.dx;tc+=g*l.chot;
    const a=byN[l.ncc]||(byN[l.ncc]={n:0,k1:0,k2:0}); a.n++; a.k1+=g*l.chot*l.k1r; a.k2+=g*l.chot*(1-l.k1r);
    byG[x.n2]=(byG[x.n2]||0)+g*l.chot;});
  const ncc=Object.entries(byN).sort((a,b)=>(b[1].k1+b[1].k2)-(a[1].k1+a[1].k2)); const mx=ncc.length?ncc[0][1].k1+ncc[0][1].k2:1;
  const totS=Object.values((S.khach&&S.khach.days)||{}).flat().reduce((a,b)=>a+(+b||0),0), refB=totS*DM_SUAT;
  const f=ED.filt, q=f.q.toLowerCase();
  const needX=(l,x)=>!!(l.e||l.note||/ngoài bảng giá/.test(x.n2||'')||!priceOf(x,l.ncc)); // dòng cần xem: sửa tay, có ghi chú, ngoài bảng giá, chưa có giá
  const nNeed=lines.filter(l=>{const x=IT.get(l.ma);return x&&needX(l,x);}).length, showAll=f.all||!nNeed||!!(f.q||f.ncc||f.n2);
  const shown=lines.map((l,i)=>[l,i]).filter(([l])=>{const x=IT.get(l.ma);return x&&(showAll||needX(l,x))&&(!f.ncc||l.ncc===f.ncc)&&(!f.n2||x.n2===f.n2)&&(!q||(x.ma+' '+x.ten).toLowerCase().includes(q));});
  const rowsH=shown.slice(0,400).map(([l,i])=>{const x=IT.get(l.ma),g=priceOf(x,l.ncc),h=x.h||[0,0,0];
    return `<tr class="${l.e?'edited':''}"><td class="num">${x.ma}</td><td class="wrap">${esc(x.ten)}<div class="muted" style="font-size:11.5px">${esc(x.n3||x.n2)}</div></td><td>${esc(x.dvt)}</td>
    <td>${x.ncc.length>1&&W()?`<select data-ln="${i}" data-f="ncc" aria-label="NCC cho ${esc(l.ma)}">${x.ncc.map(n=>`<option ${n.ncc===l.ncc?'selected':''}>${esc(n.ncc)}</option>`).join('')}</select>`:esc(l.ncc)}</td>
    <td class="r num">${Q(l.tm)}</td><td class="r num">${Q(l.ls)}</td><td class="r num">${l.ton==null?'<span class="muted">–</span>':Q(l.ton)+(l.mn?`<div class="muted" style="font-size:11px">tối thiểu ${Q(l.mn)}</div>`:'')}</td><td class="r num">${Q(l.dx)}</td>
    <td class="r"><input class="q" type="text" inputmode="decimal" id="ln-${i}" aria-label="SL chốt ${esc(l.ma)}" data-ln="${i}" data-f="chot" value="${QI(l.chot)}" ${W()?'':'disabled'}></td>
    <td class="r num">${Q(l.chot*l.k1r)} / ${Q(l.chot*(1-l.k1r))}</td>
    <td class="r num">${VND(g)}</td><td class="r num">${VND(g*l.chot)}</td><td class="r num muted">${Q(h[0])} / ${Q(h[1])}</td>
    <td><input type="text" id="lnn-${i}" aria-label="Ghi chú ${esc(l.ma)}" data-ln="${i}" data-f="note" value="${esc(l.note)}" placeholder="ghi chú" ${W()?'':'disabled'} style="min-width:120px"></td></tr>`;}).join('');
  const nccs=[...new Set(lines.map(l=>l.ncc))].sort();
  return head('1.4 Dự trù tự động & ngân sách','Bấm “Tính dự trù” để cộng: (a) nguyên liệu theo thực đơn × số suất × định mức, (b) phần còn lại theo mức tiêu thụ/suất của tháng 9, rồi cộng dự phòng; đồ khô trừ thêm tồn kho thực tế từ app Kho khô. Cột <b>SL chốt</b> là số bếp quyết định, sửa trực tiếp; dòng đã sửa tay được giữ nguyên khi tính lại.','tinh')+
  `<div class="toolbar">${W()?`<details class="opts"${OPTS_OPEN?' open':''}><summary>Tuỳ chọn · dự phòng ${pr.dpT}% / ${pr.dpK}%</summary><div class="optbody"><label class="meta" for="dpT">Dự phòng hàng tươi %</label><input class="n" type="number" id="dpT" value="${pr.dpT}" min="0" max="50" style="width:64px">
    <label class="meta" for="dpK">Đồ khô %</label><input class="n" type="number" id="dpK" value="${pr.dpK}" min="0" max="50" style="width:64px">
    <label class="meta" title="Suất chưa có thực đơn tính theo mức tiêu thụ/suất tháng 9"><input type="checkbox" id="wHist" ${pr.withHist!==false?'checked':''}> Suất chưa có thực đơn: theo T9</label>
    <label class="meta" title="Bỏ tích khi định mức đã gồm gia vị, dầu ăn… (vd thực đơn Excel tuần 1) để không cộng trùng"><input type="checkbox" id="wNgoai" ${(pr.histNgoai!=null?pr.histNgoai:!monthMenu().some(i=>i.nl&&i.nl.length))?'checked':''}> Cộng mặt hàng ngoài định mức theo T9</label>
    <label class="meta"><input type="checkbox" id="uKho" ${S.kho?(pr.useKho!==false?'checked':''):'disabled'}> Trừ tồn kho đồ khô</label></div></details>
    <button class="btn primary" data-act="calc">${dt?'Tính lại dự trù':'Tính dự trù tự động'}</button>`:''}
    <span class="meta">${dt?`Tính lúc ${tstr(dt.luc)} · ${esc(nm(dt.ai))}`:'Chưa tính.'} ${ED.lnDirty?'<span class="dirty">· Có số chốt chưa lưu</span>':''}</span><span class="sp"></span>
    ${W()?`<button class="btn primary" data-act="ln-save" ${ED.lnDirty?'':'disabled'}>Lưu số chốt</button>`:''}</div>
  <div class="warns">${warns.map(w=>`<div class="warn ${w[1]?'ok':''}">${w[0]}</div>`).join('')}</div>
  ${khoPanel()}
  ${lines.length?`
  <div class="kpis"><div class="kv big"><b>Ngân sách dự tính (SL chốt)</b><div class="num">${TR(tc)}</div><small>${VND(tc)} đ, gồm VAT</small></div>
    <div class="kv"><b>Đề xuất tự động</b><div class="num">${TR(tx)}</div><small>chênh ${tc>=tx?'+':''}${TR(tc-tx)} do sửa tay</small></div>
    <div class="kv"><b>Tham chiếu theo lượng khách</b><div class="num">${TR(refB)}</div><small>${refB?'Dự tính '+(tc>=refB?'cao hơn ':'thấp hơn ')+Math.abs((tc/refB-1)*100).toFixed(1).replace('.',',')+'% · '+VND(totS)+' suất':'cần lưu lượng khách'}</small></div>
    <div class="kv"><b>Hàng tươi / Đồ khô</b><div class="num">${TR(byG['Thực phẩm tươi sống']||0)} / ${TR(byG['Đồ khô']||0)}</div><small>${lines.length} mã hàng · ${ncc.length} NCC</small></div></div>
  <div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>Nhà cung cấp</th><th class="r">Số mã</th><th class="r k12">Kỳ 1 (01–15)</th><th class="r k12">Kỳ 2 (16–cuối)</th><th class="r">Cả tháng</th><th style="width:30%"><span class="sr">Tỷ trọng</span></th></tr></thead><tbody>
   ${(ED.nccAll?ncc:ncc.slice(0,6)).map(([n,a])=>`<tr><td><button class="btn sm lnk" data-act="f-ncc" data-n="${esc(n)}" title="Lọc bảng dưới theo NCC này">${esc(n)}</button></td><td class="r num">${a.n}</td><td class="r num k12">${VND(a.k1)}</td><td class="r num k12">${VND(a.k2)}</td><td class="r num"><b>${VND(a.k1+a.k2)}</b></td><td><div class="bar" style="width:${((a.k1+a.k2)/mx*100).toFixed(1)}%"></div></td></tr>`).join('')}
  </tbody></table></div>
  ${ncc.length>6?`<button class="btn sm lnk morebtn" data-act="ncc-all">${ED.nccAll?'Thu gọn':'Xem cả '+ncc.length+' NCC'} </button>`:''}
  <div class="toolbar">${nNeed?`<div class="seg" role="group" aria-label="Dòng hiển thị"><button data-act="f-all" data-v="0" aria-pressed="${!f.all}" title="Dòng sửa tay, có ghi chú, ngoài bảng giá hoặc chưa có giá">Cần xem ${nNeed}</button><button data-act="f-all" data-v="1" aria-pressed="${!!f.all}">Tất cả ${lines.length}</button></div>`:''}<input type="text" id="f-q" aria-label="Tìm mã hoặc tên hàng" placeholder="Tìm mã / tên hàng" value="${esc(f.q)}" style="width:200px">
    <select id="f-ncc" aria-label="Lọc theo NCC"><option value="">Tất cả NCC</option>${nccs.map(n=>`<option ${n===f.ncc?'selected':''}>${esc(n)}</option>`).join('')}</select>
    <select id="f-n2" aria-label="Lọc theo nhóm hàng"><option value="">Tất cả nhóm</option>${['Thực phẩm tươi sống','Đồ khô','(ngoài bảng giá T10)'].map(n=>`<option ${n===f.n2?'selected':''}>${n}</option>`).join('')}</select>
    <span class="meta">${shown.length} dòng${shown.length>400?' (hiện 400 dòng đầu, lọc để xem thêm)':''}</span></div>
  <div class="tbl-wrap"><table class="t"><thead><tr><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th>NCC</th><th class="r">Theo thực đơn</th><th class="r">Theo lịch sử</th><th class="r">Tồn kho</th><th class="r">Đề xuất</th><th class="r">SL chốt</th><th class="r">Giao K1 / K2</th><th class="r">Đơn giá</th><th class="r">Thành tiền</th><th class="r">Thực nhận T9 / K1 T10</th><th>Ghi chú</th></tr></thead><tbody>${rowsH}</tbody></table></div>`
  :'<div class="empty"><b>Chưa có số dự trù</b>Bấm “Tính dự trù tự động”. Có thể tính ngay cả khi thực đơn chưa đủ: phần thiếu lấy theo lịch sử.</div>'}`;
}

function khoPanel(){
  const k=S.kho, map=khoMapAll(), saved=(S.khomap&&S.khomap.map)||{};
  const ids=k?Object.keys(k.items||{}):[]; const mapped=ids.filter(id=>map[id]&&IT.has(map[id])).length;
  return `<details class="box" id="khoBox" ${ED.khoOpen?'open':''}><summary><b>Tồn kho đồ khô · app Kho khô</b> <span class="meta">${k?`đọc ${tstr(k.luc)} · ${ids.length} mặt hàng · ${mapped} đã ghép mã · ${k.ngd||0} giao dịch`:'chưa đọc'}</span></summary>
  <div class="toolbar"><button class="btn sm" data-act="kho-read" ${mcp&&!ED.khoBusy?'':'disabled'}>${ED.khoBusy?'Đang đọc…':'Đọc tồn kho từ app'}</button>${mcp?'':'<span class="meta">Cần mở trang trên claude.ai và cho phép kết nối Google Sheets.</span>'}<span class="sp"></span>${W()&&ED.mapDirty?'<button class="btn sm primary" data-act="map-save">Lưu ghép mã</button>':''}</div>
  <p class="meta" style="margin:0">Chỉ đọc file Items_AppSheet_KhoKho (Items + Transactions), không ghi gì vào app. Tồn = tồn đầu + nhập − xuất. Đồ khô có ghép mã: đề xuất = nhu cầu × (1 + dự phòng) − tồn + tồn tối thiểu.</p>
  ${k?`<div class="tbl-wrap" style="max-height:320px"><table class="t"><thead><tr><th>Mã kho</th><th>Tên trong app</th><th>ĐVT</th><th class="r">Tồn</th><th class="r">Tối thiểu</th><th>Mã hàng ghép</th><th>Tên trong bảng giá</th></tr></thead><tbody>
  ${ids.map(id=>{const v=k.items[id],ma=map[id]||'',x=IT.get(ma),sc=(KHOMAP0[id]||[])[1],chosen=(ED.map&&id in ED.map)||(id in saved),weak=!chosen&&sc!=null&&sc<0.35;
    return `<tr><td class="num">${esc(id)}</td><td class="wrap">${esc(v.ten)}</td><td>${esc(v.dvt)}</td><td class="r num">${Q(v.ton)}</td><td class="r num">${Q(v.min)}</td>
    <td><input type="text" list="dl-item" id="map-${esc(id)}" data-map="${esc(id)}" value="${esc(ma)}" style="width:110px" ${W()?'':'disabled'}></td>
    <td class="wrap">${x?esc(x.ten)+(x.dvt.toUpperCase()!==String(v.dvt).toUpperCase()?` <span class="pill run">ĐVT giá: ${esc(x.dvt)}</span>`:''):'<span class="muted">chưa ghép</span>'}${weak?' <span class="pill bad">cần kiểm tra</span>':''}</td></tr>`;}).join('')}
  </tbody></table></div>`:''}</details>`;
}
async function readKho(){
  if(!mcp){toast('Trang chưa được kết nối Google Sheets.');return;}
  ED.khoBusy=true; ED.khoOpen=true; render();
  try{
    const get=async range=>{ const r=await call(KHO.server,'get_values',{spreadsheetId:KHO.id,range},'Đọc app Kho khô · '+range.split('!')[0]); let p=r&&r.payload;
      if((!p||!p.values)&&r&&r.content){ try{p=JSON.parse((r.content.find(c=>c.type==='text')||{}).text);}catch(e){} } return (p&&p.values)||[]; };
    const items=await get('Items!A1:H2000'), tx=await get('Transactions!A1:L20000'), it={};
    items.slice(1).forEach(r=>{ if(!r||!r[0]) return; it[String(r[0]).trim()]={ten:r[1]||'',dvt:r[2]||'',ton:pnum(r[3]),min:pnum(r[4]),act:String(r[7]||'').toUpperCase()!=='FALSE'}; });
    const h=(tx[0]||[]).map(x=>String(x).trim()), ci=n=>h.indexOf(n), iT=ci('TxnType'),iI=ci('ItemID'),iQ=ci('Quantity'),iS=ci('Status'); let n=0;
    tx.slice(1).forEach(r=>{ const id=String(r[iI]||'').trim(); if(!it[id]) return; if(iS>=0&&/hủy|huỷ|huy|xóa|xoá/i.test(String(r[iS]||''))) return;
      const q=pnum(r[iQ]), ty=String(r[iT]||'').toLowerCase(); if(ty.startsWith('nh')){it[id].ton+=q;n++;} else if(ty.startsWith('xu')){it[id].ton-=q;n++;} });
    Object.values(it).forEach(v=>v.ton=Math.round(v.ton*1000)/1000);
    await store.set('kho/ton',{items:it,luc:Date.now(),ai:uid,ngd:n});
    await log('1.4',`Đọc tồn kho đồ khô từ app Kho khô: ${Object.keys(it).length} mặt hàng, ${n} giao dịch nhập/xuất`); toast('Đã đọc tồn kho. Bấm Tính lại dự trù để trừ tồn.');
  }catch(e){ toast(e&&e.code==='not_granted'?'Chưa cho phép trang đọc Google Sheets.':'Không đọc được app kho: '+((e&&(e.message||e.code))||e)); }
  ED.khoBusy=false; render();
}

/* ---------- 1.5 CHỐT ---------- */
function vChot(){
  const lines=curLines(); let tc=0; lines.forEach(l=>{const x=IT.get(l.ma);if(x)tc+=priceOf(x,l.ncc)*l.chot;});
  const req=['td','khach','dm','tinh'], okAll=req.every(k=>S.tt[k]&&S.tt[k].ok);
  return head('1.5 Chốt dự trù tháng','Sau khi chốt, số lượng được khóa và chuyển sang bước 2 để lập PR theo tháng cho cả hàng tươi và đồ khô (cột Giao K1/K2 chỉ để tham khảo lịch giao). Có thể mở lại để sửa trước khi lập PR.','chot')+
  `<div class="kpis"><div class="kv big"><b>Ngân sách dự tính</b><div class="num">${TR(tc)}</div><small>${lines.length} mã hàng</small></div>
   ${SUB.slice(0,4).map(([n,t,k])=>`<div class="kv"><b>${n} ${t}</b><div>${S.tt[k]&&S.tt[k].ok?'<span class="pill ok">✓ Đã xác nhận</span>':'<span class="pill run">Chưa</span>'}</div><small>${S.tt[k]&&S.tt[k].ok?tstr(S.tt[k].luc)+' · '+esc(nm(S.tt[k].ai)):''}</small></div>`).join('')}</div>
  ${!okAll&&!locked()?'<div class="warn">Nên xác nhận đủ 4 bước con trước khi chốt. Vẫn có thể chốt nếu anh chấp nhận số hiện tại.</div>':''}
  <div class="toolbar">${canWrite?(locked()?`<span class="pill ok">✓ Đã chốt ${tstr(S.tt.chot.luc)} · ${esc(nm(S.tt.chot.ai))}</span><button class="btn" data-act="unlock">Mở lại để sửa</button>`:
    `<button class="btn primary" data-act="lock" ${lines.length&&!ED.lnDirty?'':'disabled'}>Chốt dự trù tháng ${S.thang.slice(5)}/${S.thang.slice(0,4)}</button>${ED.lnDirty?'<span class="dirty">Lưu số chốt ở bước 1.4 trước</span>':''}`):''}
    <button class="btn" data-act="export" ${lines.length?'':'disabled'}>Xuất file cho bước 2 (CSV)</button></div>`;
}
function csv(){
  const L=curLines(), r=[['Thang','Ky_PR','Ma_hang','Ten_hang','DVT','Nhom','NCC','Ma_NCC','SL_chot','Giao_ky1','Giao_ky2','Don_gia_gom_VAT','Thanh_tien','Ghi_chu']];
  L.forEach(l=>{const x=IT.get(l.ma);if(!x)return;const n=x.ncc.find(n=>n.ncc===l.ncc)||x.ncc[0]||{};const g=n.g||0;
    const g1=Math.min(l.chot,roundQ(l.chot*l.k1r,x.dvt)); r.push([S.thang,'Thang',x.ma,x.ten,x.dvt,x.n2,l.ncc,n.m||'',l.chot,g1,+(l.chot-g1).toFixed(3),g,Math.round(g*l.chot),l.note||'']);});
  return '﻿'+r.map(a=>a.map(v=>{v=String(v??'');return /[",\n;]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;}).join(',')).join('\r\n');
}

