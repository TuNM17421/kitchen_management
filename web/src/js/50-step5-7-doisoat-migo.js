/* ================= BƯỚC 5 · ĐỐI SOÁT NCC THEO CHU KỲ ================= */
const SUB5=[['5.1','Bảng kê & gửi NCC','bk'],['5.2','So file NCC','sofile'],['5.3','Chốt khớp theo kỳ','dsk']];
const SUB6=[['6.1','Đề xuất chỉnh PO','cpo'],['6.2','Gắn PR bổ sung vào PO','gpo'],['6.3','PO so với số đã khớp','kpo']];
const SUB7=[['7.1','Lập & xuất MIGO','mg'],['7.2','Lịch sử MIGO','mgls']];
const NCCINFO={"NCC MẪU 02": {"email": ["ncc02@example.com"], "tel": "Liên hệ 4373004820", "mancc": "2900008", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 03": {"email": ["ncc03@example.com"], "tel": "Liên hệ 2469667754", "mancc": "2900021", "ck": 15, "note": ""}, "NCC MẪU 04": {"email": ["ncc04@example.com"], "tel": "Liên hệ 5148247615", "mancc": "2900024", "ck": 15, "note": ""}, "NCC MẪU 35": {"email": ["ncc35@example.com"], "tel": "Liên hệ 3270902444", "mancc": "2900028", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 06": {"email": ["ncc06@example.com"], "tel": "Liên hệ 6691105590", "mancc": "2900016", "ck": 15, "note": "Tên gọi khác: HTX Hữu cơ NCC MẪU 06; Hữu cơ NCC MẪU 06"}, "NCC MẪU 36": {"email": ["ncc36@example.com"], "tel": "Liên hệ 0221749317", "mancc": "2900011", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 12": {"email": ["ncc12@example.com"], "tel": "Liên hệ 4365971533", "mancc": "2900030", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng | Gộp từ mã 2000009 (giá khung đồ khô) ngày 1"}, "NCC MẪU 09": {"email": ["ncc09@example.com"], "tel": "Liên hệ 8235991963", "mancc": "2900020", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 10": {"email": ["ncc10@example.com"], "tel": "Liên hệ 7671174409", "mancc": "2900019", "ck": 15, "note": ""}, "NCC MẪU 11": {"email": ["ncc11@example.com"], "tel": "Liên hệ 0335117849", "mancc": "2900015", "ck": 15, "note": ""}, "NCC MẪU 13": {"email": ["ncc13@example.com"], "tel": "Liên hệ 7689391837", "mancc": "2900004", "ck": 15, "note": ""}, "NCC MẪU 14": {"email": ["ncc14@example.com"], "tel": "Liên hệ 3009457506", "mancc": "2900002", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 16": {"email": ["ncc16@example.com"], "tel": "Liên hệ 7131950159", "mancc": "2900031", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 15": {"email": ["ncc15@example.com"], "tel": "Liên hệ 1064235546", "mancc": "2900026", "ck": 15, "note": ""}, "NCC MẪU 18": {"email": ["ncc18@example.com"], "tel": "Liên hệ 6576980805", "mancc": "2900013", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 19": {"email": ["ncc19@example.com"], "tel": "Liên hệ 5386748949", "mancc": "2900027", "ck": 15, "note": ""}, "NCC MẪU 21": {"email": ["ncc21@example.com"], "tel": "Liên hệ 1300933915", "mancc": "2900018", "ck": 15, "note": "Giá khung ghi tên 'MẦM XANH' — cùng mã NCC 2900018 | Tên gọi khác: MẦM XANH"}, "NCC MẪU 22": {"email": ["ncc22@example.com"], "tel": "Liên hệ 8043832823", "mancc": "2900022", "ck": 15, "note": "Giá khung ghi tên 'T VÀ P' — cùng mã NCC 2900022 | Tên gọi khác: T VÀ P"}, "NCC MẪU 26": {"email": ["ncc26@example.com"], "tel": "Liên hệ 9307469646", "mancc": "2900014", "ck": 15, "note": ""}, "NCC MẪU 23": {"email": ["ncc23@example.com"], "tel": "Liên hệ 3574493471", "mancc": "2900005", "ck": 15, "note": ""}, "NCC MẪU 24": {"email": ["ncc24@example.com"], "tel": "Liên hệ 8946590880", "mancc": "2900023", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 27": {"email": ["ncc27@example.com"], "tel": "Liên hệ 3038269652", "mancc": "2900025", "ck": 15, "note": "Tên gọi khác: TP Cao cấp"}, "NCC MẪU 28": {"email": ["ncc28@example.com"], "tel": "Liên hệ 4267681070", "mancc": "2900029", "ck": 15, "note": ""}, "NCC MẪU 17": {"email": ["ncc17@example.com"], "tel": "Liên hệ 7550598971", "mancc": "2900035", "ck": 15, "note": ""}, "NCC MẪU 05": {"email": ["ncc05@example.com"], "tel": "", "mancc": "2900032", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 32": {"email": [], "tel": "", "mancc": "2900001", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 07": {"email": [], "tel": "", "mancc": "2900003", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 01": {"email": [], "tel": "", "mancc": "2900006", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 30": {"email": [], "tel": "", "mancc": "2900007", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 31": {"email": [], "tel": "", "mancc": "2900009", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 33": {"email": ["ncc33@example.com"], "tel": "", "mancc": "2900010", "ck": 30, "note": "Chỉ báo công nợ ngày 30 hàng tháng"}, "NCC MẪU 34": {"email": [], "tel": "", "mancc": "2900012", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 29": {"email": [], "tel": "", "mancc": "2900017", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 08": {"email": [], "tel": "", "mancc": "2900034", "ck": 0, "note": "Nhà cung cấp này không cần báo công nợ"}, "NCC MẪU 20": {"email": [], "tel": "", "mancc": "2900036", "ck": 15, "note": "Mã NCC tạm 2900036 (chưa có mã thật); chưa có email/SĐT — điền để báo công nợ"}, "NCC MẪU 25": {"email": ["ncc25@example.com"], "tel": "", "mancc": "2900033", "ck": 15, "note": ""}};
const CKN={15:'15 ngày (2 kỳ)',30:'Tháng (1 kỳ)',0:'Không báo NCC'};
const TTDS={chua:['Chưa gửi',''],gui:['Đã gửi NCC','run'],lech:['Lệch','bad'],khop:['Đã khớp','ok']};
const stripId=o=>{const x={...o};delete x.id;return x;};
function ckOf(ncc){ const m=S.ck&&S.ck.map; if(m&&m[ncc]!=null) return +m[ncc]; const i=NCCINFO[ncc]; return i&&i.ck!=null?i.ck:15; }
function kyList(t,ncc){ const ds=monthDays(t),end=ds[ds.length-1]; return ckOf(ncc)===15?[{k:1,from:t+'-01',to:t+'-15'},{k:2,from:t+'-16',to:end}]:[{k:0,from:t+'-01',to:end}]; }
const kyName=ky=>ky.k===0?`Cả tháng ${mmY(ky.from.slice(0,7))}`:`Kỳ ${ky.k} (${dm(ky.from)}–${dm(ky.to)})`;
const dsId=(t,ncc,k)=>`${t}_${slug(ncc)}_k${k}`;
const dsDoc=(ncc,k)=>(S.ds||{})[dsId(S.thang,ncc,k)]||null;
function priceFor(ncc,ma){ const p=poOf(S.thang).find(p=>p.ncc===ncc); const l=p&&p.lines.find(l=>l.ma===ma); if(l) return l.g; const x=IT.get(ma); return x?priceOf(x,ncc):0; }
function recvFull(){ const out=[];
  ((S.nhan&&S.nhan.rows)||[]).forEach((r,i)=>out.push({d:r.d,ma:r.ma,ncc:String(r.ncc||'').toUpperCase(),sl:+r.sl||0,po:r.po||'',phieu:r.nguon||'dán tay',src:{t:'nh',i}}));
  pnOf(S.thang).forEach(p=>p.lines.forEach((l,j)=>out.push({d:p.d,ma:l.ma,ncc:p.ncc,sl:l.gach?0:+l.sl||0,po:p.soPO,phieu:p.ma,src:{t:'pn',id:p.id,j}})));
  return out.filter(r=>r.sl>0||r.src.t==='pn').map(r=>{const x=IT.get(r.ma);return {...r,ten:x?x.ten:r.ma,dvt:x?x.dvt:'',g:priceFor(r.ncc,r.ma)};}); }
const kyRows=(ncc,ky)=>recvFull().filter(r=>r.ncc===ncc&&r.d>=ky.from&&r.d<=ky.to&&r.sl>0).sort((a,b)=>a.d.localeCompare(b.d)||a.ma.localeCompare(b.ma));
function lockedDay(ncc,d){ return kyList(d.slice(0,7),ncc).some(k=>d>=k.from&&d<=k.to&&((S.ds||{})[dsId(d.slice(0,7),ncc,k.k)]||{}).tt==='khop'); }
function dsNccs(){ const s=new Set(recvFull().map(r=>r.ncc)); poOf(S.thang).forEach(p=>s.add(p.ncc)); return [...s].filter(Boolean).sort((a,b)=>a.localeCompare(b,'vi')); }
async function editRecv(src,sl,note){
  if(src.t==='nh'){ const rows=[...S.nhan.rows]; rows[src.i]={...rows[src.i],sl}; await store.set('nhanhang/'+S.thang,{...S.nhan,rows,luc:Date.now()}); return; }
  const p=S.pn[src.id]; const lines=p.lines.map((l,j)=>j===src.j?{...l,sl,gach:sl===0?l.gach:false,note:[l.note,note].filter(Boolean).join(' · ')}:l);
  await store.set('phieunhan/'+p.id,{...stripId(p),lines,tong:Math.round(lines.reduce((a,l)=>a+(+l.sl||0)*l.g,0)),hist:[...(p.hist||[]),{ts:Date.now(),tt:p.tt,ai:uid,note}]}); }
async function addRecv(ncc,d,ma,sl,note){ const x=IT.get(ma)||{ten:ma,dvt:''}; const po=poOf(d.slice(0,7)).find(p=>p.ncc===ncc); const pid=d.slice(0,7)+'-'+Date.now().toString(36);
  await store.set('phieunhan/'+pid,{thang:d.slice(0,7),ma:`DS${d.slice(8)}${d.slice(5,7)}-${slug(ncc).slice(0,8).toUpperCase()}`,d,ncc,poId:po?po.id:'',soPO:po?po.soPO:'',dh:'',lines:[{ma,ten:x.ten,dvt:x.dvt,g:priceFor(ncc,ma),slDat:null,sl,gach:false,note,alloc:[],over:0}],tong:Math.round(sl*priceFor(ncc,ma)),nguon:'Thêm khi đối soát NCC',tt:'nhan',migo:'',luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'nhan',ai:uid,note}]}); }
async function setDS(ncc,ky,patch,note){ const id=dsId(S.thang,ncc,ky.k), o=(S.ds||{})[id]||{thang:S.thang,ncc,k:ky.k,from:ky.from,to:ky.to,tt:'chua',hist:[]};
  const rows=kyRows(ncc,ky); await store.set('doisoat/'+id,{...stripId(o),...patch,nDong:rows.length,tong:Math.round(rows.reduce((a,r)=>a+r.sl*r.g,0)),hist:[...(o.hist||[]),{ts:Date.now(),tt:patch.tt||o.tt,ai:uid,note:note||''}]}); }
function dsSel(){ const N=dsNccs(); if(!ED.dsN||!N.includes(ED.dsN)) ED.dsN=N[0]||''; const K=ED.dsN?kyList(S.thang,ED.dsN):[]; let ky=K.find(k=>k.k===ED.dsK); if(!ky){ const td=todayVN(); ky=K.find(k=>td>=k.from&&td<=k.to)||K[0]; ED.dsK=ky?ky.k:0; } return {N,K,ky}; }
function dsPicker(N,K,ky){ return `<div class="toolbar"><select id="ds-n">${N.map(n=>`<option ${n===ED.dsN?'selected':''}>${esc(n)}</option>`).join('')}</select>
  <select id="ds-k">${K.map(k=>`<option value="${k.k}" ${ky&&k.k===ky.k?'selected':''}>${kyName(k)}</option>`).join('')}</select>
  ${ky?`<span class="pill ${(TTDS[(dsDoc(ED.dsN,ky.k)||{}).tt||'chua'])[1]}">${(TTDS[(dsDoc(ED.dsN,ky.k)||{}).tt||'chua'])[0]}</span><span class="meta">Chu kỳ ${CKN[ckOf(ED.dsN)]}</span>`:''}</div>`; }
function bkEmail(ncc,ky,rows){ const tong=rows.reduce((a,r)=>a+r.sl*r.g,0), td='style="border:1px solid #ccc;padding:4px 8px"', tr='style="border:1px solid #ccc;padding:4px 8px;text-align:right"';
  const subj=`Báo cáo công nợ ${ncc} - ${kyName(ky)} - Bếp Vinmec Ocean Park 2`;
  const html=`<p>Kính gửi Quý nhà cung cấp <b>${esc(ncc)}</b>,</p><p>Bộ phận Bếp – Bệnh viện ĐKQT Vinmec Ocean Park 2 gửi bảng kê hàng đã nhận <b>${kyName(ky)}</b>, theo biên bản giao nhận hằng ngày. Tổng: <b>${VND(tong)} đ</b> (${rows.length} dòng).</p>
<table style="border-collapse:collapse;font-size:13px"><thead><tr><th ${td}>STT</th><th ${td}>Ngày giao</th><th ${td}>Mã hàng</th><th ${td}>Tên hàng</th><th ${td}>ĐVT</th><th ${td}>Số lượng</th><th ${td}>Đơn giá</th><th ${td}>Thành tiền</th></tr></thead><tbody>
${rows.map((r,i)=>`<tr><td ${td}>${i+1}</td><td ${td}>${r.d.split('-').reverse().join('/')}</td><td ${td}>${r.ma}</td><td ${td}>${esc(r.ten)}</td><td ${td}>${esc(r.dvt)}</td><td ${tr}>${Q(r.sl)}</td><td ${tr}>${VND(r.g)}</td><td ${tr}>${VND(r.sl*r.g)}</td></tr>`).join('')}
<tr><td ${td} colspan="7"><b>Tổng cộng</b></td><td ${tr}><b>${VND(tong)}</b></td></tr></tbody></table>
<p>Đề nghị Quý đơn vị kiểm tra và phản hồi email này: xác nhận khớp, hoặc gửi file chốt công nợ để Bếp đối chiếu. Sau khi khớp, Bếp nhập kho trên SAP theo đúng số này và Quý đơn vị xuất hóa đơn theo số PO.</p><p>Trân trọng,<br>Bộ phận Bếp – Vinmec Ocean Park 2</p>`;
  return {subj,html,text:`${subj}\n`+rows.map((r,i)=>`${i+1}. ${r.d.split('-').reverse().join('/')} ${r.ma} ${r.ten}: ${Q(r.sl)} ${r.dvt} x ${VND(r.g)} = ${VND(r.sl*r.g)}`).join('\n')+`\nTổng: ${VND(tong)} đ`}; }

/* ---- 5.1 ---- */
function vBK(){
  const {N,K,ky}=dsSel(); if(!N.length) return head2('5.1 Bảng kê & gửi NCC','')+'<div class="empty"><b>Chưa có số nhận hàng tháng này</b>Ghi phiếu nhận ở bước 4.2.</div>';
  const rows=kyRows(ED.dsN,ky), doc=dsDoc(ED.dsN,ky.k), lock=doc&&doc.tt==='khop'; ED.bk=ED.bk||{};
  const tong=rows.reduce((a,r)=>a+r.sl*r.g,0), em=(S.ck&&S.ck.email&&S.ck.email[ED.dsN])||((NCCINFO[ED.dsN]||{}).email||[]);
  const body=rows.map((r,i)=>{const k=JSON.stringify(r.src); return `<tr class="${ED.bk[k]!=null?'edited':''}"><td>${dm(r.d)}</td><td class="num meta">${esc(r.phieu)}</td><td class="num">${r.ma}</td><td class="wrap">${esc(r.ten)}</td><td>${esc(r.dvt)}</td>
    <td class="r">${lock||!canWrite?Q(r.sl):`<input class="q" type="text" inputmode="decimal" id="bk-${i}" data-bk='${esc(k)}' value="${QI(ED.bk[k]!=null?ED.bk[k]:r.sl)}">`}</td><td class="r num">${VND(r.g)}</td><td class="r num">${VND(r.sl*r.g)}</td><td class="num">${esc(r.po||'')}</td></tr>`;}).join('');
  return head2('5.1 Bảng kê nhận hàng & gửi NCC','Số nhận hàng là dữ liệu chuẩn. Cuối mỗi chu kỳ (15 ngày hoặc tháng, theo từng NCC), gửi bảng kê cho NCC đối chiếu như dự án công nợ cũ: bảng nhúng trong email nháp, NCC xác nhận hoặc gửi file chốt. Sửa số ở đây là sửa thẳng vào phiếu nhận.')+
  dsPicker(N,K,ky)+
  `<div class="kpis"><div class="kv big"><b>${esc(ED.dsN)} · ${kyName(ky)}</b><div class="num">${TR(tong)}</div><small>${rows.length} dòng nhận</small></div>
   <div class="kv"><b>Email NCC</b><div style="font-size:14px;word-break:break-all">${esc(em.join(', ')||'chưa có')}</div><small>${esc((NCCINFO[ED.dsN]||{}).note||'')}</small></div>
   <div class="kv"><b>Hạn đối soát</b><div class="num">${dm(ky.to)}</div><small>${todayVN()>=ky.to?'Đã đến hạn':'Chưa hết kỳ'}</small></div></div>
  ${lock?'<div class="warn ok">Kỳ này đã chốt khớp, số nhận bị khóa. Mở lại ở 5.3 nếu cần sửa.</div>':''}
  <div class="toolbar">${canWrite&&!lock?`<button class="btn" data-act5="bk-save" ${Object.keys(ED.bk).length?'':'disabled'}>Lưu sửa (${Object.keys(ED.bk).length})</button>
    <button class="btn primary" data-act5="bk-draft" ${rows.length&&em.length&&mcp&&!ED.drafting?'':'disabled'}>${ED.drafting?'Đang tạo nháp…':'Tạo nháp email bảng kê'}</button><button class="btn" data-act5="bk-copy" ${rows.length?'':'disabled'}>Copy bảng kê</button>
    <button class="btn" data-act5="bk-sent" ${rows.length?'':'disabled'}>Đánh dấu đã gửi</button>`:''}${doc&&doc.draftUrl?`<a class="meta" href="${esc(doc.draftUrl)}" target="_blank" rel="noopener">Mở nháp Gmail</a>`:''}</div>
  ${mcp?'':'<div class="meta">Tạo nháp Gmail cần mở trang trên claude.ai và cho phép kết nối Gmail. Trang chỉ tạo nháp, anh tự bấm Gửi.</div>'}
  ${rows.length?`<div class="tbl-wrap"><table class="t"><thead><tr><th>Ngày</th><th>Phiếu</th><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th class="r">Số lượng</th><th class="r">Đơn giá</th><th class="r">Thành tiền</th><th>PO</th></tr></thead><tbody>${body}</tbody><tfoot><tr><td colspan="7">Tổng</td><td class="r num">${VND(tong)}</td><td></td></tr></tfoot></table></div>`:'<div class="empty"><b>Không có hàng nhận trong kỳ này</b></div>'}`;
}
/* ---- 5.2 So file NCC ---- */
function parseNccTable(aoa,ky){
  const nb=s=>String(s??'').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/đ/g,'d');
  let hr=-1,col={};
  for(let i=0;i<Math.min(20,aoa.length)&&hr<0;i++){ const r=aoa[i].map(nb); const f=re=>r.findIndex(c=>re.test(c));
    const c={d:f(/ngay|date/),ma:f(/ma ?(hang|vt|vat tu|sp|san pham)|^ma$|code/),ten:f(/ten/),sl:f(/so luong|^sl|qty|khoi luong/),g:f(/don gia|^gia/),tt:f(/thanh tien|doanh so|thanh toan/)};
    if(c.ma>=0&&(c.sl>=0||c.tt>=0)){hr=i;col=c;} }
  if(hr<0) return {err:'Không tìm thấy dòng tiêu đề có cột Mã hàng và Số lượng/Thành tiền trong 20 dòng đầu.'};
  let lastD=''; const rows=[];
  aoa.slice(hr+1).forEach(r=>{ let d=col.d>=0?r[col.d]:''; if(typeof d==='number'){const t=new Date(Date.UTC(1899,11,30)+d*864e5);d=t.toISOString().slice(0,10);} else if(d){const m=String(d).match(/(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/); d=m?`${m[3]?(m[3].length===2?'20'+m[3]:m[3]):ky.from.slice(0,4)}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`:'';}
    if(d) lastD=d; else d=lastD; const ma=String(r[col.ma]||'').trim().toUpperCase(); if(!/^[A-Z]{2,4}\d{3,}/.test(ma)) return;
    const sl=col.sl>=0?pnum(r[col.sl]):0, g=col.g>=0?pnum(String(r[col.g]).replace(/\./g,'').replace(',','.')):0, tt=col.tt>=0?pnum(String(r[col.tt]).replace(/\./g,'').replace(',','.')):sl*g;
    if(!d||d<ky.from||d>ky.to) return; rows.push({d,ma,ten:col.ten>=0?String(r[col.ten]||''):'',sl:sl||(g?tt/g:0),g:g||(sl?tt/sl:0),tt}); });
  return {rows,hr,col};
}
function compareNcc(ours,theirs){
  const agg=a=>{const m=new Map();a.forEach(r=>{const k=r.d+'|'+r.ma;const o=m.get(k)||{d:r.d,ma:r.ma,ten:r.ten,sl:0,tt:0,g:r.g,src:[]};o.sl+=r.sl;o.tt+=r.sl*r.g;if(r.src)o.src.push(r.src);m.set(k,o);});return m;};
  const A=agg(ours), B=agg(theirs.map(r=>({...r,sl:r.sl,g:r.g}))), out=[];
  for(const [k,a] of A){ const b=B.get(k); if(!b) continue; A.delete(k); B.delete(k);
    if(Math.abs(a.sl-b.sl)>0.001) out.push({loai:'LỆCH SỐ LƯỢNG',a,b}); else if(Math.abs(a.g-b.g)>1) out.push({loai:'LỆCH ĐƠN GIÁ',a,b}); }
  for(const [ka,a] of [...A]){ const bm=[...B.entries()].find(([kb,b])=>b.d===a.d&&Math.abs(b.tt-a.tt)<=1); if(bm){out.push({loai:'KHÁC MÃ HÀNG',a,b:bm[1]});A.delete(ka);B.delete(bm[0]);continue;}
    const bd=[...B.entries()].find(([kb,b])=>b.ma===a.ma&&Math.abs(b.tt-a.tt)<=1&&Math.abs((P(b.d)-P(a.d))/864e5)<=3); if(bd){out.push({loai:'LỆCH NGÀY',a,b:bd[1]});A.delete(ka);B.delete(bd[0]);} }
  A.forEach(a=>out.push({loai:'MÌNH CÓ – NCC KHÔNG TÍNH',a,b:null})); B.forEach(b=>out.push({loai:'NCC TÍNH – MÌNH KHÔNG CÓ',a:null,b}));
  return out.sort((x,y)=>(x.a||x.b).d.localeCompare((y.a||y.b).d)); }
function vSoFile(){
  const {N,K,ky}=dsSel(); if(!N.length) return head2('5.2 So file chốt của NCC','')+'<div class="empty"><b>Chưa có số nhận hàng</b></div>';
  const doc=dsDoc(ED.dsN,ky.k), R=ED.sf&&ED.sf.key===ED.dsN+'|'+ky.k?ED.sf:null, lock=doc&&doc.tt==='khop';
  const ours=kyRows(ED.dsN,ky), tO=ours.reduce((a,r)=>a+r.sl*r.g,0);
  let res=''; if(R){ if(R.err) res=`<div class="warn">${esc(R.err)}</div>`; else { const D=compareNcc(ours,R.rows), tN=R.rows.reduce((a,r)=>a+r.tt,0);
    res=`<div class="kpis"><div class="kv ${D.length?'':'big'}"><b>Kết luận</b><div>${D.length?'<span class="pill bad">LỆCH</span>':'<span class="pill ok">KHỚP</span>'}</div><small>${D.length} chênh lệch</small></div>
      <div class="kv"><b>Bếp ghi nhận</b><div class="num">${VND(tO)}</div><small>${ours.length} dòng</small></div><div class="kv"><b>NCC tính</b><div class="num">${VND(tN)}</div><small>${R.rows.length} dòng · ${esc(R.ten||'')}</small></div><div class="kv"><b>Chênh tiền</b><div class="num">${VND(tN-tO)}</div></div></div>
    ${D.length?`<div class="tbl-wrap"><table class="t"><thead><tr><th>Loại</th><th>Ngày</th><th>Mã</th><th>Tên</th><th class="r">SL Bếp</th><th class="r">SL NCC</th><th class="r">Giá Bếp</th><th class="r">Giá NCC</th><th class="r">Tiền Bếp</th><th class="r">Tiền NCC</th><th></th></tr></thead><tbody>
      ${D.map((x,i)=>{const a=x.a||{},b=x.b||{};const act=lock||!canWrite?'':x.loai==='LỆCH SỐ LƯỢNG'?`<button class="btn sm" data-act5="sf-fix" data-i="${i}">Sửa theo NCC</button>`:x.loai==='NCC TÍNH – MÌNH KHÔNG CÓ'?`<button class="btn sm" data-act5="sf-add" data-i="${i}">Thêm vào phiếu nhận</button>`:x.loai==='MÌNH CÓ – NCC KHÔNG TÍNH'?`<button class="btn sm warn" data-act5="sf-zero" data-i="${i}">Bỏ dòng của Bếp</button>`:'<span class="meta">trao đổi NCC</span>';
        return `<tr><td><span class="pill ${x.loai.startsWith('LỆCH ĐƠN')||x.loai.startsWith('LỆCH NGÀY')?'run':'bad'}">${x.loai}</span></td><td>${dm((x.a||x.b).d)}${x.a&&x.b&&x.a.d!==x.b.d?' / '+dm(x.b.d):''}</td><td class="num">${a.ma||''}${x.a&&x.b&&a.ma!==b.ma?' / '+b.ma:(!x.a?b.ma:'')}</td><td class="wrap">${esc(a.ten||b.ten||'')}</td>
          <td class="r num">${x.a?Q(a.sl):''}</td><td class="r num">${x.b?Q(b.sl):''}</td><td class="r num">${x.a?VND(a.g):''}</td><td class="r num">${x.b?VND(b.g):''}</td><td class="r num">${x.a?VND(a.tt):''}</td><td class="r num">${x.b?VND(b.tt):''}</td><td>${act}</td></tr>`;}).join('')}</tbody></table></div>`:''}
    ${canWrite&&!lock?`<div class="toolbar"><span class="sp"></span>${D.length?'<button class="btn" data-act5="sf-mark-lech">Ghi nhận lệch</button><button class="btn warn" data-act5="sf-force">Cho phép chốt dù lệch</button>':'<button class="btn primary" data-act5="sf-khop">Chốt KHỚP theo file NCC</button>'}</div>`:''}`; ED.sfD=D; } }
  return head2('5.2 So file chốt của NCC','Tải file Excel NCC gửi (sổ chi tiết / bảng kê), trang tự nhận cột Ngày, Mã hàng, Tên, Số lượng, Đơn giá, Thành tiền trong 20 dòng đầu và so từng dòng theo (ngày, mã hàng): lệch số lượng, lệch đơn giá, khác mã hàng, lệch ngày (≤ 3 ngày), Bếp có – NCC không tính, NCC tính – Bếp không có. Sửa đến khi khớp thì chốt.')+
  dsPicker(N,K,ky)+(lock?'<div class="warn ok">Kỳ này đã chốt khớp.</div>':'')+
  `<div class="toolbar"><label class="btn" for="sf-file">Chọn file NCC (.xlsx, .xls, .csv)</label><input type="file" id="sf-file" accept=".xlsx,.xls,.csv" hidden>
   <button class="btn sm" data-act5="sf-paste">Hoặc dán bảng</button>${R?'<button class="btn sm" data-act5="sf-clear">Bỏ file</button>':''}</div>
  ${ED.sfPaste?`<div class="box"><label class="meta" for="sf-txt">Dán bảng copy từ Excel, gồm cả dòng tiêu đề (Ngày, Mã hàng, Số lượng, Đơn giá, Thành tiền).</label><textarea class="paste" id="sf-txt"></textarea><div class="toolbar"><button class="btn sm primary" data-act5="sf-paste-ok">So sánh</button></div></div>`:''}
  ${res||'<div class="empty"><b>Chưa có file NCC</b>Chọn file hoặc dán bảng NCC gửi để so.</div>'}`;
}
/* ---- 5.3 Chốt khớp ---- */
function vDSK(){
  const N=dsNccs(), td=todayVN();
  const rows=N.flatMap(n=>kyList(S.thang,n).map(ky=>({n,ky,doc:dsDoc(n,ky.k),rows:kyRows(n,ky)})));
  const due=rows.filter(r=>td>=r.ky.to&&r.rows.length&&(!r.doc||r.doc.tt!=='khop'));
  return head2('5.3 Chốt khớp theo kỳ','Mỗi NCC mỗi kỳ chỉ được dùng cho bước 6 (chỉnh PO, PR bổ sung) và bước 7 (MIGO) khi đã chốt KHỚP: NCC xác nhận, file NCC khớp, hoặc anh cho phép. Chốt xong số nhận của kỳ bị khóa. Chu kỳ lấy theo Danh mục NCC, sửa được ở cột Chu kỳ.')+
  `<div class="kpis"><div class="kv big"><b>Đến hạn chưa khớp</b><div class="num">${due.length}</div><small>NCC × kỳ</small></div>
   <div class="kv"><b>Đã khớp</b><div class="num">${rows.filter(r=>r.doc&&r.doc.tt==='khop').length}/${rows.filter(r=>r.rows.length).length}</div><small>kỳ có hàng nhận</small></div></div>
  <div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>NCC</th><th>Chu kỳ</th><th>Kỳ</th><th class="r">Dòng</th><th class="r">Tiền</th><th>Trạng thái</th><th>Căn cứ</th><th></th></tr></thead><tbody>
  ${rows.map(r=>{const st=(r.doc&&r.doc.tt)||'chua', tien=r.rows.reduce((a,x)=>a+x.sl*x.g,0), due=td>=r.ky.to, ck=ckOf(r.n);
    return `<tr><td><b>${esc(r.n)}</b></td><td>${r.ky.k<=1&&canWrite?`<select data-ck="${esc(r.n)}">${[15,30,0].map(v=>`<option value="${v}" ${v===ck?'selected':''}>${CKN[v]}</option>`).join('')}</select>`:''}</td><td>${kyName(r.ky)}${due?'':' <span class="pill">chưa hết kỳ</span>'}</td>
    <td class="r num">${r.rows.length}</td><td class="r num">${VND(tien)}</td><td><span class="pill ${TTDS[st][1]}">${TTDS[st][0]}</span></td><td class="meta">${esc((r.doc&&r.doc.canCu)||'')}</td>
    <td>${canWrite?(st==='khop'?`<button class="btn sm" data-act5="ds-open" data-n="${esc(r.n)}" data-k="${r.ky.k}">Mở lại</button>`:(r.rows.length?`<select id="cc-${slug(r.n)}-${r.ky.k}"><option>NCC xác nhận qua email</option><option>${ck===0?'Không báo NCC – Bếp tự xác nhận':'Người dùng cho phép'}</option></select> <button class="btn sm primary" data-act5="ds-khop" data-n="${esc(r.n)}" data-k="${r.ky.k}">Chốt khớp</button>`:'')):''}</td></tr>`;}).join('')}
  </tbody></table></div>`;
}

/* ================= BƯỚC 6 · CHỈNH PO & PR BỔ SUNG (CUỐI CHU KỲ) ================= */
function mocList(){ const ds=monthDays(S.thang); return [{m:'15',d:S.thang+'-15',ten:'Mốc ngày 15 (kỳ 1 của NCC 15 ngày)'},{m:'cuoi',d:ds[ds.length-1],ten:`Mốc cuối tháng ${dm(ds[ds.length-1])}`}]; }
function mocKy(ncc,moc){ return kyList(S.thang,ncc).filter(k=>k.to<=moc.d); }
function cpoCompute(moc){
  const last=moc.m==='cuoi', out=[];
  dsNccs().forEach(ncc=>{ const kys=mocKy(ncc,moc); if(!kys.length) return; const endK=kys[kys.length-1]; if(endK.to!==moc.d) return;
    const ok=kys.every(k=>{const d=dsDoc(ncc,k.k);return (d&&d.tt==='khop')||!kyRows(ncc,k).length;});
    const recv=new Map(); kys.forEach(k=>kyRows(ncc,k).forEach(r=>recv.set(r.ma,(recv.get(r.ma)||0)+r.sl)));
    const po=poOf(S.thang).find(p=>p.ncc===ncc&&p.tt!=='huy'); const poQ=new Map(); if(po) po.lines.forEach(l=>poQ.set(l.ma,(poQ.get(l.ma)||0)+l.sl));
    const prQ=new Map(); batchesOf(S.thang).filter(b=>b.tt==='duyet').forEach(b=>b.lines.filter(l=>l.ncc===ncc).forEach(l=>prQ.set(l.ma,(prQ.get(l.ma)||0)+l.sl)));
    const pendBS=new Map(); batchesOf(S.thang).filter(b=>b.loai==='Bổ sung'&&b.tt!=='huy'&&b.tt!=='duyet').forEach(b=>b.lines.filter(l=>l.ncc===ncc).forEach(l=>pendBS.set(l.ma,(pendBS.get(l.ma)||0)+l.sl)));
    new Set([...recv.keys(),...poQ.keys()]).forEach(ma=>{ const r=recv.get(ma)||0, q=poQ.get(ma)||0, pr=prQ.get(ma)||0, x=IT.get(ma)||{ten:ma,dvt:''};
      let act='khop', tang=0, bs=0, giam=0; const free=Math.max(0,pr-q);
      if(r>q+1e-9){ const need=r-q; tang=Math.min(need,free); bs=+(need-tang).toFixed(3); act=bs>0?'bs':'tang'; }
      else if(last&&q>r+1e-9){ giam=+(q-r).toFixed(3); act='giam'; }
      if(act!=='khop'||r>0) out.push({ncc,po,ma,ten:x.ten,dvt:x.dvt,g:priceFor(ncc,ma),r,q,pr,free,tang:+tang.toFixed(3),bs,giam,act,ok,pend:pendBS.get(ma)||0}); }); });
  return out; }
const CPO={bs:['Cần PR bổ sung','bad'],tang:['Tăng PO (PR đủ)','run'],giam:['Giảm PO = số nhận','run'],khop:['Khớp','ok']};
function vCPO(){
  const M=mocList(); if(!ED.moc) ED.moc=todayVN()>M[0].d?'cuoi':'15'; const moc=M.find(m=>m.m===ED.moc);
  const R=cpoCompute(moc), nOk=R.filter(r=>r.ok), notOk=[...new Set(R.filter(r=>!r.ok).map(r=>r.ncc))];
  const bsDoc=(S.prbs||{})[S.thang+'_m'+moc.m], cnt={}; nOk.forEach(r=>cnt[r.act]=(cnt[r.act]||0)+1);
  const f=ED.cpoF||''; const shown=R.filter(r=>(!f||r.act===f)&&(r.act!=='khop'||f==='khop'));
  return head2('6.1 Đề xuất chỉnh PO cuối chu kỳ','Khi số nhận của NCC đã chốt khớp (bước 5), so lũy kế số nhận với PO: thiếu mà PR còn trống thì tăng PO; PR không đủ thì gom vào MỘT đợt PR bổ sung duy nhất của mốc này; mốc cuối tháng thì giảm PO bằng đúng số nhận để PO = MIGO. Sửa PO xong nhớ sửa trên SAP (ME22N).')+
  `<div class="toolbar"><div class="seg">${M.map(m=>`<button data-act6="moc" data-m="${m.m}" aria-pressed="${m.m===ED.moc}">${m.ten}</button>`).join('')}</div></div>
  ${notOk.length?`<div class="warn">Chưa chốt khớp nên chưa chỉnh: <b>${notOk.map(esc).join(', ')}</b> (bước 5.3).</div>`:''}
  ${bsDoc?`<div class="warn ok">Mốc này đã tạo PR bổ sung lúc ${tstr(bsDoc.luc)}: ${(bsDoc.batches||[]).map(id=>esc((S.prdot[id]||{}).nhom||id)).join(', ')}. Mỗi mốc chỉ tạo 1 lần; muốn làm lại thì hủy các đợt đó ở 2.3.</div>`:''}
  <div class="kpis"><div class="kv big"><b>Cần PR bổ sung</b><div class="num">${cnt.bs||0} dòng</div><small>${TR(nOk.filter(r=>r.act==='bs').reduce((a,r)=>a+r.bs*r.g,0))}</small></div>
   <div class="kv"><b>Tăng PO (đủ PR)</b><div class="num">${cnt.tang||0}</div><small>${TR(nOk.filter(r=>r.act==='tang').reduce((a,r)=>a+r.tang*r.g,0))}</small></div>
   <div class="kv"><b>Giảm PO</b><div class="num">${cnt.giam||0}</div><small>${TR(nOk.filter(r=>r.act==='giam').reduce((a,r)=>a+r.giam*r.g,0))}</small></div>
   <div class="kv"><b>Đã khớp</b><div class="num">${cnt.khop||0}</div></div></div>
  <div class="toolbar"><div class="seg"><button data-act6="cpo-f" data-f="" aria-pressed="${!f}">Cần xử lý</button>${Object.entries(CPO).map(([k,v])=>`<button data-act6="cpo-f" data-f="${k}" aria-pressed="${f===k}">${v[0]} ${R.filter(r=>r.act===k).length}</button>`).join('')}</div><span class="sp"></span>
   ${canWrite?`<button class="btn" data-act6="cpo-apply" ${nOk.some(r=>r.act==='tang'||r.act==='giam')?'':'disabled'}>Áp dụng tăng/giảm PO</button><button class="btn primary" data-act6="cpo-bs" ${nOk.some(r=>r.act==='bs'&&r.bs-r.pend>0)&&!bsDoc?'':'disabled'}>Tạo PR bổ sung mốc này</button>`:''}</div>
  ${shown.length?`<div class="tbl-wrap"><table class="t"><thead><tr><th>NCC · PO</th><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th class="r">Nhận đã khớp</th><th class="r">SL PO</th><th class="r">PR đã duyệt</th><th class="r">Tăng PO</th><th class="r">PR bổ sung</th><th class="r">Giảm PO</th><th>Đề xuất</th></tr></thead><tbody>
  ${shown.map(r=>`<tr${r.ok?'':' style="opacity:.55"'}><td>${esc(r.ncc)}<div class="meta num">${r.po?esc(r.po.soPO||'chưa có số'):'chưa có PO'}</div></td><td class="num">${r.ma}</td><td class="wrap">${esc(r.ten)}</td><td>${esc(r.dvt)}</td><td class="r num">${Q(r.r)}</td><td class="r num">${Q(r.q)}</td><td class="r num">${Q(r.pr)}</td><td class="r num">${r.tang?'+'+Q(r.tang):''}</td><td class="r num">${r.bs?Q(r.bs)+(r.pend?` <span class="meta">(đang chờ ${Q(r.pend)})</span>`:''):''}</td><td class="r num">${r.giam?'−'+Q(r.giam):''}</td><td><span class="pill ${CPO[r.act][1]}">${CPO[r.act][0]}</span>${r.ok?'':' <span class="pill">chưa khớp NCC</span>'}</td></tr>`).join('')}
  </tbody></table></div>`:'<div class="empty"><b>Không có dòng nào cần chỉnh ở mốc này</b>Hoặc chưa có NCC nào kết thúc kỳ tại mốc này.</div>'}`;
}
function addToPO(p,adds,note){ // adds:[{ma,sl,g,soPR,dt,ten,dvt}]
  const lines=p.lines.map(l=>({...l}));
  adds.forEach(a=>{ const ex=lines.filter(l=>l.ma===a.ma&&(a.soPR?l.soPR===a.soPR:true)); if(ex.length&&!a.soPR){ ex[0].sl=+(ex[0].sl+a.sl).toFixed(3); } else lines.push({ma:a.ma,ten:a.ten,dvt:a.dvt,sl:a.sl,g:a.g,soPR:a.soPR||'',dt:a.dt||''}); });
  return lines; }
function vGPO(){
  const used=new Set(poAll().filter(p=>p.tt!=='huy').flatMap(p=>p.src||[]));
  const B=batchesOf(S.thang).filter(b=>b.loai==='Bổ sung'&&b.tt!=='huy'), pend=[];
  B.forEach(b=>b.lines.forEach(l=>{const k=b.id+'|'+l.ncc; if(!used.has(k)) pend.push({b,l,k});}));
  const ready=pend.filter(x=>x.b.tt==='duyet');
  return head2('6.2 Gắn PR bổ sung vào PO','Khi đợt PR bổ sung đã duyệt (2.3), thêm các dòng đó vào PO tháng sẵn có của từng NCC (không lập PO mới). NCC chưa có PO thì lập PO mới. Sau đó sửa PO trên SAP (ME22N, thêm dòng tham chiếu PR).')+
  `<div class="toolbar">${canWrite?`<button class="btn primary" data-act6="gpo-apply" ${ready.length?'':'disabled'}>Gắn ${ready.length} dòng PR bổ sung đã duyệt vào PO</button>`:''}<span class="meta">${pend.length-ready.length} dòng PR bổ sung chưa duyệt</span></div>
  ${pend.length?`<div class="tbl-wrap"><table class="t"><thead><tr><th>Đợt</th><th>Số PR</th><th>Trạng thái</th><th>NCC</th><th>PO hiện có</th><th>Mã</th><th>Tên</th><th class="r">SL</th><th class="r">Tiền</th></tr></thead><tbody>
   ${pend.map(x=>{const po=poOf(S.thang).find(p=>p.ncc===x.l.ncc);return `<tr><td class="num">${esc(x.b.nhom)}</td><td class="num">${esc(x.b.soPR||'')}</td><td>${ttPill(x.b)}</td><td>${esc(x.l.ncc)}</td><td class="num">${po?esc(po.soPO||'(chưa có số)'):'<span class="pill run">sẽ lập PO mới</span>'}</td><td class="num">${x.l.ma}</td><td class="wrap">${esc(x.l.ten)}</td><td class="r num">${Q(x.l.sl)}</td><td class="r num">${VND(x.l.sl*x.l.g)}</td></tr>`;}).join('')}</tbody></table></div>`
  :'<div class="empty"><b>Không có PR bổ sung nào chờ gắn vào PO</b></div>'}`;
}
function vKPO(){
  const migoQ=new Map(); migoAll().filter(m=>m.thang===S.thang).forEach(m=>m.lines.forEach(l=>migoQ.set(m.poId+'|'+l.item,(migoQ.get(m.poId+'|'+l.item)||0)+l.sl)));
  const rows=[]; poOf(S.thang).forEach(po=>{ const kys=kyList(S.thang,po.ncc), okK=kys.filter(k=>(dsDoc(po.ncc,k.k)||{}).tt==='khop');
    const r=new Map(); okK.forEach(k=>kyRows(po.ncc,k).forEach(x=>r.set(x.ma,(r.get(x.ma)||0)+x.sl)));
    const q=new Map(),mg=new Map(); poItems(po).forEach(l=>{q.set(l.ma,(q.get(l.ma)||0)+l.sl); mg.set(l.ma,(mg.get(l.ma)||0)+(migoQ.get(po.id+'|'+l.item)||0));});
    const t={po,dong:q.size,q:0,r:0,mg:0,lech:0}; q.forEach((v,ma)=>{t.q+=v*priceFor(po.ncc,ma);t.r+=(r.get(ma)||0)*priceFor(po.ncc,ma);t.mg+=(mg.get(ma)||0)*priceFor(po.ncc,ma); if(Math.abs(v-(mg.get(ma)||0))>1e-6) t.lech++;});
    t.allK=okK.length===kys.length; rows.push(t); });
  return head2('6.3 PO so với số đã khớp và MIGO','Kiểm tra cuối tháng: mỗi PO phải bằng tổng MIGO (và bằng số nhận đã khớp với NCC) trên từng mã. ĐNTT lập theo số PO.')+
  `<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th>NCC</th><th>PO</th><th class="r">Tiền PO</th><th class="r">Nhận đã khớp</th><th class="r">Đã MIGO</th><th class="r">Mã lệch PO–MIGO</th><th>Kết luận</th></tr></thead><tbody>
  ${rows.map(t=>`<tr><td><b>${esc(t.po.ncc)}</b>${t.po.sua?' <span class="pill bad">cần sửa ME22N</span>':''}</td><td class="num">${esc(t.po.soPO||'')}</td><td class="r num">${VND(t.q)}</td><td class="r num">${VND(t.r)}</td><td class="r num">${VND(t.mg)}</td><td class="r num">${t.lech}</td><td>${!t.lech&&t.allK&&!t.po.sua?'<span class="pill ok">PO = MIGO · sẵn sàng ĐNTT</span>':t.allK?'<span class="pill run">Chưa khớp hết</span>':'<span class="pill">Còn kỳ chưa đối soát</span>'}</td></tr>`).join('')||'<tr><td colspan="7" class="muted">Chưa có PO.</td></tr>'}
  </tbody></table></div>`;
}

/* ================= BƯỚC 7 · MIGO CUỐI KỲ (1 MIGO / PO / KỲ) ================= */
const migoAll=()=>Object.values(S.migo||{}).filter(m=>m.tt!=='huy');
const TTMG={cho:['Chờ nhập kho','run'],xuat:['Đã gửi / xuất cho macro','run'],migo:['Đã MIGO','ok'],huy:['Đã hủy','bad']};
function mgCandidates(){
  const out=[]; const done=new Set(migoAll().map(m=>m.poId+'|'+m.k));
  poOf(S.thang).filter(p=>p.soPO).forEach(po=>kyList(S.thang,po.ncc).forEach(ky=>{ const d=dsDoc(po.ncc,ky.k); if(!d||d.tt!=='khop'||done.has(po.id+'|'+ky.k)) return;
    const rows=kyRows(po.ncc,ky); if(!rows.length) return;
    const used={}; migoAll().filter(m=>m.poId===po.id).forEach(m=>m.lines.forEach(l=>used[l.item]=(used[l.item]||0)+l.sl));
    const byMa=new Map(); rows.forEach(r=>byMa.set(r.ma,(byMa.get(r.ma)||0)+r.sl));
    const lines=[],short=[]; byMa.forEach((q,ma)=>{const a=allocItems(po,ma,q,used); a.alloc.forEach(x=>lines.push({item:x.item,ma,sl:x.sl})); if(a.over>0) short.push({ma,over:a.over});});
    out.push({po,ky,lines:lines.sort((a,b)=>a.item-b.item),short,tong:Math.round(rows.reduce((a,r)=>a+r.sl*r.g,0))}); }));
  return out; }
function vMG(){
  const C=mgCandidates(), W=migoAll().filter(m=>m.thang===S.thang&&(m.tt==='cho'||m.tt==='xuat'));
  if(!ED.mgsel) ED.mgsel=new Set(W.map(m=>m.id)); const sel=W.filter(m=>ED.mgsel.has(m.id));
  const rows=sel.flatMap(m=>m.lines.map(l=>[m.ma,m.soPO,m.ngay.split('-').reverse().join('.'),l.item,l.ma,l.sl,`Mua hàng phục vụ Bếp tháng ${+m.thang.slice(5)}`]));
  return head2('7.1 Lập & xuất MIGO cuối kỳ','Mỗi PO mỗi kỳ đã chốt khớp = 1 chứng từ MIGO, số lượng = số nhận đã khớp, ngày hạch toán = ngày cuối kỳ. Nếu PO chưa đủ số lượng thì chỉnh PO ở bước 6 trước (SAP không cho MIGO vượt PO). Bấm <b>Nhập kho trên SAP</b>: trạm SAP chạy MIGO cho từng chứng từ đã chọn, số chứng từ tự ghi về. PO phải đã cập nhật SAP (không còn sửa chờ ME22N).')+tramBox('MIGO')+
  `${C.length?`<h3 class="eyebrow" style="margin:0">Kỳ đã khớp chưa lập MIGO</h3><div class="tbl-wrap" style="max-height:300px"><table class="t"><thead><tr><th>NCC</th><th>PO</th><th>Kỳ</th><th class="r">Dòng MIGO</th><th class="r">Tiền</th><th>Kiểm tra PO</th></tr></thead><tbody>
   ${C.map(c=>`<tr><td>${esc(c.po.ncc)}</td><td class="num">${esc(c.po.soPO)}</td><td>${kyName(c.ky)}</td><td class="r num">${c.lines.length}</td><td class="r num">${VND(c.tong)}</td><td>${c.short.length?`<span class="pill bad">PO thiếu ${c.short.length} mã – chỉnh ở bước 6</span>`:c.po.sua?'<span class="pill run">PO đã sửa – cập nhật SAP ở 3.4 trước</span>':'<span class="pill ok">PO đủ</span>'}</td></tr>`).join('')}</tbody></table></div>
   ${canWrite?`<div class="toolbar"><button class="btn primary" data-act7="mg-make" ${C.some(c=>!c.short.length)?'':'disabled'}>Lập ${C.filter(c=>!c.short.length).length} chứng từ MIGO (PO đủ)</button></div>`:''}`:'<div class="meta">Không có kỳ đã khớp nào chờ lập MIGO.</div>'}
  ${W.length?`<h3 class="eyebrow" style="margin:6px 0 0">Chứng từ chờ chạy macro</h3><div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th></th><th>Ma_phieu</th><th>NCC</th><th>PO</th><th>Ngày hạch toán</th><th class="r">Dòng</th><th class="r">Tiền</th><th>Trạng thái</th></tr></thead><tbody>
   ${W.map(m=>`<tr><td><input type="checkbox" data-mgsel="${m.id}" ${ED.mgsel.has(m.id)?'checked':''} aria-label="Chọn"></td><td class="num">${esc(m.ma)}</td><td>${esc(m.ncc)}</td><td class="num">${esc(m.soPO)}</td><td>${m.ngay.split('-').reverse().join('/')}</td><td class="r num">${m.lines.length}</td><td class="r num">${VND(m.tong)}</td><td><span class="pill ${TTMG[m.tt][1]}">${TTMG[m.tt][0]}</span>${lenhPill(m.lenh)}${m.ghichu?`<div class="meta">${esc(m.ghichu)}</div>`:''}</td></tr>`).join('')}</tbody></table></div>
   ${canWrite?`<div class="toolbar"><button class="btn primary" data-tram="migo" ${sel.filter(m=>!lenhBusy(m.lenh)).length?'':'disabled'}>Nhập kho trên SAP (${sel.filter(m=>!lenhBusy(m.lenh)).length} chứng từ đã chọn)</button></div>`:''}
   <div class="toolbar"><button class="btn" data-act7="mg-xlsx" ${rows.length?'':'disabled'}>Tải file cho sheet MIGO_TP</button><button class="btn" data-act7="mg-copy" ${rows.length?'':'disabled'}>Copy để dán</button><span class="meta">${rows.length} dòng</span></div>
   ${canWrite?`<div class="box"><label class="meta" for="mg-txt">Dán kết quả từ sheet MIGO_TP: <b>Ma_phieu</b> ⇥ <b>So_chung_tu_MIGO</b>.</label><textarea class="paste" id="mg-txt" style="min-height:60px"></textarea><div class="toolbar"><button class="btn sm primary" data-act7="mg-ok">Ghi số MIGO</button></div></div>`:''}`:''}`;
}
function vMGls(){
  const L=migoAll().filter(m=>m.thang===S.thang).sort((a,b)=>b.luc-a.luc);
  return head2('7.2 Lịch sử MIGO','Các chứng từ MIGO cuối kỳ theo PO. Hủy được khi chưa có số chứng từ.')+
  (L.length?`<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th></th><th>Ma_phieu</th><th>NCC</th><th>PO</th><th>Kỳ</th><th>Ngày HT</th><th class="r">Tiền</th><th>Trạng thái</th><th>Số MIGO</th><th></th></tr></thead><tbody>
   ${L.map(m=>`<tr><td><button class="x" data-act7="mg-open" data-id="${m.id}">${ED.mgOpen===m.id?'▾':'▸'}</button></td><td class="num">${esc(m.ma)}</td><td>${esc(m.ncc)}</td><td class="num">${esc(m.soPO)}</td><td>${m.k===0?'Tháng':'Kỳ '+m.k}</td><td>${m.ngay.split('-').reverse().join('/')}</td><td class="r num">${VND(m.tong)}</td><td><span class="pill ${TTMG[m.tt][1]}">${TTMG[m.tt][0]}</span></td><td class="num">${esc(m.soMIGO||'')}</td><td>${canWrite&&m.tt!=='migo'?`<button class="btn sm warn" data-act7="mg-huy" data-id="${m.id}">Hủy</button>`:''}</td></tr>
   ${ED.mgOpen===m.id?`<tr><td></td><td colspan="9"><table class="t"><tbody>${m.lines.map(l=>`<tr><td class="num">Dòng ${l.item}</td><td class="num">${l.ma}</td><td>${esc((IT.get(l.ma)||{}).ten||'')}</td><td class="r num">${Q(l.sl)}</td></tr>`).join('')}</tbody></table></td></tr>`:''}`).join('')}</tbody></table></div>`:'<div class="empty"><b>Chưa có chứng từ MIGO</b>Lập ở 7.1.</div>');
}

/* ---- sự kiện 5–7 ---- */
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act5],[data-act6],[data-act7]'); if(!t) return; const a=t.dataset.act5||t.dataset.act6||t.dataset.act7;
  try{
  const {ky}=(t.dataset.act5?dsSel():{});
  if(a==='bk-save'){ for(const [k,v] of Object.entries(ED.bk)) await editRecv(JSON.parse(k),v,'Sửa khi đối soát '+tstr(Date.now())); await log('5.1',`Sửa ${Object.keys(ED.bk).length} dòng nhận hàng ${ED.dsN} (${kyName(ky)})`); ED.bk={}; }
  else if(a==='bk-copy'){ const E=bkEmail(ED.dsN,ky,kyRows(ED.dsN,ky)); try{ await navigator.clipboard.writeText(E.text); toast('Đã copy bảng kê.'); }catch(err){ toast('Không copy được.'); } }
  else if(a==='bk-draft'){ const rows=kyRows(ED.dsN,ky), E=bkEmail(ED.dsN,ky,rows), em=(NCCINFO[ED.dsN]||{}).email||[]; ED.drafting=true; render();
    try{ const r=await call('Gmail','create_draft',{to:em,subject:E.subj,htmlBody:E.html,body:E.text,includeSignature:true},'Tạo nháp Gmail bảng kê '+ED.dsN); const pl=(r&&r.payload)||{};
      await setDS(ED.dsN,ky,{tt:'gui',guiLuc:Date.now(),draftUrl:pl.viewUrl||''},'Tạo nháp email bảng kê'); await log('5.1',`Tạo nháp email bảng kê ${ED.dsN} ${kyName(ky)}`); toast('Đã tạo nháp trong Gmail. Kiểm tra rồi bấm Gửi.'); }
    catch(err){ toast(err&&err.code==='not_granted'?'Chưa cho phép trang dùng Gmail.':'Không tạo được nháp: '+((err&&(err.message||err.code))||err)); }
    ED.drafting=false; render(); }
  else if(a==='bk-sent'){ await setDS(ED.dsN,ky,{tt:'gui',guiLuc:Date.now()},'Đã gửi bảng kê'); await log('5.1',`Đã gửi bảng kê ${ED.dsN} ${kyName(ky)}`); }
  else if(a==='sf-paste'){ ED.sfPaste=!ED.sfPaste; render(); }
  else if(a==='sf-clear'){ ED.sf=null; render(); }
  else if(a==='sf-paste-ok'){ const aoa=$('#sf-txt').value.split(/\r?\n/).filter(Boolean).map(l=>l.split('\t')); ED.sf={key:ED.dsN+'|'+ky.k,ten:'bảng dán',...parseNccTable(aoa,ky)}; ED.sfPaste=false; render(); }
  else if(a==='sf-fix'||a==='sf-zero'||a==='sf-add'){ const x=ED.sfD[+t.dataset.i];
    if(a==='sf-add'){ await addRecv(ED.dsN,x.b.d,x.b.ma,x.b.sl,'Thêm theo file NCC'); }
    else { const target=a==='sf-zero'?0:x.b.sl; const src=x.a.src; let rem=target; for(let i=0;i<src.length;i++){ const r=recvFull().find(z=>JSON.stringify(z.src)===JSON.stringify(src[i])); const v=i===src.length-1?rem:Math.min(rem,r?r.sl:0); await editRecv(src[i],+(+v).toFixed(3),a==='sf-zero'?'Bỏ theo file NCC':'Sửa theo file NCC'); rem-=v; } }
    await log('5.2',`${a==='sf-add'?'Thêm':a==='sf-zero'?'Bỏ':'Sửa'} ${(x.a||x.b).ma} ngày ${dm((x.a||x.b).d)} theo file ${ED.dsN}`); }
  else if(a==='sf-khop'||a==='sf-force'){ await setDS(ED.dsN,ky,{tt:'khop',khopLuc:Date.now(),canCu:a==='sf-khop'?'Khớp file NCC':'Người dùng cho phép dù lệch'},a==='sf-khop'?'Khớp file NCC':'Cho phép chốt dù lệch'); await log('5.2',`Chốt khớp ${ED.dsN} ${kyName(ky)} (${a==='sf-khop'?'file NCC khớp':'cho phép dù lệch'})`); }
  else if(a==='sf-mark-lech'){ await setDS(ED.dsN,ky,{tt:'lech'},`Lệch ${ED.sfD.length} dòng so với file NCC`); }
  else if(a==='ds-khop'){ const n=t.dataset.n, kk=kyList(S.thang,n).find(k=>k.k===+t.dataset.k), cc=$(`#cc-${slug(n)}-${kk.k}`).value; await setDS(n,kk,{tt:'khop',khopLuc:Date.now(),canCu:cc},'Chốt khớp: '+cc); await log('5.3',`Chốt khớp ${n} ${kyName(kk)} – ${cc}`); }
  else if(a==='ds-open'){ const n=t.dataset.n, kk=kyList(S.thang,n).find(k=>k.k===+t.dataset.k); if(migoAll().some(m=>m.ncc===n&&m.k===kk.k&&m.thang===S.thang)){toast('Kỳ này đã lập MIGO; hủy MIGO ở 7.2 trước.');return;} await setDS(n,kk,{tt:'lech'},'Mở lại để sửa'); await log('5.3',`Mở lại kỳ ${n} ${kyName(kk)}`); }
  else if(a==='moc'){ ED.moc=t.dataset.m; render(); }
  else if(a==='cpo-f'){ ED.cpoF=t.dataset.f; render(); }
  else if(a==='cpo-apply'){ const moc=mocList().find(m=>m.m===ED.moc), R=cpoCompute(moc).filter(r=>r.ok&&(r.act==='tang'||r.act==='giam')&&r.po); const by=new Map(); R.forEach(r=>(by.get(r.po.id)||by.set(r.po.id,[]).get(r.po.id)).push(r));
    for(const [id,list] of by){ const po=S.po[id]; const lines=po.lines.map(l=>({...l})); const notes=[];
      list.forEach(r=>{ const L=lines.filter(l=>l.ma===r.ma); if(r.act==='tang'){ L[L.length-1].sl=+(L[L.length-1].sl+r.tang).toFixed(3); notes.push(`${r.ma} +${Q(r.tang)}`); }
        else { let rem=r.giam; for(let i=L.length-1;i>=0&&rem>0;i--){const d=Math.min(L[i].sl,rem);L[i].sl=+(L[i].sl-d).toFixed(3);rem-=d;} notes.push(`${r.ma} −${Q(r.giam)}`); } });
      await setPO(id,{lines,tong:Math.round(lines.reduce((s,l)=>s+l.sl*l.g,0)),sua:true},`Chỉnh cuối kỳ (${moc.ten}): `+notes.join('; ')); await log('6.1',`Chỉnh PO ${po.soPO} (${po.ncc}): ${notes.join('; ')} – cần sửa ME22N`); }
    toast('Đã chỉnh PO. Sang 3.4 bấm “Cập nhật trên SAP” để trạm sửa PO (ME22N).'); }
  else if(a==='cpo-bs'){ const moc=mocList().find(m=>m.m===ED.moc), key=S.thang+'_m'+moc.m; if((S.prbs||{})[key]){toast('Mốc này đã tạo PR bổ sung.');return;}
    const items=cpoCompute(moc).filter(r=>r.ok&&r.act==='bs'&&r.bs-r.pend>0).map(r=>({ma:r.ma,ncc:r.ncc,g:r.g,dvt:r.dvt,q:roundQ(r.bs-r.pend,r.dvt)}));
    const before=new Set(Object.keys(S.prdot)); await makeBSFrom(items,`PR bổ sung cuối chu kỳ – ${moc.ten}`,'6.1');
    await new Promise(r=>setTimeout(r,300)); const ids=Object.keys(S.prdot).filter(id=>!before.has(id));
    await store.set('prbs/'+key,{thang:S.thang,moc:moc.m,batches:ids,luc:Date.now(),ai:uid,nDong:items.length}); toast('Đã tạo PR bổ sung (1 lần cho mốc này). Xuất ở 2.2, duyệt ở 2.3, rồi gắn vào PO ở 6.2.'); }
  else if(a==='gpo-apply'){ const used=new Set(poAll().filter(p=>p.tt!=='huy').flatMap(p=>p.src||[])); const by=new Map();
    batchesOf(S.thang).filter(b=>b.loai==='Bổ sung'&&b.tt==='duyet').forEach(b=>b.lines.forEach(l=>{const k=b.id+'|'+l.ncc; if(used.has(k)) return; const o=by.get(l.ncc)||{src:new Set(),adds:[],pr:new Set()}; o.src.add(k); o.pr.add(b.soPR); o.adds.push({...l,soPR:b.soPR,dt:b.dt}); by.set(l.ncc,o);}));
    for(const [ncc,o] of by){ const po=poOf(S.thang).find(p=>p.ncc===ncc);
      if(po){ const lines=addToPO(po,o.adds); await setPO(po.id,{lines,tong:Math.round(lines.reduce((s,l)=>s+l.sl*l.g,0)),src:[...(po.src||[]),...o.src],pr:[...new Set([...(po.pr||[]),...o.pr])],sua:true},`Thêm ${o.adds.length} dòng từ PR bổ sung ${[...o.pr].join(', ')}`); }
      else { const pid=S.thang+'-'+slug(ncc)+'-'+Date.now().toString(36); await store.set('po/'+pid,{thang:S.thang,ncc,lifnr:LIFNR[ncc.toUpperCase()]||'',loai:'Bổ sung',src:[...o.src],pr:[...o.pr],lines:o.adds.map(({ncc:_,...l})=>l),tong:Math.round(o.adds.reduce((s,l)=>s+l.sl*l.g,0)),tt:'cho',soPO:'',tongSAP:null,luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'cho',ai:uid,note:'Lập từ PR bổ sung'}]}); } }
    await log('6.2',`Gắn PR bổ sung vào PO: ${[...by.keys()].join(', ')}`); toast('Đã gắn vào PO. Sang 3.4 bấm “Cập nhật trên SAP” để trạm thêm dòng (ME22N).'); }
  else if(a==='mg-make'){ const C=mgCandidates().filter(c=>!c.short.length); let n=0;
    for(const c of C){ const id=`${S.thang}-${slug(c.po.ncc)}-k${c.ky.k}-${Date.now().toString(36)}${n}`; const ma=`MG${tK(S.thang)}${c.ky.k?'K'+c.ky.k:''}-${slug(c.po.ncc).slice(0,8).toUpperCase()}`;
      await store.set('migo/'+id,{thang:S.thang,k:c.ky.k,ma,ncc:c.po.ncc,poId:c.po.id,soPO:c.po.soPO,ngay:c.ky.to,lines:c.lines,tong:c.tong,tt:'cho',soMIGO:'',luc:Date.now(),ai:uid}); n++; }
    ED.mgsel=null; await log('7.1',`Lập ${n} chứng từ MIGO cuối kỳ`); }
  else if(a==='mg-copy'||a==='mg-xlsx'){ const W=migoAll().filter(m=>m.thang===S.thang&&(m.tt==='cho'||m.tt==='xuat')&&ED.mgsel.has(m.id));
    const rows=W.flatMap(m=>m.lines.map(l=>[m.ma,m.soPO,m.ngay.split('-').reverse().join('.'),l.item,l.ma,l.sl,`Mua hàng phục vụ Bếp tháng ${+m.thang.slice(5)}`]));
    if(a==='mg-copy'){ try{ await navigator.clipboard.writeText(rows.map(r=>r.map((v,j)=>j===5?String(v).replace('.',','):String(v)).join('\t')).join('\r\n')); toast('Đã copy '+rows.length+' dòng. Dán vào sheet MIGO_TP ô A3.'); }catch(err){ toast('Không copy được.'); return; } }
    else { if(!window.XLSX){toast('Không tạo được file; dùng nút Copy.');return;} const ws=XLSX.utils.aoa_to_sheet([['Ma_phieu','So_PO','Ngay_hach_toan','Dong_PO','Ma_vat_tu','So_luong','Text'],...rows]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'MIGO_TP');
      try{ await saveOut(`MIGO_TP_${S.thang}_${todayVN()}.xlsx`,new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); }catch(err){ if(err&&err.code!=='cancelled') toast('Không lưu được file.'); return; } }
    for(const m of W) if(m.tt==='cho') await store.set('migo/'+m.id,{...stripId(m),tt:'xuat',xuatLuc:Date.now()}); await log('7.1',`Xuất ${rows.length} dòng MIGO (${W.length} chứng từ)`); }
  else if(a==='mg-ok'){ const map=new Map(migoAll().filter(m=>m.thang===S.thang).map(m=>[m.ma,m])); let n=0; const bad=[];
    for(const line of $('#mg-txt').value.split(/\r?\n/)){ if(!line.trim()) continue; const c=line.split('\t').map(x=>x.trim()); const m=map.get(c[0]); if(!m||!/^\d{6,12}$/.test(c[1]||'')){bad.push(line);continue;} await store.set('migo/'+m.id,{...stripId(m),tt:'migo',soMIGO:c[1],migoLuc:Date.now()}); n++; }
    if(n) await log('7.1',`Ghi ${n} số chứng từ MIGO`); $('#mg-txt').value=bad.join('\n'); toast(`Đã ghi ${n} chứng từ${bad.length?`, ${bad.length} dòng không khớp`:''}.`); }
  else if(a==='mg-open'){ ED.mgOpen=ED.mgOpen===t.dataset.id?null:t.dataset.id; render(); }
  else if(a==='mg-huy'){ const m=S.migo[t.dataset.id]; if(ED.mgHuy!==m.id){ED.mgHuy=m.id;toast('Bấm Hủy lần nữa để hủy '+m.ma+'.');return;} ED.mgHuy=null; await store.set('migo/'+m.id,{...stripId(m),tt:'huy'}); await log('7.2',`Hủy chứng từ MIGO ${m.ma}`); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target; if(t.dataset.bk){ ED.bk=ED.bk||{}; ED.bk[t.dataset.bk]=Math.max(0,pnum(t.value)); } });
document.addEventListener('change',async e=>{ const t=e.target;
  if(t.id==='ds-n'){ ED.dsN=t.value; ED.dsK=null; ED.bk={}; ED.sf=null; render(); }
  else if(t.id==='ds-k'){ ED.dsK=+t.value; ED.bk={}; ED.sf=null; render(); }
  else if(t.dataset.ck){ const m={...((S.ck&&S.ck.map)||{})}; m[t.dataset.ck]=+t.value; await store.set('cauhinh/chuky',{...(S.ck||{}),map:m,luc:Date.now(),ai:uid}); await log('5.3',`Đổi chu kỳ ${t.dataset.ck}: ${CKN[+t.value]}`); }
  else if(t.dataset.mgsel){ ED.mgsel=ED.mgsel||new Set(); t.checked?ED.mgsel.add(t.dataset.mgsel):ED.mgsel.delete(t.dataset.mgsel); render(); }
  else if(t.id==='sf-file'&&t.files&&t.files[0]){ const {ky}=dsSel(); const f=t.files[0];
    try{ if(!window.XLSX) throw new Error('chưa tải được thư viện Excel'); const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:false}); let best=null;
      for(const sn of wb.SheetNames){ const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:''}); const r=parseNccTable(aoa,ky); if(!r.err&&(!best||r.rows.length>best.rows.length)) best={...r,ten:f.name+' · '+sn}; }
      ED.sf={key:ED.dsN+'|'+ky.k,...(best||{err:'Không đọc được bảng nào có cột Mã hàng và Số lượng trong file.'})}; await log('5.2',`So file ${f.name} cho ${ED.dsN} ${kyName(ky)}`); }
    catch(err){ ED.sf={key:ED.dsN+'|'+ky.k,err:'Không mở được file: '+err.message}; } render(); }
});

