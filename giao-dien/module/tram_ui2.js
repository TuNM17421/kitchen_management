/* ---- Trạm SAP: sửa PO (ME22N) và MIGO cuối kỳ ---- */
function poSapDiff(p){ // so PO hiện tại với ảnh chụp trên SAP (sapLines)
  const sap=new Map((p.sapLines||[]).map(l=>[+l.ebelp,l])), out=[], seen=new Set();
  p.lines.forEach(l=>{ if(l.ebelp){ const s=sap.get(+l.ebelp); seen.add(+l.ebelp); if(!s) return;
      if(Math.abs((+s.sl)-(+l.sl))>1e-6||Math.abs((+s.g)-(+l.g))>0.5) out.push({loai:'SUA',ebelp:+l.ebelp,l,cu:s}); }
    else out.push({loai:'THEM',l}); });
  sap.forEach((s,eb)=>{ if(!seen.has(eb)) out.push({loai:'SUA',ebelp:eb,l:{...s,sl:0},cu:s}); });
  return out; }
async function lenhPOSua(id){
  const p=S.po[id]; if(!p||!p.soPO){ toast('PO chưa có số trên SAP.'); return; }
  if(!p.sapLines){ toast('PO này chưa có ảnh chụp dòng trên SAP (tạo trước khi có trạm). Sửa tay ME22N rồi bấm “Đã cập nhật SAP”.'); return; }
  const D=poSapDiff(p); if(!D.length){ await setPO(id,{sua:false},'Không còn khác biệt với SAP'); toast('PO trên giao diện đã khớp SAP.'); return; }
  const miss=D.filter(d=>d.loai==='THEM'&&(!d.l.soPR||!prItemOf(d.l.soPR,d.l.ma,p.ncc)));
  if(miss.length){ toast('Dòng thêm chưa có số PR / dòng PR: '+miss.map(d=>d.l.ma).join(', ')+'. PR bổ sung phải được tạo qua trạm (2.2) và duyệt.'); return; }
  const rows=D.map(d=>({c:[p.soPO,d.loai,d.loai==='SUA'?String(d.ebelp):'',d.loai==='THEM'?d.l.soPR:'',d.loai==='THEM'?prItemOf(d.l.soPR,d.l.ma,p.ncc):'',d.l.ma,+d.l.sl,+d.l.g,d.l.ten||''],ref:d.loai==='SUA'?{ebelp:d.ebelp}:{ma:d.l.ma,soPR:d.l.soPR}}));
  const lid=await guiLenh('POSUA',rows,{po:id},`Sửa PO ${p.soPO} – ${p.ncc} (${D.filter(d=>d.loai==='SUA').length} sửa, ${D.filter(d=>d.loai==='THEM').length} thêm)`); if(!lid) return;
  await setPO(id,{lenh:lid},'Gửi trạm SAP sửa PO (ME22N)'); }
const migoDate=d=>d.split('-').reverse().join('.');
async function lenhMIGO(ids){
  let n=0;
  for(const id of ids){ const m=S.migo[id]; if(!m||!(m.tt==='cho'||m.tt==='xuat')||lenhBusy(m.lenh)) continue; const po=S.po[m.poId];
    if(po&&(po.sua||lenhBusy(po.lenh))){ toast(`PO ${m.soPO} (${m.ncc}) còn sửa chưa cập nhật SAP – bấm “Cập nhật trên SAP” ở 3.4 trước.`); continue; }
    const rows=m.lines.map(l=>({c:[m.ma,m.soPO,migoDate(m.ngay),l.item,l.ma,+l.sl,`Mua hàng phục vụ Bếp tháng ${+m.thang.slice(5)}`]}));
    const lid=await guiLenh('MIGO',rows,{migo:id},`${m.ma} – PO ${m.soPO} (${rows.length} dòng)`); if(!lid) continue;
    await store.set('migo/'+id,{...stripId(m),lenh:lid,tt:'xuat',xuatLuc:m.xuatLuc||Date.now()}); n++; await new Promise(r=>setTimeout(r,5)); }
  if(n) toast(`Đã gửi ${n} lệnh MIGO cho trạm SAP.`); }
async function applyLenh2(L,K,note){
  if(L.loai==='POSUA'){ const pid=L.ref&&L.ref.po, p=pid&&S.po[pid]; if(!p) return;
    if(L.tt==='xong'&&L.luuThat){ const lines=p.lines.map(l=>({...l}));
      (L.rows||[]).forEach((r,i)=>{ if(r.ref&&r.ref.ma){ const k=K.find(x=>+x.i===i)||{}; const l=lines.find(z=>!z.ebelp&&z.ma===r.ref.ma&&z.soPR===r.ref.soPR); if(l&&k.ebelp) l.ebelp=+k.ebelp; } });
      await setPO(pid,{lines,sapLines:lines.filter(l=>l.ebelp).map(l=>({ebelp:+l.ebelp,ma:l.ma,sl:+l.sl,g:+l.g})),sua:lines.some(l=>!l.ebelp),lenh:'',sapSuaLuc:Date.now()},'Trạm SAP đã sửa PO (ME22N)'); note.push(`PO ${p.soPO}: đã cập nhật SAP`); }
    else { await setPO(pid,{lenh:''},(L.luuThat?'Trạm SAP chưa sửa được PO: ':'Chạy thử ME22N (không lưu): ')+((K.find(k=>k.tt)||{}).tt||'xem nhật ký lệnh')); note.push(`PO ${p.soPO}: ${L.luuThat?'chưa sửa được':'chạy thử xong, chưa lưu'}`); } }
  else if(L.loai==='MIGO'){ const mid=L.ref&&L.ref.migo, m=mid&&S.migo[mid]; if(!m) return; const so=[...new Set(K.map(k=>k.soMIGO).filter(Boolean))];
    if(L.tt==='xong'&&L.luuThat&&so.length===1){ await store.set('migo/'+mid,{...stripId(m),tt:'migo',soMIGO:so[0],migoLuc:Date.now(),lenh:''}); note.push(`${m.ma} → MIGO ${so[0]}`); }
    else { await store.set('migo/'+mid,{...stripId(m),lenh:'',ghichu:(K.find(k=>k.tt)||{}).tt||''}); note.push(`${m.ma}: ${L.luuThat?'chưa nhập kho được':'chạy thử xong, chưa Post'}`); } } }
