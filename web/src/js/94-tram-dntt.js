/* ---- Trạm SAP: tạo file ĐNTT (macro TaoDNTT, không chạm SAP) ---- */
async function lenhDNTT(ids){
  const R=dnRows().filter(r=>r.dn&&ids.includes(r.dn.id)&&(r.dn.tt==='lap'||r.dn.tt==='xuat')&&!lenhBusy(r.dn.lenh));
  if(!R.length){ toast('Chọn ĐNTT đã lập (chưa trình ký).'); return; }
  const miss=R.filter(r=>!SHEETMAU[r.po.ncc.toUpperCase()]); if(miss.length){ toast('Chưa có sheet mẫu ĐNTT cho: '+miss.map(r=>r.po.ncc).join(', ')+' (bảng NCC của DNTT_CauHinh).'); return; }
  const rows=dnSheetRows(R).map(c=>({c}));
  const lid=await guiLenh('DNTT',rows,{dn:R.map(r=>r.dn.id)},`${R.length} ĐNTT – ${R.map(r=>r.po.ncc).join(', ')}`); if(!lid) return;
  for(const r of R) await store.set('dntt/'+r.dn.id,{...stripId(r.dn),lenh:lid}); }
async function applyLenh3(L,K,note){
  const ids=(L.ref&&L.ref.dn)||[], up=(L.ketqua&&L.ketqua.po)||[];
  for(let i=0;i<ids.length;i++){ const d=S.dn[ids[i]]; if(!d) continue; const k=K.find(x=>+x.i===i)||{};
    if(k.file){ const base=String(k.file).split('\\').pop().replace(/\.xlsx$/i,''), pdf=up.find(u=>u.name===base+'.pdf'), xl=up.find(u=>u.name===base+'.xlsx');
      await store.set('dntt/'+d.id,{...stripId(d),lenh:'',tt:d.tt==='lap'?'xuat':d.tt,xuatLuc:Date.now(),fileDNTT:k.file,linkPdf:pdf?pdf.url:'',linkXlsx:xl?xl.url:'',thuHuong:k.thuHuong||'',hist:[...(d.hist||[]),{ts:Date.now(),tt:'xuat',ai:uid,note:'Trạm tạo file '+base}]});
      note.push(`${d.ncc}: đã tạo file`); }
    else { await store.set('dntt/'+d.id,{...stripId(d),lenh:''}); note.push(`${d.ncc}: chưa tạo được (${k.tt||'xem nhật ký'})`); } } }
const dnLinks=d=>(d.linkPdf||d.linkXlsx)?` <span class="meta">${d.linkPdf?`<a href="${esc(d.linkPdf)}" target="_blank" rel="noopener">PDF</a>`:''}${d.linkPdf&&d.linkXlsx?' · ':''}${d.linkXlsx?`<a href="${esc(d.linkXlsx)}" target="_blank" rel="noopener">Excel</a>`:''}</span>`:'';

window.__tram={pollTram:()=>pollTram(),S:()=>S,ED:()=>ED};

