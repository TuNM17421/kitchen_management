/* ---------- NẠP GÓI DỮ LIỆU TỪ DRIVE (thực đơn + định lượng + khách) ---------- */
var PACK_TUAN1=CFG.packTuan1Id;
async function bulkSet(col,entries){ // entries: [[id,doc]] – dòng mới gom 1 lần append, dòng có sẵn ghi từng dòng
  const neu=[];
  for(const [id,doc] of entries){ const path=col+'/'+id, d=JSON.parse(JSON.stringify(doc)); MEM.set(path,d); if(ROW.has(path)||!SH){ if(SH) await enqueue(()=>shWrite(path,d)); } else neu.push([path,id,d]); }
  notify();
  if(SH&&neu.length) await enqueue(async()=>{ const rows=neu.map(([p,id,d])=>rowVals(id,d)), np=rows.map(r=>r.length-2), w=Math.max(...rows.map(r=>r.length)); rows.forEach(r=>{while(r.length<w)r.push('');});
    const s=await appendRows(col,rows,`Thêm ${rows.length} bản ghi vào sheet ${col}`); if(s) neu.forEach(([p],i)=>ROW.set(p,{row:s+i,n:np[i]})); });
}
async function napGoi(ref){
  const DEMO=window.__PACK_DEMO&&(!mcp||/demo/i.test(String(ref))); // bản demo (build.py --demo) có gói test nhúng sẵn
  const id=(String(ref).match(/[-\w]{25,}/)||[])[0]; if(!id&&!DEMO) throw new Error('ID file không hợp lệ');
  let P; if(DEMO) P=JSON.parse(JSON.stringify(window.__PACK_DEMO)); else { const r=pay(await call('Google Drive','download_file_content',{fileId:id},'Tải gói dữ liệu từ Drive'));
  P=JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(r.content||''),c=>c.charCodeAt(0)))); }
  const n={td:0,dm:0,kh:0,bs:0};
  if(P.giabs&&P.giabs.items){ const old=MEM.get('cauhinh/giabs')||{items:[]}, have=new Map(old.items.map(a=>[a[0],a])); P.giabs.items.forEach(a=>have.set(a[0],a)); await store.set('cauhinh/giabs',{items:[...have.values()],ghichu:P.giabs.ghichu||'',luc:Date.now()}); n.bs=P.giabs.items.length; applyConfig(); }
  if(P.dinhmuc){ const E=Object.values(P.dinhmuc).filter(d=>{const k=slug(d.mon), cur=S.dm[k]; return !cur||cur.nguon==='AI'||cur.nguon==='Excel tuần 1';}).map(d=>[slug(d.mon),{...d,luc:Date.now(),ai:uid}]); await bulkSet('dinhmuc',E); n.dm=E.length; }
  for(const [path,doc] of Object.entries(P.docs||{})){ if(!path.startsWith('thucdon/')) continue; await store.set(path,{...doc,luc:Date.now(),ai:uid}); n.td++; }
  if(P.lapTuan){ const L=P.lapTuan, src=L.nguon;
    for(const g of L.nhom){ const S0=(P.docs||{})['thucdon/'+g+'_'+src]; if(!S0) continue;
      for(const w of L.tuan){ const off=Math.round((Date.parse(w)-Date.parse(src))/864e5), inR=d=>(!L.tu||d>=L.tu)&&(!L.den||d<=L.den);
        const moi=S0.items.map(x=>({...x,d:addD(x.d,off)})).filter(x=>inR(x.d)); if(!moi.length) continue;
        const cu=(S.td[g+'_'+w]||{}).items||[], giu=cu.filter(x=>!inR(x.d));
        await store.set('thucdon/'+g+'_'+w,{nhom:g,tuan:w,items:[...giu,...moi],luc:Date.now(),ai:uid,nguon:`${S0.nguon||'gói dữ liệu'} – lặp từ tuần ${dm(src)}`}); n.td++; } } }
  if(P.khach&&P.khach.thang){ const k='khach/'+P.khach.thang, old=MEM.get(k)||{v:2,days:{}}; await store.set(k,{...old,v:2,days:{...(old.days||{}),...P.khach.days},luc:Date.now(),ai:uid}); n.kh=Object.keys(P.khach.days).length; }
  await log('1.1',`Nạp gói “${P.ten||id}”: ${n.td} thực đơn tuần, ${n.dm} định mức, ${n.kh} ngày lượng khách, ${n.bs} mặt hàng bổ sung`);
  return n; }
document.addEventListener('click',async e=>{ const t=e.target.closest('[data-act="pack-load"],[data-act="pack-open"]'); if(!t) return;
  if(t.dataset.act==='pack-open'){ ED.packOpen=!ED.packOpen; render(); if(ED.packOpen) setTimeout(()=>{const x=$('#pack-id'); if(x){x.scrollIntoView({block:'center'});}},50); return; }
  if(ED.tdDirty){ toast('Lưu hoặc bỏ thay đổi thực đơn đang sửa trước.'); return; }
  ED.packId=$('#pack-id').value.trim(); ED.packing='Đang nạp…'; render();
  try{ const n=await napGoi(ED.packId); ED.td=null; toast(`Đã nạp: ${n.td} thực đơn tuần, ${n.dm} định mức, ${n.kh} ngày lượng khách. Sang 1.4 bấm tính lại dự trù.`); }
  catch(err){ toast('Nạp gói lỗi: '+((err&&(err.message||err.code))||err)); }
  finally{ ED.packing=''; render(); } });

window.__pack={napGoi};
async function demoStart(){ // chỉ chạy ở bản demo: nạp gói test, tính dự trù, mở bước 1.4
  try{ await napGoi('demo'); ED.td=null; render();
    const tab=[...document.querySelectorAll('button,a,[role=tab]')].find(e=>/Dự trù & ngân sách/.test(e.textContent)); if(tab){ tab.click(); await new Promise(r=>setTimeout(r,400)); }
    const c=document.querySelector('[data-act="calc"]'); if(c) c.click();
    $('#previewNote').textContent='Bản demo: đã nạp sẵn gói dữ liệu test tháng 11/2026 (dữ liệu giả) và tính dự trù. Dữ liệu chỉ nằm trên trang, không lưu, không gửi lệnh SAP.';
  }catch(e){ console.error('demo pack',e); } }
