/* ---------- SỰ KIỆN ---------- */
document.addEventListener('click',async e=>{
  const st=e.target.closest('[data-step]'); if(st){ const n=+st.dataset.step; if(n!==S.step){S.step=n;S.tab=curSub()[0][2];render();} return; }
  const t=e.target.closest('[data-tab],[data-act]'); if(!t) return;
  if(t.dataset.tab){S.tab=t.dataset.tab;render();return;}
  const a=t.dataset.act;
  try{
  if(a==='td-nhom'){ if(ED.tdDirty&&!t.dataset.force){toast('Lưu tuần đang sửa trước khi đổi nhóm.');return;} ED.nhom=t.dataset.k;ED.week=null;ED.td=null;render(); }
  else if(a==='td-add'){ const r=tdRows(); const last=r[r.length-1]; r.push({d:last?last.d:ED.week,b:last?last.b:NHOM[ED.nhom].bua[0],loai:'',mon:''}); ED.tdDirty=true; render(); const i=r.length-1; const el=document.getElementById('td-m-'+i); el&&el.focus(); }
  else if(a==='td-to'){ if(ED.tdDirty){toast('Lưu tuần đang sửa trước khi chép.');return;} const g=$('#td-to').value, src=tdRows(), from=ED.nhom;
    const cp=src.filter(r=>NHOM[g].bua.includes(r.b)).map(r=>({...r})); ED.nhom=g; ED.td=null; const rows=tdRows(); rows.push(...cp); ED.tdDirty=true; ED.copied=null; render();
    toast(`Đã chép ${cp.length} món từ ${NHOM[from].ten} sang ${NHOM[g].ten}. Kiểm tra rồi bấm Lưu tuần.`); }
  else if(a==='td-del'){ tdRows().splice(+t.dataset.i,1); ED.tdDirty=true; render(); }
  else if(a==='td-paste'){ ED.paste=!ED.paste; render(); }
  else if(a==='td-paste-ok'){ const {out,bad}=parsePaste($('#td-paste-txt').value); tdRows().push(...out); if(out.length) ED.tdDirty=true; ED.paste=bad.length>0; render(); if(ED.paste) $('#td-paste-txt').value=bad.join('\n'); toast(`Đã thêm ${out.length} món${bad.length?`, ${bad.length} dòng không đọc được (còn lại trong ô dán)`:''}.`); }
  else if(a==='td-copy'){ const prevW=addD(ED.week,-7), src=S.td[ED.nhom+'_'+prevW]; if(!src){toast(`Tuần ${dm(prevW)} chưa có thực đơn để chép.`);return;}
    const rows=tdRows(); rows.push(...src.items.map(x=>({...x,d:addD(x.d,7),monGoc:x.mon}))); ED.tdDirty=true; ED.copied=prevW; render(); toast(`Đã chép ${src.items.length} món từ tuần ${dm(prevW)}. Sửa rồi bấm Lưu tuần.`); }
  else if(a==='td-save'){ const rows=tdRows().filter(r=>r.mon&&r.mon.trim()).map(r=>{ const o={d:r.d,b:r.b,loai:(r.loai||'').trim(),mon:r.mon.trim()}; const tl=String(r.tl??'').trim(); if(tl!==''&&pnum(tl)>=0) o.tl=pnum(tl); if(r.nl&&r.nl.length&&r.mon===r.monGoc) o.nl=r.nl; return o; });
    const k=tdKey(), old=S.td[k]; await store.set('thucdon/'+k,{nhom:ED.nhom,tuan:ED.week,items:rows,luc:Date.now(),ai:uid,nguon:ED.copied?`chép từ tuần ${dm(ED.copied)} rồi sửa`:(old&&old.nguon)||''});
    ED.tdDirty=false;ED.copied=null; await log('1.1',`Lưu thực đơn ${NHOM[ED.nhom].ten.toLowerCase()} tuần ${dm(ED.week)}–${dm(addD(ED.week,6))}: ${rows.length} món${old?` (trước ${(old.items||[]).length})`:''}`); toast('Đã lưu thực đơn tuần.'); }
  else if(a==='kh-fill'){ const K=ED.khach; monthDays(S.thang).forEach(d=>{const w=P(d).getUTCDay();K[d]=[NV_TB[w][0],NV_TB[w][1],...NT_TB.TH,...NT_TB.VIP,...NT_TB.AN];}); ED.khDirty=true; render(); toast('Đã điền theo trung bình tháng 9. Sửa các ngày đặc biệt rồi bấm Lưu.'); }
  else if(a==='kh-save'){ const days={}; Object.entries(ED.khach).forEach(([d,r])=>days[d]=Array.from({length:NK},(_,i)=>Math.max(0,Math.round(+r[i]||0)))); const tot=Object.values(days).flat().reduce((x,y)=>x+y,0);
    await store.set('khach/'+S.thang,{v:2,days,luc:Date.now(),ai:uid}); ED.khDirty=false; await log('1.2',`Lưu dự báo lượng khách: ${VND(tot)} suất cả tháng`); toast('Đã lưu lượng khách.'); }
  else if(a==='dm-pick'){ if(ED.dmDirty){toast('Lưu định mức món đang sửa trước.');return;} ED.dmMon=t.dataset.k; ED.dmRows=null; render(); }
  else if(a==='dm-add'){ ED.dmRows.rows.push({ma:'',q:0}); ED.dmDirty=true; render(); const el=document.getElementById('dm-ma-'+(ED.dmRows.rows.length-1)); el&&el.focus(); }
  else if(a==='dm-del'){ ED.dmRows.rows.splice(+t.dataset.i,1); ED.dmDirty=true; render(); }
  else if(a==='dm-save'){ const k=ED.dmMon, rec=S.dm[k], mon=(dishesOfMonth().find(d=>d.k===k)||{}).mon||(rec&&rec.mon);
    const nl=ED.dmRows.rows.filter(r=>IT.has(r.ma)&&r.q>0).map(r=>({ma:r.ma,q:+(+r.q).toFixed(4)}));
    if(!nl.length){toast('Cần ít nhất 1 nguyên liệu có mã đúng và định lượng > 0.');return;}
    await store.set('dinhmuc/'+k,{mon,nl,nguon:'Bếp',ghichu:rec&&rec.ghichu||'',luc:Date.now(),ai:uid}); ED.dmDirty=false;
    await log('1.3',`${rec&&rec.nguon==='AI'?'Duyệt':'Lưu'} định mức “${mon}”: ${nl.length} nguyên liệu`); toast('Đã lưu định mức.'); }
  else if(a==='dm-clear'){ const rec=S.dm[ED.dmMon]; if(!confirm(`Xóa toàn bộ định mức của “${rec&&rec.mon}”? Món này sẽ tính theo lịch sử cho tới khi nhập lại.`)) return; await store.del('dinhmuc/'+ED.dmMon); ED.dmRows=null;ED.dmDirty=false; await log('1.3',`Xóa định mức “${rec&&rec.mon}”`); }
  else if(a==='dm-paste'){ ED.dmPaste=!ED.dmPaste; render(); }
  else if(a==='dm-paste-ok'){ const by=new Map(), bad=[];
    for(const line of $('#dm-paste-txt').value.split(/\r?\n/)){ if(!line.trim()) continue; const c=line.split(/[\t⇥]/).map(x=>x.trim()); const x=IT.get((c[1]||'').toUpperCase()); const v=pnum(c[2]);
      if(c.length<3||!c[0]||!x||!(v>0)){bad.push(line);continue;} const q=/^(KG|L)$/.test(x.dvt)?v/1000:v; if(!by.has(c[0])) by.set(c[0],[]); by.get(c[0]).push({ma:x.ma,q:+q.toFixed(4)}); }
    for(const [mon,nl] of by) await store.set('dinhmuc/'+slug(mon),{mon,nl,nguon:'Bếp',ghichu:'nạp từ Excel',luc:Date.now(),ai:uid});
    if(by.size) await log('1.3',`Nạp định mức từ Excel: ${by.size} món`); ED.dmPaste=bad.length>0; render(); if(ED.dmPaste) $('#dm-paste-txt').value=bad.join('\n');
    toast(`Đã nạp ${by.size} món${bad.length?`, ${bad.length} dòng sai mã hoặc định lượng (còn lại trong ô dán)`:''}.`); }
  else if(a==='kho-read'){ await readKho(); }
  else if(a==='map-save'){ const map={...((S.khomap&&S.khomap.map)||{}),...(ED.map||{})}; await store.set('kho/map',{map,luc:Date.now(),ai:uid}); await log('1.4',`Lưu ghép mã kho: ${Object.keys(ED.map||{}).length} dòng sửa`); ED.map=null; ED.mapDirty=false; toast('Đã lưu ghép mã.'); }
  else if(a==='dm-ai'){ await runAI(); }
  else if(a==='calc'){ if(ED.lnDirty){toast('Lưu số chốt đang sửa trước khi tính lại.');return;}
    const dpT=pnum($('#dpT').value),dpK=pnum($('#dpK').value),wh=$('#wHist').checked; const r=compute(dpT,dpK,wh,!!S.kho&&$('#uKho').checked,$('#wNgoai').checked);
    let tc=0; r.lines.forEach(l=>{const x=IT.get(l.ma);tc+=priceOf(x,l.ncc)*l.chot;});
    await store.set('dutru/'+S.thang,{lines:r.lines,warn:r.warn,params:r.params,luc:Date.now(),ai:uid,tong:Math.round(tc)}); ED.lines=null;
    await log('1.4',`Tính dự trù: ${r.lines.length} mã, ngân sách ${TR(tc)} (dự phòng tươi ${dpT}%, khô ${dpK}%; ${VND(r.warn.tot-r.warn.unc)}/${VND(r.warn.tot)} suất theo thực đơn)`); toast('Đã tính dự trù.'); }
  else if(a==='ln-save'){ const L=curLines(); let tc=0,n=0; L.forEach(l=>{const x=IT.get(l.ma);if(x)tc+=priceOf(x,l.ncc)*l.chot;if(l.e)n++;});
    await store.set('dutru/'+S.thang,{...S.dutru,lines:L,tong:Math.round(tc)}); ED.lnDirty=false; await log('1.4',`Lưu số chốt: ${n} dòng sửa tay, ngân sách ${TR(tc)}`); toast('Đã lưu số chốt.'); }
  else if(a==='f-ncc'){ ED.filt.ncc=t.dataset.n; render(); }
  else if(a==='tt'){ const k=t.dataset.k; if((k==='td'&&ED.tdDirty)||(k==='khach'&&ED.khDirty)||(k==='dm'&&ED.dmDirty)||(k==='tinh'&&ED.lnDirty)){toast('Lưu thay đổi trước khi xác nhận.');return;}
    await setTT(k,true); await log(SUB.find(s=>s[2]===k)[0],`Xác nhận ${SUB.find(s=>s[2]===k)[1].toLowerCase()}`); }
  else if(a==='lock'){ const L=curLines(); let tc=0; L.forEach(l=>{const x=IT.get(l.ma);if(x)tc+=priceOf(x,l.ncc)*l.chot;}); await setTT('tinh',true); await setTT('chot',true,{tong:Math.round(tc)}); await log('1.5',`Chốt dự trù tháng: ${L.length} mã, ${TR(tc)}`); toast('Đã chốt dự trù.'); }
  else if(a==='unlock'){ await setTT('chot',false); await log('1.5','Mở lại dự trù để sửa'); }
  else if(a==='export'){ const data=csv(), fn=`DuTru_${S.thang}.csv`;
    if(mcp||dlcap){ try{ await saveOut(fn,new Blob([data],{type:'text/csv'})); await log('1.5',`Xuất ${fn}`);}catch(err){ if(err&&err.code!=='cancelled') toast('Không lưu được file: '+(err.message||err.code)); } }
    else { try{ await navigator.clipboard.writeText(data); toast('Đã copy nội dung CSV, dán vào Excel.'); }catch(err){ toast('Trang này không tải file được.'); } } }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{
  const t=e.target;
  if(t.dataset.td!=null){ const r=tdRows()[+t.dataset.td]; r[t.dataset.f]=t.value; if(!ED.tdDirty){ED.tdDirty=true;render();} return; }
  if(t.dataset.kh){ const K=ED.khach; K[t.dataset.kh]=K[t.dataset.kh]||Z(); K[t.dataset.kh][+t.dataset.i]=Math.max(0,Math.round(pnum(t.value))); ED.khDirty=true; clearTimeout(vKhach.h); vKhach.h=setTimeout(render,500); return; }
  if(t.dataset.dm!=null){ const r=ED.dmRows.rows[+t.dataset.dm]; if(t.dataset.f==='ma'){r.ma=t.value.trim().toUpperCase();} else { const x=IT.get(r.ma); const v=pnum(t.value); r.q=x&&/^(KG|L)$/.test(x.dvt)?v/1000:v; }
    ED.dmDirty=true; clearTimeout(vDM.h); vDM.h=setTimeout(render,600); return; }
  if(t.dataset.ln!=null){ const l=curLines()[+t.dataset.ln]; if(t.dataset.f==='chot'){l.chot=Math.max(0,pnumVN(t.value));l.e=true;} else if(t.dataset.f==='note'){l.note=t.value;} ED.lnDirty=true; clearTimeout(vTinh.h); vTinh.h=setTimeout(render,700); return; }
  if(t.dataset.map){ ED.map=ED.map||{}; ED.map[t.dataset.map]=t.value.trim().toUpperCase(); if(!ED.mapDirty){ED.mapDirty=true;render();} return; }
  if(t.id==='f-q'){ ED.filt.q=t.value; clearTimeout(vTinh.f); vTinh.f=setTimeout(render,250); }
});
document.addEventListener('change',e=>{
  const t=e.target;
  if(t.id==='thang'){ if(ED.khDirty||ED.lnDirty||ED.tdDirty||ED.dmDirty||(ED.pb&&ED.pb.dirty)){toast('Đang có thay đổi chưa lưu.');t.value=S.thang;return;} S.thang=t.value; ED.week=null;ED.td=null; subscribeMonth(); render(); }
  else if(t.id==='td-week'){ if(ED.tdDirty){toast('Lưu tuần đang sửa trước khi đổi tuần.');t.value=ED.week;return;} ED.week=t.value;ED.td=null;render(); }
  else if(t.dataset.td!=null&&t.tagName==='SELECT'){ tdRows()[+t.dataset.td][t.dataset.f]=t.value; ED.tdDirty=true; render(); }
  else if(t.dataset.ln!=null&&t.dataset.f==='ncc'){ const l=curLines()[+t.dataset.ln]; l.ncc=t.value; l.e=true; ED.lnDirty=true; render(); }
  else if(t.id==='f-ncc'){ ED.filt.ncc=t.value; render(); } else if(t.id==='f-n2'){ ED.filt.n2=t.value; render(); }
  else if(t.dataset.dm!=null||t.dataset.ln!=null||t.dataset.kh) render();
});

document.addEventListener('toggle',e=>{ if(e.target&&e.target.id==='khoBox') ED.khoOpen=e.target.open; },true);
