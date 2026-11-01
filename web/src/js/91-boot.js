/* ---------- KHỞI ĐỘNG ---------- */
$('#reload').addEventListener('click',async()=>{ if(!SH){toast('Chưa kết nối file dữ liệu.');return;} if(ED.tdDirty||ED.khDirty||ED.lnDirty||ED.dmDirty){toast('Đang có thay đổi chưa lưu.');return;} try{ await WQ; await loadAll(); render(); }catch(e){} });
$('#seedBtn').addEventListener('click',async()=>{ const b=$('#seedBtn'); b.disabled=true; b.textContent='Đang nạp…'; try{ await seedAll(); render(); toast('Đã nạp dữ liệu khởi tạo vào file.'); }catch(e){ toast('Nạp lỗi: '+((e&&(e.message||e.code))||e)); b.disabled=false; b.textContent='Nạp dữ liệu khởi tạo'; } });
$('#dl-item').innerHTML=ITEMS.map(x=>`<option value="${x.ma}">${esc(x.ten)} (${esc(x.dvt)})</option>`).join('');
document.addEventListener('keydown',e=>{ if((e.key==='Enter'||e.key===' ')&&e.target.dataset&&e.target.dataset.step){e.preventDefault();e.target.click();} });
$('#thang').innerHTML=['2026-10','2026-11','2026-12','2027-01'].map(t=>`<option value="${t}">${t.slice(5)}/${t.slice(0,4)}</option>`).join('');
$('#thang').value=S.thang;
render();
