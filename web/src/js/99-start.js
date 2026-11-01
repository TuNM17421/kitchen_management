(async()=>{
  const C=window.claude;
  if(C&&C.use){ [user,sample,dlcap,mcp]=await Promise.all(['user','sample','downloads','mcp'].map(n=>C.use(n).catch(()=>null))); gmail=mcp; }
  if(user){ try{ uid=(await user.id())||''; const me=await user.me(); uname=me&&me.name||''; $('#who').textContent=uname; $('#who').hidden=!uname; }catch(e){} } else $('#who').hidden=true;
  subscribeAll(); renderCmd();
  if(mcp){ try{ SH=true; await loadAll(); pollTram(); }catch(e){ SH=false; $('#conn').textContent='Không đọc được file dữ liệu'; $('#conn').className='pill bad'; $('#previewNote').hidden=false; toast('Không đọc được file dữ liệu: '+((e&&(e.message||e.code))||e)); } }
  else { $('#conn').textContent='Chạy thử (không lưu)'; $('#conn').classList.add('run'); $('#previewNote').hidden=false; $('#reload').hidden=true; $('#openSheet').hidden=true; if(window.__PACK_DEMO) setTimeout(demoStart,300); }
  render();
})();
