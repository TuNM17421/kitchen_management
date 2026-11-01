/* ---------- giao diện gọn: mục "Cách làm" thu gọn, nhật ký thu gọn trên màn hình nhỏ ---------- */
document.addEventListener('click',e=>{ document.querySelectorAll('details.menu[open]').forEach(d=>{ if(!d.contains(e.target)) d.open=false; }); });
document.addEventListener('input',e=>{ if(e.target.id==='dm-find'){ ED.dmFind=e.target.value; const q=fold(e.target.value); document.querySelectorAll('.dlist .ditem').forEach(b=>{b.hidden=!!q&&!fold(b.textContent).includes(q);}); } });
let HELP_OPEN=false, OPTS_OPEN=false, LAST_VIEW='', LAST_VIEW_SEEN=false;
document.addEventListener('toggle',e=>{ const d=e.target; if(d.matches&&d.matches('details.help')) HELP_OPEN=d.open; if(d.matches&&d.matches('details.opts')) OPTS_OPEN=d.open; },true);
document.addEventListener('click',e=>{ const t=e.target.closest('[data-act]'); if(!t) return; const a=t.dataset.act;
  if(a==='ncc-all'){ ED.nccAll=!ED.nccAll; render(); }
  else if(a==='f-all'){ ED.filt.all=t.dataset.v==='1'; render(); }
  else if(a==='go-next'&&t.dataset.step&&t.dataset.tab){ e.stopPropagation(); S.step=+t.dataset.step; S.tab=t.dataset.tab; render(); window.scrollTo({top:Math.max(0,($('#trackSec').getBoundingClientRect().top+scrollY)-12),behavior:'smooth'}); }
  else if(a==='go-next'){ e.stopPropagation(); const el=t.dataset.tab?document.querySelector(`#tabs [data-tab="${t.dataset.tab}"]`):document.querySelector(`.node[data-step="${t.dataset.step}"]`); if(el){ el.click(); window.scrollTo({top:Math.max(0,($('#trackSec').getBoundingClientRect().top+scrollY)-12),behavior:'smooth'}); } } },true);
if(window.matchMedia&&matchMedia('(max-width:1000px)').matches) document.querySelectorAll('details.lg[data-lg="log"]').forEach(d=>d.open=false);
/* ---------- RENDER CHUNG ---------- */
function renderTrack(){
  const m=prMain(), st2=m.length?(m.every(b=>b.tt==='duyet')?['ok','PR đã duyệt']:['run','Đang làm']):['','Chưa bắt đầu'];
  $('#nodes').innerHTML=STEPS.map((s,i)=>{const sel=S.step===i+1, done=i===0&&locked();
    const pm=poOf(S.thang), st3=pm.length?(pm.every(p=>p.soPO)?['ok','Đã có số PO']:['run','Đang làm']):['','Chưa bắt đầu'];
    const dsAll=Object.values(S.ds||{}).filter(d=>d.thang===S.thang), st4=pnOf(S.thang).length?['run',pnOf(S.thang).length+' phiếu nhận']:['','Chưa có phiếu nhận'], st5=dsAll.length?['run',dsAll.filter(d=>d.tt==='khop').length+' kỳ đã khớp']:['','Chưa bắt đầu'], st6=['','Làm cuối chu kỳ'], mgs=migoAll().filter(m=>m.thang===S.thang), st7=mgs.length?['run',mgs.filter(m=>m.tt==='migo').length+'/'+mgs.length+' đã MIGO']:['','Chưa bắt đầu'];
    const dns=Object.values(S.dn||{}).filter(d=>d.thang===S.thang&&d.tt!=='huy'), st8=dns.length?['run',dns.filter(d=>d.tt==='tt').length+'/'+dns.length+' đã thanh toán']:['','Chưa bắt đầu'];
    const pill=i===0?(done?['ok','Đã chốt']:['run','Đang làm']):[null,st2,st3,st4,st5,st6,st7,st8][i];
    const stc=(done||pill[0]==='ok')?'done':pill[0]==='run'?'run':'todo';
    return `<div class="node ${stc}${sel?' cur':''}" ${i<8?`data-step="${i+1}" role="button" tabindex="0" aria-current="${sel?'step':'false'}" style="cursor:pointer" title="${esc(s[1])}"`:''}><div class="dot">${stc==='done'?'✓ ':''}${i+1}</div><div class="lbl">${s[0]}</div><span class="pill ${pill[0]}">${pill[1]}</span></div>`;}).join('');
}
const curSub=()=>S.step===8?SUB8:S.step===7?SUB7:S.step===6?SUB6:S.step===5?SUB5:S.step===4?SUB4:S.step===3?SUB3:S.step===2?SUB2:SUB;
function renderTabs(){
  const tt7=k=>{if(k!=='mg')return false;const A=migoAll().filter(m=>m.thang===S.thang);return A.length>0&&!mgCandidates().length&&A.every(m=>m.tt==='migo');};
  const tt6=k=>{const M=mocList(),td=todayVN(); if(k==='cpo'){const R=M.filter(m=>td>=m.d); return R.length>0&&R.every(m=>!cpoCompute(m).some(r=>r.ok&&r.act!=='khop'&&!(r.act==='bs'&&r.pend>=r.bs)));} if(k==='gpo'){const used=new Set(poAll().filter(p=>p.tt!=='huy').flatMap(p=>p.src||[]));const B=batchesOf(S.thang).filter(b=>b.loai==='Bổ sung'&&b.tt!=='huy');return B.length>0&&B.every(b=>b.lines.every(l=>used.has(b.id+'|'+l.ncc)));} return false;};
  const tt5=k=>{const td=todayVN(),R=dsNccs().flatMap(n=>kyList(S.thang,n).map(ky=>({ky,d:dsDoc(n,ky.k),c:kyRows(n,ky).length}))).filter(r=>r.c&&td>=r.ky.to);if(!R.length)return false;return k==='bk'?R.every(r=>r.d&&r.d.tt!=='chua'):k==='dsk'?R.every(r=>r.d&&r.d.tt==='khop'):false;};
  const tt4=k=>{const D=dhOf(S.thang);return k==='dat'?D.length>0:k==='nhan'?D.length>0&&D.every(d=>d.pn):false;};
  const tt3=k=>{const m=poOf(S.thang);return k==='pocan'?m.length&&!poPending().length:k==='posap'?m.length&&m.every(p=>p.soPO):k==='pols'?m.length&&m.every(p=>p.soPO&&!p.sua):false;};
  const tt2=k=>k==='pb'?!!S.pb:k==='xuat'?prMain().length&&prMain().every(b=>b.tt!=='nhap'):k==='ls'?prMain().length&&prMain().every(b=>b.tt==='duyet'):false;
  $('#tabs').innerHTML=curSub().map(([n,t,k])=>{const trk=(S.step===2&&k==='dc')||(S.step===3&&k==='podc')||(S.step===4&&k==='nls')||(S.step===5&&k==='sofile')||(S.step===6&&k==='kpo')||(S.step===7&&k==='mgls'), ok=S.step===7?tt7(k):S.step===6?tt6(k):S.step===5?tt5(k):S.step>7?false:S.step===4?tt4(k):S.step===3?tt3(k):S.step===2?tt2(k):(S.tt[k]&&S.tt[k].ok);return `<button class="tab" role="tab" data-tab="${k}" aria-selected="${S.tab===k}" title="${trk?'Theo dõi, không bắt buộc':ok?'Đã xác nhận':'Chưa xác nhận'}"><b>${n}</b>${t}${trk?'<span class="ck-note">theo dõi</span>':`<span class="ck ${ok?'ok':''}" aria-hidden="true">${ok?'✓':''}</span>`}</button>`;}).join('');
}
function renderSteps(){
  $('#stHead').textContent=`Trạng thái bước ${S.step} · ${S.thang.slice(5)}/${S.thang.slice(0,4)}`;
  if(S.step===2){ const m=batchesOf(S.thang).filter(b=>b.tt!=='huy'), C=dcCompute();
    const it=[['2.1 Phân bổ ngân sách',S.pb?`<span class="pill ok">✓ ${tstr(S.pb.luc).split(' ')[0]}</span>`:'<span class="pill">Chưa</span>'],
      ['2.2 Đợt PR chờ xuất',`<span class="pill ${m.some(b=>b.tt==='nhap')?'run':''}">${m.filter(b=>b.tt==='nhap').length}</span>`],
      ['2.3 Đã duyệt',`<span class="pill ${m.length&&m.every(b=>b.tt==='duyet')?'ok':''}">${m.filter(b=>b.tt==='duyet').length}/${m.length}</span>`],
      ['2.4 Cần PR bổ sung',`<span class="pill ${C.rows.some(r=>r.need>0)?'bad':''}">${C.rows.filter(r=>r.need>0).length} mã</span>`]];
    $('#steps').innerHTML=it.map(x=>`<div><span>${x[0]}</span>${x[1]}</div>`).join(''); return; }
  if(S.step===4){ const L=pnOf(S.thang), D=dhOf(S.thang);
    const it=[['4.1 Đơn hàng chưa nhận',`<span class="pill ${D.some(d=>!d.pn)?'run':''}">${D.filter(d=>!d.pn).length}</span>`],['4.2 Phiếu nhận tháng',`<span class="pill">${L.length}</span>`],['4.3 Phiếu chưa chốt khớp NCC',`<span class="pill">${L.filter(p=>!lockedDay(p.ncc,p.d)).length}</span>`]];
    $('#steps').innerHTML=it.map(x=>`<div><span>${x[0]}</span>${x[1]}</div>`).join(''); return; }
  if(S.step>=5){ const td=todayVN(), N=dsNccs(), rows=N.flatMap(n=>kyList(S.thang,n).map(k=>({n,k,d:dsDoc(n,k.k),c:kyRows(n,k).length}))).filter(r=>r.c);
    const mg=migoAll().filter(m=>m.thang===S.thang);
    const it=[['Kỳ đến hạn chưa khớp',`<span class="pill ${rows.some(r=>td>=r.k.to&&!(r.d&&r.d.tt==='khop'))?'bad':''}">${rows.filter(r=>td>=r.k.to&&!(r.d&&r.d.tt==='khop')).length}</span>`],['Kỳ đã khớp',`<span class="pill ok">${rows.filter(r=>r.d&&r.d.tt==='khop').length}/${rows.length}</span>`],
      ['PO đã sửa chờ ME22N',`<span class="pill ${poOf(S.thang).some(p=>p.sua)?'run':''}">${poOf(S.thang).filter(p=>p.sua).length}</span>`],['Chứng từ MIGO đã có số',`<span class="pill">${mg.filter(m=>m.tt==='migo').length}/${mg.length}</span>`],
      ['ĐNTT đã lập / thanh toán',`<span class="pill">${Object.values(S.dn||{}).filter(d=>d.thang===S.thang&&d.tt!=='huy').length} / ${Object.values(S.dn||{}).filter(d=>d.thang===S.thang&&d.tt==='tt').length}</span>`]];
    $('#steps').innerHTML=it.map(x=>`<div><span>${x[0]}</span>${x[1]}</div>`).join(''); return; }
  if(S.step===3){ const m=poOf(S.thang), pend=poPending(), C3=poRecvCompute();
    const it=[['3.1 NCC chờ lập PO',`<span class="pill ${pend.length?'run':''}">${pend.length}</span>`],['3.2 Đã có số PO',`<span class="pill ${m.length&&m.every(p=>p.soPO)?'ok':''}">${m.filter(p=>p.soPO).length}/${m.length}</span>`],
      ['3.3 Dòng vượt / ngoài PO',`<span class="pill ${C3.rows.some(r=>r.need>0)?'bad':''}">${C3.rows.filter(r=>r.need>0).length}</span>`],['3.4 PO sửa chờ cập nhật SAP',`<span class="pill ${m.some(p=>p.sua)?'run':''}">${m.filter(p=>p.sua).length}</span>`]];
    $('#steps').innerHTML=it.map(x=>`<div><span>${x[0]}</span>${x[1]}</div>`).join(''); return; }
  $('#steps').innerHTML=SUB.map(([n,t,k])=>{const x=S.tt[k];return `<div><span>${n} ${t}</span>${x&&x.ok?`<span class="pill ok" title="${esc(tstr(x.luc)+' · '+nm(x.ai))}">✓ ${tstr(x.luc).split(' ')[0]}</span>`:'<span class="pill">Chưa</span>'}</div>`;}).join('');
}
function renderLog(){
  const a=S.log.filter(l=>!l.thang||l.thang===S.thang);
  $('#log').innerHTML=a.length?a.map(l=>`<li><div class="h"><time>${tstr(l.ts)}</time><span class="pill">${esc(l.buoc)}</span><span>${esc(l.nguoi||nm(l.ai))}</span></div><div>${esc(l.viec)}</div></li>`).join(''):'<li class="muted">Chưa có thao tác nào trong tháng này. Mỗi lần lưu, tính, chốt sẽ được ghi lại ở đây.</li>';
}
function applyMode(){ const ag=ED.mode==='agent'; $('#agentPanel').hidden=!ag; $('#trackSec').hidden=ag; $('#stageRow').hidden=ag;
  document.querySelectorAll('[data-mode]').forEach(b=>b.setAttribute('aria-pressed',String((b.dataset.mode==='agent')===ag)));
  const n=Object.values(S.dx||{}).filter(d=>d.thang===S.thang&&d.tt==='cho').length, c=$('#agCount'); if(c){ c.hidden=!n; c.textContent=n; c.className='pill '+(n?'bad':''); } }
document.addEventListener('click',e=>{ const b=e.target.closest('[data-mode]'); if(!b) return; ED.mode=b.dataset.mode; applyMode(); renderAgents(); window.scrollTo({top:0}); });
function render(){ try{applyMode();renderAgents();}catch(e){} renderTrack();renderTabs();renderSteps();renderLog();
  const v={td:vTD,khach:vKhach,dm:vDM,tinh:vTinh,chot:vChot,pb:vPB,xuat:vXuat,ls:vLS,dc:vDC,pocan:vPOcan,posap:vPOsap,podc:vPOdc,pols:vPOls,dat:vDat,nhan:vNhan,nls:vNls,bk:vBK,sofile:vSoFile,dsk:vDSK,cpo:vCPO,gpo:vGPO,kpo:vKPO,mg:vMG,mgls:vMGls,hd:vHD,dn:vDN,dnls:vDNls}[S.tab]; const ae=document.activeElement, id=ae&&ae.id, pos=ae&&ae.selectionStart;
  const fa=document.activeElement, fid=fa&&fa.id&&$('#view').contains(fa)&&/^(INPUT|TEXTAREA)$/.test(fa.tagName)?fa.id:'', fs=fid&&fa.selectionStart!=null?[fa.selectionStart,fa.selectionEnd]:null;
  $('#view').innerHTML=v();
  if(fid){ const nf=document.getElementById(fid); if(nf){ nf.focus({preventScroll:true}); if(fs) try{nf.setSelectionRange(fs[0],fs[1]);}catch(e){} } }
  { const vw=$('#view'), sh=vw.querySelector(':scope > .sh'), k=vw.querySelector(':scope > .kpis'), cb=vw.querySelector(':scope > .ctxbar'); if(sh&&cb){ sh.after(cb); if(k) cb.after(k); } else if(sh&&k) sh.after(k); // kết quả lên ngay dưới tiêu đề
    const kb=vw.querySelector(':scope > #khoBox'); if(kb) vw.appendChild(kb); // tồn kho (ít dùng) xuống cuối
    const subs=curSub(), ix=subs.findIndex(x=>x[2]===S.tab), nx=subs[ix+1], ns=STEPS[S.step]; // nút đi tiếp
    if(nx||(S.step<8&&ns)){ const b=document.createElement('div'); b.className='nextbar';
      b.innerHTML=nx?`<button class="btn" data-act="go-next" data-tab="${nx[2]}">Tiếp: ${nx[0]} ${esc(nx[1])}</button>`:`<button class="btn" data-act="go-next" data-step="${S.step+1}">Sang bước ${S.step+1}: ${esc(ns[0])}</button>`; vw.appendChild(b);
      const tt=vw.querySelector('.toolbar [data-act="tt"]'); if(tt){ const tb=tt.closest('.toolbar'); tt.classList.add('primary'); b.prepend(tt); if(tb&&![...tb.children].some(c=>!c.classList.contains('sp'))) tb.remove(); }
      if(!nx&&S.step===1&&typeof locked==='function'&&!locked()){ const m=document.createElement('span'); m.className='meta nb-note'; m.textContent='Chưa chốt dự trù: bước 2 sẽ dùng số hiện tại.'; b.prepend(m); } } }
  $('#view').querySelectorAll('table.t').forEach(t=>{ if(!t.closest('.tbl-wrap')){ const w=document.createElement('div'); w.className='tbl-wrap'; t.before(w); w.appendChild(t); } });
  $('#view').querySelectorAll('.tbl-wrap').forEach(w=>{ if(!w.querySelector('input,select,button,a,textarea')){ w.tabIndex=0; w.setAttribute('role','region'); if(!w.hasAttribute('aria-label')) w.setAttribute('aria-label','Bảng (cuộn được)'); } });
  const vk=S.step+'|'+S.tab; if(vk!==LAST_VIEW){ LAST_VIEW=vk; {const tt=document.getElementById('toast'); if(tt&&LAST_VIEW_SEEN) tt.hidden=true;} LAST_VIEW_SEEN=true; requestAnimationFrame(()=>{ for(const el of [document.querySelector('.node.cur'),document.querySelector('.tab[aria-selected="true"]')]){ const sc=el&&el.closest('.track,.tabs'); if(sc&&sc.scrollWidth>sc.clientWidth+4) sc.scrollTo({left:el.offsetLeft-(sc.clientWidth-el.offsetWidth)/2,behavior:'smooth'}); } }); }
  if(id){const el=document.getElementById(id); if(el){el.focus(); try{if(pos!=null)el.setSelectionRange(pos,pos);}catch(e){}}}
}
function head(title,desc,key){ const x=S.tt[key];
  return `<div class="sh"><div><h2>${title}</h2>${desc?`<details class="help"${HELP_OPEN?' open':''}><summary>Cách làm</summary><p>${desc}</p></details>`:''}</div><div class="toolbar">${x&&x.ok?`<span class="pill ok">✓ Đã xác nhận ${tstr(x.luc)} · ${esc(nm(x.ai))}</span>`:'<span class="pill run">Chưa xác nhận</span>'}${locked()?'<span class="pill bad">Dự trù đã chốt · chỉ xem</span>':''}</div></div>`; }

