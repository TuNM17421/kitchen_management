/* ================= BƯỚC 4 · ĐẶT HÀNG, NHẬN HÀNG & MIGO TỪNG PHẦN ================= */
const _x=0;
const SUB4=[['4.1','Đặt hàng theo PO','dat'],['4.2','Nhận hàng','nhan'],['4.3','Lịch sử nhận hàng','nls']];
const MACRO_BAS="Attribute VB_Name = \"MIGO_TungPhan\"\r\nOption Explicit\r\n\r\n' ================================================================================\r\n'  MIGO_TungPhan v1 (30/10/2026) - NHAP KHO TUNG PHAN THEO THUC NHAN\r\n'  Dung CUNG file voi module SAP_Automation v22 (dung lai cac ham: GetSapSession,\r\n'  FindCtl, FindCtlSoft, MigoRowCell, TblColIndex, CheckPopup, CheckSbarError,\r\n'  ExtractDocNumber, SapNumber, ParseSapNum, WaitSap).\r\n'\r\n'  Sheet \"MIGO_TP\" (macro tu tao neu chua co), du lieu tu dong 3:\r\n'   A Ma_phieu | B So_PO | C Ngay_hach_toan (dd.mm.yyyy) | D Dong_PO (10,20..)\r\n'   E Ma_vat_tu | F So_luong | G Text | H So_chung_tu_MIGO (macro ghi) | I Status (macro ghi)\r\n'  Dan du lieu tu trang \"Quy trinh PR truoc\" (buoc 4.3) vao A3.\r\n'\r\n'  Moi Ma_phieu = 1 chung tu MIGO: Goods Receipt / Purchase Order / 101,\r\n'  ngay hach toan = ngay giao, chi tich OK cac dong PO co trong phieu,\r\n'  so luong = so thuc nhan. Dong PO khong co trong phieu se bo tich.\r\n'\r\n'  CHE DO KHONG LUU: MIGO_CHO_PHEP_POST = False -> macro dien man hinh, bam\r\n'  Check, ghi ket qua, KHONG Post. Doi thanh True khi duoc phep ap dung chinh thuc.\r\n' ================================================================================\r\n\r\nPublic Const MIGO_CHO_PHEP_POST As Boolean = False\r\nPublic Const MIGO_DATE_FMT As String = \"dd.mm.yyyy\"   ' theo SU3 > Defaults > Date format\r\nPrivate Const TP_TBL As String = \"tblSAPLMIGOTV_GOITEM\"\r\nPrivate tpStep As String\r\n\r\nSub EnsureMigoTP()\r\n    Dim ws As Worksheet\r\n    On Error Resume Next\r\n    Set ws = ThisWorkbook.Sheets(\"MIGO_TP\")\r\n    On Error GoTo 0\r\n    If ws Is Nothing Then\r\n        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))\r\n        ws.Name = \"MIGO_TP\"\r\n        ws.Range(\"A1:I1\").Value = Array(\"Ma_phieu\", \"So_PO\", \"Ngay_hach_toan\", \"Dong_PO\", \"Ma_vat_tu\", \"So_luong\", \"Text\", \"So_chung_tu_MIGO\", \"Status\")\r\n        ws.Range(\"A2:I2\").Value = Array(\"Dan tu trang web\", \"\", \"dd.mm.yyyy\", \"10,20..\", \"\", \"so thuc nhan\", \"\", \"<< macro ghi >>\", \"<< macro ghi >>\")\r\n        ws.Range(\"A1:I1\").Font.Bold = True\r\n        ws.Range(\"A1:I1\").Interior.Color = RGB(221, 235, 247)\r\n        ws.Range(\"A2:G2\").Interior.Color = RGB(255, 242, 204)\r\n        ws.Columns(\"A:I\").ColumnWidth = 16\r\n        ws.Columns(\"C\").NumberFormat = \"@\"\r\n    End If\r\nEnd Sub\r\n\r\nSub PostMIGO_TungPhan()\r\n    EnsureMigoTP\r\n    Dim ws As Worksheet, lastR As Long, r As Long, k As String\r\n    Set ws = ThisWorkbook.Sheets(\"MIGO_TP\")\r\n    lastR = ws.Cells(ws.Rows.Count, \"A\").End(xlUp).Row\r\n\r\n    Dim phieu As Object: Set phieu = CreateObject(\"Scripting.Dictionary\")\r\n    For r = 3 To lastR\r\n        k = Trim(CStr(ws.Cells(r, \"A\").Value))\r\n        If k <> \"\" And Trim(CStr(ws.Cells(r, \"H\").Value)) = \"\" Then\r\n            If Not phieu.Exists(k) Then phieu.Add k, New Collection\r\n            phieu(k).Add r\r\n        End If\r\n    Next r\r\n    If phieu.Count = 0 Then\r\n        MsgBox \"Khong co phieu nao can MIGO (cot H trong).\", vbInformation, \"MIGO tung phan v1\"\r\n        Exit Sub\r\n    End If\r\n\r\n    Dim sm As String, v As Variant\r\n    For Each v In phieu.Keys\r\n        sm = sm & \"  - \" & v & \": PO \" & ws.Cells(phieu(v)(1), \"B\").Value & \", ngay \" & ws.Cells(phieu(v)(1), \"C\").Value & \", \" & phieu(v).Count & \" dong\" & vbCrLf\r\n    Next v\r\n    If MsgBox(\"Se nhap kho tung phan \" & phieu.Count & \" phieu:\" & vbCrLf & sm & vbCrLf & _\r\n              IIf(MIGO_CHO_PHEP_POST, \"Macro se HOI truoc khi Post tung phieu.\", \"CHE DO KHONG LUU: chi dien man hinh + Check, KHONG Post.\") & vbCrLf & \"Bat dau?\", _\r\n              vbYesNo + vbQuestion, \"MIGO tung phan v1\") = vbNo Then Exit Sub\r\n\r\n    Dim session As Object: Set session = GetSapSession()\r\n    If session Is Nothing Then Exit Sub\r\n\r\n    Dim rows As Collection, msg As String, docNo As String, ans As VbMsgBoxResult, nOK As Long\r\n    For Each v In phieu.Keys\r\n        Set rows = phieu(v)\r\n        On Error GoTo LoiPhieu\r\n        nOK = DienMIGOTungPhan(session, ws, rows)\r\n        If MIGO_CHO_PHEP_POST Then\r\n            ans = MsgBox(\"Phieu \" & v & \": da dien \" & nOK & \"/\" & rows.Count & \" dong.\" & vbCrLf & \"Kiem tra tren SAP roi tra loi: POST nhap kho?\", vbYesNo + vbQuestion + vbDefaultButton2, \"MIGO tung phan v1\")\r\n            If ans = vbNo Then GhiStatus ws, rows, \"\", \"CHUA POST - dung theo yeu cau\": Exit Sub\r\n            tpStep = \"bam Post\"\r\n            session.findById(\"wnd[0]/tbar[0]/btn[11]\").press\r\n            CheckPopup session\r\n            msg = session.findById(\"wnd[0]/sbar\").Text\r\n            CheckSbarError session\r\n            docNo = ExtractDocNumber(msg)\r\n            GhiStatus ws, rows, docNo, msg\r\n        Else\r\n            tpStep = \"bam Check\"\r\n            On Error Resume Next\r\n            session.findById(\"wnd[0]/tbar[1]/btn[7]\").press   ' Check (neu co)\r\n            On Error GoTo LoiPhieu\r\n            WaitSap session\r\n            msg = session.findById(\"wnd[0]/sbar\").Text\r\n            GhiStatus ws, rows, \"\", \"DA DIEN + CHECK, CHUA POST (che do khong luu) \" & Format(Now, \"dd/mm hh:nn\") & IIf(msg <> \"\", \" | SAP: \" & msg, \"\")\r\n            If MsgBox(\"Phieu \" & v & \": da dien \" & nOK & \"/\" & rows.Count & \" dong, chua Post.\" & vbCrLf & \"Kiem tra man hinh SAP. Chuyen sang phieu tiep theo?\", vbYesNo + vbQuestion, \"MIGO tung phan v1\") = vbNo Then Exit Sub\r\n        End If\r\n        GoTo TiepTheo\r\nLoiPhieu:\r\n        GhiStatus ws, rows, \"\", \"LOI buoc [\" & tpStep & \"]: \" & Err.Description\r\n        MsgBox \"Loi phieu \" & v & \":\" & vbCrLf & Err.Description & vbCrLf & \"Buoc: \" & tpStep & vbCrLf & _\r\n               \"Man hinh SAP giu nguyen. Neu loi 'khong tim thay control': mo rong phan Header/Item cua MIGO, chay SAP_ChanDoan va gui file.\", vbCritical, \"MIGO tung phan v1\"\r\n        Exit Sub\r\nTiepTheo:\r\n        On Error GoTo 0\r\n    Next v\r\n    MsgBox \"Xong. Kiem tra cot H/I sheet MIGO_TP, roi copy cot A + H dan vao trang web (buoc 4.3). Nho Ctrl+S.\", vbInformation, \"MIGO tung phan v1\"\r\nEnd Sub\r\n\r\nPrivate Sub GhiStatus(ws As Worksheet, rows As Collection, ByVal docNo As String, ByVal st As String)\r\n    Dim r As Variant\r\n    For Each r In rows\r\n        ws.Cells(r, \"H\").NumberFormat = \"@\"\r\n        If docNo <> \"\" Then ws.Cells(r, \"H\").Value = docNo\r\n        ws.Cells(r, \"I\").Value = st\r\n    Next r\r\nEnd Sub\r\n\r\n' Dien 1 phieu vao MIGO (chua Post). Tra ve so dong da tich OK.\r\nFunction DienMIGOTungPhan(session As Object, ws As Worksheet, rows As Collection) As Long\r\n    Dim soPO As String, ngay As String, phieuId As String, c As Object, tbl As Object\r\n    Dim cOK As Long, cTxt As Long, cQty As Long, cItem As Long, nRows As Long, i As Long, r As Variant\r\n    Dim want As Object: Set want = CreateObject(\"Scripting.Dictionary\")   ' Dong_PO -> row Excel\r\n\r\n    soPO = Trim(CStr(ws.Cells(rows(1), \"B\").Value))\r\n    ngay = Trim(CStr(ws.Cells(rows(1), \"C\").Value))\r\n    phieuId = Trim(CStr(ws.Cells(rows(1), \"A\").Value))\r\n    For Each r In rows\r\n        want(CLng(Val(ws.Cells(r, \"D\").Value))) = r\r\n    Next r\r\n\r\n    tpStep = \"mo MIGO\"\r\n    session.findById(\"wnd[0]\").maximize\r\n    session.findById(\"wnd[0]/tbar[0]/okcd\").Text = \"/nmigo\"\r\n    session.findById(\"wnd[0]\").sendVKey 0\r\n    CheckPopup session\r\n    Set c = FindCtlSoft(session, \"cmbGODYNPRO-ACTION\")\r\n    If Not c Is Nothing Then\r\n        If c.key <> MIGO_ACTION Then c.key = MIGO_ACTION: CheckPopup session\r\n    End If\r\n    Set c = FindCtlSoft(session, \"cmbGODYNPRO-REFDOC\")\r\n    If Not c Is Nothing Then\r\n        If c.key <> MIGO_REFDOC Then c.key = MIGO_REFDOC: CheckPopup session\r\n    End If\r\n\r\n    tpStep = \"go So PO \" & soPO\r\n    FindCtl(session, \"ctxtGODYNPRO-PO_NUMBER\").Text = soPO\r\n    session.findById(\"wnd[0]\").sendVKey 0\r\n    CheckPopup session\r\n    CheckSbarError session\r\n\r\n    tpStep = \"ngay chung tu / ngay hach toan \" & ngay\r\n    FindCtl(session, \"ctxtGOHEAD-BLDAT\").Text = ngay\r\n    FindCtl(session, \"ctxtGOHEAD-BUDAT\").Text = ngay\r\n    Set c = FindCtlSoft(session, \"txtGOHEAD-LFSNR\")\r\n    If Not c Is Nothing Then c.Text = Left(phieuId, 16)\r\n    session.findById(\"wnd[0]\").sendVKey 0\r\n    CheckPopup session\r\n\r\n    tpStep = \"doc bang dong MIGO\"\r\n    Set tbl = FindCtl(session, TP_TBL)\r\n    nRows = tbl.RowCount\r\n    cOK = TblColIndex(tbl, \"GOITEM-TAKE_IT\")\r\n    cTxt = TblColIndex(tbl, \"GOITEM-SGTXT\")\r\n    cQty = TblColIndex(tbl, \"GOITEM-ERFMG\")\r\n    cItem = TblColIndex(tbl, \"GOITEM-EBELP\")\r\n    If cOK < 0 Or cQty < 0 Or cItem < 0 Then\r\n        Err.Raise vbObjectError + 801, \"MIGO_TP\", \"Bang dong MIGO thieu cot OK (TAKE_IT) / So luong (ERFMG) / Dong PO (EBELP). Hien cot bang Layout roi chay lai.\"\r\n    End If\r\n\r\n    Dim it As Long, cel As Object, nOK As Long, seen As Object: Set seen = CreateObject(\"Scripting.Dictionary\")\r\n    For i = 0 To nRows - 1\r\n        tpStep = \"dong MIGO \" & (i + 1)\r\n        it = CLng(Val(MigoRowCell(session, i, cItem).Text))\r\n        If it = 0 Then GoTo NextRow\r\n        Set cel = MigoRowCell(session, i, cOK)\r\n        If want.Exists(it) Then\r\n            MigoRowCell(session, i, cQty).Text = SapNumber(CDbl(ws.Cells(want(it), \"F\").Value))\r\n            If cTxt >= 0 Then MigoRowCell(session, i, cTxt).Text = CStr(ws.Cells(want(it), \"G\").Value)\r\n            Set cel = MigoRowCell(session, i, cOK)\r\n            If Not cel.Selected Then cel.Selected = True\r\n            seen(it) = True\r\n            nOK = nOK + 1\r\n        Else\r\n            If cel.Selected Then cel.Selected = False\r\n        End If\r\nNextRow:\r\n    Next i\r\n    session.findById(\"wnd[0]\").sendVKey 0\r\n    CheckPopup session\r\n    CheckSbarError session\r\n\r\n    Dim miss As String, k As Variant\r\n    For Each k In want.Keys\r\n        If Not seen.Exists(k) Then miss = miss & k & \" \"\r\n    Next k\r\n    If miss <> \"\" Then Err.Raise vbObjectError + 802, \"MIGO_TP\", \"PO \" & soPO & \" khong con (hoac khong co) dong: \" & miss & \"- kiem tra so dong PO tren ME23N / PO da nhan du?\"\r\n    DienMIGOTungPhan = nOK\r\nEnd Function\r\n";
const TTPN={nhan:['Đã nhận, chờ MIGO','run'],xuat:['Đã xuất cho macro','run'],migo:['Đã MIGO','ok'],huy:['Đã hủy','bad']};
const pnAll=()=>Object.values(S.pn||{}).filter(p=>p.tt!=='huy');
const pnOf=t=>pnAll().filter(p=>p.thang===t).sort((a,b)=>a.d.localeCompare(b.d)||a.luc-b.luc);
const dhOf=t=>Object.values(S.dh||{}).filter(d=>d.thang===t&&d.tt!=='huy').sort((a,b)=>a.d.localeCompare(b.d)||a.luc-b.luc);
const poById=id=>S.po&&S.po[id];
const poItems=p=>p.lines.map((l,i)=>({...l,prItem:l.item,item:+l.ebelp||(i+1)*10}));
function poMerge(p){ const m=new Map(); poItems(p).forEach(l=>{const o=m.get(l.ma)||{ma:l.ma,ten:l.ten,dvt:l.dvt,g:l.g,sl:0,items:[]};o.sl+=l.sl;o.items.push(l.item);m.set(l.ma,o);}); return [...m.values()]; }
function poUsed(poId,exceptId){ const u={}; pnAll().filter(p=>p.poId===poId&&p.id!==exceptId).forEach(p=>p.lines.forEach(l=>(l.alloc||[]).forEach(a=>u[a.item]=(u[a.item]||0)+a.sl))); return u; }
function poRecvByMa(poId){ const u={}; pnAll().filter(p=>p.poId===poId).forEach(p=>p.lines.forEach(l=>u[l.ma]=(u[l.ma]||0)+(+l.sl||0))); return u; }
function poOrdByMa(poId){ const u={}; dhOf(S.thang).filter(d=>d.poId===poId).forEach(d=>d.lines.forEach(l=>u[l.ma]=(u[l.ma]||0)+(+l.sl||0))); return u; }
function allocItems(p,ma,qty,used){ let rem=qty; const out=[];
  poItems(p).filter(x=>x.ma===ma).forEach(x=>{const cap=Math.max(0,x.sl-(used[x.item]||0)), t=Math.min(cap,rem); if(t>1e-9){out.push({item:x.item,sl:+t.toFixed(3)});used[x.item]=(used[x.item]||0)+t;rem-=t;}});
  return {alloc:out,over:+Math.max(0,rem).toFixed(3)}; }
/* số nhận dùng chung cho báo cáo 2.4 và 3.3: phiếu nhận (bước 4) + số dán tay / nạp từ Công nợ */
function recvRows(){ const a=((S.nhan&&S.nhan.rows)||[]).slice();
  pnOf(S.thang).forEach(p=>p.lines.forEach(l=>{ if(+l.sl>0) a.push({d:p.d,ma:l.ma,ncc:p.ncc,sl:+l.sl,migo:p.migo||'',po:p.soPO,nguon:'Phiếu '+p.ma}); })); return a; }
const poWithNo=()=>poOf(S.thang).filter(p=>p.soPO&&p.tt!=='huy');
const tomorrowVN=()=>{const d=new Date(Date.now()+7*3600e3+864e5);return d.toISOString().slice(0,10);};
function migoText(d){ return `Mua hàng phục vụ Bếp tháng ${+d.slice(5,7)}`; }

/* ---- 4.1 Đặt hàng theo PO ---- */
function vDat(){
  const P=poWithNo(); if(!P.length) return head2('4.1 Đặt hàng theo PO','Tạo đơn đặt hàng hằng ngày từ dòng PO.')+'<div class="empty"><b>Chưa có PO nào có số</b>Ghi số PO ở bước 3.2.</div>';
  if(!ED.datPO||!P.find(p=>p.id===ED.datPO)) ED.datPO=P[0].id; if(!ED.datD) ED.datD=tomorrowVN();
  const p=poById(ED.datPO), M=poMerge(p), rec=poRecvByMa(p.id), ord=poOrdByMa(p.id); ED.datQ=ED.datQ||{};
  let tong=0; const over=[];
  const rows=M.map(m=>{const q=ED.datQ[m.ma]; const con=Math.max(0,m.sl-(rec[m.ma]||0)); if(q>0){tong+=q*m.g; if(q>con+1e-9) over.push(m.ma);}
    return `<tr class="${q>0?'edited':''}"><td class="num">${m.items.join(', ')}</td><td class="num">${m.ma}</td><td class="wrap">${esc(m.ten)}</td><td>${esc(m.dvt)}</td><td class="r num">${Q(m.sl)}</td><td class="r num">${Q(ord[m.ma]||0)}</td><td class="r num">${Q(rec[m.ma]||0)}</td><td class="r num">${Q(con)}</td>
    <td class="r"><input class="q" type="text" inputmode="decimal" id="dq-${m.ma}" data-dq="${m.ma}" value="${q?QI(q):''}" ${canWrite?'':'disabled'}>${q>con+1e-9?' <span class="pill bad">vượt PO</span>':''}</td></tr>`;}).join('');
  const done=dhOf(S.thang).filter(d=>d.poId===p.id).slice(-5).reverse();
  return head2('4.1 Đặt hàng theo PO','Chọn ngày giao và nhà cung cấp, nhập số lượng đặt cho từng mã. Đơn hàng mang số PO và dòng PO để bên giao ghi lên biên bản; số đặt được dùng làm số nhận mặc định ở bước 4.2.')+
  `<div class="toolbar"><label class="meta" for="dat-d">Ngày giao</label><input type="date" id="dat-d" value="${ED.datD}">
   <select id="dat-po">${P.map(x=>`<option value="${x.id}" ${x.id===p.id?'selected':''}>${esc(x.ncc)} · PO ${esc(x.soPO)}</option>`).join('')}</select><span class="sp"></span>
   ${canWrite?`<button class="btn" data-act4="dat-copy" ${tong?'':'disabled'}>Copy đơn hàng</button><button class="btn primary" data-act4="dat-save" ${tong?'':'disabled'}>Lưu đơn hàng</button>`:''}</div>
  ${over.length?`<div class="warn">${over.length} mã đặt vượt phần còn lại của PO. SAP không cho MIGO vượt PO: làm PR bổ sung (bước 3.3 hoặc 2.4) trước khi nhận.</div>`:''}
  <div class="meta">Tổng đơn: <b class="num">${VND(tong)} đ</b>${done.length?` · Đơn gần đây của PO này: ${done.map(d=>`${dm(d.d)} (${d.lines.length} mã)`).join(', ')}`:''}</div>
  <div class="tbl-wrap"><table class="t"><thead><tr><th>Dòng PO</th><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th class="r">SL PO</th><th class="r">Đã đặt</th><th class="r">Đã nhận</th><th class="r">Còn lại</th><th class="r">SL đặt</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}
function datText(p,d,lines){ return `ĐƠN ĐẶT HÀNG – Bếp Vinmec Ocean Park 2\nNgày giao: ${dm(d)}/${d.slice(0,4)}\nNCC: ${p.ncc} · Số PO: ${p.soPO}\n`+lines.map((l,i)=>`${i+1}. [PO ${p.soPO} dòng ${l.items.join('/')}] ${l.ma} ${l.ten}: ${Q(l.sl)} ${l.dvt}`).join('\n')+`\nĐề nghị ghi số PO và mã hàng trên biên bản giao hàng.`; }

/* ---- 4.2 Nhận hàng ---- */
function nhanDraft(){
  if(ED.nd) return ED.nd;
  const P=poWithNo(); const dhs=dhOf(S.thang).filter(d=>!d.pn);
  const dh=dhs.find(d=>d.id===ED.ndDH)||null; const po=dh?poById(dh.poId):(P.find(p=>p.id===ED.ndPO)||P[0]);
  if(!po) return null;
  const M=new Map(poMerge(po).map(m=>[m.ma,m]));
  const lines=dh?dh.lines.map(l=>{const m=M.get(l.ma)||{};return {ma:l.ma,ten:m.ten||l.ten,dvt:m.dvt||l.dvt,g:m.g||l.g,slDat:l.sl,sl:l.sl,gach:false,note:''};}):[];
  ED.nd={dh:dh?dh.id:'',poId:po.id,d:dh?dh.d:todayVN(),lines,nguon:dh?'Đơn hàng '+dm(dh.d):'Nhập tay'}; return ED.nd;
}
function vNhan(){
  const P=poWithNo(); if(!P.length) return head2('4.2 Nhận hàng','Ghi số thực nhận theo từng lần giao.')+'<div class="empty"><b>Chưa có PO nào có số</b>Ghi số PO ở bước 3.2.</div>';
  const N=nhanDraft(), po=poById(N.poId), used=poUsed(po.id), dhs=dhOf(S.thang).filter(d=>!d.pn);
  const M=poMerge(po), sim={...used}; let tong=0;
  const rows=N.lines.map((l,i)=>{const q=l.gach?0:+l.sl||0; const a=allocItems(po,l.ma,q,sim); tong+=q*l.g;
    return `<tr class="${l.gach?'':(Math.abs(q-(l.slDat||0))>1e-9?'edited':'')}"><td class="num">${l.ma}</td><td class="wrap">${esc(l.ten)}</td><td>${esc(l.dvt)}</td><td class="r num">${l.slDat!=null?Q(l.slDat):'–'}</td>
    <td class="r"><input class="q" type="text" inputmode="decimal" id="nq-${i}" data-nq="${i}" value="${QI(l.sl)}" ${l.gach||!canWrite?'disabled':''}></td>
    <td><label class="meta"><input type="checkbox" data-ngach="${i}" ${l.gach?'checked':''}> gạch bỏ</label></td>
    <td class="num meta">${a.alloc.map(x=>x.item+': '+Q(x.sl)).join(', ')||'–'}</td><td>${a.over>0?`<span class="pill bad">vượt PO ${Q(a.over)}</span>`:''}</td>
    <td><input type="text" id="nn-${i}" data-nn="${i}" value="${esc(l.note||'')}" placeholder="ghi chú" style="min-width:110px"></td><td><button class="x" data-act4="nl-del" data-i="${i}" aria-label="Xóa dòng">×</button></td></tr>`;}).join('');
  const addable=M.filter(m=>!N.lines.some(l=>l.ma===m.ma));
  return head2('4.2 Nhận hàng','Mỗi lần giao là một phiếu nhận: ngày giao, PO, số thực nhận từng mã. Mặc định thực nhận = số đặt; dòng bị gạch bỏ trên biên bản = 0. Đây là dữ liệu chuẩn để đối soát NCC (bước 5); phần vượt PO được xử lý một lần cuối chu kỳ (bước 6). Có thể đọc ảnh biên bản để điền tự động.')+
  `<div class="toolbar"><select id="nd-dh"><option value="">— Nhận không theo đơn —</option>${dhs.map(d=>`<option value="${d.id}" ${d.id===N.dh?'selected':''}>Đơn ${dm(d.d)} · ${esc(d.ncc)} · ${d.lines.length} mã</option>`).join('')}</select>
   ${N.dh?'':`<select id="nd-po">${P.map(x=>`<option value="${x.id}" ${x.id===po.id?'selected':''}>${esc(x.ncc)} · PO ${esc(x.soPO)}</option>`).join('')}</select>`}
   <label class="meta" for="nd-d">Ngày giao (ngày hạch toán)</label><input type="date" id="nd-d" value="${N.d}">
   <span class="sp"></span>${canWrite?`<label class="btn sm" for="nd-img" style="${sample?'':'opacity:.5'}">${ED.ndAI?esc(ED.ndAI):'Đọc ảnh biên bản'}</label><input type="file" id="nd-img" accept="image/*" multiple hidden ${sample&&!ED.ndAI?'':'disabled'}>`:''}</div>
  <div class="meta">PO ${esc(po.soPO)} · ${esc(po.ncc)} · nguồn: ${esc(N.nguon)} · tổng <b class="num">${VND(tong)} đ</b></div>
  ${N.lines.length?`<div class="tbl-wrap"><table class="t"><thead><tr><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th class="r">SL đặt</th><th class="r">Thực nhận</th><th></th><th>Dòng PO: SL</th><th></th><th>Ghi chú</th><th></th></tr></thead><tbody>${rows}</tbody></table></div>`:'<div class="empty"><b>Phiếu chưa có dòng</b>Chọn đơn hàng, đọc ảnh biên bản, hoặc thêm mã từ PO bên dưới.</div>'}
  ${canWrite?`<div class="toolbar"><select id="nd-add"><option value="">+ Thêm mã từ PO…</option>${addable.map(m=>`<option value="${m.ma}">${m.ma} · ${esc(m.ten)}</option>`).join('')}</select><span class="sp"></span>
   <button class="btn" data-act4="nd-reset">Làm lại</button><button class="btn primary" data-act4="nd-save" ${N.lines.length?'':'disabled'}>Lưu phiếu nhận</button></div>`:''}`;
}
async function readImages(files){
  if(!sample){toast('Trang chưa được phép hỏi Claude.');return;}
  const N=nhanDraft(), po=poById(N.poId), M=poMerge(po);
  let lim=null; try{ lim=await sample.limits(); }catch(e){}
  if(lim&&!lim.images){ toast('Phiên xem này chưa cho phép gửi ảnh cho Claude.'); return; }
  ED.ndAI='Claude đang đọc ảnh…'; render();
  const prompt=`Đây là ảnh biên bản giao hàng của nhà cung cấp ${po.ncc} cho Bếp Vinmec Ocean Park 2 (PO ${po.soPO}).
Đọc từng dòng hàng và ghép với DANH SÁCH DÒNG PO bên dưới (theo mã hàng nếu có, không thì theo tên). Quy tắc:
- "sl" = số ở ô Thực nhận. Nếu ô Thực nhận để trống thì để null (sẽ lấy số đặt/số PO trên biên bản).
- "sl_dat" = số đặt / SL PO in trên biên bản nếu có.
- Dòng bị GẠCH BỎ (không giao) thì "gach_bo": true.
- Nếu không chắc mã thì "can_kiem_tra": true và ghi lý do ngắn.
Trả JSON đúng dạng: {"ngay":"dd/mm/yyyy hoặc null","lines":[{"ma":"THP…","sl":12.5,"sl_dat":12,"gach_bo":false,"can_kiem_tra":false,"ghi_chu":""}]}
DANH SÁCH DÒNG PO (mã|tên|ĐVT):
${M.map(m=>`${m.ma}|${m.ten}|${m.dvt}`).join('\n')}`;
  try{
    const res=await SJ(prompt,{images:[...files],modelTier:'default'});
    let n=0,chk=0; const ok=new Set(M.map(m=>m.ma));
    (res&&res.lines||[]).forEach(r=>{ const ma=String(r.ma||'').toUpperCase(); if(!ok.has(ma)) {chk++;return;} const m=M.find(x=>x.ma===ma);
      const slDat=r.sl_dat!=null?+r.sl_dat:null, sl=r.gach_bo?0:(r.sl!=null?+r.sl:(slDat!=null?slDat:0));
      const ex=N.lines.find(l=>l.ma===ma); const note=[r.can_kiem_tra?'⚠ cần kiểm tra':'',r.ghi_chu||''].filter(Boolean).join(' · ');
      if(ex){ ex.sl=sl; ex.gach=!!r.gach_bo; ex.note=note; if(ex.slDat==null) ex.slDat=slDat; } else N.lines.push({ma,ten:m.ten,dvt:m.dvt,g:m.g,slDat,sl,gach:!!r.gach_bo,note}); n++; if(r.can_kiem_tra) chk++; });
    const dd=res&&res.ngay&&String(res.ngay).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if(dd) N.d=`${dd[3]}-${dd[2].padStart(2,'0')}-${dd[1].padStart(2,'0')}`;
    N.nguon=(N.nguon||'')+' + ảnh ('+files.length+')'; N.anh=files.length; ED.ndAI=null; render(); toast(`Claude đọc được ${n} dòng${chk?`, ${chk} dòng cần kiểm tra`:''}. Soát lại rồi Lưu phiếu.`);
  }catch(e){ ED.ndAI=null; render(); toast(e&&e.code==='rate_limited'?'Claude đang bận, thử lại sau.':'Không đọc được ảnh: '+((e&&e.message)||e)); }
}

/* ---- 4.3 MIGO từng phần ---- */
function migoRows(list){ const out=[]; list.forEach(p=>p.lines.forEach(l=>(l.alloc||[]).forEach(a=>out.push([p.ma,p.soPO,p.d.split('-').reverse().join('.'),a.item,l.ma,a.sl,migoText(p.d)])))); return out; }
function vMigo(){
  const W=pnOf(S.thang).filter(p=>p.tt==='nhan'||p.tt==='xuat'); if(!ED.msel) ED.msel=new Set(W.filter(p=>p.tt==='nhan').map(p=>p.id));
  const sel=W.filter(p=>ED.msel.has(p.id)), rows=migoRows(sel), overs=sel.flatMap(p=>p.lines.filter(l=>l.over>0).map(l=>({p,l})));
  return head2('4.3 MIGO từng phần','Mỗi phiếu nhận là một chứng từ MIGO: nhập kho theo PO, chỉ các dòng PO có trong phiếu, số lượng = thực nhận, ngày hạch toán = ngày giao. Xuất sang sheet MIGO_TP của file macro, chạy PostMIGO_TungPhan (module mới MIGO_TungPhan_v1), rồi dán số chứng từ về đây.')+
  `<div class="box"><b>Chạy trên máy SAP</b><div class="meta">1) Import module <b>MIGO_TungPhan_v1.bas</b> vào Template_PR_v3.xlsm (cùng file có SAP_Automation v22). 2) Dán dữ liệu vào sheet MIGO_TP ô A3. 3) Alt+F8 → PostMIGO_TungPhan. Đang ở <b>chế độ không lưu</b>: macro chỉ điền màn hình và bấm Check, không Post. 4) Copy cột A (Ma_phieu) và H (So_chung_tu_MIGO) dán vào ô bên dưới.</div>
   <div class="toolbar"><button class="btn sm" data-act4="bas-dl">Tải macro MIGO_TungPhan_v1.bas</button><span class="meta">Bản này cũng nằm trong thư mục scripts của dự án.</span></div></div>
  ${W.length?`<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th></th><th>Mã phiếu</th><th>Ngày giao</th><th>NCC</th><th>PO</th><th class="r">Dòng</th><th class="r">Tiền</th><th>Trạng thái</th></tr></thead><tbody>
   ${W.map(p=>`<tr><td><input type="checkbox" data-msel="${p.id}" ${ED.msel.has(p.id)?'checked':''} aria-label="Chọn ${esc(p.ma)}"></td><td class="num">${esc(p.ma)}</td><td>${dm(p.d)}</td><td>${esc(p.ncc)}</td><td class="num">${esc(p.soPO)}</td><td class="r num">${p.lines.filter(l=>(l.alloc||[]).length).length}</td><td class="r num">${VND(p.tong)}</td><td><span class="pill ${TTPN[p.tt][1]}">${TTPN[p.tt][0]}</span>${p.lines.some(l=>l.over>0)?' <span class="pill bad">có phần vượt PO</span>':''}</td></tr>`).join('')}
   </tbody></table></div>
   ${overs.length?`<div class="warn">Phần vượt PO không đưa vào MIGO: ${overs.map(o=>`${esc(o.p.ma)} ${o.l.ma} +${Q(o.l.over)}`).join(', ')}. Làm PR/PO bổ sung rồi nhận bằng phiếu mới.</div>`:''}
   <div class="toolbar"><button class="btn primary" data-act4="migo-xlsx" ${rows.length?'':'disabled'}>Tải file cho sheet MIGO_TP</button><button class="btn" data-act4="migo-copy" ${rows.length?'':'disabled'}>Copy để dán</button><span class="meta">${rows.length} dòng MIGO · ${sel.length} phiếu</span></div>
   <div class="tbl-wrap" style="max-height:280px"><table class="t"><thead><tr>${['Ma_phieu','So_PO','Ngay_hach_toan','Dong_PO','Ma_vat_tu','So_luong','Text'].map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.slice(0,200).map(r=>`<tr>${r.map((v,j)=>`<td class="${j===5?'r num':''}">${j===5?Q(v):esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`
   :'<div class="empty"><b>Không có phiếu nào chờ MIGO</b>Lưu phiếu nhận ở 4.2.</div>'}
  ${canWrite?`<div class="box"><label class="meta" for="migo-txt">Dán kết quả từ sheet MIGO_TP: <b>Ma_phieu</b> ⇥ <b>So_chung_tu_MIGO</b> (cột A và H). Ở chế độ không lưu chưa có số chứng từ thì bỏ qua bước này.</label><textarea class="paste" id="migo-txt" style="min-height:70px"></textarea><div class="toolbar"><button class="btn sm primary" data-act4="migo-ok">Ghi số MIGO</button></div></div>`:''}`;
}

/* ---- 4.4 Lịch sử ---- */
function vNls(){
  const L=pnOf(S.thang).slice().reverse(), D=dhOf(S.thang).slice().reverse();
  const body=L.map(p=>{const open=ED.pnOpen===p.id;
    return `<tr><td><button class="x" data-act4="pn-open" data-id="${p.id}" aria-label="Xem">${open?'▾':'▸'}</button></td><td class="num">${esc(p.ma)}</td><td>${dm(p.d)}</td><td>${esc(p.ncc)}</td><td class="num">${esc(p.soPO)}</td><td class="r num">${p.lines.length}</td><td class="r num">${VND(p.tong)}</td><td><span class="pill ${TTPN[p.tt][1]}">${TTPN[p.tt][0]}</span></td><td class="num">${esc(p.migo||'')}</td><td class="meta">${esc(p.nguon||'')}</td>
    <td>${lockedDay(p.ncc,p.d)?'<span class="pill ok">đã khớp NCC</span>':canWrite?`<button class="btn sm warn" data-act4="pn-huy" data-id="${p.id}">Hủy</button>`:''}</td></tr>
    ${open?`<tr><td></td><td colspan="10"><div class="tbl-wrap" style="max-height:260px"><table class="t"><thead><tr><th>Mã</th><th>Tên</th><th class="r">Đặt</th><th class="r">Thực nhận</th><th>Dòng PO</th><th class="r">Vượt PO</th><th>Ghi chú</th></tr></thead><tbody>${p.lines.map(l=>`<tr><td class="num">${l.ma}</td><td class="wrap">${esc(l.ten)}</td><td class="r num">${l.slDat!=null?Q(l.slDat):''}</td><td class="r num">${l.gach?'<span class="pill">gạch bỏ</span>':Q(l.sl)}</td><td class="num">${(l.alloc||[]).map(a=>a.item+': '+Q(a.sl)).join(', ')}</td><td class="r num">${l.over>0?Q(l.over):''}</td><td>${esc(l.note||'')}</td></tr>`).join('')}</tbody></table></div>
      <ol class="meta" style="margin:6px 0 0;padding-left:18px">${(p.hist||[]).map(h=>`<li>${tstr(h.ts)} · ${(TTPN[h.tt]||[h.tt])[0]} · ${esc(nm(h.ai))}${h.note?' · '+esc(h.note):''}</li>`).join('')}</ol></td></tr>`:''}`;}).join('');
  return head2('4.3 Lịch sử nhận hàng','Mọi đơn đặt hàng và phiếu nhận trong tháng, kèm phân bổ dự kiến vào dòng PO. Phiếu thuộc kỳ đã chốt khớp với NCC bị khóa.')+
  `<div class="kpis"><div class="kv big"><b>Phiếu nhận tháng ${mmY(S.thang)}</b><div class="num">${L.length}</div><small>${TR(L.reduce((a,p)=>a+p.tong,0))}</small></div>
   <div class="kv"><b>Đã chốt khớp NCC</b><div class="num">${L.filter(p=>lockedDay(p.ncc,p.d)).length}/${L.length}</div><small>phiếu</small></div>
   <div class="kv"><b>Đơn đặt hàng</b><div class="num">${D.length}</div><small>${D.filter(d=>!d.pn).length} đơn chưa nhận</small></div></div>
  ${L.length?`<div class="tbl-wrap" style="max-height:none"><table class="t"><thead><tr><th></th><th>Mã phiếu</th><th>Ngày</th><th>NCC</th><th>PO</th><th class="r">Dòng</th><th class="r">Tiền</th><th>Trạng thái</th><th>Số MIGO</th><th>Nguồn</th><th></th></tr></thead><tbody>${body}</tbody></table></div>`:'<div class="empty"><b>Chưa có phiếu nhận</b>Tạo ở 4.2.</div>'}
  ${D.length?`<h3 class="eyebrow" style="margin:6px 0 0">Đơn đặt hàng</h3><div class="tbl-wrap" style="max-height:260px"><table class="t"><thead><tr><th>Ngày giao</th><th>NCC</th><th>PO</th><th class="r">Mã</th><th class="r">Tiền</th><th>Đã nhận</th><th></th></tr></thead><tbody>${D.map(d=>`<tr><td>${dm(d.d)}</td><td>${esc(d.ncc)}</td><td class="num">${esc(d.soPO)}</td><td class="r num">${d.lines.length}</td><td class="r num">${VND(d.tong)}</td><td>${d.pn?'<span class="pill ok">đã có phiếu</span>':'<span class="pill run">chưa nhận</span>'}</td><td>${canWrite&&!d.pn?`<button class="btn sm warn" data-act4="dh-huy" data-id="${d.id}">Hủy</button>`:''}</td></tr>`).join('')}</tbody></table></div>`:''}`;
}

/* ---- sự kiện bước 4 ---- */
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-act4]'); if(!t) return; const a=t.dataset.act4, id=t.dataset.id;
  try{
  if(a==='dat-save'||a==='dat-copy'){ const p=poById(ED.datPO), M=poMerge(p); const L=M.filter(m=>ED.datQ[m.ma]>0).map(m=>({ma:m.ma,ten:m.ten,dvt:m.dvt,g:m.g,sl:ED.datQ[m.ma],items:m.items}));
    if(a==='dat-copy'){ try{ await navigator.clipboard.writeText(datText(p,ED.datD,L)); toast('Đã copy đơn hàng.'); }catch(err){ toast('Không copy được.'); } return; }
    if(ED.datD.slice(0,7)!==p.thang){ toast('Ngày giao phải nằm trong tháng của PO ('+mmY(p.thang)+').'); return; }
    const did=p.thang+'-'+ED.datD.slice(8)+'-'+slug(p.ncc)+'-'+Date.now().toString(36), tong=Math.round(L.reduce((s,l)=>s+l.sl*l.g,0));
    await store.set('donhang/'+did,{thang:p.thang,d:ED.datD,ncc:p.ncc,poId:p.id,soPO:p.soPO,lines:L.map(({items,...l})=>l),tong,tt:'dat',pn:'',luc:Date.now(),ai:uid});
    await log('4.1',`Đơn hàng ${dm(ED.datD)} – ${p.ncc} (PO ${p.soPO}): ${L.length} mã, ${TR(tong)}`); ED.datQ={}; toast('Đã lưu đơn hàng. Khi hàng về, ghi nhận ở 4.2.'); }
  else if(a==='nl-del'){ ED.nd.lines.splice(+t.dataset.i,1); render(); }
  else if(a==='nd-reset'){ ED.nd=null; render(); }
  else if(a==='nd-save'){ const N=ED.nd, po=poById(N.poId); if(N.d.slice(0,7)!==po.thang){toast('Ngày giao phải nằm trong tháng của PO.');return;} if(lockedDay(po.ncc,N.d)){toast('Kỳ này của '+po.ncc+' đã chốt khớp; mở lại ở 5.3 trước.');return;}
    const used=poUsed(po.id); const lines=N.lines.map(l=>{const q=l.gach?0:Math.max(0,+l.sl||0); const r=allocItems(po,l.ma,q,used); return {...l,sl:q,alloc:r.alloc,over:r.over};});
    const tong=Math.round(lines.reduce((s,l)=>s+l.sl*l.g,0)), seq=pnOf(po.thang).filter(p=>p.d===N.d).length+1;
    const ma=`NH${N.d.slice(8)}${N.d.slice(5,7)}-${slug(po.ncc).slice(0,8).toUpperCase()}-${seq}`, pid=po.thang+'-'+Date.now().toString(36);
    await store.set('phieunhan/'+pid,{thang:po.thang,ma,d:N.d,ncc:po.ncc,poId:po.id,soPO:po.soPO,dh:N.dh,lines,tong,nguon:N.nguon,tt:'nhan',migo:'',luc:Date.now(),ai:uid,hist:[{ts:Date.now(),tt:'nhan',ai:uid,note:N.nguon}]});
    if(N.dh){ const d=S.dh[N.dh]; const x={...d,pn:pid,tt:'nhan'}; delete x.id; await store.set('donhang/'+N.dh,x); }
    const ov=lines.filter(l=>l.over>0); await log('4.2',`Phiếu nhận ${ma}: ${lines.length} mã, ${TR(tong)}${ov.length?`; vượt PO ${ov.length} mã (${ov.map(l=>l.ma).join(', ')})`:''}`);
    ED.nd=null; ED.ndDH=''; ED.msel=null; toast(ov.length?'Đã lưu phiếu. Phần vượt PO sẽ xử lý cuối chu kỳ (bước 6).':'Đã lưu phiếu nhận.'); }
  else if(a==='bas-dl'){ if(!dlcap){toast('Trang này không tải file được; lấy file trong thư mục scripts của dự án.');return;}
    try{ await dlcap.save({filename:'MIGO_TungPhan_v1.bas',data:new Blob([MACRO_BAS],{type:'text/plain'})}); }catch(err){ if(err&&err.code!=='cancelled') toast('Không lưu được file.'); } }
  else if(a==='migo-copy'||a==='migo-xlsx'){ const sel=pnOf(S.thang).filter(p=>(p.tt==='nhan'||p.tt==='xuat')&&ED.msel.has(p.id)), rows=migoRows(sel);
    if(a==='migo-copy'){ try{ await navigator.clipboard.writeText(rows.map(r=>r.map((v,j)=>j===5?String(v).replace('.',','):String(v)).join('\t')).join('\r\n')); toast('Đã copy '+rows.length+' dòng. Dán vào sheet MIGO_TP ô A3.'); }catch(err){ toast('Không copy được.'); return; } }
    else { if(!window.XLSX){toast('Không tạo được file; dùng nút Copy.');return;}
      const ws=XLSX.utils.aoa_to_sheet([['Ma_phieu','So_PO','Ngay_hach_toan','Dong_PO','Ma_vat_tu','So_luong','Text'],...rows]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'MIGO_TP');
      try{ await saveOut(`MIGO_TP_${S.thang}_${todayVN()}.xlsx`,new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); }catch(err){ if(err&&err.code!=='cancelled') toast('Không lưu được file.'); return; } }
    for(const p of sel) if(p.tt==='nhan'){ const x={...p,tt:'xuat',hist:[...(p.hist||[]),{ts:Date.now(),tt:'xuat',ai:uid,note:'Xuất cho MIGO_TP'}]}; delete x.id; await store.set('phieunhan/'+p.id,x); }
    await log('4.3',`Xuất ${rows.length} dòng MIGO (${sel.length} phiếu) cho sheet MIGO_TP`); }
  else if(a==='migo-ok'){ const map=new Map(pnOf(S.thang).map(p=>[p.ma,p])); let n=0; const bad=[];
    for(const line of $('#migo-txt').value.split(/\r?\n/)){ if(!line.trim()) continue; const c=line.split('\t').map(x=>x.trim()); const p=map.get(c[0]);
      if(!p||!/^\d{6,12}$/.test(c[1]||'')){bad.push(line);continue;} if(p.migo===c[1]) continue;
      const x={...p,tt:'migo',migo:c[1],migoLuc:Date.now(),hist:[...(p.hist||[]),{ts:Date.now(),tt:'migo',ai:uid,note:'Số MIGO '+c[1]}]}; delete x.id; await store.set('phieunhan/'+p.id,x); n++; }
    if(n) await log('4.3',`Ghi ${n} số chứng từ MIGO`); $('#migo-txt').value=bad.join('\n'); toast(`Đã ghi ${n} phiếu${bad.length?`, ${bad.length} dòng không khớp mã phiếu hoặc số MIGO`:''}.`); }
  else if(a==='pn-open'){ ED.pnOpen=ED.pnOpen===id?null:id; render(); }
  else if(a==='pn-huy'){ const p=S.pn[id]; if(ED.pnHuy!==id){ED.pnHuy=id;toast('Bấm Hủy lần nữa để hủy phiếu '+p.ma+'.');return;} ED.pnHuy=null;
    const x={...p,tt:'huy',hist:[...(p.hist||[]),{ts:Date.now(),tt:'huy',ai:uid,note:'Hủy phiếu'}]}; delete x.id; await store.set('phieunhan/'+id,x);
    if(p.dh&&S.dh[p.dh]){ const d={...S.dh[p.dh],pn:'',tt:'dat'}; delete d.id; await store.set('donhang/'+p.dh,d); } await log('4.4',`Hủy phiếu nhận ${p.ma}`); }
  else if(a==='dh-huy'){ const d={...S.dh[id],tt:'huy'}; delete d.id; await store.set('donhang/'+id,d); await log('4.4',`Hủy đơn hàng ${dm(d.d)} – ${d.ncc}`); }
  }catch(err){ toast('Không lưu được: '+(err&&err.message||err)); }
});
document.addEventListener('input',e=>{ const t=e.target;
  if(t.dataset.dq){ ED.datQ[t.dataset.dq]=Math.max(0,pnum(t.value)); clearTimeout(vDat.h); vDat.h=setTimeout(render,700); }
  else if(t.dataset.nq!=null){ ED.nd.lines[+t.dataset.nq].sl=Math.max(0,pnum(t.value)); clearTimeout(vNhan.h); vNhan.h=setTimeout(render,700); }
  else if(t.dataset.nn!=null){ ED.nd.lines[+t.dataset.nn].note=t.value; }
});
document.addEventListener('change',e=>{ const t=e.target;
  if(t.id==='dat-po'){ ED.datPO=t.value; ED.datQ={}; render(); }
  else if(t.id==='dat-d'){ ED.datD=t.value; render(); }
  else if(t.id==='nd-dh'){ ED.ndDH=t.value; ED.nd=null; render(); }
  else if(t.id==='nd-po'){ ED.ndPO=t.value; ED.nd=null; render(); }
  else if(t.id==='nd-d'){ ED.nd.d=t.value; }
  else if(t.dataset.ngach!=null){ const l=ED.nd.lines[+t.dataset.ngach]; l.gach=t.checked; render(); }
  else if(t.id==='nd-add'&&t.value){ const po=poById(ED.nd.poId), m=poMerge(po).find(x=>x.ma===t.value); ED.nd.lines.push({ma:m.ma,ten:m.ten,dvt:m.dvt,g:m.g,slDat:null,sl:0,gach:false,note:''}); render(); }
  else if(t.id==='nd-img'&&t.files&&t.files.length){ readImages([...t.files]); }
  else if(t.dataset.msel){ ED.msel=ED.msel||new Set(); t.checked?ED.msel.add(t.dataset.msel):ED.msel.delete(t.dataset.msel); render(); }
});

