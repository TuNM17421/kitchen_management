# =====================================================================
#  TRAM SAP - chay tren may cong ty co SAP GUI
#  Nhan lenh tu giao dien "Dieu hanh PR truoc" (qua file PRTruoc_DuLieu),
#  tu dien Template_PR_Tram.xlsm va chay macro CreatePR / CreatePO / sua PO (ME22N) / MIGO,
#  roi gui so chung tu / trang thai ve giao dien.
#  Chay bang ChayTramSAP.bat. Dung: dong cua so nay (Ctrl+C).
# =====================================================================
$ErrorActionPreference = 'Stop'
$Ver = 'TramSAP v4.2 (31/10/2026)'
$script:GhiChuLenh = ''
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$cfg = Get-Content (Join-Path $Root 'TramSAP.config.json') -Raw | ConvertFrom-Json
$LogFile = Join-Path $Root 'TramSAP_log.txt'
$KqDir = Join-Path $Root 'KetQua'
if (-not (Test-Path $KqDir)) { New-Item -ItemType Directory -Path $KqDir | Out-Null }
# Chay ngam: Excel luon an; cua so SAP thu nho khi tram dung (dat "anCuaSoSAP": false de giu nguyen cua so SAP)
$AnSap = $true; if ($null -ne $cfg.anCuaSoSAP) { $AnSap = [bool]$cfg.anCuaSoSAP }
# So cua so SAP chay song song cho PR nhieu nhom (vd be1 + be2). 1 = chay lan luot.
$SoPhien = 2; if ($null -ne $cfg.soPhienSAP) { $SoPhien = [Math]::Max(1, [int]$cfg.soPhienSAP) }

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try { [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials } catch {}

# ---------- Dong Excel an cho chac: Quit + giai phong COM + tat dung tien trinh Excel do tram mo (theo PID) ----------
$XlPidSrc = 'using System; using System.Runtime.InteropServices; public static class XlPid { [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int p); }'
if (-not ('XlPid' -as [type])) { Add-Type -TypeDefinition $XlPidSrc }
function LayPidExcel($xl) { $p = 0; try { [void][XlPid]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) } catch {}; return $p }
function DongExcel($xl, $xpid) {
  if ($xl) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {} }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if ($xpid) { Start-Sleep -Milliseconds 800; if (Get-Process -Id $xpid -ErrorAction SilentlyContinue) { Stop-Process -Id $xpid -Force -ErrorAction SilentlyContinue } }
}
function FileBiKhoa($path) { try { $f = [IO.File]::Open($path, 'Open', 'ReadWrite', 'None'); $f.Close(); return $false } catch { return $true } }

function Ghi($msg) {
  $line = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + '  ' + $msg
  Write-Host $line
  Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Goi($body) {
  $body.token = $cfg.token
  $body.may = $env:COMPUTERNAME
  $body.ver = $Ver
  $json = $body | ConvertTo-Json -Depth 30 -Compress
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  $r = Invoke-WebRequest -Uri $cfg.url -Method Post -Body $bytes -ContentType 'application/json; charset=utf-8' -UseBasicParsing -TimeoutSec 120
  $txt = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
  $o = $txt | ConvertFrom-Json
  if (-not $o.ok) { throw ('Cong Google bao loi: ' + $o.loi) }
  return $o
}

function TienDo($id, $t) { try { [void](Goi @{ a = 'tiendo'; id = $id; tiendo = $t }) } catch { Ghi ('  (khong gui duoc tien do: ' + $_.Exception.Message + ')') } }

function Txt($cell) { $v = $cell.Value2; if ($null -eq $v) { return '' } ; return ([string]$v).Trim() }

function DienBang($wb, $sheet, $rows, $nCol) {
  $n = @($rows).Count
  $arr = New-Object 'object[,]' $n, $nCol
  for ($i = 0; $i -lt $n; $i++) {
    $c = @($rows[$i].c)
    for ($j = 0; $j -lt $nCol; $j++) {
      $v = $(if ($j -lt $c.Count) { $c[$j] } else { $null })
      if ($v -is [decimal] -or $v -is [int] -or $v -is [long]) { $v = [double]$v }
      elseif ($v -is [string] -and $v.StartsWith('=')) { $v = "'" + $v }
      $arr[$i, $j] = $v
    }
  }
  $ws = $wb.Worksheets.Item($sheet)
  $last = [char](64 + $nCol)
  $ws.Range('A3:' + $last + (2 + $n)).Value2 = $arr
  return $ws
}

# ---------- PR song song: moi nhom PR (Nhom_PR, vd be1 / be2) chay tren 1 cua so SAP + 1 Excel an rieng ----------
$WorkerPR = {
  param($tpl, $work, $rowsJson, $phien, $luu, $anSap)
  $ErrorActionPreference = 'Stop'
  $out = @{ rows = @(); log = ''; loi = '' }
  $rows = @(($rowsJson | ConvertFrom-Json) | ForEach-Object { $_ })
  $xl = $null; $wb = $null
  try {
    Copy-Item -Path $tpl -Destination $work -Force
    if (-not ('XlPid' -as [type])) { Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class XlPid { [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int p); }' }
    $xl = New-Object -ComObject Excel.Application
    $xpid = 0; try { [void][XlPid]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xpid) } catch {}
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($work)
    [void]$xl.Run('TramCauHinh', [int]$phien, [bool]$anSap)
    [void]$xl.Run('TramChuanBi')
    $n = $rows.Count
    $arr = New-Object 'object[,]' $n, 11
    for ($i = 0; $i -lt $n; $i++) {
      $c = @($rows[$i].c)
      for ($j = 0; $j -lt 11; $j++) {
        $v = $(if ($j -lt $c.Count) { $c[$j] } else { $null })
        if ($v -is [decimal] -or $v -is [int] -or $v -is [long]) { $v = [double]$v }
        elseif ($v -is [string] -and $v.StartsWith('=')) { $v = "'" + $v }
        $arr[$i, $j] = $v
      }
    }
    $wb.Worksheets.Item('CongNo').Range('A3:K' + (2 + $n)).Value2 = $arr
    $xl.Calculate()
    $out.log = [string]$xl.Run('TramChayPR', [bool]$luu)
    $ws = $wb.Worksheets.Item('PR')
    $T = { param($cell) $v = $cell.Value2; if ($null -eq $v) { '' } else { ([string]$v).Trim() } }
    for ($i = 0; $i -lt $n; $i++) {
      $r = 3 + $i
      $out.rows += @{ soPR = (& $T $ws.Range("L$r")); tt = (& $T $ws.Range("M$r")); item = (& $T $ws.Range("P$r")); giaSAP = (& $T $ws.Range("R$r")); slSAP = (& $T $ws.Range("S$r")) }
    }
    $wb.Save()
  }
  catch { $out.loi = $_.Exception.Message }
  finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($xpid) { Start-Sleep -Milliseconds 800; if (Get-Process -Id $xpid -ErrorAction SilentlyContinue) { Stop-Process -Id $xpid -Force -ErrorAction SilentlyContinue } }
  }
  return ($out | ConvertTo-Json -Depth 10 -Compress)
}

function ChayPRSongSong($L) {
  $id = $L.id; $luu = [bool]$L.luuThat; $rows = @($L.rows)
  $groups = [ordered]@{}
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $g = [string]$rows[$i].c[8]
    if (-not $groups.Contains($g)) { $groups[$g] = New-Object System.Collections.ArrayList }
    [void]$groups[$g].Add($i)
  }
  $keys = @($groups.Keys)
  $nP = [Math]::Min($keys.Count, $SoPhien)
  TienDo $id ("Chuan bi $nP cua so SAP")
  $have = 0
  $xl = New-Object -ComObject Excel.Application
  $xpid = LayPidExcel $xl
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  try {
    $wb = $xl.Workbooks.Open([string]$cfg.fileExcel, 0, $true)
    $ver = [string]$xl.Run('TramPhienBan')
    if (-not $ver.StartsWith('v23.1')) { $wb.Close($false); throw ("MACRO_CU:File Excel tram dang co macro $ver - chua co chay song song. Import lai SAP_Automation_v23_Tram.bas ban 31/10 (PhienBan phai la v23.1-Tram)") }
    $have = [int]$xl.Run('TramDamBaoPhien', [int]$nP, [bool]$AnSap)
    $wb.Close($false)
  } finally { DongExcel $xl $xpid }
  if ($have -lt 1) { throw 'Khong ket noi duoc SAP (chua dang nhap?)' }
  if ($have -lt $nP) { Ghi "  SAP chi mo duoc $have cua so - chay $have nhom mot luc"; $nP = $have }
  Ghi ("  PR song song: " + ($keys -join ' + ') + " tren $nP cua so SAP")

  $kqRows = @(); $logs = @(); $loi = @()
  for ($s = 0; $s -lt $keys.Count; $s += $nP) {
    $jobs = @()
    for ($k = 0; $k -lt $nP -and ($s + $k) -lt $keys.Count; $k++) {
      $g = $keys[$s + $k]; $idxs = @($groups[$g])
      $sub = @($idxs | ForEach-Object { $rows[$_] })
      $work = Join-Path $KqDir ('{0}_{1}.xlsm' -f $id, ($s + $k + 1))
      $json = ConvertTo-Json -InputObject $sub -Depth 10 -Compress
      $jobs += @{ g = $g; idxs = $idxs; phien = $k; job = (Start-Job -ScriptBlock $WorkerPR -ArgumentList ([string]$cfg.fileExcel), $work, $json, $k, $luu, $AnSap) }
    }
    TienDo $id ('Dang nhap song song: ' + (($jobs | ForEach-Object { $_.g }) -join ' + '))
    foreach ($j in $jobs) {
      $raw = Receive-Job -Job $j.job -Wait -AutoRemoveJob
      $o = $null; try { $o = ([string]($raw | Select-Object -Last 1)) | ConvertFrom-Json } catch {}
      if (-not $o) { $loi += ($j.g + ': tien trinh loi'); continue }
      if ($o.loi) { $loi += ($j.g + ': ' + $o.loi) }
      $logs += ('[Cua so SAP ' + ($j.phien + 1) + ' - ' + $j.g + ']' + "`n" + [string]$o.log)
      $rr = @($o.rows)
      for ($t = 0; $t -lt $j.idxs.Count; $t++) {
        $x = $(if ($t -lt $rr.Count) { $rr[$t] } else { $null })
        $kqRows += @{ i = $j.idxs[$t]; soPR = [string]$x.soPR; tt = [string]$x.tt; item = [string]$x.item; giaSAP = [string]$x.giaSAP; slSAP = [string]$x.slSAP }
      }
    }
  }
  if ($luu) { $ok = (@($kqRows | Where-Object { $_.soPR }).Count -eq $rows.Count) } else { $ok = (@($kqRows | Where-Object { $_.tt -match 'LOI|LECH' }).Count -eq 0) }
  if ($loi.Count) { $ok = $false; $logs += ('LOI: ' + ($loi -join ' | ')) }
  Ghi ("  Xong PR song song: " + $(if ($ok) { 'thanh cong' } else { 'CHUA DAT - xem Status' }))
  return @{ tt = $(if ($ok) { 'xong' } else { 'loi' }); ketqua = @{ rows = $kqRows; po = @(); fileKetQua = $KqDir }; log = ($logs -join "`n") }
}

function ChayLenh($L) {
  $id = $L.id
  $loai = [string]$L.loai
  $luu = [bool]$L.luuThat
  $n = @($L.rows).Count
  Ghi ("Nhan lenh $id : $loai, " + $(if ($luu) { 'LUU THAT' } else { 'KHONG LUU' }) + ", $n dong")
  if (@('PR', 'PO', 'POSUA', 'MIGO', 'DNTT') -notcontains $loai) { throw "Loai lenh khong ho tro: $loai" }
  if ($n -lt 1) { throw 'Lenh khong co dong nao' }
  if ($loai -eq 'PR' -and $SoPhien -gt 1) {
    $nhom = @($L.rows | ForEach-Object { [string]$_.c[8] } | Select-Object -Unique)
    if ($nhom.Count -gt 1) {
      try { return (ChayPRSongSong $L) }
      catch {
        if ($_.Exception.Message -like 'MACRO_CU*') { Ghi ('  ' + $_.Exception.Message.Substring(9) + ' -> nhap lan luot'); $script:GhiChuLenh = $_.Exception.Message.Substring(9) }
        else { throw }
      }
    }
  }

  TienDo $id 'Mo Excel Template_PR_Tram'
  if (FileBiKhoa ([string]$cfg.fileExcel)) { throw 'File Template_PR_Tram.xlsm dang bi mo/khoa (co the do Excel an con sot lai) - dong Excel trong Task Manager roi gui lai lenh' }
  $xl = New-Object -ComObject Excel.Application
  $xpid = LayPidExcel $xl
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $wb = $null
  try {
    $wb = $xl.Workbooks.Open([string]$cfg.fileExcel)
    try { [void]$xl.Run('TramCauHinh', 0, $AnSap) } catch { Ghi '  (macro cu: chua co che do chay ngam cho SAP - import lai SAP_Automation_v23_Tram.bas ban 31/10)' }
    $ver = $xl.Run('TramPhienBan')
    if (-not ([string]$ver).StartsWith('v23')) { throw "File Excel chua co macro v23-Tram (dang la: $ver)" }
    $kqRows = @(); $kqPO = @()

    if ($loai -eq 'PR' -or $loai -eq 'PO') {
      [void]$xl.Run('TramChuanBi')
      [void](DienBang $wb 'CongNo' $L.rows 11)
      $xl.Calculate()
      if ($loai -eq 'PO') {
        if ($L.lifnr) { foreach ($p in $L.lifnr.PSObject.Properties) { if ([string]$p.Value) { [void]$xl.Run('TramDatMaNCC', [string]$p.Name, [string]$p.Value) } } }
        for ($i = 0; $i -lt $n; $i++) {
          $sp = [string]$L.rows[$i].soPR
          if ($sp) { [void]$xl.Run('TramGanPR', [int](3 + $i), $sp, [string]$L.rows[$i].item) }
        }
        $xl.Calculate()
      }
      TienDo $id ('Dang nhap ' + $loai + ' vao SAP')
      if ($loai -eq 'PR') { $mlog = $xl.Run('TramChayPR', $luu) } else { $mlog = $xl.Run('TramChayPO', $luu) }
      $wsPR = $wb.Worksheets.Item('PR')
      for ($i = 0; $i -lt $n; $i++) {
        $r = 3 + $i
        $kqRows += @{ i = $i; soPR = (Txt $wsPR.Range("L$r")); tt = (Txt $wsPR.Range("M$r")); item = (Txt $wsPR.Range("P$r")); soPO = (Txt $wsPR.Range("Q$r")); giaSAP = (Txt $wsPR.Range("R$r")); slSAP = (Txt $wsPR.Range("S$r")) }
      }
      if ($loai -eq 'PO') {
        $wsPO = $wb.Worksheets.Item('PO')
        $last = $wsPO.Cells.Item($wsPO.Rows.Count, 2).End(-4162).Row
        for ($r = 3; $r -le $last; $r++) {
          $ncc = Txt $wsPO.Range("B$r")
          if ($ncc) { $kqPO += @{ ncc = $ncc; lifnr = (Txt $wsPO.Range("C$r")); tongSAP = (Txt $wsPO.Range("G$r")); soPO = (Txt $wsPO.Range("J$r")); tt = (Txt $wsPO.Range("K$r")) } }
        }
        $ok = @($kqRows | Where-Object { $_.soPO }).Count -eq $n
      } else { $ok = @($kqRows | Where-Object { $_.soPR }).Count -eq $n }
    }
    elseif ($loai -eq 'POSUA') {
      [void]$xl.Run('TramChuanBiPOSua')
      [void](DienBang $wb 'PO_SUA' $L.rows 9)
      TienDo $id 'Dang sua PO tren SAP (ME22N)'
      $mlog = $xl.Run('TramChayPOSua', $luu)
      $ws = $wb.Worksheets.Item('PO_SUA')
      for ($i = 0; $i -lt $n; $i++) { $r = 3 + $i; $kqRows += @{ i = $i; tt = (Txt $ws.Range("J$r")); ebelp = (Txt $ws.Range("K$r")) } }
      $ok = @($kqRows | Where-Object { $_.tt -like 'DA LUU*' }).Count -eq $n
    }
    elseif ($loai -eq 'MIGO') {
      [void]$xl.Run('TramChuanBiMIGO')
      [void](DienBang $wb 'MIGO_TP' $L.rows 7)
      TienDo $id 'Dang nhap kho (MIGO) tren SAP'
      $mlog = $xl.Run('TramChayMIGO', $luu)
      $ws = $wb.Worksheets.Item('MIGO_TP')
      for ($i = 0; $i -lt $n; $i++) { $r = 3 + $i; $kqRows += @{ i = $i; soMIGO = (Txt $ws.Range("H$r")); tt = (Txt $ws.Range("I$r")) } }
      $ok = @($kqRows | Where-Object { $_.soMIGO }).Count -eq $n
    }
    elseif ($loai -eq 'DNTT') {
      [void]$xl.Run('TramChuanBiDNTT')
      [void](DienBang $wb 'DNTT' $L.rows 16)
      TienDo $id 'Dang tao file DNTT'
      $mlog = $xl.Run('TramChayDNTT')
      $ws = $wb.Worksheets.Item('DNTT')
      $files = @{}
      for ($i = 0; $i -lt $n; $i++) {
        $r = 3 + $i
        $f = Txt $ws.Range("O$r")
        $kqRows += @{ i = $i; file = $f; tt = (Txt $ws.Range("P$r")); thuHuong = (Txt $ws.Range("N$r")) }
        if ($f) { $files[$f] = 1 }
      }
      $up = @()
      foreach ($f in $files.Keys) {
        foreach ($g in @($f, ($f -replace '\.xlsx$', '.pdf'))) {
          if (Test-Path $g) {
            try {
              TienDo $id ('Chep len Drive: ' + (Split-Path $g -Leaf))
              $mime = $(if ($g -like '*.pdf') { 'application/pdf' } else { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
              $o2 = Goi @{ a = 'file'; name = (Split-Path $g -Leaf); mime = $mime; b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($g)) }
              $up += @{ path = $g; name = (Split-Path $g -Leaf); url = $o2.url }
            } catch { Ghi ('  Khong chep duoc len Drive: ' + $_.Exception.Message) }
          }
        }
      }
      $kqPO = $up
      $ok = @($kqRows | Where-Object { $_.file }).Count -eq $n
    }
    if (-not $luu) { $ok = @($kqRows | Where-Object { $_.tt -match 'LOI|LECH' }).Count -eq 0 }

    $copy = Join-Path $KqDir ($id + '.xlsm')
    try { $wb.SaveCopyAs($copy) } catch {}
    Ghi ("  Xong: " + $(if ($ok) { 'thanh cong' } else { 'CHUA DAT - xem Status' }))
    $nl = [string]$mlog; if ($script:GhiChuLenh) { $nl = 'LUU Y: ' + $script:GhiChuLenh + "`n" + $nl; $script:GhiChuLenh = '' }
    return @{ tt = $(if ($ok) { 'xong' } else { 'loi' }); ketqua = @{ rows = $kqRows; po = $kqPO; fileKetQua = $copy }; log = $nl }
  }
  finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    DongExcel $xl $xpid
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

Ghi "=== $Ver khoi dong tren may $env:COMPUTERNAME. File Excel: $($cfg.fileExcel)"
if (-not (Test-Path ([string]$cfg.fileExcel))) { Ghi 'KHONG THAY file Excel trong cau hinh. Sua TramSAP.config.json roi chay lai.'; Read-Host 'Enter de thoat'; exit 1 }
if (-not $cfg.url) { Ghi 'Chua dien url (Apps Script /exec) trong TramSAP.config.json.'; Read-Host 'Enter de thoat'; exit 1 }
try { [void](Goi @{ a = 'ping' }); Ghi 'Ket noi cong Google: OK. Dang cho lenh...' } catch { Ghi ('Khong ket noi duoc cong Google: ' + $_.Exception.Message) }

$loiLienTiep = 0
while ($true) {
  try {
    $o = Goi @{ a = 'poll' }
    $loiLienTiep = 0
    if ($o.lenh) {
      $L = $o.lenh
      try { $res = ChayLenh $L }
      catch { $res = @{ tt = 'loi'; ketqua = $null; log = ('LOI TRAM: ' + $_.Exception.Message) }; Ghi ('  LOI: ' + $_.Exception.Message) }
      for ($k = 0; $k -lt 5; $k++) {
        try { [void](Goi @{ a = 'ketqua'; id = $L.id; tt = $res.tt; ketqua = $res.ketqua; log = $res.log }); break }
        catch { Ghi ('  Gui ket qua loi, thu lai: ' + $_.Exception.Message); Start-Sleep -Seconds 10 }
      }
      continue
    }
  }
  catch {
    $loiLienTiep++
    if ($loiLienTiep -le 3 -or $loiLienTiep % 20 -eq 0) { Ghi ('Loi ket noi: ' + $_.Exception.Message) }
  }
  Start-Sleep -Seconds ([int]$cfg.chuKyGiay)
}
