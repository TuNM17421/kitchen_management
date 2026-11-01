<#
.SYNOPSIS
  Cài hoặc cập nhật Trạm SAP vào C:\TramSAP (hoặc thư mục khác).

.DESCRIPTION
  1. Chép TramSAP.ps1, ChayTramSAP.bat, TramSAP.config.example.json, VERSION và các module VBA vào thư mục trạm.
  2. Bỏ đánh dấu "tải từ Internet" (Unblock-File) để chạy được với ExecutionPolicy RemoteSigned.
  3. Tạo TramSAP.config.json từ file mẫu nếu chưa có (không ghi đè file đang dùng).
  4. Nếu có -TemplateXlsm: chép thành Template_PR_Tram.xlsm (nếu chưa có) rồi nạp 3 module VBA bằng tools\vba\Import-Vba.ps1.

  Chạy từ thư mục gốc repo hoặc thư mục gói phát hành đã giải nén:
    powershell -NoProfile -ExecutionPolicy Bypass -File sap-station\install.ps1 -TemplateXlsm "D:\...\Template_PR_v3.xlsm"
  (Bypass chỉ dùng cho lần cài này; trạm chạy hằng ngày bằng ChayTramSAP.bat với RemoteSigned.)

.PARAMETER Dest
  Thư mục trạm. Mặc định C:\TramSAP.
.PARAMETER TemplateXlsm
  File Template_PR đang dùng, để tạo bản riêng cho trạm. Bỏ qua nếu đã có Template_PR_Tram.xlsm.
.PARAMETER WithSimulator
  Nạp thêm bộ giả lập SAP (sap-simulator) để chạy thử không cần SAP thật.
.PARAMETER SkipVba
  Không nạp VBA (khi Excel chưa bật "Trust access to the VBA project object model").
#>
[CmdletBinding()]
param(
  [string]$Dest = 'C:\TramSAP',
  [string]$TemplateXlsm,
  [switch]$WithSimulator,
  [switch]$SkipVba
)
$ErrorActionPreference = 'Stop'
$Src = $PSScriptRoot                                   # ...\sap-station
$Root = (Resolve-Path (Join-Path $Src '..')).Path      # thư mục gốc repo / gói

New-Item -ItemType Directory -Force -Path $Dest, (Join-Path $Dest 'vba') | Out-Null
foreach ($f in 'TramSAP.ps1', 'ChayTramSAP.bat', 'TramSAP.config.example.json') {
  Copy-Item (Join-Path $Src $f) $Dest -Force
}
Copy-Item (Join-Path $Src 'vba\*.bas') (Join-Path $Dest 'vba') -Force
$verFile = Join-Path $Root 'VERSION'
if (Test-Path $verFile) { Copy-Item $verFile $Dest -Force }
Get-ChildItem $Dest -Recurse -File | Unblock-File
Write-Host "Đã chép file trạm vào $Dest"

$cfg = Join-Path $Dest 'TramSAP.config.json'
if (-not (Test-Path $cfg)) {
  Copy-Item (Join-Path $Dest 'TramSAP.config.example.json') $cfg
  Write-Host "Đã tạo $cfg từ file mẫu: mở bằng Notepad, điền url (Apps Script /exec) và token."
} else {
  Write-Host "Giữ nguyên $cfg đang dùng."
}

$tram = Join-Path $Dest 'Template_PR_Tram.xlsm'
if (-not (Test-Path $tram)) {
  if (-not $TemplateXlsm) {
    Write-Warning "Chưa có $tram. Chạy lại với -TemplateXlsm <file Template_PR đang dùng>."
    return
  }
  Copy-Item $TemplateXlsm $tram
  Write-Host "Đã tạo $tram (bản riêng cho trạm)."
}
if ($SkipVba) { Write-Host 'Bỏ qua nạp VBA (-SkipVba).'; return }

$imp = Join-Path $Root 'tools\vba\Import-Vba.ps1'
& $imp -Workbook $tram -RepoRoot $Root -WithSimulator:$WithSimulator
Write-Host ''
Write-Host 'Xong. Còn lại: điền TramSAP.config.json, đăng nhập SAP, nháy đúp ChayTramSAP.bat.'
