<#
.SYNOPSIS
  Nạp (import) các module VBA từ file .bas/.cls trong repo vào một file Excel (.xlsm).

.DESCRIPTION
  File .bas/.cls trong repo là mã nguồn gốc; file .xlsm chỉ là nơi chạy. Script này thay thế
  module cùng tên trong file Excel bằng bản trong repo, rồi lưu file.
  Danh sách module lấy từ sap-station/vba/modules.json (theo tên file Excel) hoặc tham số -Modules.

  Cần: Excel trên Windows, và bật "Trust access to the VBA project object model"
  (File > Options > Trust Center > Trust Center Settings > Macro Settings).

.EXAMPLE
  .\tools\vba\Import-Vba.ps1 -Workbook C:\TramSAP\Template_PR_Tram.xlsm
.EXAMPLE
  .\tools\vba\Import-Vba.ps1 -Workbook C:\TramSAP\Template_PR_Tram.xlsm -WithSimulator
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Workbook,
  [string[]]$Modules,
  [switch]$WithSimulator,
  [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
)
$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path $Workbook).Path

if (-not $Modules) {
  $manifest = Get-Content (Join-Path $RepoRoot 'sap-station\vba\modules.json') -Raw -Encoding UTF8 | ConvertFrom-Json
  $entry = $manifest.PSObject.Properties | Where-Object { $_.Name -eq (Split-Path $Workbook -Leaf) } | Select-Object -First 1
  if (-not $entry) { $entry = $manifest.PSObject.Properties | Where-Object { $_.Name -eq 'Template_PR_Tram.xlsm' } }
  $Modules = @($entry.Value.modules)
  if ($WithSimulator) { $Modules += @($entry.Value.simulator) }
}
$files = foreach ($m in $Modules) {
  $p = if ([IO.Path]::IsPathRooted($m)) { $m } else { Join-Path $RepoRoot ($m -replace '/', '\') }
  if (-not (Test-Path $p)) { throw "Không thấy file module: $p" }
  (Resolve-Path $p).Path
}

function Get-VbName([string]$path) {
  $line = Get-Content $path -TotalCount 20 | Where-Object { $_ -match '^Attribute VB_Name = "(.+)"' } | Select-Object -First 1
  if ($line -match '"(.+)"') { return $Matches[1] }
  return [IO.Path]::GetFileNameWithoutExtension($path)
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Workbook)
  try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count }
  catch { throw 'Excel chặn truy cập VBA. Bật "Trust access to the VBA project object model" trong Trust Center rồi chạy lại.' }
  foreach ($f in $files) {
    $name = Get-VbName $f
    $old = $null; foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $name) { $old = $c } }
    if ($old) { $vbp.VBComponents.Remove($old); Write-Host "  - gỡ $name" }
    $null = $vbp.VBComponents.Import($f)
    Write-Host "  + nạp $name  <=  $f"
  }
  $wb.Save()
  try { $ver = [string]$xl.Run("'" + $wb.Name + "'!TramPhienBan") } catch { $ver = '(không gọi được TramPhienBan)' }
  Write-Host "Xong. PhienBan: $ver"
  $wb.Close($true)
}
finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
