<#
.SYNOPSIS
  Xuất (export) module VBA từ file Excel (.xlsm) ra file .bas/.cls trong repo.

.DESCRIPTION
  Dùng khi đã sửa macro trực tiếp trong Excel (Alt+F11) và muốn đưa thay đổi về repo để commit.
  Chỉ xuất các module có trong sap-station/vba/modules.json (cả phần simulator nếu file Excel có),
  ghi đè đúng file nguồn tương ứng. Xem thay đổi bằng "git diff" trước khi commit.

.EXAMPLE
  .\tools\vba\Export-Vba.ps1 -Workbook C:\TramSAP\Template_PR_Tram.xlsm
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Workbook,
  [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
)
$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path $Workbook).Path
$manifest = Get-Content (Join-Path $RepoRoot 'sap-station\vba\modules.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$map = @{}
foreach ($prop in $manifest.PSObject.Properties) {
  if ($prop.Name -like '_*') { continue }
  foreach ($m in @($prop.Value.modules) + @($prop.Value.simulator)) {
    $p = Join-Path $RepoRoot ($m -replace '/', '\')
    $name = [IO.Path]::GetFileNameWithoutExtension($p)
    if (Test-Path $p) {
      $line = Get-Content $p -TotalCount 20 | Where-Object { $_ -match '^Attribute VB_Name = "(.+)"' } | Select-Object -First 1
      if ($line -match '"(.+)"') { $name = $Matches[1] }
    }
    $map[$name] = $p
  }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Workbook, 0, $true)
  try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count }
  catch { throw 'Excel chặn truy cập VBA. Bật "Trust access to the VBA project object model" trong Trust Center rồi chạy lại.' }
  $n = 0
  foreach ($c in $vbp.VBComponents) {
    if ($map.ContainsKey($c.Name)) { $c.Export($map[$c.Name]); Write-Host "  xuất $($c.Name)  =>  $($map[$c.Name])"; $n++ }
  }
  $wb.Close($false)
  Write-Host "Xong: $n module. Chạy 'git diff' để xem thay đổi."
}
finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
