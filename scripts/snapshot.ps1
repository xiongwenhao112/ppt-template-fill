# snapshot.ps1 — 用 PowerPoint COM 把 pptx 每页导出为 PNG。
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File snapshot.ps1 -Path <pptx> -OutDir <dir> [-Width 1440]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$Path,
    [Parameter(Mandatory = $true)] [string]$OutDir,
    [int]$Width = 1440
)

$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path $Path).Path
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    # msoTrue=-1, msoFalse=0; ReadOnly, Untitled, WithWindow=false
    $pres = $pp.Presentations.Open($Path, -1, 0, 0)
    $ratio = $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth
    $h = [int]($Width * $ratio)
    $count = $pres.Slides.Count
    for ($i = 1; $i -le $count; $i++) {
        $out = Join-Path $OutDir ("slide-{0:d2}.png" -f $i)
        $pres.Slides.Item($i).Export($out, 'PNG', $Width, $h)
    }
    Write-Output "exported $count slides -> $OutDir"
}
finally {
    if ($pres) { $pres.Close() }
    if ($pp) { $pp.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
