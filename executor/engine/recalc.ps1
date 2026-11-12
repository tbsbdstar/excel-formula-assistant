param([Parameter(Mandatory = $true)][string]$Path)
# Recalculate a workbook with local Excel or WPS (COM), then save it.
# Output: "OK <ProgId>" | "NOAPP" | "ERR <message>"
$ErrorActionPreference = 'Stop'
$app = $null
$used = ''
foreach ($p in @('Excel.Application', 'KET.Application', 'ET.Application')) {
    try { $app = New-Object -ComObject $p; $used = $p; break } catch { }
}
if ($null -eq $app) { Write-Output 'NOAPP'; exit 2 }
$wb = $null
try {
    try { $app.Visible = $false } catch { }
    try { $app.DisplayAlerts = $false } catch { }
    try { $app.AskToUpdateLinks = $false } catch { }
    try { $app.ScreenUpdating = $false } catch { }
    $wb = $app.Workbooks.Open($Path, 0, $false)
    try { $app.CalculateFull() } catch { $app.Calculate() }
    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Write-Output ('OK ' + $used)
    $code = 0
} catch {
    Write-Output ('ERR ' + $_.Exception.Message)
    $code = 1
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    try { $app.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit $code
