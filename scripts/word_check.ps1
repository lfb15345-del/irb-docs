# Open a .docx in hidden Word (read-only) and report tracked-change and comment counts.
# Use it to confirm Word accepts a generated file when PDF export hangs (e.g. redirected printers).
#   powershell -NoProfile -ExecutionPolicy Bypass -File word_check.ps1 -Path "C:\...\file.docx" [-TimeoutSec 60]
# Prints: OK revisions=N comments=N replies=N resolved=N  /  ERROR ...  /  TIMEOUT
# Only the Word started by this script is stopped; Word windows the user opened are left alone.
# ASCII only (Windows PowerShell 5.1 reads BOM-less UTF-8 as ANSI).
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$TimeoutSec = 60
)
function Get-AutomationWord {
    @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
        Where-Object { $_.CommandLine -match '/Automation' } |
        ForEach-Object { [int]$_.ProcessId })
}
$before = Get-AutomationWord
$job = Start-Job -ScriptBlock {
    param($p)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        $doc = $word.Documents.Open($p, $false, $true, $false)
        $rev = $doc.Revisions.Count
        $com = $doc.Comments.Count
        $replies = 0
        $done = 0
        foreach ($c in $doc.Comments) {
            if ($c.Ancestor -ne $null) { $replies++ }
            if ($c.Done) { $done++ }
        }
        $doc.Close([ref]0)
        "OK revisions=$rev comments=$com replies=$replies resolved=$done"
    } catch {
        "ERROR " + $_.Exception.Message
    } finally {
        try { $word.Quit([ref]0) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
} -ArgumentList $Path
if (Wait-Job $job -Timeout $TimeoutSec) {
    Receive-Job $job
} else {
    "TIMEOUT"
}
Remove-Job $job -Force
Start-Sleep -Milliseconds 800
foreach ($id in (Get-AutomationWord)) {
    if ($before -notcontains $id) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
