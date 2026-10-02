# Export .docx to PDF with Word (COM). With -ToDocx, convert legacy .doc forms to .docx instead.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File docx_to_pdf.ps1 -Files "a.docx|b.docx" -OutDir "C:\...\pdf"
#   powershell -NoProfile -ExecutionPolicy Bypass -File docx_to_pdf.ps1 -Files "form.doc" -OutDir "C:\...\forms" -ToDocx
#
# With -File, PowerShell cannot pass an array, so give several files as one "|"-separated string.
#
# The Word work runs in a background job. If it does not finish within -TimeoutSec, the script stops only the
# automation Word process it started (command line "/Automation -Embedding", not present before it began),
# prints TIMEOUT and exits with code 2. It never touches Word windows the user opened.
#
# Why exports hang: ExportAsFixedFormat asks the default printer for page metrics. Over Remote Desktop the
# default printer is often a redirected client printer (port TS001, name "... (redirect N)"). When the client
# printer is asleep or the RDP client is gone, the call blocks forever with no dialog. Changing the default
# printer is a user setting: ask the user (or have them reconnect / wake the printer) instead of changing it.
#
# Keep this file ASCII only. Windows PowerShell 5.1 reads BOM-less UTF-8 as the ANSI code page, and
# Japanese text then breaks the parser.
param(
    [Parameter(Mandatory = $true)][string[]]$Files,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [switch]$ToDocx,
    [int]$TimeoutSec = 120
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $OutDir | Out-Null
$names = @($Files | ForEach-Object { $_ -split '\|' } | Where-Object { $_.Trim() })
$full = @($names | ForEach-Object { (Resolve-Path -LiteralPath $_.Trim()).Path })

function Get-AutomationWord {
    @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
        Where-Object { $_.CommandLine -match '/Automation -Embedding' } |
        ForEach-Object { [int]$_.ProcessId })
}
$before = Get-AutomationWord

# Paths go in as one "|"-joined string: arrays passed to Start-Job arrive nested. "|" cannot appear in a path.
$job = Start-Job -ArgumentList ($full -join "|"), $OutDir, ([bool]$ToDocx) -ScriptBlock {
    param($joined, $outDir, $toDocx)
    $paths = $joined -split '\|'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {
        foreach ($p in $paths) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($p)
            $doc = $word.Documents.Open($p, $false, $true)
            try {
                if ($toDocx) {
                    $out = Join-Path $outDir ($base + ".docx")
                    $doc.SaveAs2($out, 16)
                }
                else {
                    $out = Join-Path $outDir ($base + ".pdf")
                    $doc.ExportAsFixedFormat($out, 17)
                }
                Write-Output "saved $out"
            }
            finally {
                $doc.Close($false)
            }
        }
    }
    finally {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

$done = Wait-Job $job -Timeout $TimeoutSec
if (-not $done) {
    # Kill our automation Word first, then the job's host process (Stop-Job alone can block for minutes).
    $mine = Get-AutomationWord | Where-Object { $before -notcontains $_ }
    foreach ($id in $mine) {
        Stop-Process -Id $id -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-Output "stopped automation Word $id"
    }
    Get-CimInstance Win32_Process -Filter "ParentProcessId=$PID" |
        Where-Object { $_.CommandLine -match '-s -NoLogo' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -Confirm:$false -ErrorAction SilentlyContinue }
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    $printer = Get-CimInstance Win32_Printer | Where-Object { $_.Default } | Select-Object -First 1
    $msg = "TIMEOUT after {0}s. Word automation stopped responding. Default printer: {1} (port {2}). " +
        "Check for leftover Word processes started with /Automation, the default printer (Remote Desktop " +
        "redirected printers can block exports), and add-ins. Retry once; if it still hangs, ask the user."
    Write-Output ($msg -f $TimeoutSec, $printer.Name, $printer.PortName)
    exit 2
}
Receive-Job $job
Remove-Job $job -Force
