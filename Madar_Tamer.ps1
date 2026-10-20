<#
Madar_Tamer.ps1  (v4)

Finds the newest report email (Inbox or the archive folder), saves its attachment
under a fixed name, and replaces the matching file in the OneDrive-synced SharePoint folder.
Skips the copy when the file is already identical, so OneDrive doesn't re-upload for nothing.

Requirements:
  - Classic Outlook desktop app, open and signed in (the "new Outlook" does not support this).

Run manually (keeps the window open so you can read the output):
  powershell -ExecutionPolicy Bypass -File E:\Projects\Madar_Tamer.ps1 -Pause
#>
param([switch]$Pause)

# ---------- Settings ----------
$Destination = 'C:\Users\Test\OneDrive - Procter and Gamble\Saudi PD - Transportation\Last Mile Operations Officer\Source\Transporters Reports'

# Only attachments with these extensions are considered (skips signature images etc.)
$AllowedExtensions = '.xlsx', '.xlsm', '.xlsb', '.xls', '.csv'

# How many of the newest emails to scan in each place
$ScanLimit = 150

# Seconds to wait for Outlook to download new mail before reading it
# (longer when Outlook was closed and had to be started by this script)
$SyncWaitSeconds    = 45
$StartupWaitSeconds = 120

# Folder = Outlook archive folder name | NewName = saved file name | SubjectLike = subject filter (* = any text)
$Jobs = @(
    @{ Folder = 'Madar Control Tower'; NewName = 'madar';                     SubjectLike = '*Control Data Update*' }
    @{ Folder = 'Tamer Control Tower'; NewName = 'TAMER Control Tower - P&G'; SubjectLike = '*Control Tower X P&G*' }
)

$LogFile = Join-Path $PSScriptRoot 'Madar_Tamer.log'
# ------------------------------

function Write-Log([string]$Message) {
    $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

# Returns EVERY folder with this name (a mailbox and its archive can both have one)
function Find-OutlookFolders($Parent, [string]$Name) {
    foreach ($child in $Parent.Folders) {
        if ($child.Name -eq $Name) { $child }
        Find-OutlookFolders $child $Name
    }
}

# Newest email in a folder whose subject matches and that has a real file attachment
function Get-LatestReport($Folder, [string]$SubjectLike, [string[]]$Extensions, [int]$Limit) {
    $items = $Folder.Items
    $items.Sort('[ReceivedTime]', $true)
    $n = 0
    foreach ($mail in $items) {
        $n++
        if ($n -gt $Limit) { break }
        if ($mail.Class -ne 43) { continue }               # 43 = regular email
        if ($mail.Subject -notlike $SubjectLike) { continue }
        foreach ($att in $mail.Attachments) {
            $ext = [IO.Path]::GetExtension($att.FileName).ToLower()
            if ($att.Type -eq 1 -and $Extensions -contains $ext) {
                return [pscustomobject]@{
                    Mail = $mail; Attachment = $att; Extension = $ext; Received = $mail.ReceivedTime
                }
            }
        }
    }
    return $null
}

$wasRunning = [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)

try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
} catch {
    Write-Log "Could not connect to Outlook: $_"
    if ($Pause) { Read-Host 'Press Enter to close' }
    exit 1
}

# Ask Outlook to download new mail first, otherwise we only see its old local copy
try {
    if ($namespace.Offline) { Write-Log 'WARNING: Outlook is in Offline mode - new emails will NOT download' }
    $namespace.SendAndReceive($false)
    $wait = $SyncWaitSeconds
    $state = 'already running'
    if (-not $wasRunning) { $wait = $StartupWaitSeconds; $state = 'was closed, started by script' }
    Write-Log "Outlook $state - syncing, waiting $wait s..."
    Start-Sleep -Seconds $wait
} catch {
    Write-Log "Could not trigger Outlook sync: $_"
}

foreach ($job in $Jobs) {
    $temp = $null
    Write-Log "---- $($job.Folder) -> $($job.NewName)"
    try {
        # Find every folder with this name, then use the one holding the newest email
        $candidates = @(foreach ($store in $namespace.Folders) { Find-OutlookFolders $store $job.Folder })
        $sources = @()
        if ($candidates.Count -gt 0) {
            $archive = $candidates[0]
            $newest  = [datetime]::MinValue
            foreach ($c in $candidates) {
                $its = $c.Items
                $its.Sort('[ReceivedTime]', $true)
                $first = $its.GetFirst()
                $when = [datetime]::MinValue
                if ($first -and $first.ReceivedTime) { $when = $first.ReceivedTime }
                if ($when -gt $newest) { $newest = $when; $archive = $c }
            }
            $sources += $archive
        } else {
            Write-Log "Archive folder not found: $($job.Folder) (checking Inbox only)"
        }
        $sources += $namespace.GetDefaultFolder(6)         # 6 = Inbox

        # Look in each place and keep the newest match overall
        $best = $null
        foreach ($src in $sources) {
            $hit = Get-LatestReport $src $job.SubjectLike $AllowedExtensions $ScanLimit
            if ($hit) {
                Write-Log "  $($src.FolderPath): newest match $($hit.Received) | $($hit.Mail.Subject) | $($hit.Attachment.FileName)"
                if (-not $best -or $hit.Received -gt $best.Received) { $best = $hit }
            } else {
                Write-Log "  $($src.FolderPath): no match in newest $ScanLimit emails. Newest emails there:"
                $peek = $src.Items
                $peek.Sort('[ReceivedTime]', $true)
                $k = 0
                foreach ($m in $peek) {
                    if ($k -ge 3) { break }
                    if ($m.Class -eq 43) { $k++; Write-Log "      $($m.ReceivedTime) | $($m.Subject)" }
                }
            }
        }

        if (-not $best) {
            Write-Log "Nothing to save for $($job.Folder)"
            continue
        }

        # Save to temp first so a failed download never leaves a half-written file
        $temp   = Join-Path $env:TEMP ([guid]::NewGuid().ToString() + $best.Extension)
        $best.Attachment.SaveAsFile($temp)
        $target = Join-Path $Destination ($job.NewName + $best.Extension)

        $same = $false
        try {
            if (Test-Path -LiteralPath $target) {
                $same = (Get-FileHash -LiteralPath $temp).Hash -eq (Get-FileHash -LiteralPath $target).Hash
            }
        } catch { $same = $false }

        if ($same) {
            Write-Log "SKIP  already up to date (email received $($best.Received)): $target"
            Remove-Item -LiteralPath $temp -Force
            continue
        }

        Move-Item -LiteralPath $temp -Destination $target -Force
        Write-Log "OK  saved '$($best.Attachment.FileName)' (email received $($best.Received)) -> $target"
    } catch {
        Write-Log "FAILED [$($job.Folder)]: $_"
        if ($temp -and (Test-Path -LiteralPath $temp)) { Remove-Item -LiteralPath $temp -Force }
    }
}

if ($Pause) { Read-Host 'Done - press Enter to close' }
