# ============================================================
# G NEXUS - Google Drive Migration Engine (backup engine + text menu)
# Destination: this folder, or "Destination" in settings.json
#
# Start it with Backup.bat (double-click) or:
#   powershell -ExecutionPolicy Bypass -File .\backup-drive.ps1
#
# MENU
#   <n>   download account n (signs in first if needed), then verify it
#   V<n>  verify account n only: byte check against Drive + duplicates report
#         (several files with the same name are all saved as "name (dup N).ext")
#   H<n>  write a SHA-256 checksum file of account n's local backup
#   J<n>  remove Mac junk files (._name, .DS_Store, __MACOSX) from account n's local backup
#   G<n>  back up account n's Gmail with Got Your Back (GYB)
#   C<n>  open Google's Contacts + Calendar export for account n
#   A     download + verify every AUTHENTICATED account
#   I<n>  show the size on Drive of each part (My Drive, Shared with me, Trash, shared drives)
#   X     cross-check "Shared with me": files owned by these accounts vs. shared from outside
#   S     calculate the Drive sizes of every authenticated account (MyDrive / Shared / Total columns)
#   W     Shared with me on/off for all accounts (when on, each download asks Y/n with its size)
#   L     toggle overwrite mode (re-download every file, even unchanged ones)
#         Without L, re-downloading an account asks: S = skip existing, O = overwrite all
#   R     check accounts now (sign-in + sizes)   Q  quit
#
# Start with -Check to run the checking before the menu; without it the table
# opens at once with the last saved results (marked *).
#
# Per account the Drive goes into its GT_ / INV_ folder:
#   Mi_unidad\               - My Drive
#   Compartidos_conmigo\     - "Shared with me"
#   Papelera\                - files in the trash
#   Unidades_compartidas\X\  - every shared drive the account can see
#
# VERIFY = "rclone check": every regular file is compared with Drive by
# size and MD5 checksum (Drive's own hash). Google Docs/Sheets/Slides have
# no hash on Drive (they are exported), so they are verified by file count.
# Results and problem lists go to _logs\; the menu table shows the summary.
#
# It only copies: nothing is ever deleted or changed on Drive.
# ============================================================

param(
    # -Check: check every account's sign-in and Drive sizes before showing the menu.
    # Without it the table opens at once with the last saved results (marked *).
    [switch]$Check,
    # Used by the Windows app (backup-app.ps1): run one action and close, no menu.
    #   -Action Export | T | R | N | S | X | A | Download | SignIn | Sizes | Verify | Sha | Contacts
    #   -Select "1,3"  -OverwriteAll  -NoShared
    [string]$Action = "",
    # (names must not match script variables: PowerShell names ignore case, so a
    #  parameter called Accounts would overwrite the $accounts list)
    [string]$Select = "",     # account numbers, comma separated: "1,3,5"
    [switch]$OverwriteAll,
    [switch]$NoShared,
    [string]$SharedMode = "", # E = only files from outside these accounts, A = all, N = skip
    [switch]$NoVerify,        # download without the byte check afterwards
    [switch]$KeepJunk,        # also download Mac system files (._*, .DS_Store); skipped by default
    [switch]$CleanJunk,       # after each download, delete Mac junk files already in the account folder
    [switch]$NoShortcuts,     # skip Drive shortcuts (their files are not downloaded a second time)
    [switch]$OwnedOnly,       # My Drive: only files owned by the account itself
    [string]$Parts = "",      # which parts: "My,Shared,Drives,Trash" (empty = default from settings)
    [string]$DataDir = "",    # backup folder; its _gnexus folder holds settings.json, accounts.csv, session.json
    [switch]$Resume           # continue a paused download: My Drive folders already finished are not done again
)
# App mode: started by G Nexus in the background. No questions are asked (the app passes
# the answers) and everything printed goes to _logs\app_output.log for the app to show.
$script:appMode = [bool]($Action -and $Action -ne "Export")

# ---------- settings ----------
# Everything that changes per organization lives in two files next to this script:
#   accounts.csv   Name, Email, Password, Folder (+ optional Remote, Legacy), one row per account
#   settings.json  Destination, ClientId, ClientSecret, IncludeTrash, Transfers, Checkers, ExportFormats
# To back up another organization: copy the folder, edit those two files.
$scriptDir = $PSScriptRoot

function Stop-WithError($msg) {
    Write-Host "ERROR: $msg" -ForegroundColor Red
    # The Windows app shows this file when the engine can't start
    try { Set-Content -LiteralPath (Join-Path $scriptDir "_app_error.txt") -Value $msg -Encoding UTF8 } catch { }
    if (-not $Action) { $null = Read-Host "Press Enter to exit" }
    exit 1
}
Remove-Item -LiteralPath (Join-Path $scriptDir "_app_error.txt") -Force -ErrorAction SilentlyContinue

# Defaults (settings.json overrides them)
$destination   = $scriptDir
$exportFormats = "docx,xlsx,pptx,pdf,odt,ods,odp,csv,tsv,rtf,txt,html,epub,md,json,svg"
$retries       = 2
$transfers     = 8     # files downloaded at the same time (lower to 4 if Google says "Quota exceeded")
$checkers      = 8
$includeTrash  = $false
# Your own Google "OAuth client" (see HOW TO at the end of this file). Empty = rclone's shared one,
# which is overloaded. After changing it, every account must sign in again (table: OLD SIGN-IN).
$clientId      = ""
$clientSecret  = ""
$bwLimit       = ""    # download speed limit, e.g. "20M" (20 MB/s); empty = no limit
$skipJunk      = $true # skip Mac system files: ._name (AppleDouble), .DS_Store, __MACOSX, ...

# Data folder: the backup location the user chose. Its _gnexus folder holds the settings and
# accounts, so the app folder only has the app. The app passes it (-DataDir); the text menu
# reads it from %APPDATA%\G Nexus\location.json. Older copies kept everything next to the scripts.
$cfgDir = $scriptDir
if (-not $DataDir) {
    try { $DataDir = "$((Get-Content -LiteralPath (Join-Path $env:APPDATA 'G Nexus\location.json') -Raw -Encoding UTF8 | ConvertFrom-Json).DataFolder)" } catch { }
}
if ($DataDir) {
    $cfgDir = Join-Path $DataDir "_gnexus"
    New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
}
# Settings and accounts of older copies (next to the scripts) are copied over once
if ($cfgDir -ne $scriptDir) {
    foreach ($f in @("settings.json", "accounts.csv")) {
        $oldF = Join-Path $scriptDir $f
        if (-not (Test-Path -LiteralPath (Join-Path $cfgDir $f)) -and (Test-Path -LiteralPath $oldF)) {
            Copy-Item -LiteralPath $oldF -Destination (Join-Path $cfgDir $f)
        }
    }
}
$settingsFile = Join-Path $cfgDir "settings.json"
$settingsExample = Join-Path $scriptDir "settings.example.json"
if (-not (Test-Path -LiteralPath $settingsFile) -and (Test-Path -LiteralPath $settingsExample)) {
    Copy-Item -LiteralPath $settingsExample -Destination $settingsFile
}
if (Test-Path -LiteralPath $settingsFile) {
    try { $cfg = Get-Content -LiteralPath $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { Stop-WithError "settings.json is not valid JSON: $($_.Exception.Message)" }
    if ($cfg.Destination)   { $destination   = "$($cfg.Destination)" }
    if ($DataDir)           { $destination   = $DataDir }
    if ($cfg.ExportFormats) { $exportFormats = "$($cfg.ExportFormats)" }
    if ($cfg.Transfers)     { $transfers     = [int]$cfg.Transfers }
    if ($cfg.Checkers)      { $checkers      = [int]$cfg.Checkers }
    if ($null -ne $cfg.IncludeTrash) { $includeTrash = [bool]$cfg.IncludeTrash }
    if ($cfg.ClientId)      { $clientId      = "$($cfg.ClientId)" }
    if ($cfg.ClientSecret)  { $clientSecret  = "$($cfg.ClientSecret)" }
    if ($cfg.BwLimit)       { $bwLimit       = "$($cfg.BwLimit)".Trim() }
    if ($cfg.Retries)       { $retries       = [int]$cfg.Retries }
    if ($null -ne $cfg.SkipJunk) { $skipJunk = [bool]$cfg.SkipJunk }
    if ($cfg.MoreLocations) { $script:moreLocations = @($cfg.MoreLocations | ForEach-Object { "$_".Trim() } | Where-Object { $_ }) }
}
# More places to save to when the first disk is full (My Drive goes folder by folder:
# each folder is saved where it fits). Every place gets the same <account>\Mi_unidad layout.
if (-not $script:moreLocations) { $script:moreLocations = @() }
if ($KeepJunk) { $skipJunk = $false }
# Mac system files copied to Drive from a Mac: no user data, only Finder metadata.
# Skipped in downloads, verify and sizes so the numbers stay comparable.
$script:junkArgs = @()
if ($skipJunk) {
    foreach ($pat in @("._*", ".DS_Store", "__MACOSX/**", ".AppleDouble/**", ".Spotlight-V100/**",
                       ".Trashes/**", ".fseventsd/**", ".TemporaryItems/**", ".DocumentRevisions-V100/**")) {
        $script:junkArgs += @("--exclude", $pat)
    }
}

# Parts to work on (the app's checkboxes): My Drive, Shared with me, Shared drives, Trash
$script:wantMy = $true; $script:wantShared = $true; $script:wantDrives = $true
if ($Parts) {
    $pl = @("$Parts" -split '[,; ]+' | ForEach-Object { $_.ToLower() })
    $script:wantMy     = $pl -contains "my"
    $script:wantShared = $pl -contains "shared"
    $script:wantDrives = $pl -contains "drives"
    $includeTrash      = $pl -contains "trash"
}

# Accounts. Password = reminder for the Google sign-in (rclone does not use it).
# Folder / Remote can be left empty: they are made from the Name (or the email).
# Legacy = remote name used by older versions of this script (reused if it exists).
$accountsFile = Join-Path $cfgDir "accounts.csv"
# First run from a fresh copy (e.g. from GitHub): start from the example file
$accountsExample = Join-Path $scriptDir "accounts.example.csv"
if (-not (Test-Path -LiteralPath $accountsFile) -and (Test-Path -LiteralPath $accountsExample)) {
    Copy-Item -LiteralPath $accountsExample -Destination $accountsFile
}
if (-not (Test-Path -LiteralPath $accountsFile)) { Stop-WithError "accounts.csv not found in $cfgDir" }
$accounts = @()
foreach ($r in @(Import-Csv -LiteralPath $accountsFile -Encoding UTF8)) {
    $email = "$($r.Email)".Trim()
    if (-not $email) { continue }
    $name = "$($r.Name)".Trim()
    if (-not $name) { $name = "$($r.Company)".Trim() }
    if (-not $name) { $name = $email.Split("@")[0] }
    $folder = "$($r.Folder)".Trim()
    if (-not $folder) {
        $folder = (($name.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '') -replace '[^A-Za-z0-9]+', '_').Trim('_')
        if (-not $folder) { $folder = ($email.Split("@")[0]) -replace '[^A-Za-z0-9]', '' }
    }
    $remote = "$($r.Remote)".Trim()
    if (-not $remote) { $remote = ($folder -replace '[^A-Za-z0-9_]', '').ToLower() }
    $accounts += @{ Name = $name; Company = $name; Remote = $remote; Legacy = "$($r.Legacy)".Trim(); Email = $email; Password = "$($r.Password)"; Folder = $folder }
}
if ($accounts.Count -eq 0) { Stop-WithError "accounts.csv has no accounts" }

# ---------- checks ----------
if (-not (Get-Command rclone -ErrorAction SilentlyContinue)) {
    Stop-WithError "rclone was not found in PATH. Install it with: winget install Rclone.Rclone"
}
$destRoot = [IO.Path]::GetPathRoot($destination)
if (-not (Test-Path -LiteralPath $destRoot)) { Stop-WithError "Drive $destRoot is not accessible." }
New-Item -ItemType Directory -Force -Path $destination | Out-Null

$logDir = Join-Path $destination "_logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

# ---------- error log ----------
# Every PowerShell error goes to _logs\errors.log (with where it happened), so a
# problem can be sent for support with the app's "Report for support" button.
$script:errSaved = 0
function Save-ErrorLog($note = "") {
    try {
        $file = Join-Path $logDir "errors.log"
        $new = $Error.Count - $script:errSaved
        if ($new -le 0 -and -not $note) { return }
        $lines = @()
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        if ($note) { $lines += "$stamp  [$Action] $note" }
        for ($k = [math]::Min($new, 50) - 1; $k -ge 0; $k--) {
            $e = $Error[$k]
            $lines += "$stamp  [$Action] $($e.Exception.Message)"
            if ($e.InvocationInfo) { $lines += "      at line $($e.InvocationInfo.ScriptLineNumber): $("$($e.InvocationInfo.Line)".Trim())" }
            if ($e.ScriptStackTrace) { $lines += ("      " + ($e.ScriptStackTrace -replace "`r?`n", "`n      ")) }
        }
        Add-Content -LiteralPath $file -Value $lines -Encoding UTF8
        $script:errSaved = $Error.Count
    } catch { }
}
trap {
    Save-ErrorLog "STOPPED BY ERROR: $($_.Exception.Message)"
    Write-Host "ERROR: $($_.Exception.Message)  (details in _logs\errors.log)" -ForegroundColor Red
    continue
}

# Never run two backup actions at once (they would compete and slow each other down)
if ($script:appMode) {
    $others = @()
    try {
        $others = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
            Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -match 'backup-drive\.ps1' -and $_.CommandLine -match '-Action' -and $_.CommandLine -notmatch 'Export' })
    } catch { }
    if ($others.Count -gt 0) {
        Add-Content -LiteralPath (Join-Path $logDir "app_output.log") -Value "ANOTHER BACKUP ACTION IS ALREADY RUNNING (process $($others[0].ProcessId)). Close G Nexus, open it again and answer Yes to stop it." -Encoding UTF8
        exit 2
    }
}

if ($script:appMode) {
    $script:appOut = Join-Path $logDir "app_output.log"
    Set-Content -LiteralPath $script:appOut -Value "" -Encoding UTF8
    Remove-Item -LiteralPath (Join-Path $logDir "skip.flag") -Force -ErrorAction SilentlyContinue
    # No console to look at: send every printed line to the file the app shows
    function Write-Host {
        param([Parameter(Position = 0, ValueFromPipeline = $true)] $Object, [switch]$NoNewline, $ForegroundColor, $BackgroundColor, $Separator)
        process {
            # Also to the console, for when the app shows it ("Show console")
            $hp = @{ Object = $Object; NoNewline = $NoNewline }
            if ($ForegroundColor) { $hp.ForegroundColor = $ForegroundColor }
            try { Microsoft.PowerShell.Utility\Write-Host @hp } catch { }
            if ($NoNewline) { return }
            try { Add-Content -LiteralPath $script:appOut -Value ("$Object") -Encoding UTF8 } catch { }
        }
    }
}
$runLog = Join-Path $logDir "summary_$(Get-Date -Format 'yyyy-MM-dd').log"
$stateDir = Join-Path $logDir "state"
New-Item -ItemType Directory -Force -Path $stateDir | Out-Null
$script:overwrite = $false
$script:runOverwrite = $false
$script:choiceMadeForAll = $false
$script:includeShared = $true
$script:online = @{}
$script:whoami = @{}

# ---------- helpers ----------
# Functions that run rclone in the foreground (sign-in, copy) never return values:
# they set $script: variables instead, so rclone's prompts and progress go straight
# to the console instead of being captured as the function's output.

# Whole-window background color shows what the script is doing:
#   black = menu, blue = checking/calculating, cyan = signing in,
#   green = downloading, magenta = verifying/hashing
$activityColors = @{
    Menu      = @{ Code = "07"; Title = "Google Drive Backup - menu" }
    Checking  = @{ Code = "1F"; Title = "Google Drive Backup - CHECKING" }
    SignIn    = @{ Code = "3F"; Title = "Google Drive Backup - SIGNING IN" }
    Download  = @{ Code = "2F"; Title = "Google Drive Backup - DOWNLOADING" }
    Verify    = @{ Code = "5F"; Title = "Google Drive Backup - VERIFYING" }
}

function Set-Activity($name, $detail = "") {
    $c = $activityColors[$name]
    if (-not $c) { return }
    try {
        # "color" recolors the whole window at once without clearing it
        cmd /c "color $($c.Code)"
        $title = $c.Title
        if ($detail) { $title = "$title - $detail" }
        $Host.UI.RawUI.WindowTitle = $title
    } catch { }
}

function Write-Log($text, $color = "Gray") {
    Write-Host $text -ForegroundColor $color
    Add-Content -LiteralPath $runLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $text"
}

function Get-Remotes {
    return @(& rclone listremotes 2>$null | ForEach-Object { $_.Trim().TrimEnd(":") })
}

# Use the new remote name, or the old one if only that exists
function Resolve-Remote($account, $remotes) {
    if (($remotes -notcontains $account.Remote) -and ($remotes -contains $account.Legacy)) {
        $account.Remote = $account.Legacy
    }
}

# AUTHENTICATED / NOT CONNECTED / EXPIRED
function Get-AuthStatus($account, $remotes) {
    if ($remotes -notcontains $account.Remote) { return "NOT CONNECTED" }
    if ($clientId) {
        $cfg = (& rclone config show $account.Remote 2>$null) | Out-String
        if ($cfg -notmatch [regex]::Escape($clientId)) { return "OLD SIGN-IN" }
    }
    # "about" is instant: Google's own storage numbers for the account (also proves the login works)
    $json = & rclone about "$($account.Remote):" --json --contimeout 15s --timeout 30s 2>$null
    if ($LASTEXITCODE -ne 0) { return "EXPIRED" }
    try { $script:online[$account.Folder] = ($json | Out-String | ConvertFrom-Json) } catch { }
    # Make sure the browser sign-in was really this account and not another one
    $who = & rclone config userinfo "$($account.Remote):" --json --contimeout 15s --timeout 30s 2>$null
    if ($LASTEXITCODE -eq 0 -and $who) {
        try {
            $email = ($who | Out-String | ConvertFrom-Json).Email
            if ($email) { $script:whoami[$account.Folder] = $email }
            if ($email -and $email -ne $account.Email) {
                $script:online.Remove($account.Folder)
                Write-Log "WRONG ACCOUNT: $($account.Remote) is signed in as $email, not $($account.Email)" "Red"
                return "WRONG ACCOUNT ($email)"
            }
        } catch { }
    }
    return "AUTHENTICATED"
}

function Get-SafeName($name) {
    return ($name -replace '[\\/:*?"<>|]', '_').Trim()
}

# Files rclone sees on Drive for this source (Google Docs counted as exported files)
function Get-RemoteCount($src, $extra) {
    $json = Invoke-RcloneQuiet (@("size", $src, "--json", "--drive-export-formats", $exportFormats) + $extra + $script:junkArgs) "Counting files on Drive..."
    if ($script:rcExit -ne 0 -or -not $json) { return -1 }
    return [int](($json | Out-String | ConvertFrom-Json).count)
}

function Get-LocalCount($dst) {
    if (-not (Test-Path -LiteralPath $dst)) { return 0 }
    $files = @(Get-ChildItem -LiteralPath $dst -Recurse -File -Force -ErrorAction SilentlyContinue)
    if ($skipJunk) { $files = @($files | Where-Object { $_.Name -notlike "._*" -and $_.Name -ne ".DS_Store" }) }
    return $files.Count
}

# Sign in (new remote) or re-authorize (expired remote). Sets $script:connectOk.
function Connect-Account($account, $status) {
    $script:connectOk = $false
    Set-Activity SignIn $account.Email
    Write-Host ""
    Write-Host "--------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Sign in to Google as: $($account.Email)" -ForegroundColor Yellow
    Write-Host "Password:             $($account.Password)" -ForegroundColor Yellow
    Write-Host "Tip: if Google picks another account, choose 'Use another account'"
    Write-Host "     or paste the link rclone prints into a private window."
    Write-Host "--------------------------------------------" -ForegroundColor DarkGray
    if ($script:appMode) { Write-Host "Opening the browser: sign in as $($account.Email)  (password: $($account.Password))" }
    else { $null = Read-Host "Press Enter to open the browser" }

    if ($status -ne "NOT CONNECTED") {
        # A broken or half-finished remote (e.g. "Client missing a project id") is
        # rebuilt from scratch instead of reconnected. This only touches rclone's
        # local settings, never the Google account.
        & rclone config delete $account.Remote
    }
    $createArgs = @("config", "create", $account.Remote, "drive", "scope=drive.readonly")
    if ($clientId) { $createArgs += @("client_id=$clientId", "client_secret=$clientSecret") }
    & rclone @createArgs

    $newStatus = Get-AuthStatus $account (Get-Remotes)
    if ($newStatus -eq "AUTHENTICATED") {
        Write-Log "CONNECTED: $($account.Email) as $($account.Remote)" "Green"
        $script:connectOk = $true
        # Sizes saved from an earlier sign-in may belong to another account: forget them
        $st = Read-State $account
        foreach ($k in @("GoogleUsed", "MyDriveBytes", "MyDriveFiles", "SharedBytes", "SharedFiles", "OtherBytes", "OtherFiles",
                         "TrashBytes", "TrashFiles", "SharedDrivesBytes", "SharedDrivesFiles", "SharedDrivesCount")) { $st.Remove($k) }
        Write-State $account $st
    } else {
        Write-Log "NOT CONNECTED: $($account.Email) ($newStatus)" "Red"
    }
}

function Format-Bytes($b) {
    if ($null -eq $b) { return "?" }
    $b = [double]$b
    if ($b -ge 1GB) { return "{0:N2} GB" -f ($b / 1GB) }
    if ($b -ge 1MB) { return "{0:N1} MB" -f ($b / 1MB) }
    if ($b -ge 1KB) { return "{0:N0} KB" -f ($b / 1KB) }
    return "{0:N0} B" -f $b
}

# True when the user pressed S or Esc (used to skip a step that hangs)
function Test-SkipKey {
    # The app's "Skip step" button leaves this file
    $flag = Join-Path $logDir "skip.flag"
    if (Test-Path -LiteralPath $flag) {
        Remove-Item -LiteralPath $flag -Force -ErrorAction SilentlyContinue
        return $true
    }
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq 'S' -or $k.Key -eq 'Escape') { return $true }
        }
    } catch { }
    return $false
}

# Runs rclone in the background and shows a percentage bar (Write-Progress)
# from the stats rclone writes to its JSON log every 2 seconds. Returns the exit code.
# Live progress for the Windows app (backup-app.ps1 reads this file)
function Write-AppProgress($activity, $status, $pct) {
    if (-not $Action) { return }
    try {
        $o = @{ Activity = "$activity"; Status = "$status"; Percent = [int]$pct; Time = (Get-Date).ToString("o") }
        [IO.File]::WriteAllText((Join-Path $logDir "progress.json"), ($o | ConvertTo-Json -Compress), (New-Object System.Text.UTF8Encoding($false)))
    } catch { }
}

# Free bytes on the disk that holds $path ($null if unknown, e.g. network paths)
function Get-FreeBytes($path) {
    try {
        $root = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($path))
        return [double](New-Object System.IO.DriveInfo($root)).AvailableFreeSpace
    } catch { return $null }
}
$script:diskFull = $false
$script:curDest = $null   # where the running download writes (for the disk-full guard)

# All places backups are saved in: the main one first, then the extra ones (Settings)
function Get-Locations { return @(@($destination) + @($script:moreLocations)) }

# Bytes of the files already in a folder (0 when it does not exist)
function Get-DirBytes($dir) {
    if (-not (Test-Path -LiteralPath $dir)) { return [double]0 }
    $sum = (Get-ChildItem -LiteralPath $dir -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    return [double]$(if ($sum) { $sum } else { 0 })
}

# Place for one My Drive folder of $need bytes: the first place with room for what is still
# missing there (+3 GB spare). $preferred (where it was saved before) is tried first.
# When it fits nowhere as a whole, the place with the most free space is used anyway (the
# disk-full guard stops it if it really fills up). $null = no place has even 5 GB free.
function Select-Location($need, $preferred, $sub) {
    $list = @(Get-Locations)
    if ($preferred) { $list = @($preferred) + @($list | Where-Object { $_ -ne $preferred }) }
    $best = $null; $bestFree = [double]0
    foreach ($loc in $list) {
        $free = Get-FreeBytes $loc
        if ($null -eq $free) { continue }   # disk not plugged in
        $left = [double]$need
        if ($sub) { $left -= Get-DirBytes (Join-Path $loc $sub) }
        if ($free -ge ([math]::Max(0, $left) + 3GB)) { return $loc }
        if ($free -gt $bestFree) { $best = $loc; $bestFree = $free }
    }
    if ($best -and $bestFree -ge 5GB) {
        Write-Log "     Not enough room for all of it on one disk: using $best ($(Format-Bytes $bestFree) free). If it fills up, the download stops and you can add a place and press Continue." "Yellow"
        return $best
    }
    return $null
}

function Invoke-RcloneWithProgress($argList, $log, $activity) {
    $exe = (Get-Command rclone).Source
    $argString = ($argList | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join " "
    $errFile = [IO.Path]::GetTempFileName()
    $proc = Start-Process -FilePath $exe -ArgumentList $argString -NoNewWindow -PassThru -RedirectStandardError $errFile
    $null = $proc.Handle   # keeps ExitCode readable after the process ends
    $frames = @('|', '/', '-', '\')
    $tick = 0
    $started = Get-Date

    $script:partSkipped = $false
    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 700
        # Disk full: stop instead of failing on every remaining file
        if ($tick % 10 -eq 0) {
            $free = Get-FreeBytes $(if ($script:curDest) { $script:curDest } else { $destination })
            if ($null -ne $free -and $free -lt 2GB) {
                try { & taskkill.exe /PID $proc.Id /T /F | Out-Null } catch { }
                $script:diskFull = $true
                Write-Host ""
                Write-Log "DISK FULL: only $(Format-Bytes $free) left on the disk of $(if ($script:curDest) { $script:curDest } else { $destination }). Stopped. Free up space or add another place in Settings > More places, then press Continue (files already saved are kept)." "Red"
                break
            }
        }
        if (Test-SkipKey) {
            try { $proc.Kill() } catch { }
            $script:partSkipped = $true
            Write-Host ""
            Write-Host "     Skipped by user (S)." -ForegroundColor Yellow
            break
        }
        $tick++
        $spin = "$($frames[$tick % 4]) $((New-TimeSpan $started (Get-Date)).ToString('hh\:mm\:ss'))  |  "
        $stats = $null
        try {
            $line = Get-Content -LiteralPath $log -Tail 200 -ErrorAction Stop |
                Where-Object { $_ -match '"stats"' } | Select-Object -Last 1
            if ($line) { $stats = ($line | ConvertFrom-Json).stats }
        } catch { }

        if ($stats) {
            $pct = 0
            $eta = "?"
            if ($stats.eta) { $eta = [timespan]::FromSeconds([double]$stats.eta).ToString("hh\:mm\:ss") }
            if ($stats.totalBytes -gt 0) {
                $pct = [math]::Min(100, [math]::Floor(100 * $stats.bytes / $stats.totalBytes))
                $status = "$pct%  |  $(Format-Bytes $stats.bytes) of $(Format-Bytes $stats.totalBytes)  |  " +
                          "files $($stats.transfers) of $($stats.totalTransfers)  |  $(Format-Bytes $stats.speed)/s  |  ETA $eta"
            } else {
                if ($stats.totalChecks -gt 0) { $pct = [math]::Min(100, [math]::Floor(100 * $stats.checks / $stats.totalChecks)) }
                $status = "$pct%  |  checked $($stats.checks) of $($stats.totalChecks) files"
            }
            if ($stats.errors -gt 0) { $status += "  |  errors $($stats.errors)" }
            Write-Progress -Activity "$activity   (press S to skip)" -Status ($spin + $status) -PercentComplete $pct
            Write-AppProgress $activity ($spin + $status) $pct
        } else {
            Write-Progress -Activity "$activity   (press S to skip)" -Status ($spin + "Listing files on Drive...") -PercentComplete 0
            Write-AppProgress $activity ($spin + "Listing files on Drive...") 0
        }
    }
    $proc.WaitForExit()
    Write-Progress -Activity "$activity   (press S to skip)" -Completed
    # Why rclone stopped, when it failed (errors before the log file opens only go here)
    if ($proc.ExitCode -ne 0 -and -not $script:diskFull -and -not $script:partSkipped) {
        $err = @(Get-Content -LiteralPath $errFile -Encoding UTF8 -ErrorAction SilentlyContinue |
            Where-Object { $_ -match 'ERROR|CRITICAL|Fatal|Failed|error' -and $_ -notmatch "shared Google Drive client_id" } |
            Select-Object -Last 3)
        if (-not $err) {
            $err = @(Get-Content -LiteralPath $log -Tail 400 -ErrorAction SilentlyContinue |
                Where-Object { $_ -match '"level":"error"' } | Select-Object -Last 3 |
                ForEach-Object { try { $j = $_ | ConvertFrom-Json; "$($j.object): $($j.msg)" } catch { $_ } })
        }
        if (-not $err) { $err = @("rclone exit code $($proc.ExitCode)") }
        foreach ($e in $err) { Write-Log "     rclone: $(if ("$e".Length -gt 300) { "$e".Substring(0, 300) + '...' } else { $e })" "Red" }
    }
    Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    if ($script:diskFull) { return 98 }
    if ($script:partSkipped) { return 99 }
    return $proc.ExitCode
}

# Runs rclone, shows a spinner with elapsed time while it works, and returns its
# output lines. The exit code is left in $script:rcExit.
# Listing calls give up quickly on files that are gone or not accessible
# (e.g. deleted "Shared with me" items) instead of retrying for a long time.
# Stops after $maxMinutes or when S / Esc is pressed ($script:rcExit = -2).
# Google's "Quota exceeded ... per minute" error clears after a minute: wait and try again.
function Invoke-RcloneQuiet($argList, $label, $maxMinutes = 20) {
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        $lines = Invoke-RcloneQuietOnce $argList $label $maxMinutes
        if ($script:rcExit -eq 0 -or -not $script:rcQuota -or $attempt -eq 4) { return $lines }
        Write-Host "     Google says too many requests. Waiting 70 seconds, then trying again ($attempt of 3)... (S = skip)" -ForegroundColor Yellow
        for ($w = 0; $w -lt 70; $w++) {
            if (Test-SkipKey) { $script:rcExit = -2; return $lines }
            Start-Sleep -Seconds 1
        }
    }
}

function Invoke-RcloneQuietOnce($argList, $label, $maxMinutes = 20) {
    $argList = @($argList) + @("--retries", "1", "--low-level-retries", "3", "--timeout", "2m", "--contimeout", "30s")
    $exe = (Get-Command rclone).Source
    $argString = ($argList | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join " "
    $outFile = [IO.Path]::GetTempFileName()
    $errFile = [IO.Path]::GetTempFileName()
    $proc = Start-Process -FilePath $exe -ArgumentList $argString -NoNewWindow -PassThru `
        -RedirectStandardOutput $outFile -RedirectStandardError $errFile
    $null = $proc.Handle
    $frames = @('|', '/', '-', '\')
    $tick = 0
    $started = Get-Date
    $stopped = $false
    while (-not $proc.HasExited) {
        $span = New-TimeSpan $started (Get-Date)
        if ((Test-SkipKey) -or $span.TotalMinutes -ge $maxMinutes) {
            try { $proc.Kill() } catch { }
            $stopped = $true
            break
        }
        $elapsed = $span.ToString('hh\:mm\:ss')
        Write-Host -NoNewline "`r     $($frames[$tick % 4]) $label  $elapsed  (S = skip)   " -ForegroundColor DarkGray
        if ($tick % 4 -eq 0) { Write-AppProgress $label "$elapsed" -1 }
        $tick++
        Start-Sleep -Milliseconds 250
    }
    $proc.WaitForExit()
    Write-Host -NoNewline ("`r" + (" " * ($label.Length + 30)) + "`r")
    $script:rcExit = $proc.ExitCode
    if ($stopped) {
        $script:rcExit = -2
        Write-Host "     Skipped: $label (not available or taking too long)" -ForegroundColor Yellow
    }
    $lines = @(Get-Content -LiteralPath $outFile -Encoding UTF8 -ErrorAction SilentlyContinue)
    $script:rcQuota = $false
    if ($script:rcExit -ne 0 -and -not $stopped) {
        # Show why rclone failed instead of a silent "?"
        $errText = [System.Net.WebUtility]::HtmlDecode(((Get-Content -LiteralPath $errFile -Encoding UTF8 -ErrorAction SilentlyContinue) -join "`n"))
        if ($errText -match 'Quota exceeded|rateLimitExceeded|userRateLimitExceeded') { $script:rcQuota = $true }
        $err = @($errText -split "`n" |
            Where-Object { ($_ -match 'CRITICAL|Failed to') -and $_ -notmatch "shared Google Drive client_id" } |
            ForEach-Object { if ($_.Length -gt 220) { $_.Substring(0, 220) + "..." } else { $_ } }) | Select-Object -Last 2
        if (-not $err) { $err = @("rclone exit code $($script:rcExit)") }
        foreach ($e in $err) { Write-Log "     rclone error ($label): $e" "Red" }
    }
    Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue
    return $lines
}

# ---------- saved state per account (shown in the menu table) ----------
function Get-StatePath($account) { return (Join-Path $stateDir "$($account.Folder).json") }

function Read-State($account) {
    $h = @{}
    $p = Get-StatePath $account
    if (Test-Path -LiteralPath $p) {
        try {
            $o = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
            foreach ($prop in $o.PSObject.Properties) { $h[$prop.Name] = $prop.Value }
        } catch { }
    }
    return $h
}

function Write-State($account, $h) {
    $h | ConvertTo-Json | Set-Content -LiteralPath (Get-StatePath $account) -Encoding UTF8
}

function Get-StateValue($h, $key, $default) {
    if ($h.ContainsKey($key) -and $null -ne $h[$key] -and "$($h[$key])" -ne "") { return $h[$key] }
    return $default
}

# Files, folders and bytes in the account's Drive folders (Gmail and Contacts not counted)
function Update-LocalStats($account, $state) {
    $files = 0; $folders = 0; $bytes = [double]0
    $subs = @("Mi_unidad", "Papelera", "Unidades_compartidas")
    if ($script:includeShared -and -not (Get-StateValue $state "SharedSkipped" $false)) { $subs += "Compartidos_conmigo" }
    foreach ($dir in @(foreach ($loc in Get-Locations) { foreach ($sub in $subs) { Join-Path (Join-Path $loc $account.Folder) $sub } })) {
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($item in Get-ChildItem -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue) {
            if ($item.PSIsContainer) { $folders++ } else { $files++; $bytes += $item.Length }
        }
    }
    $state.LocalFiles   = $files
    $state.LocalFolders = $folders
    $state.LocalBytes   = $bytes
    $drive = [int](Get-StateValue $state "DriveFiles" -1)
    if ($drive -gt 0) { $state.Percent = [math]::Min(100, [math]::Round(100 * $files / $drive, 1)) }
    elseif ($drive -eq 0) { $state.Percent = 100 }
}

# The parts of one account: My Drive, Shared with me, Trash, and each shared drive
# "Shared with me" is left out when it is switched off (W) or was skipped for this account,
# unless $all is $true.
function Get-Parts($account, $all = $false) {
    $root = Join-Path $destination $account.Folder
    $src  = "$($account.Remote):"
    $state = Read-State $account
    $withShared = $script:wantShared -and ($all -or ($script:includeShared -and -not (Get-StateValue $state "SharedSkipped" $false)))
    $parts = @()
    $scArgs = @(); if ($NoShortcuts) { $scArgs = @("--drive-skip-shortcuts") }
    if ($script:wantMy) {
        $myExtra = @() + $scArgs
        $chunked = $false
        if ($OwnedOnly) {
            # Downloading: folder by folder (Backup-MyDriveChunked). Verify / sizes: one owner list.
            if ($script:downloading) { $chunked = $true }
            else {
                $ol = Update-OwnedList $account
                if ($ol) { $myExtra += @("--files-from-raw", $ol) }
            }
        }
        $parts += [pscustomobject]@{ Name = "Mi_unidad"; Dst = (Join-Path $root "Mi_unidad"); Extra = $myExtra; Chunked = $chunked }
    }
    if ($withShared) {
        $sharedExtra = @("--drive-shared-with-me") + $scArgs
        $extList = Get-ExternalListPath $account
        if (-not $all -and (Get-StateValue $state "SharedOnlyExternal" $false) -and (Test-Path -LiteralPath $extList)) {
            $sharedExtra += @("--files-from-raw", $extList)
        }
        $parts += [pscustomobject]@{ Name = "Compartidos_conmigo"; Dst = (Join-Path $root "Compartidos_conmigo"); Extra = $sharedExtra }
    }
    if ($includeTrash) {
        $parts += [pscustomobject]@{ Name = "Papelera"; Dst = (Join-Path $root "Papelera"); Extra = (@("--drive-trashed-only") + $scArgs) }
    }
    if (-not $script:wantDrives) { return $parts }
    $drivesJson = Invoke-RcloneQuiet @("backend", "drives", $src) "Looking for shared drives..."
    if ($script:rcExit -eq 0 -and $drivesJson) {
        try {
            foreach ($d in @(($drivesJson | Out-String | ConvertFrom-Json) | ForEach-Object { $_ })) {
                if (-not $d.id) { continue }
                $name = Get-SafeName $d.name
                $parts += [pscustomobject]@{
                    Name = "Unidad_compartida_$name"; Dst = (Join-Path $root "Unidades_compartidas\$name"); Extra = (@("--drive-team-drive", $d.id) + $scArgs)
                }
            }
        } catch { }
    }
    return $parts
}

# ---------- duplicates: several files with the same name in one Drive folder ----------
# Drive lets one folder hold several files with the same name; Windows does not.
# Copies with the same content (same MD5) are kept once. When the contents differ, the
# first copy keeps the normal name and the others are saved by Drive ID next to it as
# "name (dup 2).ext", "name (dup 3).ext", ...

# One copy for each different content. Google Docs have no MD5: each one counts as different.
function Select-DistinctCopies($copies) {
    $seen = @{}
    $out = New-Object System.Collections.ArrayList
    foreach ($c in $copies) {
        if ($c.Md5) {
            if ($seen.ContainsKey($c.Md5)) { continue }
            $seen[$c.Md5] = $true
        }
        [void]$out.Add($c)
    }
    return , $out
}

# Reads one lsjson line: $null for folders and other lines
function Read-LsjsonFile($line) {
    $t = "$line".Trim().TrimEnd(',')
    if (-not $t.StartsWith("{")) { return $null }
    try { $o = $t | ConvertFrom-Json } catch { return $null }
    $md5 = ""; if ($o.Hashes -and $o.Hashes.md5) { $md5 = "$($o.Hashes.md5)" }
    $size = [double]$o.Size; if ($size -lt 0) { $size = 0 }
    $owner = $(if ($o.Metadata -and $o.Metadata.owner) { "$($o.Metadata.owner)" } else { "" })
    return [pscustomobject]@{ Path = "$($o.Path)"; Id = "$($o.ID)"; Md5 = $md5; Size = $size; IsDir = [bool]$o.IsDir; Owner = $owner }
}

# Returns @{ "folder/name.ext" = @(copy1, copy2, ...) } for names that have 2 or more
# different contents. Each copy is @{ Id; Md5; Size }.
function Get-DuplicateGroups($src, $extra) {
    $groups = @{}
    $lines = Invoke-RcloneQuiet (@("lsjson", $src, "-R", "--files-only", "--hash", "--hash-type", "MD5",
                                   "--drive-export-formats", $exportFormats) + $extra + $script:junkArgs) "Looking for duplicate names..."
    foreach ($line in @($lines)) {
        $f = Read-LsjsonFile $line
        if ($null -eq $f -or $f.IsDir) { continue }
        if (-not $groups.ContainsKey($f.Path)) { $groups[$f.Path] = New-Object System.Collections.ArrayList }
        [void]$groups[$f.Path].Add($f)
    }
    $dups = @{}
    foreach ($k in $groups.Keys) {
        if ($groups[$k].Count -lt 2) { continue }
        $d = Select-DistinctCopies $groups[$k]
        if ($d.Count -gt 1) { $dups[$k] = $d }
    }
    return $dups
}

# Copy 1 keeps the normal name; copy 2, 3, ... get " (dup N)"
function Get-DupTarget($dst, $path, $k) {
    $rel  = $path -replace '/', '\'
    if ($k -le 1) { return (Join-Path $dst $rel) }
    $dir  = [IO.Path]::GetDirectoryName($rel)
    $leaf = [IO.Path]::GetFileName($rel)
    $name = "$([IO.Path]::GetFileNameWithoutExtension($leaf)) (dup $k)$([IO.Path]::GetExtension($leaf))"
    if ($dir) { return (Join-Path (Join-Path $dst $dir) $name) }
    return (Join-Path $dst $name)
}

function Save-Duplicates($account, $p, $src) {
    Save-DupGroups $account $p $src (Get-DuplicateGroups $src $p.Extra)
}

function Save-DupGroups($account, $p, $src, $dups) {
    if ($dups.Count -eq 0) { return }
    $log = Join-Path $logDir "$($account.Folder)_$(Get-SafeName $p.Name)_duplicates_saved.log"
    $extra = @($p.Extra | Where-Object { $_ })
    Write-Host "     $($dups.Count) name(s) have several different files on Drive: saving each different one..." -ForegroundColor Yellow
    foreach ($path in $dups.Keys) {
        $k = 1
        foreach ($c in $dups[$path]) {
            $target = Get-DupTarget $p.Dst $path $k
            $need = $script:overwrite -or $script:runOverwrite -or -not (Test-Path -LiteralPath $target)
            # The normal name may hold another copy (rclone picks one): replace it when the size is not copy 1's
            if (-not $need -and $k -eq 1 -and $c.Md5) { $need = ((Get-Item -LiteralPath $target).Length -ne [long]$c.Size) }
            if ($need) {
                New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($target)) | Out-Null
                $out = & rclone backend copyid $src $c.Id $target --drive-export-formats $exportFormats @extra 2>&1 | ForEach-Object { "$_" }
                Add-Content -LiteralPath $log -Value "$path  id=$($c.Id)  ->  $target  $($out -join ' ')"
            }
            $k++
        }
    }
    Write-Log "DUPLICATES saved: $($account.Email) / $($p.Name), $($dups.Count) name(s)" "Cyan"
}

# ---------- sizes on Drive ----------
# Files and bytes on Drive for one part (Google Docs counted as exported files). $null on error.
function Measure-Part($src, $extra) {
    $json = Invoke-RcloneQuiet (@("size", $src, "--json", "--drive-export-formats", $exportFormats) + $extra + $script:junkArgs) "Calculating size on Drive..."
    if ($script:rcExit -ne 0 -or -not $json) { return $null }
    $o = $json | Out-String | ConvertFrom-Json
    return [pscustomobject]@{ Files = [int]$o.count; Bytes = [double]$o.bytes }
}

function Show-DriveSizes($account) {
    Set-Activity Checking "sizes $($account.Email)"
    $src = "$($account.Remote):"
    $state = Read-State $account
    $rows = @()
    Write-Host ""
    Write-Host "=========== SIZE ON DRIVE: $($account.Email) ===========" -ForegroundColor Cyan
    $other = [double]0; $otherFiles = 0
    $trash = [double]0; $trashFiles = 0; $teams = [double]0; $teamFiles = 0; $teamCount = 0
    $measuredNow = @()
    foreach ($p in @(Get-Parts $account $true)) {
        Write-Host "  $($p.Name)"
        if ($p.Name -eq "Mi_unidad") { $info = Measure-MyDriveDetail $account $state }
        else { $info = Measure-Part $src $p.Extra }
        if ($info) {
            $measuredNow += [double]$info.Bytes
            if ($p.Name -eq "Mi_unidad") { }   # saved by Measure-MyDriveDetail
            elseif ($p.Name -eq "Compartidos_conmigo") { $state.SharedFiles = $info.Files; $state.SharedBytes = $info.Bytes }
            else {
                $other += $info.Bytes; $otherFiles += $info.Files
                if ($p.Name -eq "Papelera") { $trash += $info.Bytes; $trashFiles += $info.Files }
                else { $teams += $info.Bytes; $teamFiles += $info.Files }
            }
        }
        if ($p.Name -like "Unidad_compartida_*") { $teamCount++ }
        $rows += [pscustomobject]@{
            Part  = $p.Name
            Files = $(if ($info) { $info.Files } else { "?" })
            Size  = $(if ($info) { Format-Bytes $info.Bytes } else { "?" })
        }
    }
    # Only parts that were measured this time replace the saved numbers
    if ($includeTrash) { $state.TrashBytes = $trash; $state.TrashFiles = $trashFiles }
    if ($script:wantDrives) { $state.SharedDrivesBytes = $teams; $state.SharedDrivesFiles = $teamFiles; $state.SharedDrivesCount = $teamCount }
    $state.OtherBytes = [double](Get-StateValue $state "SharedDrivesBytes" 0) + $(if ($includeTrash) { [double](Get-StateValue $state "TrashBytes" 0) } else { 0 })
    $state.OtherFiles = [int](Get-StateValue $state "SharedDrivesFiles" 0) + $(if ($includeTrash) { [int](Get-StateValue $state "TrashFiles" 0) } else { 0 })
    Write-State $account $state
    $total = [double](($measuredNow | Measure-Object -Sum).Sum)
    $rows += [pscustomobject]@{ Part = "TOTAL"; Files = ""; Size = Format-Bytes $total }
    $rows | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
}

# My Drive in detail: everything, without shortcuts, own files, files of other owners.
# Saves all of it in the state; returns what the current options download.
function Measure-MyDriveDetail($account, $state) {
    $src = "$($account.Remote):"
    $all = Measure-Part $src @()
    $noSc = Measure-Part $src @("--drive-skip-shortcuts")
    if ($all) { $state.MyAllFiles = $all.Files; $state.MyAllBytes = $all.Bytes }
    if ($all -and $noSc) {
        $state.MyShortcutFiles = [math]::Max(0, $all.Files - $noSc.Files)
        $state.MyShortcutBytes = [math]::Max(0, $all.Bytes - $noSc.Bytes)
        $state.MyNoShortcutFiles = $noSc.Files; $state.MyNoShortcutBytes = $noSc.Bytes
    }
    Write-State $account $state
    $null = Update-OwnedList $account
    $fresh = Read-State $account
    foreach ($k in @("MyOwnedFiles", "MyOwnedBytes", "MyOthersFiles", "MyOthersBytes")) { if ($fresh.ContainsKey($k)) { $state[$k] = $fresh[$k] } }
    $pick = Get-MyDriveChoice $state
    if ($null -eq $pick) { return $null }
    $state.MyDriveFiles = $pick.Files; $state.MyDriveBytes = $pick.Bytes
    Write-Host "     all $(Format-Bytes (Get-StateValue $state 'MyAllBytes' 0))  |  own $(Format-Bytes (Get-StateValue $state 'MyOwnedBytes' 0))  |  other owners $(Format-Bytes (Get-StateValue $state 'MyOthersBytes' 0))  |  shortcuts $(Format-Bytes (Get-StateValue $state 'MyShortcutBytes' 0))" -ForegroundColor DarkGray
    return $pick
}

# Files and bytes of My Drive for the ticked options (own only / no shortcuts / all)
function Get-MyDriveChoice($s) {
    if ($OwnedOnly -and $s.ContainsKey("MyOwnedBytes")) { return [pscustomobject]@{ Files = [int]$s.MyOwnedFiles; Bytes = [double]$s.MyOwnedBytes } }
    if ($NoShortcuts -and $s.ContainsKey("MyNoShortcutBytes")) { return [pscustomobject]@{ Files = [int]$s.MyNoShortcutFiles; Bytes = [double]$s.MyNoShortcutBytes } }
    if ($s.ContainsKey("MyAllBytes")) { return [pscustomobject]@{ Files = [int]$s.MyAllFiles; Bytes = [double]$s.MyAllBytes } }
    if ($s.ContainsKey("MyDriveBytes")) { return [pscustomobject]@{ Files = [int](Get-StateValue $s "MyDriveFiles" 0); Bytes = [double]$s.MyDriveBytes } }
    return $null
}
function Get-MyDriveBytes($s) { $c = Get-MyDriveChoice $s; if ($c) { return $c.Bytes } return [double]0 }
function Format-SizeFiles($s, $bytesKey, $filesKey) {
    if (-not $s.ContainsKey($bytesKey)) { return "-" }
    return "$(Format-Bytes $s[$bytesKey]) ($($s[$filesKey]))"
}

# ---------- Shared with me cross-check ----------
# Drive gives every file one ID, whoever looks at it. A "Shared with me" file whose
# ID is in the My Drive / shared drives / trash of one of these accounts is already
# backed up there; only the rest (shared from outside) needs its own copy.

function Get-ExternalListPath($account) { return (Join-Path $stateDir "$($account.Folder)_shared_external.txt") }
function Get-OwnedListPath($account) { return (Join-Path $stateDir "$($account.Folder)_owned_files.txt") }

# My Drive files owned by the account itself (paths, one per line, for --files-from-raw).
# Made once per run. Returns the list path, or $null when the listing failed (then nothing is filtered).
$script:ownedDone = @{}
function Update-OwnedList($account) {
    $file = Get-OwnedListPath $account
    if ($script:ownedDone.ContainsKey($account.Email)) { return $script:ownedDone[$account.Email] }
    # Without shortcuts: a shortcut's target lives somewhere else (its own owner's place)
    $args2 = @("lsjson", "$($account.Remote):", "-R", "--files-only", "--metadata", "--drive-skip-shortcuts",
               "--drive-export-formats", $exportFormats) + $script:junkArgs
    $lines = Invoke-RcloneQuiet $args2 "Finding the files owned by $($account.Email)..." 120
    $result = $null
    if ($script:rcExit -eq 0) {
        $own = New-Object System.Collections.Generic.List[string]
        $other = 0; $otherB = 0.0; $ownB = 0.0
        foreach ($line in @($lines)) {
            $t = "$line".Trim().TrimEnd(',')
            if (-not $t.StartsWith("{")) { continue }
            try { $o = $t | ConvertFrom-Json } catch { continue }
            $size = [double]$o.Size; if ($size -lt 0) { $size = 0 }
            $owner = $(if ($o.Metadata -and $o.Metadata.owner) { "$($o.Metadata.owner)" } else { "" })
            if ($owner -eq "" -or $owner -ieq $account.Email) { $own.Add($o.Path); $ownB += $size }
            else { $other++; $otherB += $size }
        }
        [IO.File]::WriteAllLines($file, $own, (New-Object System.Text.UTF8Encoding($false)))
        $st = Read-State $account
        $st.MyOwnedFiles = $own.Count; $st.MyOwnedBytes = $ownB
        $st.MyOthersFiles = $other; $st.MyOthersBytes = $otherB
        Write-State $account $st
        Write-Log "OWNERS: $($account.Email) My Drive: $($own.Count) own files ($(Format-Bytes $ownB)), $other files of other owners ($(Format-Bytes $otherB))" "Cyan"
        $result = $file
    } else {
        Write-Log "OWNERS: could not list the owners for $($account.Email); all of My Drive is used instead." "Yellow"
    }
    $script:ownedDone[$account.Email] = $result
    return $result
}

# @{ id = @{ Size; Path } } for every file of one part
function Get-IdList($src, $extra, $label) {
    $ids = @{}
    $lines = Invoke-RcloneQuiet (@("lsf", $src, "-R", "--files-only", "--format", "isp", "--separator", "|", "--drive-export-formats", $exportFormats) + $extra) $label 120
    foreach ($line in @($lines)) {
        $f = "$line".Split('|', 3)
        if ($f.Count -lt 3) { continue }
        $size = [double]0
        [void][double]::TryParse($f[1], [ref]$size)
        if ($size -lt 0) { $size = 0 }
        $ids[$f[0]] = @{ Size = $size; Path = $f[2] }
    }
    return $ids
}

# Shared with me with the owner of each file: @{ id = @{ Size; Path; Owner } }.
# rclone lsjson prints one file per line, so it is read line by line (no size limit).
function Get-SharedList($src, $label) {
    $list = @{}
    $lines = Invoke-RcloneQuiet @("lsjson", $src, "-R", "--files-only", "--metadata", "--drive-shared-with-me",
        "--drive-export-formats", $exportFormats) $label 120
    foreach ($line in @($lines)) {
        $t = "$line".Trim().TrimEnd(',')
        if (-not $t.StartsWith("{")) { continue }
        try { $o = $t | ConvertFrom-Json } catch { continue }
        $owner = "(unknown)"
        if ($o.Metadata -and $o.Metadata.owner) { $owner = $o.Metadata.owner }
        $size = [double]$o.Size
        if ($size -lt 0) { $size = 0 }
        $list[$o.ID] = @{ Size = $size; Path = $o.Path; Owner = $owner }
    }
    return $list
}

function Write-SharedReport($account, $items, $intN, $intB, $extN, $extB) {
    $root = Join-Path $destination $account.Folder
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $txt = Join-Path $root "_REPORTE_Compartidos_conmigo.txt"
    $csv = Join-Path $root "_REPORTE_Compartidos_conmigo.csv"
    $out = New-Object System.Collections.ArrayList
    [void]$out.Add("SHARED WITH ME - $($account.Email)")
    [void]$out.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$out.Add("")
    [void]$out.Add("Total files: $($intN + $extN)")
    [void]$out.Add("REDUNDANT (owned by one of these accounts, already in its backup): $intN files, $(Format-Bytes $intB)")
    [void]$out.Add("FROM OUTSIDE (shared by other people, kept in Compartidos_conmigo): $extN files, $(Format-Bytes $extB)")
    [void]$out.Add("")
    [void]$out.Add("==================== FROM OUTSIDE, by who shared it ====================")
    foreach ($g in ($items | Where-Object { -not $_.Redundant } | Group-Object Owner | Sort-Object Count -Descending)) {
        $gb = ($g.Group | Measure-Object Size -Sum).Sum
        [void]$out.Add("")
        [void]$out.Add("--- $($g.Name): $($g.Count) files, $(Format-Bytes $gb)")
        foreach ($i in ($g.Group | Sort-Object Path)) { [void]$out.Add("    $($i.Path)  ($(Format-Bytes $i.Size))") }
    }
    [void]$out.Add("")
    [void]$out.Add("==================== REDUNDANT, by owner account ====================")
    foreach ($g in ($items | Where-Object { $_.Redundant } | Group-Object Owner | Sort-Object Count -Descending)) {
        [void]$out.Add("")
        [void]$out.Add("--- $($g.Name): $($g.Count) files (backed up in its own folder)")
        foreach ($i in ($g.Group | Sort-Object Path)) { [void]$out.Add("    $($i.Path)  ->  $($i.BackedUpIn)") }
    }
    $out | Set-Content -LiteralPath $txt -Encoding UTF8
    $items | Sort-Object Redundant, Owner, Path |
        Select-Object @{ n = "Status"; e = { if ($_.Redundant) { "REDUNDANT" } else { "OUTSIDE" } } }, Owner, Path,
                      @{ n = "Size"; e = { Format-Bytes $_.Size } }, @{ n = "Bytes"; e = { $_.Size } }, BackedUpIn |
        Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
    return $txt
}

function Compare-SharedWithMe {
    Set-Activity Checking "shared with me"
    # A shared file is redundant when its owner is one of these accounts: it is
    # backed up in the owner's own folder. The owner comes from Drive's metadata,
    # so the owner account does not need to be signed in for this check.
    $emailToFolder = @{}
    foreach ($a in $accounts) { $emailToFolder[$a.Email.ToLower()] = $a.Folder }
    $checked = @()
    Write-Host ""
    Write-Host "=========== CROSS-CHECK: Shared with me vs. these accounts ===========" -ForegroundColor Cyan
    for ($k = 0; $k -lt $accounts.Count; $k++) {
        if ($script:statuses[$k] -eq "AUTHENTICATED") { $checked += $accounts[$k] }
    }

    $rows = @()
    foreach ($a in $checked) {
        $shared = Get-SharedList "$($a.Remote):" "Reading Shared with me + owners: $($a.Email)..."
        $intN = 0; $intB = [double]0; $extN = 0; $extB = [double]0
        $extPaths = New-Object System.Collections.ArrayList
        $report = New-Object System.Collections.ArrayList
        $items = New-Object System.Collections.ArrayList
        foreach ($id in $shared.Keys) {
            $f = $shared[$id]
            $ownerKey = "$($f.Owner)".ToLower()
            $isRedundant = $emailToFolder.ContainsKey($ownerKey)
            if ($isRedundant) {
                $intN++; $intB += $f.Size
                [void]$report.Add("INSIDE   (owner: $($f.Owner))  $($f.Path)")
            } else {
                $extN++; $extB += $f.Size
                [void]$extPaths.Add($f.Path)
                [void]$report.Add("OUTSIDE  (owner: $($f.Owner))  $($f.Path)  ($(Format-Bytes $f.Size))")
            }
            [void]$items.Add([pscustomobject]@{
                Redundant  = $isRedundant
                Owner      = $f.Owner
                Path       = $f.Path
                Size       = $f.Size
                BackedUpIn = $(if ($isRedundant) { "$($emailToFolder[$ownerKey]) (owner's own backup)" } else { "$($a.Folder)\Compartidos_conmigo" })
            })
        }
        $reportFile = Write-SharedReport $a $items $intN $intB $extN $extB
        # Plain UTF-8 without BOM, one path per line, so rclone --files-from-raw reads it exactly
        [IO.File]::WriteAllLines((Get-ExternalListPath $a), [string[]]@($extPaths | Sort-Object -Unique), (New-Object System.Text.UTF8Encoding($false)))
        $report | Sort-Object | Set-Content -LiteralPath (Join-Path $logDir "$($a.Folder)_shared_crosscheck.txt") -Encoding UTF8

        $state = Read-State $a
        $state.SharedInternalFiles = $intN
        $state.SharedExternalFiles = $extN
        $state.SharedExternalBytes = $extB
        Write-State $a $state
        $rows += [pscustomobject]@{
            Account = $a.Email
            'Shared files' = $shared.Count
            'From these accounts' = "$intN ($(Format-Bytes $intB))"
            'From outside' = "$extN ($(Format-Bytes $extB))"
            Report = $reportFile
        }
    }
    $rows | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    Write-Host "Report in each account folder: _REPORTE_Compartidos_conmigo.txt (by who shared it) and .csv (for Excel)."
    $notChecked = @(for ($k = 0; $k -lt $accounts.Count; $k++) { if ($script:statuses[$k] -ne "AUTHENTICATED") { $accounts[$k].Email } })
    if ($notChecked.Count -gt 0) {
        Write-Host "Not signed in, so their own Shared with me was not read yet: $($notChecked -join ', ')." -ForegroundColor DarkGray
    }
    Write-Host "When downloading, choose E to get only the files from outside." -ForegroundColor DarkGray
}

# ---------- My Drive in chunks (own files only) ----------
# Reads lsjson lines with owners. Returns the own paths (with $prefix in front), the counts,
# the folders, and the own names that appear more than once (path -> IDs) for the duplicates step.
function Split-ByOwner($account, $lines, $prefix) {
    $r = [pscustomobject]@{ Own = (New-Object System.Collections.Generic.List[string]); OwnBytes = 0.0
                            Others = 0; OthersBytes = 0.0; Dirs = @(); Dups = @{} }
    $byPath = [ordered]@{}
    foreach ($line in @($lines)) {
        $f = Read-LsjsonFile $line
        if ($null -eq $f) { continue }
        if ($f.IsDir) { $r.Dirs += $f.Path; continue }
        if ($f.Owner -ne "" -and $f.Owner -ine $account.Email) { $r.Others++; $r.OthersBytes += $f.Size; continue }
        $path = "$prefix$($f.Path)"
        if (-not $byPath.Contains($path)) { $byPath[$path] = New-Object System.Collections.ArrayList }
        [void]$byPath[$path].Add($f)
    }
    foreach ($path in $byPath.Keys) {
        $copies = Select-DistinctCopies $byPath[$path]
        foreach ($c in $copies) { $r.OwnBytes += $c.Size }
        # One content: rclone downloads it. Several: the duplicates step saves each one by ID.
        if ($copies.Count -gt 1) { $r.Dups[$path] = $copies } else { $r.Own.Add($path) }
    }
    return $r
}

# My Drive, own files only, in chunks: first the files at the top, then one top folder at a time.
# Each chunk lists the owners of its own folder only and downloads right after, so the
# download starts in seconds instead of after a listing of all of My Drive.
# rclone --files-from-raw goes straight to the listed files and does not scan other folders.
function Backup-MyDriveChunked($account, $p, $src) {
    $enc = New-Object System.Text.UTF8Encoding($false)
    $chunkFile = Join-Path $stateDir "$($account.Folder)_owned_chunk.txt"
    $listArgs = @("--files-only", "--metadata", "--hash", "--hash-type", "MD5", "--drive-skip-shortcuts", "--drive-export-formats", $exportFormats) + $script:junkArgs
    $base = @($p.Extra | Where-Object { $_ })

    $top = Invoke-RcloneQuiet (@("lsjson", $src, "--max-depth", "1", "--metadata", "--hash", "--hash-type", "MD5", "--drive-skip-shortcuts",
                                "--drive-export-formats", $exportFormats) + $script:junkArgs) "Reading the top of My Drive..."
    if ($script:rcExit -ne 0) {
        Write-Log "OWNERS: could not read My Drive of $($account.Email) folder by folder; downloading all of My Drive instead." "Yellow"
        Invoke-Part $account $p.Name $src $p.Dst $base
        if (-not $script:diskFull) { Save-Duplicates $account $p $src }
        return
    }
    $rootPart = Split-ByOwner $account $top ""
    # Finished chunks are written down, so a paused download continues where it was
    $doneFile = Join-Path $stateDir "$($account.Folder)_chunks_done.txt"
    $done = New-Object 'System.Collections.Generic.HashSet[string]'
    # Kept between runs (also after a normal Download), so finished folders can be zipped and
    # deleted from the disk without being downloaded again. Overwrite starts from zero.
    $fresh = $script:overwrite -or $script:runOverwrite
    if (-not $fresh -and (Test-Path -LiteralPath $doneFile)) { foreach ($l in Get-Content -LiteralPath $doneFile -Encoding UTF8) { [void]$done.Add("$l") } }
    else { Remove-Item -LiteralPath $doneFile -Force -ErrorAction SilentlyContinue }
    # Which place each folder was saved in (kept, so the next run checks the same place)
    $whereFile = Join-Path $stateDir "$($account.Folder)_locations.txt"
    $where = @{}
    if (Test-Path -LiteralPath $whereFile) {
        foreach ($l in Get-Content -LiteralPath $whereFile -Encoding UTF8) { $t = "$l".Split("`t", 2); if ($t.Count -eq 2) { $where[$t[1]] = $t[0] } }
    }
    $chunks = @([pscustomobject]@{ Name = "(files at the top)"; Data = $rootPart }) +
              @($rootPart.Dirs | Sort-Object | ForEach-Object { [pscustomobject]@{ Name = $_; Data = $null } })

    $all = New-Object System.Collections.Generic.List[string]
    $ownN = 0; $ownB = 0.0; $othN = 0; $othB = 0.0; $complete = $true
    $n = 0
    foreach ($c in $chunks) {
        $n++
        if ($script:diskFull) { $complete = $false; break }
        $label = "Mi_unidad [$n of $($chunks.Count)] $($c.Name)"
        if ($done.Contains($c.Name)) {
            Write-Host "     $($c.Name): already finished (skipped)" -ForegroundColor DarkGray
            $complete = $false   # its owners were not listed this time
            continue
        }
        $d = $c.Data
        if ($null -eq $d) {
            $lines = Invoke-RcloneQuiet (@("lsjson", "$src$($c.Name)", "-R") + $listArgs) "Owners in $($c.Name) ($n of $($chunks.Count))..." 120
            if ($script:rcExit -ne 0) {
                # Owners unknown: download the whole folder rather than miss files
                Write-Log "OWNERS: could not list $($c.Name) for $($account.Email); downloading the whole folder." "Yellow"
                $complete = $false
                $glob = $c.Name -replace '([\\*?\[\]{}])', '\$1'
                Invoke-Part $account $label $src $p.Dst ($base + @("--include", "/$glob/**"))
                continue
            }
            $d = Split-ByOwner $account $lines "$($c.Name)/"
        }
        $ownN += $d.Own.Count + $d.Dups.Count; $ownB += $d.OwnBytes; $othN += $d.Others; $othB += $d.OthersBytes
        $all.AddRange($d.Own); foreach ($k in $d.Dups.Keys) { $all.Add($k) }
        Write-Host "     $($c.Name): $($d.Own.Count + $d.Dups.Count) own files ($(Format-Bytes $d.OwnBytes)), $($d.Others) of other owners left out" -ForegroundColor DarkGray
        # Where this folder fits (the place it started in, if it was paused there)
        $loc = Select-Location $d.OwnBytes $where[$c.Name] $(if ($null -eq $c.Data) { Join-Path (Join-Path $account.Folder "Mi_unidad") $c.Name } else { "" })
        if (-not $loc) {
            Write-Log "NO SPACE: $($c.Name) needs about $(Format-Bytes $d.OwnBytes) and no place has even 5 GB free. Free up space or add a place in Settings > More places, then press Continue." "Red"
            $script:diskFull = $true
            continue
        }
        if ($where[$c.Name] -ne $loc) { $where[$c.Name] = $loc; Add-Content -LiteralPath $whereFile -Value "$loc`t$($c.Name)" -Encoding UTF8 }
        $pc = [pscustomobject]@{ Name = $p.Name; Dst = (Join-Path (Join-Path $loc $account.Folder) "Mi_unidad"); Extra = $p.Extra }
        if ($loc -ne $destination) { Write-Host "     $($c.Name) -> saved in $loc" -ForegroundColor Cyan }
        $chunkOk = $true
        if ($d.Own.Count -gt 0) {
            [IO.File]::WriteAllLines($chunkFile, $d.Own, $enc)
            Invoke-Part $account $label $src $pc.Dst ($base + @("--files-from-raw", $chunkFile))
            $chunkOk = (@($script:partResults)[-1].Download -eq "OK")
        }
        if ($script:diskFull) { continue }
        Save-DupGroups $account $pc $src $d.Dups
        if ($chunkOk) { Add-Content -LiteralPath $doneFile -Value $c.Name -Encoding UTF8 }
    }
    Remove-Item -LiteralPath $chunkFile -Force -ErrorAction SilentlyContinue

    if (-not $complete) { return }
    # Every folder was listed: save the owner numbers and the full list (Verify reuses it)
    $file = Get-OwnedListPath $account
    [IO.File]::WriteAllLines($file, $all, $enc)
    $st = Read-State $account
    $st.MyOwnedFiles = $ownN; $st.MyOwnedBytes = $ownB; $st.MyOthersFiles = $othN; $st.MyOthersBytes = $othB
    Write-State $account $st
    $script:ownedDone[$account.Email] = $file
    Write-Log "OWNERS: $($account.Email) My Drive: $ownN own files ($(Format-Bytes $ownB)), $othN files of other owners ($(Format-Bytes $othB))" "Cyan"
}

# ---------- download ----------
# Copy one part with retries. Adds one row to $script:partResults.
function Invoke-Part($account, $part, $src, $dst, $extra) {
    Set-Activity Download "$($account.Email) $part"
    $script:curDest = $dst
    $stamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    $log   = Join-Path $logDir "$($account.Folder)_$(($part -replace '[^\w\-]+', '_').Trim('_'))_$stamp.log"

    $rcArgs = @(
        "--transfers", "$transfers",
        "--bwlimit", $(if ($bwLimit) { $bwLimit } else { "off" }),
        "--checkers", "$checkers",
        "--fast-list",
        "--multi-thread-streams", "8",
        "--multi-thread-cutoff", "64M",
        "--buffer-size", "64M",
        "--retries", "3",
        "--low-level-retries", "10",
        "--create-empty-src-dirs",
        "--drive-export-formats", $exportFormats,
        "--drive-acknowledge-abuse",
        "--log-file", $log,
        "--log-level", "INFO",
        "--use-json-log",
        "--stats", "2s",
        "--stats-log-level", "NOTICE"
    )
    if ($script:overwrite -or $script:runOverwrite) { $rcArgs += "--ignore-times" }
    $rcArgs += $extra
    # A --files-from-raw list is already without Mac files; rclone does not mix it with --exclude rules
    $filterArgs = $(if (@($extra) -contains "--files-from-raw") { @() } else { $script:junkArgs })

    Write-Host ""
    Write-Host "  -> Download: $part" -ForegroundColor Cyan
    New-Item -ItemType Directory -Force -Path $dst | Out-Null

    $ok = $false
    for ($try = 1; $try -le ($retries + 1); $try++) {
        if ($try -gt 1) {
            Write-Log "Retry $($try - 1) of ${retries}: $($account.Email) / $part" "Yellow"
            Start-Sleep -Seconds 30
        }
        $code = Invoke-RcloneWithProgress (@("copy", $src, $dst) + $rcArgs + $filterArgs) $log "Download $($account.Email) / $part"
        if ($code -eq 0) { $ok = $true; break }
        if ($script:partSkipped -or $script:diskFull) { break }
    }

    $status = "OK"
    if (-not $ok) { $status = "FAILED" }
    if ($script:partSkipped) { $status = "SKIPPED" }
    if ($script:diskFull) { $status = "STOPPED (disk full)" }
    Write-Log "DOWNLOAD $status  $($account.Email) / $part" $(if ($ok) { "Green" } else { "Red" })
    $script:partResults += [pscustomobject]@{ Part = $part; Download = $status; Log = $log }
}

# Asks what to do with files that are already downloaded. Sets $script:runOverwrite
# ($true = overwrite all, $false = skip existing) and $script:runCancel.
function Read-OverwriteChoice($label) {
    $script:runOverwrite = $false
    $script:runCancel = $false
    if ($script:appMode) { $script:runOverwrite = [bool]$OverwriteAll; return }
    Write-Host ""
    Write-Host "$label already has downloaded files. What should happen with them?" -ForegroundColor Yellow
    Write-Host "  S = Skip all files already downloaded (fast: only new or changed files)  [default]"
    Write-Host "  O = Overwrite all (download every file again)"
    Write-Host "  X = Cancel"
    $answer = "$(Read-Host 'S/O/X')".Trim().ToUpper()
    if ($answer -eq "O") { $script:runOverwrite = $true }
    elseif ($answer -eq "X") { $script:runCancel = $true }
}

function Test-HasDownload($account) {
    $dir = Join-Path (Join-Path $destination $account.Folder) "Mi_unidad"
    if (-not (Test-Path -LiteralPath $dir)) { return $false }
    return ($null -ne (Get-ChildItem -LiteralPath $dir -Recurse -File -Force -ErrorAction SilentlyContinue | Select-Object -First 1))
}

# Download everything for one account, then verify it
function Backup-Account($account) {
    if (-not $script:overwrite -and -not $script:choiceMadeForAll) {
        $script:runOverwrite = $false
        if (Test-HasDownload $account) {
            Read-OverwriteChoice $account.Email
            if ($script:runCancel) { Write-Host "Cancelled." -ForegroundColor DarkGray; return }
        }
    }

    $root  = Join-Path $destination $account.Folder
    $src   = "$($account.Remote):"
    $state = Read-State $account
    $script:partResults = @()

    Write-Host ""
    Write-Host "============================================" -ForegroundColor DarkGray
    Write-Host "[$($account.Name)] $($account.Email)  ->  $root" -ForegroundColor Cyan
    if ($script:overwrite -or $script:runOverwrite) { Write-Host "OVERWRITE ALL: every file is downloaded again" -ForegroundColor Yellow }
    else { Write-Host "SKIP EXISTING: only new or changed files are downloaded" -ForegroundColor DarkGray }
    Write-Host "============================================" -ForegroundColor DarkGray
    Write-Log "BACKUP START: $($account.Email) (overwrite=$($script:overwrite -or $script:runOverwrite))" "Cyan"

    # No full size measuring here, so the download starts right away (Refresh measures).
    # The space check uses the sizes saved last time, if any; the disk-full guard stops it anyway.
    $st = Read-State $account
    $need = [double](Get-StateValue $st "MyDriveBytes" 0) + [double](Get-StateValue $st "SharedDrivesBytes" 0)
    if (-not (Get-StateValue $st "SharedSkipped" $false)) {
        if (Get-StateValue $st "SharedOnlyExternal" $false) { $need += [double](Get-StateValue $st "SharedExternalBytes" 0) }
        else { $need += [double](Get-StateValue $st "SharedBytes" 0) }
    }
    $need -= [double](Get-StateValue $st "LocalBytes" 0)
    if ($need -gt 0 -and $script:moreLocations.Count -eq 0) { try {
        $free = [double](Get-PSDrive -Name ($destination.Substring(0, 1))).Free
        Write-Host "Free space on $($destination.Substring(0, 2)) $(Format-Bytes $free)   still to download (about): $(Format-Bytes ([math]::Max(0, $need)))" -ForegroundColor DarkGray
        if ($need -gt $free) {
            Write-Log "NOT ENOUGH SPACE for $($account.Email): needs about $(Format-Bytes $need), free $(Format-Bytes $free)" "Red"
            if ($script:appMode) { Write-Log "Stopped: free up space on the disk first." "Red"; return }
            if ("$(Read-Host 'Download anyway? (Y/N)')".Trim().ToUpper() -ne "Y") { return }
        }
    } catch { } }

    $script:downloading = $true
    $parts = @(Get-Parts $account)
    if ($script:includeShared -and -not $script:choiceMadeForAll) {
        # Ask every time (also when it was skipped before), showing its size
        $parts = @(Get-Parts $account $true)
        $state = Read-State $account
        $sizeText = $(if ($state.ContainsKey("SharedBytes")) { "$($state.SharedFiles) files, $(Format-Bytes $state.SharedBytes)" } else { "size unknown" })
        $hasExt = $state.ContainsKey("SharedExternalFiles") -and (Test-Path -LiteralPath (Get-ExternalListPath $account))
        Write-Host ""
        Write-Host "Shared with me: $sizeText" -ForegroundColor Yellow
        if ($hasExt) {
            Write-Host "  $($state.SharedInternalFiles) of them belong to the other accounts in this list (already in their own backup)."
            Write-Host "  E = only the $($state.SharedExternalFiles) files from OUTSIDE these accounts ($(Format-Bytes $state.SharedExternalBytes))  [default]"
            Write-Host "  A = all of Shared with me"
        } else {
            Write-Host "  (Tip: X in the menu finds which of these belong to your other accounts.)" -ForegroundColor DarkGray
            Write-Host "  A = download all of Shared with me  [default]"
        }
        Write-Host "  N = skip Shared with me for this account"
        if ($script:appMode) { $answer = "$SharedMode".Trim().ToUpper(); Write-Host "  -> $(if ($answer) { $answer } else { 'default' })" }
        else { $answer = "$(Read-Host 'Choice')".Trim().ToUpper() }
        $state.SharedSkipped = ($answer -eq "N")
        $state.SharedOnlyExternal = ($hasExt -and $answer -ne "A" -and $answer -ne "N")
        Write-State $account $state
        $parts = @(Get-Parts $account)
    }
    $script:downloading = $false

    foreach ($p in $parts) {
        if ($script:diskFull) { break }
        if ($p.Chunked) { Backup-MyDriveChunked $account $p $src; continue }
        Invoke-Part $account $p.Name $src $p.Dst $p.Extra
        if (-not $script:diskFull) { Save-Duplicates $account $p $src }
    }

    $failed = @($script:partResults | Where-Object { $_.Download -ne "OK" }).Count
    $state.BackupDate   = (Get-Date -Format "yyyy-MM-dd HH:mm")
    $state.BackupStatus = $(if ($script:diskFull) { "STOPPED (no space)" } elseif ($failed -eq 0) { "DONE" } else { "$failed FAILED" })
    Write-State $account $state
    Write-Log "BACKUP END: $($account.Email) ($($state.BackupStatus))" "Cyan"

    if ($script:diskFull) { Write-Host "Verify skipped: the disk is full." -ForegroundColor Red }
    elseif ($NoVerify) { Write-Host "Verify skipped (turned off). Use Verify later." }
    else { Verify-Account $account }
}

# ---------- verify ----------
# Byte check of one part against Drive + duplicates list. Adds one row to $script:verifyResults.
function Verify-Part($account, $p, $src, $problemsFile) {
    $stamp    = Get-Date -Format "yyyy-MM-dd_HHmm"
    $safe     = Get-SafeName $p.Name
    $log      = Join-Path $logDir "$($account.Folder)_${safe}_check_$stamp.log"
    $combined = Join-Path $logDir "$($account.Folder)_${safe}_check_$stamp.txt"
    $dupFile  = Join-Path $logDir "$($account.Folder)_${safe}_duplicates.txt"
    New-Item -ItemType Directory -Force -Path $p.Dst | Out-Null

    Write-Host ""
    Write-Host "  -> Verify: $($p.Name)" -ForegroundColor Cyan

    # 1. Byte check: size + MD5 of every regular file, Drive vs local
    $rcArgs = @(
        "check", $src, $p.Dst,
        "--one-way",
        "--drive-skip-gdocs",
        "--checkers", "$checkers",
        "--fast-list",
        "--combined", $combined,
        "--log-file", $log,
        "--log-level", "NOTICE",
        "--use-json-log",
        "--stats", "2s",
        "--stats-log-level", "NOTICE"
    ) + $p.Extra + $script:junkArgs
    $null = Invoke-RcloneWithProgress $rcArgs $log "Verify $($account.Email) / $($p.Name)"
    if ($script:partSkipped) {
        Add-Content -LiteralPath $problemsFile -Value "SKIPPED  [$($p.Name)] verify was skipped"
        $script:verifyResults += [pscustomobject]@{
            Part = $p.Name; Status = "SKIPPED"; DriveFiles = -1; LocalFiles = (Get-LocalCount $p.Dst)
            MD5_OK = 0; Missing = 0; Corrupt = 0; Errors = 1; Dupes = 0; DupUnsaved = 0
        }
        return
    }

    $matched = 0; $missing = 0; $differ = 0; $errors = 0
    if (Test-Path -LiteralPath $combined) {
        foreach ($line in Get-Content -LiteralPath $combined -Encoding UTF8) {
            if ($line.Length -lt 2) { continue }
            switch ($line[0]) {
                '=' { $matched++ }
                '-' { $missing++; Add-Content -LiteralPath $problemsFile -Value "MISSING  [$($p.Name)] $($line.Substring(2))" }
                '*' { $differ++;  Add-Content -LiteralPath $problemsFile -Value "CORRUPT  [$($p.Name)] $($line.Substring(2))" }
                '!' { $errors++;  Add-Content -LiteralPath $problemsFile -Value "ERROR    [$($p.Name)] $($line.Substring(2))" }
            }
        }
    } else {
        $errors++
        Add-Content -LiteralPath $problemsFile -Value "ERROR    [$($p.Name)] check did not run, see $log"
    }

    # 2. Duplicates: same name, different files in one Drive folder. Every copy must
    #    exist locally as "name (dup N).ext" (saved by the download step).
    $dups = Get-DuplicateGroups $src $p.Extra
    $dupes = $dups.Count
    $unsaved = 0
    $dupReport = @()
    foreach ($path in $dups.Keys) {
        $k = 1; $allThere = $true
        foreach ($id in $dups[$path]) {
            $target = Get-DupTarget $p.Dst $path $k
            if (-not (Test-Path -LiteralPath $target)) { $allThere = $false }
            $k++
        }
        $dupReport += "$(if ($allThere) { 'SAVED    ' } else { 'NOT SAVED' })  $path  ($($dups[$path].Count) files)"
        if (-not $allThere) {
            $unsaved++
            Add-Content -LiteralPath $problemsFile -Value "DUPLICATE NOT SAVED [$($p.Name)] $path ($($dups[$path].Count) files with this name)"
        }
    }
    $dupReport | Set-Content -LiteralPath $dupFile -Encoding UTF8

    # 3. File counts (Google Docs included, as exported files)
    $driveCount = Get-RemoteCount $src $p.Extra
    $localCount = Get-LocalCount $p.Dst

    $status = "OK"
    if ($missing + $differ + $errors -gt 0) { $status = "PROBLEMS" }
    elseif ($driveCount -gt $localCount) { $status = "CHECK" }
    elseif ($unsaved -gt 0) { $status = "DUPLICATES" }
    $color = "Green"; if ($status -ne "OK") { $color = "Yellow" }
    Write-Log "VERIFY $status  $($account.Email) / $($p.Name)  matched=$matched missing=$missing corrupt=$differ errors=$errors dupes=$dupes unsaved=$unsaved drive=$driveCount local=$localCount" $color

    $script:verifyResults += [pscustomobject]@{
        Part = $p.Name; Status = $status; DriveFiles = $driveCount; LocalFiles = $localCount
        MD5_OK = $matched; Missing = $missing; Corrupt = $differ; Errors = $errors; Dupes = $dupes; DupUnsaved = $unsaved
    }
}

function Verify-Account($account) {
    Set-Activity Verify $account.Email
    $src   = "$($account.Remote):"
    $state = Read-State $account
    $stamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    $problemsFile = Join-Path $logDir "$($account.Folder)_PROBLEMS_$stamp.txt"
    $script:verifyResults = @()

    Write-Host ""
    Write-Host "=========== VERIFY: $($account.Email) ===========" -ForegroundColor Cyan
    foreach ($p in @(Get-Parts $account)) {
        Verify-Part $account $p $src $problemsFile
    }

    $r = $script:verifyResults
    $missing = ($r | Measure-Object Missing -Sum).Sum
    $differ  = ($r | Measure-Object Corrupt -Sum).Sum
    $errors  = ($r | Measure-Object Errors  -Sum).Sum
    $dupes   = ($r | Measure-Object Dupes   -Sum).Sum
    $unsaved = ($r | Measure-Object DupUnsaved -Sum).Sum
    $driveFiles = ($r | Where-Object { $_.DriveFiles -ge 0 } | Measure-Object DriveFiles -Sum).Sum

    $state.DriveFiles = [int]$driveFiles
    Update-LocalStats $account $state
    $state.VerifyDate = (Get-Date -Format "yyyy-MM-dd HH:mm")
    $state.Duplicates = $(if ($dupes -gt 0) { "$dupes ($([int]$unsaved) unsaved)" } else { "0" })
    if ($missing + $differ + $errors -eq 0) { $state.VerifyStatus = "OK" }
    else { $state.VerifyStatus = "$missing miss/$differ bad/$errors err" }
    Write-State $account $state

    Write-Host ""
    Write-Host "=========== RESULT: $($account.Email) ===========" -ForegroundColor Cyan
    $r | Format-Table Part, Status, DriveFiles, LocalFiles, MD5_OK, Missing, Corrupt, Errors, Dupes, DupUnsaved -AutoSize | Out-String -Width 300 | Write-Host
    Add-Content -LiteralPath $runLog -Value ($r | Format-Table -AutoSize | Out-String -Width 300)
    $r | Export-Csv -LiteralPath (Join-Path $logDir "$($account.Folder)_verify_$stamp.csv") -NoTypeInformation -Encoding UTF8

    Write-Host "Local: $($state.LocalFiles) files, $($state.LocalFolders) folders, $(Format-Bytes $state.LocalBytes)"
    if ($state.VerifyStatus -eq "OK" -and [int]$unsaved -eq 0 -and ($r | Where-Object { $_.Status -ne "OK" }).Count -eq 0) {
        Write-Host "VERIFIED: every file matches Drive. This account's Drive is safe to delete." -ForegroundColor Green
    } else {
        Write-Host "NOT fully verified. Run the account again (it downloads only what is missing), then verify." -ForegroundColor Yellow
        if (Test-Path -LiteralPath $problemsFile) { Write-Host "Problem list: $problemsFile" -ForegroundColor Yellow }
    }
    Write-Host "Google Docs/Sheets/Slides have no checksum on Drive: they are checked by file count only." -ForegroundColor DarkGray
    if ($dupes -gt 0) { Write-Host "Same-name files are saved as 'name (dup 1)', 'name (dup 2)'... next to the original." -ForegroundColor DarkGray }
}

# ---------- SHA-256 checksum file ----------
function Write-Sha256File($account) {
    Set-Activity Verify "SHA-256 $($account.Email)"
    $root = Join-Path $destination $account.Folder
    $out  = Join-Path $root "_SHA256SUMS.txt"
    $state = Read-State $account
    Write-Host ""
    Write-Host "Writing SHA-256 checksums of every local file in $root" -ForegroundColor Cyan
    Write-Host "(it reads every file, so it takes a while for big folders)"
    & rclone hashsum SHA256 $root --output-file $out --exclude "/_SHA256SUMS.txt" --progress
    if ($LASTEXITCODE -eq 0) {
        $state.ShaDate = (Get-Date -Format "yyyy-MM-dd HH:mm")
        Write-State $account $state
        Write-Log "SHA256 file written: $out" "Green"
        Write-Host "To re-check the files later: rclone checksum SHA256 `"$out`" `"$root`"" -ForegroundColor DarkGray
    } else {
        Write-Log "SHA256 file FAILED for $($account.Email)" "Red"
    }
}

# ---------- remove Mac junk files already downloaded ----------
# ._name (AppleDouble), .DS_Store and Mac system folders: Finder info only, no content.
function Remove-MacJunk($account) {
    Set-Activity Verify "Clean Mac files $($account.Email)"
    $root = Join-Path $destination $account.Folder
    Write-Host ""
    Write-Host "Removing Mac junk files from $root" -ForegroundColor Cyan
    if (-not (Test-Path -LiteralPath $root)) { Write-Host "  Nothing downloaded yet." -ForegroundColor DarkGray; return }
    $junkDirs = @("__MACOSX", ".AppleDouble", ".Spotlight-V100", ".Trashes", ".fseventsd", ".TemporaryItems", ".DocumentRevisions-V100")
    $n = 0; $bytes = 0.0; $failed = 0
    foreach ($d in @(Get-ChildItem -LiteralPath $root -Recurse -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $junkDirs -contains $_.Name })) {
        if (-not (Test-Path -LiteralPath $d.FullName)) { continue }
        $inside = @(Get-ChildItem -LiteralPath $d.FullName -Recurse -File -Force -ErrorAction SilentlyContinue)
        try {
            Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction Stop
            $n += $inside.Count; $bytes += ($inside | Measure-Object Length -Sum).Sum
        } catch { $failed++ }
    }
    Get-ChildItem -LiteralPath $root -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "._*" -or $_.Name -eq ".DS_Store" } |
        ForEach-Object {
            $len = $_.Length
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                $n++; $bytes += $len
                if ($n % 2000 -eq 0) { Write-Host "  $n removed..." }
            } catch { $failed++ }
        }
    $mb = [math]::Round($bytes / 1MB, 1)
    Write-Log "MAC JUNK removed: $($account.Email), $n file(s), $mb MB$(if ($failed) { ", $failed could not be removed" })" $(if ($failed) { "Yellow" } else { "Green" })
}

# ---------- Gmail with Got Your Back ----------
function Find-Gyb {
    $cmd = Get-Command gyb -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    foreach ($p in @("$env:USERPROFILE\bin\gyb\gyb.exe", "C:\GYB\gyb.exe", "$PSScriptRoot\gyb\gyb.exe")) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Backup-Gmail($account) {
    $gyb = Find-Gyb
    if (-not $gyb) {
        Write-Host ""
        Write-Host "Got Your Back (GYB) is not installed." -ForegroundColor Yellow
        Write-Host "Install it from https://github.com/GAM-team/got-your-back/releases (Windows .msi),"
        Write-Host "then open Backup.bat again and use G<number>."
        return
    }
    $gybDir = Split-Path $gyb
    if (-not (Test-Path -LiteralPath (Join-Path $gybDir "client_secrets.json"))) {
        Write-Host ""
        Write-Host "GYB is installed but has no Google API project yet." -ForegroundColor Yellow
        Write-Host "It needs one-time Google Cloud API access. When you have it, run:"
        Write-Host "  `"$gyb`" --action create-project --email $($account.Email)" -ForegroundColor Cyan
        Write-Host "then use G<number> again."
        return
    }

    $state = Read-State $account
    $dir = Join-Path (Join-Path $destination $account.Folder) "Gmail"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Write-Host ""
    Write-Host "--------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Gmail backup for: $($account.Email)" -ForegroundColor Yellow
    Write-Host "Password:         $($account.Password)" -ForegroundColor Yellow
    Write-Host "The first time, GYB asks you to sign in to this account in the browser."
    Write-Host "--------------------------------------------" -ForegroundColor DarkGray
    if (-not $script:appMode) { $null = Read-Host "Press Enter to start" }

    & $gyb --email $account.Email --action backup --local-folder $dir --config-folder $gybDir
    $code = $LASTEXITCODE
    $emails = @(Get-ChildItem -LiteralPath $dir -Recurse -File -Filter *.eml -ErrorAction SilentlyContinue).Count
    $state.GmailDate   = (Get-Date -Format "yyyy-MM-dd HH:mm")
    $state.GmailStatus = $(if ($code -eq 0) { "OK" } else { "FAILED" })
    $state.GmailEmails = $emails
    Write-State $account $state
    Write-Log "GMAIL $($state.GmailStatus): $($account.Email), $emails emails in $dir" $(if ($code -eq 0) { "Green" } else { "Red" })
}

# ---------- Contacts + Calendar (browser) ----------
# Opens Google's own export pages for one account and the folder to save them in.
function Export-ContactsCalendar($account) {
    $dir = Join-Path (Join-Path $destination $account.Folder) "Contactos_Calendario"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $user = [uri]::EscapeDataString($account.Email)

    Write-Host ""
    Write-Host "--------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Contacts + Calendar for: $($account.Email)" -ForegroundColor Yellow
    Write-Host "Password:                $($account.Password)" -ForegroundColor Yellow
    Write-Host "--------------------------------------------" -ForegroundColor DarkGray
    Write-Host "1. CALENDAR: a .zip with all calendars (.ics) downloads by itself."
    Write-Host "2. CONTACTS: click 'Export' -> 'Google CSV' or 'vCard' -> Export."
    Write-Host "3. Move both downloaded files into the folder that opens:"
    Write-Host "   $dir" -ForegroundColor Cyan
    Write-Host "If the browser opens the wrong account, sign in with this one first and press C again."
    if (-not $script:appMode) { $null = Read-Host "Press Enter to open the pages" }

    Start-Process "https://calendar.google.com/calendar/exporticalzip?authuser=$user"
    Start-Process "https://contacts.google.com/?authuser=$user"
    Start-Process explorer.exe -ArgumentList "`"$dir`""
    Write-Log "CONTACTS/CALENDAR export opened: $($account.Email)" "Cyan"
}

# ---------- menu ----------
# Signs in if needed. Sets $script:authOk.
function Confirm-Auth($i) {
    $script:authOk = $true
    if ($script:statuses[$i] -eq "AUTHENTICATED") { return }
    if ($script:statuses[$i] -ne "NOT CONNECTED" -and $script:statuses[$i] -ne "EXPIRED") {
        # Not checked in this session yet: check just this account
        Set-Activity Checking $accounts[$i].Email
        Write-Host "Checking $($accounts[$i].Email)..." -ForegroundColor DarkGray
        $script:statuses[$i] = Get-AuthStatus $accounts[$i] (Get-Remotes)
        Save-AuthState $i
        if ($script:statuses[$i] -eq "AUTHENTICATED") { return }
    }
    Connect-Account $accounts[$i] $script:statuses[$i]
    if ($script:connectOk) {
        $script:statuses[$i] = "AUTHENTICATED"
    } else {
        $script:statuses[$i] = Get-AuthStatus $accounts[$i] (Get-Remotes)
        $script:authOk = $false
    }
    Save-AuthState $i
}

# Pause before going back to the menu (not when the Windows app runs one action)
function Wait-Back($msg) {
    if (-not $Action) { $null = Read-Host $msg }
}

# Table rows (one per account + TOTAL ALL), used by the menu and the Windows app
# Online size of the parts that are ticked (My Drive / Shared with me / Shared drives / Trash)
function Get-CheckedTotal($s) {
    $t = [double]0
    if ($script:wantMy)     { $t += Get-MyDriveBytes $s }
    if ($script:wantShared) { $t += [double](Get-StateValue $s "SharedBytes" 0) }
    if ($script:wantDrives) { $t += [double](Get-StateValue $s "SharedDrivesBytes" 0) }
    if ($includeTrash)      { $t += [double](Get-StateValue $s "TrashBytes" 0) }
    return $t
}

function Get-TableRows {
    $sum = @{ Own = [double]0; Others = [double]0; Sc = [double]0; Used = [double]0; My = [double]0; Shared = [double]0; Trash = [double]0; Drives = [double]0; Total = [double]0; Local = [double]0; Folders = 0; Files = 0 }
    $rows = for ($i = 0; $i -lt $accounts.Count; $i++) {
        $a = $accounts[$i]
        $s = Read-State $a
        if (-not $s.ContainsKey("BackupDate") -and (Test-HasDownload $a)) {
            # Downloaded with an older version of this script: pick it up from the disk
            $mi = Get-Item -LiteralPath (Join-Path (Join-Path $destination $a.Folder) "Mi_unidad")
            $s.BackupDate = "$($mi.LastWriteTime.ToString('yyyy-MM-dd HH:mm')) (found)"
            Update-LocalStats $a $s
            Write-State $a $s
        }
        $pct = Get-StateValue $s "Percent" $null
        if ($script:online.ContainsKey($a.Folder)) { $sum.Used += [double]$script:online[$a.Folder].used }
        elseif ($s.ContainsKey("GoogleUsed")) { $sum.Used += [double]$s.GoogleUsed }
        $sum.My      += Get-MyDriveBytes $s
        $sum.Own     += [double](Get-StateValue $s "MyOwnedBytes" 0)
        $sum.Others  += [double](Get-StateValue $s "MyOthersBytes" 0)
        $sum.Sc      += [double](Get-StateValue $s "MyShortcutBytes" 0)
        $sum.Shared  += [double](Get-StateValue $s "SharedBytes" 0)
        $sum.Trash   += [double](Get-StateValue $s "TrashBytes" 0)
        $sum.Drives  += [double](Get-StateValue $s "SharedDrivesBytes" 0)
        $checkedTotal = Get-CheckedTotal $s
        $sum.Total   += $checkedTotal
        if ($checkedTotal -gt 0 -and $s.ContainsKey("LocalBytes")) { $pct = [math]::Min(100, [math]::Round(100 * [double]$s.LocalBytes / $checkedTotal, 1)) }
        $sum.Local   += [double](Get-StateValue $s "LocalBytes" 0)
        $sum.Folders += [int](Get-StateValue $s "LocalFolders" 0)
        $sum.Files   += [int](Get-StateValue $s "LocalFiles" 0)
        $gmail = Get-StateValue $s "GmailStatus" "-"
        if ($gmail -eq "OK") { $gmail = "$(Get-StateValue $s 'GmailEmails' 0) emails" }
        [pscustomobject]@{
            '#'        = $i + 1
            Auth       = $script:statuses[$i]
            Name       = $a.Name
            Email      = $a.Email
            'Signed in as' = $(if ($script:whoami.ContainsKey($a.Folder)) { $script:whoami[$a.Folder] } else { Get-StateValue $s "SignedInAs" "?" })
            Password   = $a.Password
            Downloaded = Get-StateValue $s "BackupDate" "never"
            'Drive used (Google)' = $(if ($script:online.ContainsKey($a.Folder)) { Format-Bytes $script:online[$a.Folder].used } elseif ($s.ContainsKey("GoogleUsed")) { "$(Format-Bytes $s.GoogleUsed)*" } else { "-" })
            'Online My' = $(if ((Get-MyDriveChoice $s)) { Format-Bytes (Get-MyDriveBytes $s) } else { "-" })
            'My: own'       = Format-SizeFiles $s "MyOwnedBytes" "MyOwnedFiles"
            'My: others'    = Format-SizeFiles $s "MyOthersBytes" "MyOthersFiles"
            'My: shortcuts' = Format-SizeFiles $s "MyShortcutBytes" "MyShortcutFiles"
            'Online Shared' = $(if ($s.ContainsKey("SharedBytes")) { "$(Format-Bytes $s.SharedBytes)$(if (Get-StateValue $s 'SharedSkipped' $false) { ' (skip)' } elseif (Get-StateValue $s 'SharedOnlyExternal' $false) { ' (ext only)' })" } else { "-" })
            'Shared outside' = $(if ($s.ContainsKey("SharedExternalFiles")) { "$(Format-Bytes $s.SharedExternalBytes) ($($s.SharedExternalFiles))" } else { "-" })
            'Online Trash' = $(if (-not $includeTrash) { "skip" } elseif ($s.ContainsKey("TrashBytes")) { "$(Format-Bytes $s.TrashBytes) ($($s.TrashFiles))" } else { "-" })
            'Online Shared drives' = $(if ($s.ContainsKey("SharedDrivesBytes")) { "$(Format-Bytes $s.SharedDrivesBytes) ($($s.SharedDrivesCount) drives)" } else { "-" })
            'Online Total' = $(if ($checkedTotal -gt 0) { Format-Bytes $checkedTotal } else { "-" })
            Local      = $(if ($s.ContainsKey("LocalBytes")) { Format-Bytes $s.LocalBytes } else { "-" })
            'Done%'    = $(if ($null -ne $pct) { "$pct%" } else { "-" })
            Folders    = Get-StateValue $s "LocalFolders" "-"
            Files      = Get-StateValue $s "LocalFiles" "-"
            Verified   = Get-StateValue $s "VerifyStatus" "-"
            Dupes      = Get-StateValue $s "Duplicates" "-"
            Gmail      = $gmail
        }
    }
    $rows = @($rows) + [pscustomobject]@{
        '#' = ""; Auth = "TOTAL ALL"; Name = ""; Email = "(all $($accounts.Count) accounts)"; Password = ""; Downloaded = ""
        'Drive used (Google)' = Format-Bytes $sum.Used; 'Online My' = Format-Bytes $sum.My;
        'My: own' = Format-Bytes $sum.Own; 'My: others' = Format-Bytes $sum.Others; 'My: shortcuts' = Format-Bytes $sum.Sc; 'Online Shared' = Format-Bytes $sum.Shared
        'Online Trash' = Format-Bytes $sum.Trash; 'Online Shared drives' = Format-Bytes $sum.Drives
        'Online Total' = Format-Bytes $sum.Total; Local = Format-Bytes $sum.Local
        'Done%' = $(if ($sum.Total -gt 0) { "$([math]::Min(100, [math]::Round(100 * $sum.Local / $sum.Total, 1)))%" } else { "-" })
        Folders = $sum.Folders; Files = $sum.Files; Verified = ""; Dupes = ""; Gmail = ""
    }
    # Show only the columns of the ticked parts
    $hide = @()
    if (-not $script:wantMy)     { $hide += @("Online My", "My: own", "My: others", "My: shortcuts") }
    if (-not $script:wantShared) { $hide += @("Online Shared", "Shared outside") }
    if (-not $script:wantDrives) { $hide += "Online Shared drives" }
    if (-not $includeTrash)      { $hide += "Online Trash" }
    foreach ($r in $rows) { foreach ($h in $hide) { $r.PSObject.Properties.Remove($h) } }
    return $rows
}

function Show-Menu {
    Set-Activity Menu
    Clear-Host
    Write-Host "=========== GOOGLE DRIVE BACKUP - ACCOUNTS ===========" -ForegroundColor Cyan
    Write-Host "Destination: $destination"
    if ($script:overwrite) { Write-Host "OVERWRITE MODE: ON (downloads re-copy every file)" -ForegroundColor Yellow }
    else { Write-Host "Overwrite mode: off (downloads only copy new or changed files)" -ForegroundColor DarkGray }
    if ($script:includeShared) { Write-Host "Shared with me: ON (asked per account, with its size)" -ForegroundColor DarkGray }
    else { Write-Host "Shared with me: OFF for all accounts (W to turn on)" -ForegroundColor Yellow }

    $rows = Get-TableRows
    $script:tableRows = $rows
    $rows | Format-Table -AutoSize | Out-String -Width 400 | Write-Host
    Write-Host "  <n>  download account n (sign-in first if needed), then verify    A = download + verify ALL authenticated"
    Write-Host "  V<n> verify only (MD5 byte check vs Drive + duplicates)          L = overwrite mode on/off"
    Write-Host "  H<n> SHA-256 checksum file of the local backup                    G<n> Gmail backup (GYB)"
    Write-Host "  C<n> Contacts + Calendar export                                   W = Shared with me on/off"
    Write-Host "  J<n> remove Mac junk files (._name, .DS_Store) from the local backup"
    Write-Host "  I<n> size of each part on Drive (My Drive, Shared with me, ...)   S = sizes of ALL authenticated accounts"
    Write-Host "  X = cross-check Shared with me: which files belong to these accounts and which come from outside"
    Write-Host "  R = check accounts now (sign-in status + Google used + sizes)    T = check sign-in only, all accounts (fast)"
    Write-Host "  N = sign in ALL accounts that need it, one after another          O = open the table in a window (click to choose)"
    Write-Host "  Q = quit"
    if (-not $script:checked) { Write-Host "  * = result saved from the last check. Press R to check now (or start with -Check)." -ForegroundColor Yellow }
    Write-Host "Drive used (Google) = what Google says this sign-in owns in Drive. Signed in as = the real Google account behind it. Online My / Shared / Trash / Shared drives / Total = file sizes on Drive (S or I<n> recalculates)." -ForegroundColor DarkGray
    Write-Host "Tip: an account downloaded before this version shows 'never' / '-': use V<n> to fill its columns." -ForegroundColor DarkGray
}

# Saves the sign-in result and Google's storage number so the next start can show them
function Save-AuthState($i) {
    $a = $accounts[$i]
    $st = Read-State $a
    $st.LastAuth = $script:statuses[$i]
    $st.LastAuthDate = (Get-Date -Format "yyyy-MM-dd HH:mm")
    if ($script:online.ContainsKey($a.Folder)) { $st.GoogleUsed = [double]$script:online[$a.Folder].used }
    if ($script:whoami.ContainsKey($a.Folder)) { $st.SignedInAs = $script:whoami[$a.Folder] }
    Write-State $a $st
}

# The checking: sign-in status of every account + Drive sizes the first time
function Invoke-Checks($authOnly = $false) {
    Set-Activity Checking "all accounts"
    Write-Host "Checking accounts..." -ForegroundColor DarkGray
    $remotes = Get-Remotes
    $script:statuses = @()
    for ($k = 0; $k -lt $accounts.Count; $k++) {
        Resolve-Remote $accounts[$k] $remotes
        Write-Host "  $($accounts[$k].Email)..." -ForegroundColor DarkGray
        $script:statuses += Get-AuthStatus $accounts[$k] $remotes
        Save-AuthState $k
    }
    $script:checked = $true
    if ($authOnly) { return }
    for ($k = 0; $k -lt $accounts.Count; $k++) {
        if ($script:statuses[$k] -ne "AUTHENTICATED") { continue }
        if ((Read-State $accounts[$k]).ContainsKey("MyDriveBytes")) { continue }
        Write-Host "Calculating online sizes for $($accounts[$k].Email) (one time, can take a few minutes for big drives)..." -ForegroundColor DarkGray
        Show-DriveSizes $accounts[$k]
    }
    $script:checked = $true
}

# Without checking: last saved statuses, marked with *
function Set-SavedStatuses {
    $remotes = Get-Remotes
    # Saved sign-in tokens from rclone's config file (local, no internet needed)
    $config = $null
    try { $config = (& rclone config dump 2>$null) | Out-String | ConvertFrom-Json } catch { }
    $script:statuses = @()
    foreach ($acc in $accounts) {
        Resolve-Remote $acc $remotes
        $st = Read-State $acc
        $hasToken = $false
        if ($config -and $config.($acc.Remote) -and $config.($acc.Remote).token -match 'refresh_token') { $hasToken = $true }
        if ($st.ContainsKey("LastAuth")) { $script:statuses += "$($st.LastAuth)*" }
        elseif ($hasToken) { $script:statuses += "AUTHENTICATED*" }
        elseif ($remotes -contains $acc.Remote) { $script:statuses += "NOT CHECKED" }
        else { $script:statuses += "NOT CONNECTED" }
    }
    $script:checked = $false
}

function Confirm-AllChecked {
    if (-not $script:checked) { Invoke-Checks }
}

# One action on one account: "" download, V verify, H SHA-256, G Gmail, C contacts, I sizes
function Invoke-AccountAction($letter, $i) {
    $account = $accounts[$i]
    New-Item -ItemType Directory -Force -Path (Join-Path $destination $account.Folder) | Out-Null
    switch ($letter) {
        "" {
            if ($script:diskFull) { Write-Log "Skipped $($account.Email): the disk is full." "Red"; return }
            Confirm-Auth $i
            if ($script:authOk) {
                Backup-Account $account
                if ($CleanJunk) { Remove-MacJunk $account }
            }
        }
        "V" {
            Confirm-Auth $i
            if ($script:authOk) { Verify-Account $account }
        }
        "H" { Write-Sha256File $account }
        "J" { Remove-MacJunk $account }
        "G" { Backup-Gmail $account }
        "C" { Export-ContactsCalendar $account }
        "I" {
            Confirm-Auth $i
            if ($script:authOk) { Show-DriveSizes $account }
        }
        "N" { Confirm-Auth $i }
    }
}

# The table in a Windows window: select accounts with the mouse, then pick what to do
function Show-WindowPicker {
    $rows = @($script:tableRows | Where-Object { "$($_.'#')" -ne "" })
    $picked = @($rows | Out-GridView -Title "Google Drive Backup - select one or more accounts (Ctrl+click), then OK" -OutputMode Multiple)
    if ($picked.Count -eq 0) { return }
    $actions = @(
        [pscustomobject]@{ Key = "";  Action = "Download (sign in first if needed), then verify" }
        [pscustomobject]@{ Key = "N"; Action = "Sign in only" }
        [pscustomobject]@{ Key = "I"; Action = "Sizes on Drive" }
        [pscustomobject]@{ Key = "V"; Action = "Verify only (byte check vs Drive)" }
        [pscustomobject]@{ Key = "H"; Action = "SHA-256 checksum file" }
        [pscustomobject]@{ Key = "C"; Action = "Contacts + Calendar export" }
    )
    $names = ($picked | ForEach-Object { $_.Email }) -join ", "
    $act = $actions | Out-GridView -Title "What to do with: $names" -OutputMode Single
    if (-not $act) { return }
    if ($act.Key -eq "") {
        $script:runOverwrite = $false
        if ($picked.Count -gt 1 -and -not $script:overwrite) {
            Read-OverwriteChoice "Selected accounts"
            if ($script:runCancel) { return }
            $script:choiceMadeForAll = $true
        }
    }
    foreach ($row in $picked) { Invoke-AccountAction $act.Key ([int]$row.'#' - 1) }
    $script:choiceMadeForAll = $false
    $script:runOverwrite = $false
    $null = Read-Host "Done. Press Enter to go back to the menu"
}

if ($Check) { Invoke-Checks } else { Set-SavedStatuses }

# ---------- one action from the Windows app ----------
$pending = New-Object System.Collections.Queue
if ($Action) {
    $script:overwrite = [bool]$OverwriteAll
    $script:includeShared = -not $NoShared
    if ($Action -eq "Export") {
        # Table for the app: last saved results, no internet needed
        $out = Join-Path $logDir "table.json"
        $json = @(Get-TableRows) | ConvertTo-Json -Depth 3
        [IO.File]::WriteAllText($out, $json, (New-Object System.Text.UTF8Encoding($false)))
        exit 0
    }
    $perAccount = @{ Download = ""; SignIn = "N"; Sizes = "I"; Verify = "V"; Sha = "H"; Contacts = "C"; Clean = "J" }
    if ($perAccount.ContainsKey($Action)) {
        $letter = $perAccount[$Action]
        $acctList = @("$Select" -split '[,; ]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
        if ($letter -eq "" -and $acctList.Count -gt 1 -and -not $script:overwrite) {
            Read-OverwriteChoice "Selected accounts"
            if ($script:runCancel) { exit 0 }
            $script:choiceMadeForAll = $true
        }
        foreach ($n in $acctList) { $pending.Enqueue("$letter$n") }
    } else {
        $pending.Enqueue($Action.ToUpper())
    }
    $pending.Enqueue("Q")
}

while ($true) {

    if ($pending.Count -gt 0) {
        $choice = "$($pending.Dequeue())"
        if ($choice -eq "Q") {
            Remove-Item -LiteralPath (Join-Path $logDir "progress.json") -Force -ErrorAction SilentlyContinue
            Save-ErrorLog
            if ($script:appMode) { Write-Host "=== Finished ===" }
            else { $null = Read-Host "Finished. Press Enter to close this window" }
            break
        }
    } else {
        Show-Menu
        $choice = "$(Read-Host 'Choice')".Trim().ToUpper()
    }
    if ($choice -notmatch '^([A-Z]?)\s*(\d*)$') {
        Write-Host "'$choice' is not a valid option." -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue
    }
    $letter = $Matches[1]
    $num = 0
    if ($Matches[2]) { $num = [int]$Matches[2] }

    if ($choice -eq "Q") { break }
    if ($choice -eq "R") { Invoke-Checks; continue }
    if ($choice -eq "T") { Invoke-Checks $true; continue }
    if ($choice -eq "O") { Show-WindowPicker; continue }
    if ($choice -eq "N") {
        # Sign in, one after another, every account that is not signed in correctly
        Invoke-Checks $true
        for ($k = 0; $k -lt $accounts.Count; $k++) {
            if ($script:statuses[$k] -eq "AUTHENTICATED") { continue }
            Confirm-Auth $k
        }
        Wait-Back "Sign-ins done. Press Enter to go back to the menu"
        continue
    }
    if ($choice -eq "L") { $script:overwrite = -not $script:overwrite; continue }
    if ($choice -eq "W") { $script:includeShared = -not $script:includeShared; continue }
    if ($choice -eq "X") {
        Confirm-AllChecked
        Compare-SharedWithMe
        Wait-Back "Press Enter to go back to the menu"
        continue
    }
    if ($choice -eq "S") {
        Confirm-AllChecked
        for ($k = 0; $k -lt $accounts.Count; $k++) {
            if ($script:statuses[$k] -eq "AUTHENTICATED") { Show-DriveSizes $accounts[$k] }
        }
        Wait-Back "Sizes saved in the table. Press Enter to go back to the menu"
        continue
    }

    if ($choice -eq "A") {
        Confirm-AllChecked
        $script:runOverwrite = $false
        if (-not $script:overwrite) {
            Read-OverwriteChoice "Some accounts"
            if ($script:runCancel) { continue }
            $script:choiceMadeForAll = $true
        }
        $skipped = @()
        for ($i = 0; $i -lt $accounts.Count; $i++) {
            if ($script:diskFull) { $skipped += $accounts[$i].Email; continue }
            if ($script:statuses[$i] -eq "AUTHENTICATED") {
                Backup-Account $accounts[$i]
                if ($CleanJunk) { Remove-MacJunk $accounts[$i] }
            }
            else { $skipped += $accounts[$i].Email }
        }
        $script:choiceMadeForAll = $false
        $script:runOverwrite = $false
        if ($skipped.Count -gt 0) { Write-Host "Skipped (not authenticated): $($skipped -join ', ')" -ForegroundColor Yellow }
        Wait-Back "All done. Press Enter to go back to the menu"
        continue
    }

    if ($num -lt 1 -or $num -gt $accounts.Count -or ($letter -ne "" -and "VHGCINJ" -notmatch $letter)) {
        Write-Host "'$choice' is not a valid option." -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue
    }

    Invoke-AccountAction $letter ($num - 1)
    Wait-Back "Press Enter to go back to the menu"
}

Set-Activity Menu
Save-ErrorLog
Write-Host "Bye. Logs are in $logDir"
# Exit code for the app: 3 = stopped because the disk is full (the download can be continued)
if ($script:diskFull) { exit 3 }

# ---------- HOW TO: your own Google OAuth client (fixes "Quota exceeded") ----------
# Any Google account works (the admin one is best). Takes about 10 minutes, free.
#  1. https://console.cloud.google.com/projectcreate  -> name "drive-backup" -> Create
#  2. https://console.cloud.google.com/apis/library/drive.googleapis.com -> Enable
#  3. https://console.cloud.google.com/auth/branding -> Get started -> app name "drive-backup",
#     your email -> Audience: "Internal" (Workspace) or "External" -> Create
#     If External: Audience -> "Publish app" (or add the 8 emails as test users)
#  4. https://console.cloud.google.com/auth/clients -> Create client -> "Desktop app" -> Create
#  5. Copy the Client ID and Client secret into settings.json (ClientId / ClientSecret).
#  6. Reopen Backup.bat: every account shows OLD SIGN-IN; type its number to sign in again.
