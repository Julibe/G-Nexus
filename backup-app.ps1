# ============================================================
# G Nexus - Google Drive Backup (Windows app)
# A window with the accounts table, buttons and a live log.
# The real work is done by backup-drive.ps1 (same folder): each button
# opens it in its own window for that one action (sign-in prompts and
# the progress bar show there), and the table refreshes when it ends.
# Start it with "Backup App.bat".
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$here      = Split-Path -Parent $MyInvocation.MyCommand.Path
$engine    = Join-Path $here "backup-drive.ps1"
$errorFile = Join-Path $here "_app_error.txt"

# ---------- data folder ----------
# Everything G Nexus makes goes to the backup folder the user chose (the "data folder"):
# the account folders, _logs, and _gnexus with settings.json, accounts.csv and session.json.
# The app folder keeps only the app. %APPDATA%\G Nexus\location.json remembers the data folder.
$pointerDir  = Join-Path $env:APPDATA "G Nexus"
$pointerFile = Join-Path $pointerDir "location.json"

# Makes <data>\_gnexus and fills what is missing: from $fromDir (an older copy or the previous
# data folder) or else from the example files next to the scripts.
function Initialize-DataFolder($data, $fromDir) {
    $cfg = Join-Path $data "_gnexus"
    New-Item -ItemType Directory -Force -Path $cfg | Out-Null
    foreach ($pair in @(@("settings.json", "settings.example.json"), @("accounts.csv", "accounts.example.csv"), @("session.json", ""))) {
        $target = Join-Path $cfg $pair[0]
        if (Test-Path -LiteralPath $target) { continue }
        $from = $(if ($fromDir) { Join-Path $fromDir $pair[0] } else { "" })
        if ($from -and (Test-Path -LiteralPath $from)) { Copy-Item -LiteralPath $from -Destination $target }
        elseif ($pair[1] -and (Test-Path -LiteralPath (Join-Path $here $pair[1]))) { Copy-Item -LiteralPath (Join-Path $here $pair[1]) -Destination $target }
    }
}

function Set-DataFolder($data, $fromDir) {
    Initialize-DataFolder $data $fromDir
    New-Item -ItemType Directory -Force -Path $pointerDir | Out-Null
    [ordered]@{ DataFolder = $data } | ConvertTo-Json | Set-Content -LiteralPath $pointerFile -Encoding UTF8
    $script:dataDir = $data
    $script:cfgDir = Join-Path $data "_gnexus"
    $script:settingsPath = Join-Path $script:cfgDir "settings.json"
    $script:accountsPath = Join-Path $script:cfgDir "accounts.csv"
}

$script:dataDir = $null
try { $script:dataDir = "$((Get-Content -LiteralPath $pointerFile -Raw -Encoding UTF8 | ConvertFrom-Json).DataFolder)" } catch { }
if (-not $script:dataDir) {
    # First start: an older copy kept its settings next to the scripts, with the location inside
    $legacy = $null
    try { $legacy = Get-Content -LiteralPath (Join-Path $here "settings.json") -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    if ($legacy -and "$($legacy.Destination)".Trim()) { $script:dataDir = "$($legacy.Destination)".Trim() }
    elseif ($legacy) { $script:dataDir = $here }
    else {
        [System.Windows.Forms.MessageBox]::Show("Welcome to G Nexus.`n`nChoose the folder where the backups go. G Nexus also keeps its settings and accounts there (in a _gnexus folder).", "G Nexus") | Out-Null
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Where should G Nexus save the backups and its settings?"
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq "OK") { $script:dataDir = $dlg.SelectedPath }
        else { $script:dataDir = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "G Nexus" }
    }
}
Set-DataFolder $script:dataDir $here
$settingsPath = $script:settingsPath

function Get-Settings {
    try { return (Get-Content -LiteralPath $script:settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}
function Set-Paths {
    $script:dest = $script:dataDir
    $script:logDir       = Join-Path $script:dest "_logs"
    $script:tableFile    = Join-Path $script:logDir "table.json"
    $script:progressFile = Join-Path $script:logDir "progress.json"
    $script:outputFile   = Join-Path $script:logDir "app_output.log"
    $script:skipFile     = Join-Path $script:logDir "skip.flag"
}
Set-Paths

function Set-Destination {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Where should G Nexus save the backup? (each account gets its own folder inside)"
    $dlg.SelectedPath = $script:dest
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog($form) -ne "OK") { return }
    Switch-DataFolder $dlg.SelectedPath
    $txtDest.Text = $script:dest
    $title.Text = "Saving to:  $($script:dest)"
    Update-Table
    $lblStatus.Text = "Backup folder: $($script:dest)  (settings and accounts are in its _gnexus folder; accounts downloaded elsewhere start fresh here)"
}
# New data folder: settings, accounts and session come along when it has none yet
function Switch-DataFolder($newDir) {
    if (-not $newDir -or $newDir -eq $script:dataDir) { return }
    Set-DataFolder $newDir $script:cfgDir
    Set-Paths
    $txtDest.Text = $script:dest
    $title.Text = "Saving to:  $($script:dest)"
}
$script:proc = $null
$script:running = ""

# ---------- dark theme ----------
function New-Color($r, $g, $b) { return [System.Drawing.Color]::FromArgb($r, $g, $b) }
$cBack     = New-Color 14 30 60      # window
$cPanel    = New-Color 20 42 82      # button bar, dialogs
$cGrid     = New-Color 17 36 70      # table
$cText     = New-Color 235 240 250
$cDim      = New-Color 150 172 208
$cLine     = New-Color 46 74 124
$cBtn      = New-Color 30 58 108
$cGreen    = New-Color 34 120 70
$cBlue     = New-Color 40 90 160
$cPurple   = New-Color 105 70 150
$cGold     = New-Color 170 130 30
$cRowOk    = New-Color 28 64 40
$cRowBad   = New-Color 80 32 36
$cRowWait  = New-Color 70 62 28
$cRowTotal = New-Color 34 60 104

# ---------- engine ----------
function Get-EngineArgs($action, $numbers, $resume) {
    $a = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$engine`"", "-Action", $action, "-DataDir", "`"$($script:dataDir)`"")
    if ($numbers -and @($numbers).Count -gt 0) { $a += @("-Select", (@($numbers) -join ",")) }
    if ($resume) { $a += "-Resume" }   # continue: never overwrite what is already there
    elseif ($chkOverwrite.Checked) { $a += "-OverwriteAll" }
    if (-not $chkVerify.Checked) { $a += "-NoVerify" }
    if (-not $chkJunk.Checked) { $a += "-KeepJunk" }
    if ($chkNoShortcuts.Checked) { $a += "-NoShortcuts" }
    if ($chkOwned.Checked) { $a += "-OwnedOnly" }
    $mode = @("E", "A")[[math]::Max(0, $cmbShared.SelectedIndex)]
    if (-not $chkPartShared.Checked) { $a += "-NoShared" } else { $a += @("-SharedMode", $mode) }
    $a += @("-Parts", (Get-PartsArg))
    return $a
}

# The ticked parts, for the engine: "My,Shared,Drives,Trash"
function Get-PartsArg {
    $pp = @()
    if ($chkPartMy.Checked) { $pp += "My" }
    if ($chkPartShared.Checked) { $pp += "Shared" }
    if ($chkPartDrives.Checked) { $pp += "Drives" }
    if ($chkPartTrash.Checked) { $pp += "Trash" }
    if ($pp.Count -eq 0) { $pp += "none" }
    return ($pp -join ",")
}

# ---------- pause / continue ----------
# A download writes resume.json when it starts and removes it when it ends normally.
# Pause (or closing the app, a crash, a full disk) leaves it, so the download can continue later.
function Get-ResumePath { return (Join-Path $script:cfgDir "resume.json") }
function Read-Resume {
    $f = Get-ResumePath
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    try { return (Get-Content -LiteralPath $f -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}
function Write-Resume($action, $numbers, $text) {
    try {
        $emails = @(foreach ($n in @($numbers)) { if ($script:rowsByNum.ContainsKey([int]$n)) { "$(Get-RowValue $script:rowsByNum[[int]$n] 'Email')" } })
        [ordered]@{ Action = $action; Emails = $emails; Text = $text; Time = (Get-Date -Format "yyyy-MM-dd HH:mm") } |
            ConvertTo-Json | Set-Content -LiteralPath (Get-ResumePath) -Encoding UTF8
    } catch { Save-AppErrors }
}
function Clear-Resume { Remove-Item -LiteralPath (Get-ResumePath) -Force -ErrorAction SilentlyContinue }
function Update-ResumeButton {
    $btnContinue.Enabled = ($null -ne (Read-Resume)) -and -not ($script:proc -and -not $script:proc.HasExited)
}
function Resume-Download {
    $r = Read-Resume
    if ($null -eq $r) { return }
    $nums = @()
    if ("$($r.Action)" -ne "A") {
        foreach ($k in @($script:rowsByNum.Keys)) { if (@($r.Emails) -contains "$(Get-RowValue $script:rowsByNum[$k] 'Email')") { $nums += [int]$k } }
        if ($nums.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("The paused accounts are not in the list any more:`n$(@($r.Emails) -join ', ')", "Continue") | Out-Null
            Clear-Resume; Update-ResumeButton; return
        }
    }
    Start-Engine "$($r.Action)" @($nums | Sort-Object) "$($r.Text) (continued)" -resume
}

# Runs one action in its own window; the table refreshes when it ends
function Start-Engine($action, $numbers, $text, [switch]$resume) {
    if ($script:proc -and -not $script:proc.HasExited) {
        [System.Windows.Forms.MessageBox]::Show("Something is already running: $($script:running)`nWait for it to finish, or press Stop.", "Busy") | Out-Null
        return
    }
    Remove-Item -LiteralPath $progressFile, $skipFile -Force -ErrorAction SilentlyContinue
    $txtLog.Text = ""
    $style = $(if ($chkConsole.Checked) { "Normal" } else { "Hidden" })
    $script:isDownload = ($action -eq "Download" -or $action -eq "A")
    if ($script:isDownload) { Write-Resume $action $numbers ($text -replace ' \(continued\)$', '') }
    $script:paused = $false
    $script:proc = Start-Process powershell -ArgumentList (Get-EngineArgs $action $numbers $resume) -WindowStyle $style -PassThru
    $null = $script:proc.Handle   # keeps ExitCode readable after it ends
    $script:running = $text
    Set-Busy $true
    Show-Tab "Home"   # Skip / Stop are there
}

# Reads the saved results (no internet needed) and fills the table
function Update-Table {
    $lblStatus.Text = "Reading saved results..."
    $form.Refresh()
    $a = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$engine`"", "-Action", "Export", "-Parts", (Get-PartsArg), "-DataDir", "`"$($script:dataDir)`"")
    if ($chkNoShortcuts.Checked) { $a += "-NoShortcuts" }
    if ($chkOwned.Checked) { $a += "-OwnedOnly" }
    $p = Start-Process powershell -ArgumentList $a -WindowStyle Hidden -PassThru
    if (-not $p.WaitForExit(90000)) {
        try { $p.Kill() } catch { }
        $lblStatus.Text = "Could not read the results (took too long). Try 'Old menu' to see what happens."
        return
    }
    if (Test-Path -LiteralPath $errorFile) {
        $msg = (Get-Content -LiteralPath $errorFile -Raw -Encoding UTF8)
        $lblStatus.Text = "ERROR: $msg"
        $lblStatus.ForeColor = New-Color 255 120 120
        return
    }
    if (-not (Test-Path -LiteralPath $tableFile)) { $lblStatus.Text = "No results yet."; return }
    # Windows PowerShell returns a JSON array as one object: unroll it into rows
    $parsed = Get-Content -LiteralPath $tableFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $rows = @($parsed | ForEach-Object { $_ })
    $dt = New-Object System.Data.DataTable
    if ($rows.Count -gt 0) {
        foreach ($prop in $rows[0].PSObject.Properties) { [void]$dt.Columns.Add($prop.Name, [string]) }
        foreach ($r in $rows) {
            $dr = $dt.NewRow()
            foreach ($prop in $r.PSObject.Properties) {
                if ($dt.Columns.Contains($prop.Name)) { $dr[$prop.Name] = "$($prop.Value)" }
            }
            $dt.Rows.Add($dr)
        }
    }
    $grid.DataSource = $dt
    try { Update-Dashboard $rows } catch { Save-AppErrors }
    foreach ($col in $grid.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    Set-RowColors
    $lblStatus.Text = "Ready. Results marked * are from the last check (press 'Check sign-in' to refresh)."
}

function Set-RowColors {
    if (-not $grid.Columns.Contains("Auth")) { return }
    foreach ($row in $grid.Rows) {
        $auth = "$($row.Cells['Auth'].Value)"
        $color = $cRowWait
        if ($auth -eq "TOTAL ALL") {
            $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
            $color = $cRowTotal
        } elseif ($auth -like "AUTHENTICATED*") { $color = $cRowOk }
        elseif ($auth -match "EXPIRED|WRONG|NOT CONNECTED|OLD") { $color = $cRowBad }
        $row.DefaultCellStyle.BackColor = $color
        $row.DefaultCellStyle.ForeColor = $cText
        if ($grid.Columns.Contains("Verified") -and "$($row.Cells['Verified'].Value)" -like "VERIFIED*") {
            $row.Cells['Verified'].Style.BackColor = $cGreen
        }
    }
}

function Get-SelectedNumbers {
    if ($dash.Visible) { return @($script:selected | Sort-Object) }
    $nums = @()
    foreach ($row in $grid.SelectedRows) {
        $n = "$($row.Cells['#'].Value)"
        if ($n -match '^\d+$') { $nums += [int]$n }
    }
    return @($nums | Sort-Object)
}

function Get-SelectedEmails {
    if ($dash.Visible) { return ((@($script:selected | Sort-Object) | ForEach-Object { Get-RowValue $script:rowsByNum[$_] "Email" }) -join ", ") }
    $e = @()
    foreach ($row in $grid.SelectedRows) { if ("$($row.Cells['#'].Value)" -match '^\d+$') { $e += "$($row.Cells['Email'].Value)" } }
    return ($e -join ", ")
}

function Set-Busy($busy) {
    foreach ($b in $script:actionButtons) { $b.Enabled = -not $busy }
    $btnSkip.Enabled = $busy
    $btnStop.Enabled = $busy
    $btnPause.Enabled = $busy -and $script:isDownload
    Update-ResumeButton
    if ($busy) {
        $lblStatus.Text = "RUNNING: $($script:running)   -   'Pause' stops it so you can continue later, 'Stop' cancels"
        $lblStatus.ForeColor = New-Color 120 220 140
    } else {
        $lblStatus.ForeColor = $cText
    }
}

function Update-Log {
    # Output of the running (or last) action; the daily summary when nothing ran yet
    $f = $null
    if (Test-Path -LiteralPath $outputFile) { $f = Get-Item -LiteralPath $outputFile }
    if (-not $f -or $f.Length -lt 5) {
        $f = Get-ChildItem -LiteralPath $logDir -Filter "summary_*.log" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime | Select-Object -Last 1
    }
    if (-not $f) { return }
    $text = (Get-Content -LiteralPath $f.FullName -Tail 300 -Encoding UTF8 -ErrorAction SilentlyContinue) -join "`r`n"
    if ($txtLog.Text -ne $text) {
        $txtLog.Text = $text
        $txtLog.SelectionStart = $txtLog.Text.Length
        $txtLog.ScrollToCaret()
    }
}

# ---------- errors and support report ----------
function Save-AppErrors {
    try {
        if ($Error.Count -eq 0) { return }
        $lines = @()
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        for ($k = [math]::Min($Error.Count, 50) - 1; $k -ge 0; $k--) {
            $e = $Error[$k]
            $lines += "$stamp  $($e.Exception.Message)"
            if ($e.InvocationInfo) { $lines += "      at line $($e.InvocationInfo.ScriptLineNumber): $("$($e.InvocationInfo.Line)".Trim())" }
        }
        Add-Content -LiteralPath (Join-Path $logDir "app_errors.log") -Value $lines -Encoding UTF8
        $Error.Clear()
    } catch { }
}

# Packs logs, errors and system info (no passwords) into one .zip to send for support
function New-SupportReport {
    $lblStatus.Text = "Making the support report..."
    $form.Refresh()
    Save-AppErrors
    $stamp = Get-Date -Format "yyyyMMdd_HHmm"
    $tmp = Join-Path $env:TEMP "GNexus_report_$stamp"
    $zip = Join-Path $script:dataDir "G-Nexus-report_$stamp.zip"
    try {
        New-Item -ItemType Directory -Force -Path $tmp | Out-Null
        $info = @()
        $info += "G Nexus support report  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $info += "App folder: $here"
        $info += "Download location: $($script:dest)"
        try { $os = Get-CimInstance Win32_OperatingSystem; $info += "Windows: $($os.Caption) $($os.Version)" } catch { }
        $info += "PowerShell: $($PSVersionTable.PSVersion)"
        try { $info += "rclone: " + ((& rclone version 2>&1 | Select-Object -First 1) -join "") } catch { $info += "rclone: NOT FOUND" }
        try {
            $d = Get-PSDrive -Name $script:dest.Substring(0, 1)
            $info += "Disk $($script:dest.Substring(0, 2)) free: $([math]::Round($d.Free / 1GB, 1)) GB, used: $([math]::Round($d.Used / 1GB, 1)) GB"
        } catch { }
        if ($script:proc -and -not $script:proc.HasExited) { $info += "Running now: $($script:running)" }
        $info += ""
        $info += "Accounts (no passwords):"
        try { foreach ($r in @(Import-Csv -LiteralPath $script:accountsPath -Encoding UTF8)) { $info += "  $($r.Name) | $($r.Email) | $($r.Folder) | $($r.Remote)" } } catch { $info += "  accounts.csv unreadable: $($_.Exception.Message)" }
        $info += ""
        $info += "settings.json (secret hidden):"
        try { $info += ((Get-Content -LiteralPath $script:settingsPath -Raw -Encoding UTF8) -replace '("ClientSecret"\s*:\s*")[^"]*', '$1***') } catch { }
        $info += ""
        $info += "rclone remotes (tokens hidden):"
        try { $info += @(& rclone config redacted 2>&1 | ForEach-Object { "$_" }) } catch { }
        Set-Content -LiteralPath (Join-Path $tmp "system.txt") -Value $info -Encoding UTF8

        $logsOut = Join-Path $tmp "_logs"
        New-Item -ItemType Directory -Force -Path (Join-Path $logsOut "state") | Out-Null
        foreach ($n in @("app_output.log", "errors.log", "app_errors.log", "table.json", "progress.json")) {
            $f = Join-Path $logDir $n
            if (Test-Path -LiteralPath $f) { Copy-Item -LiteralPath $f -Destination $logsOut }
        }
        Get-ChildItem -LiteralPath (Join-Path $logDir "state") -Filter *.json -ErrorAction SilentlyContinue | Copy-Item -Destination (Join-Path $logsOut "state")
        Get-ChildItem -LiteralPath $logDir -Filter "summary_*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 3 | Copy-Item -Destination $logsOut
        Get-ChildItem -LiteralPath $logDir -Filter "*_PROBLEMS_*.txt" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 3 | Copy-Item -Destination $logsOut
        # Last rclone logs, only their end (they can be huge)
        Get-ChildItem -LiteralPath $logDir -Filter "*.log" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch '^(summary_|app_|errors)' } | Sort-Object LastWriteTime | Select-Object -Last 8 |
            ForEach-Object { Get-Content -LiteralPath $_.FullName -Tail 600 -ErrorAction SilentlyContinue | Set-Content -LiteralPath (Join-Path $logsOut $_.Name) -Encoding UTF8 }

        Compress-Archive -Path (Join-Path $tmp "*") -DestinationPath $zip -Force
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        Start-Process explorer.exe "/select,`"$zip`""
        $lblStatus.Text = "Support report saved: $zip  (send this file)"
    } catch {
        $lblStatus.Text = "Could not make the report: $($_.Exception.Message)"
    }
}

# ---------- instructions and accounts ----------
# Small Markdown -> HTML for INSTRUCTIONS.md (headings, lists, tables, checkboxes, quotes, links, bold, code)
function Convert-InlineMd($t) {
    $t = [System.Net.WebUtility]::HtmlEncode($t)
    $t = [regex]::Replace($t, '`([^`]+)`', '<code>$1</code>')
    $t = [regex]::Replace($t, '\*\*([^*]+)\*\*', '<b>$1</b>')
    $t = [regex]::Replace($t, '(?<![\w*])\*([^*]+)\*(?![\w*])', '<i>$1</i>')
    $t = [regex]::Replace($t, '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>')
    return $t
}
function Convert-MdToHtml($lines) {
    $sb = New-Object System.Text.StringBuilder
    $list = ""; $inTable = $false; $para = @()
    foreach ($raw in $lines) {
        $line = "$raw"
        $isItem = $line -match '^\s*(-|\d+\.)\s+'
        $isRow  = $line -match '^\s*\|'
        if (-not $isItem -and $list -and $line -match '^\s{2,}\S' ) { [void]$sb.Append(" " + (Convert-InlineMd $line.Trim())); continue }
        if ($para.Count -gt 0 -and ($line.Trim() -eq "" -or $isItem -or $isRow -or $line -match '^(#|>)')) {
            [void]$sb.Append("<p>" + (Convert-InlineMd ($para -join " ")) + "</p>"); $para = @()
        }
        if ($list -and -not $isItem -and $line -notmatch '^\s+-') {
            if ($line.Trim() -eq "" -or $line -match '^\S') { [void]$sb.Append("</li></$list>"); $list = "" }
        }
        if ($inTable -and -not $isRow) { [void]$sb.Append("</table>"); $inTable = $false }
        if ($line.Trim() -eq "") { continue }
        if ($line -match '^(#{1,3})\s+(.*)$') {
            $h = $matches[1].Length
            [void]$sb.Append("<h$h>" + (Convert-InlineMd $matches[2]) + "</h$h>"); continue
        }
        if ($line -match '^>\s?(.*)$') { [void]$sb.Append("<blockquote>" + (Convert-InlineMd $matches[1]) + "</blockquote>"); continue }
        if ($isRow) {
            if ($line -match '^\s*\|[\s:|-]+\|\s*$') { continue }
            $cells = @($line.Trim().Trim('|') -split '\|' | ForEach-Object { Convert-InlineMd $_.Trim() })
            if (-not $inTable) {
                [void]$sb.Append("<table><tr>" + (($cells | ForEach-Object { "<th>$_</th>" }) -join "") + "</tr>"); $inTable = $true
            } else {
                [void]$sb.Append("<tr>" + (($cells | ForEach-Object { "<td>$_</td>" }) -join "") + "</tr>")
            }
            continue
        }
        if ($line -match '^\s*(-|\d+\.)\s+(.*)$') {
            $kind = $(if ($matches[1] -eq "-") { "ul" } else { "ol" })
            $body = $matches[2]
            $nested = $line -match '^\s{2,}-'
            if ($nested) { [void]$sb.Append("<br>&bull; " + (Convert-InlineMd $body)); continue }
            if ($list -and $list -ne $kind) { [void]$sb.Append("</li></$list>"); $list = "" }
            if (-not $list) { [void]$sb.Append("<$kind>"); $list = $kind } else { [void]$sb.Append("</li>") }
            if ($body -match '^\[ \]\s*(.*)$') { $body = "&#9744; " + (Convert-InlineMd $matches[1]) } else { $body = Convert-InlineMd $body }
            [void]$sb.Append("<li>$body"); continue
        }
        $para += $line.Trim()
    }
    if ($para.Count -gt 0) { [void]$sb.Append("<p>" + (Convert-InlineMd ($para -join " ")) + "</p>") }
    if ($list) { [void]$sb.Append("</li></$list>") }
    if ($inTable) { [void]$sb.Append("</table>") }
    return $sb.ToString()
}
function ConvertTo-HtmlColor($c) { return "#{0:X2}{1:X2}{2:X2}" -f $c.R, $c.G, $c.B }

function Show-Instructions {
    $f = Join-Path $here "INSTRUCTIONS.md"
    $lines = @("# Instructions not found", "", "The file INSTRUCTIONS.md is missing from $here.")
    if (Test-Path -LiteralPath $f) { $lines = @(Get-Content -LiteralPath $f -Encoding UTF8) }
    $bg = ConvertTo-HtmlColor $cGrid; $fg = ConvertTo-HtmlColor $cText; $dim = ConvertTo-HtmlColor $cDim
    $line = ConvertTo-HtmlColor $cLine; $panel = ConvertTo-HtmlColor $cPanel
    $css = "body{background:$bg;color:$fg;font-family:'Segoe UI',sans-serif;font-size:14px;line-height:1.5;margin:24px 36px;}" +
           "h1{font-size:26px;color:#F0B840;margin:0 0 8px;}h2{font-size:19px;color:#5AC8E6;border-bottom:1px solid $line;padding-bottom:4px;margin-top:28px;}" +
           "a{color:#64AAFF;}code{background:$panel;padding:1px 5px;font-family:Consolas,monospace;font-size:13px;}" +
           "table{border-collapse:collapse;margin:8px 0;}th,td{border:1px solid $line;padding:6px 10px;text-align:left;vertical-align:top;}th{background:$panel;}" +
           "blockquote{border-left:4px solid #F0B840;margin:12px 0;padding:6px 14px;background:$panel;color:$fg;}li{margin:3px 0;}i{color:$dim;}"
    $html = "<!DOCTYPE html><html><head><meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=edge'><style>$css</style></head><body>" +
            (Convert-MdToHtml $lines) + "</body></html>"
    $hf = New-Object System.Windows.Forms.Form
    $hf.Text = "G Nexus - How to use and connect"
    $hf.Size = New-Object System.Drawing.Size(980, 780)
    $hf.StartPosition = "CenterParent"
    if ($script:appIcon) { $hf.Icon = $script:appIcon }
    $hf.BackColor = $cGrid
    $wb = New-Object System.Windows.Forms.WebBrowser
    $wb.Dock = "Fill"
    $wb.IsWebBrowserContextMenuEnabled = $false
    $wb.ScriptErrorsSuppressed = $true
    # Links open in the normal browser
    $wb.add_Navigating({
        param($s, $e)
        if ($e.Url.Scheme -like "http*") { $e.Cancel = $true; Start-Process $e.Url.AbsoluteUri }
    })
    $hf.Controls.Add($wb)
    $wb.DocumentText = $html
    [void]$hf.ShowDialog($form)
}

function Show-AddAccount {
    $af = New-Object System.Windows.Forms.Form
    $af.Text = "G Nexus - Add account"
    $af.Size = New-Object System.Drawing.Size(520, 285)
    $af.StartPosition = "CenterParent"
    if ($script:appIcon) { $af.Icon = $script:appIcon }
    $af.FormBorderStyle = "FixedDialog"
    $af.MaximizeBox = $false
    $af.MinimizeBox = $false
    $af.BackColor = $cPanel
    $af.ForeColor = $cText
    $fields = @(
        @("Name", "e.g. GT Diseno"),
        @("Email", "name@company.com"),
        @("Password", "reminder only, shown during sign-in"),
        @("Folder", "optional: backup folder name (made from Name if empty)")
    )
    $boxes = @{}
    $y = 15
    foreach ($fd in $fields) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $fd[0]
        $l.Location = New-Object System.Drawing.Point(15, ($y + 3))
        $l.Size = New-Object System.Drawing.Size(80, 20)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(100, $y)
        $t.Size = New-Object System.Drawing.Size(390, 22)
        $h = New-Object System.Windows.Forms.Label
        $h.Text = $fd[1]
        $h.ForeColor = $cDim
        $t.BackColor = $cGrid
        $t.ForeColor = $cText
        $t.BorderStyle = "FixedSingle"
        $h.Location = New-Object System.Drawing.Point(100, ($y + 24))
        $h.Size = New-Object System.Drawing.Size(390, 16)
        $af.Controls.AddRange(@($l, $t, $h))
        $boxes[$fd[0]] = $t
        $y += 46
    }
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Add"
    $ok.Location = New-Object System.Drawing.Point(310, ($y + 2))
    $ok.Size = New-Object System.Drawing.Size(85, 28)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Location = New-Object System.Drawing.Point(405, ($y + 2))
    $cancel.Size = New-Object System.Drawing.Size(85, 28)
    $cancel.DialogResult = "Cancel"
    foreach ($bb in @($ok, $cancel)) { $bb.FlatStyle = "Flat"; $bb.ForeColor = [System.Drawing.Color]::White; $bb.BackColor = $cBtn }
    $ok.BackColor = $cGreen
    $af.CancelButton = $cancel
    $af.AcceptButton = $ok
    $af.Controls.AddRange(@($ok, $cancel))
    $ok.add_Click({
        $email = $boxes["Email"].Text.Trim()
        if ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            [System.Windows.Forms.MessageBox]::Show("Write a valid email.", "Add account") | Out-Null
            return
        }
        $csv = $script:accountsPath
        $list = @()
        if (Test-Path -LiteralPath $csv) { $list = @(Import-Csv -LiteralPath $csv -Encoding UTF8) }
        if (@($list | Where-Object { "$($_.Email)".Trim() -eq $email }).Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("$email is already in the list.", "Add account") | Out-Null
            return
        }
        $list += [pscustomobject]@{
            Name     = $boxes["Name"].Text.Trim()
            Email    = $email
            Password = $boxes["Password"].Text
            Folder   = $boxes["Folder"].Text.Trim()
            Remote   = ""
            Legacy   = ""
        }
        $list | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $af.DialogResult = "OK"
        $af.Close()
    })
    if ($af.ShowDialog($form) -eq "OK") {
        Update-Table
        $lblStatus.Text = "Account added. Select it and press 'Sign in'."
    }
}

# ---------- dashboard (cards) ----------
# One card per account + summary tiles. Click a card to select it (click again to unselect).
$script:selected = New-Object System.Collections.Generic.HashSet[int]
$script:rowsByNum = @{}
$script:cardFrames = @{}

function Get-RowValue($r, $name) {
    $p = $r.PSObject.Properties[$name]
    if ($p) { return "$($p.Value)" }
    return ""
}

function Set-CardLook($num) {
    $frame = $script:cardFrames[$num]
    if (-not $frame) { return }
    $frame.BackColor = $(if ($script:selected.Contains($num)) { $icGold } else { $cLine })
}

function Switch-CardSelection($num) {
    if ($script:selected.Contains($num)) { [void]$script:selected.Remove($num) } else { [void]$script:selected.Add($num) }
    Set-CardLook $num
    $n = $script:selected.Count
    $lblSel.Text = $(if ($n -eq 0) { "Click a card to select it" } else { "$n selected" })
}

function New-Tile($title, $value, $color) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Size = New-Object System.Drawing.Size(190, 70)
    $p.Margin = New-Object System.Windows.Forms.Padding(6)
    $p.BackColor = $cPanel
    $v = New-Object System.Windows.Forms.Label
    $v.Text = $value
    $v.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $v.ForeColor = $color
    $v.Location = New-Object System.Drawing.Point(12, 6)
    $v.Size = New-Object System.Drawing.Size(175, 32)
    $t = New-Object System.Windows.Forms.Label
    $t.Text = $title
    $t.ForeColor = $cDim
    $t.Location = New-Object System.Drawing.Point(13, 42)
    $t.Size = New-Object System.Drawing.Size(175, 20)
    $p.Controls.AddRange(@($v, $t))
    return $p
}

function New-Card($r) {
    $num = [int](Get-RowValue $r "#")
    $auth = Get-RowValue $r "Auth"
    $ok = $auth -like "AUTHENTICATED*"
    $bad = $auth -match "EXPIRED|WRONG|NOT CONNECTED|OLD"
    $dotColor = $(if ($ok) { $icGreen } elseif ($bad) { $icRed } else { $icGold })
    $stateText = $(if ($ok) { "Signed in" } elseif ($auth -like "WRONG*") { "Wrong account - sign in again" } elseif ($auth -like "OLD*") { "Sign in again (new Google app)" } elseif ($bad) { "Needs sign-in" } else { "Not checked yet" })

    $frame = New-Object System.Windows.Forms.Panel
    $frame.Size = New-Object System.Drawing.Size(330, 220)
    $frame.Margin = New-Object System.Windows.Forms.Padding(8)
    $frame.Padding = New-Object System.Windows.Forms.Padding(2)
    $card = New-Object System.Windows.Forms.Panel
    $card.Dock = "Fill"
    $card.BackColor = $cPanel
    $frame.Controls.Add($card)

    function Add-L($text, $x, $y, $w, $size, $bold, $color) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Location = New-Object System.Drawing.Point($x, $y)
        $l.Size = New-Object System.Drawing.Size($w, ([int]($size * 2.2) + 2))
        $style = $(if ($bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular })
        $l.Font = New-Object System.Drawing.Font("Segoe UI", $size, $style)
        $l.ForeColor = $(if ($color) { $color } else { $cText })
        $l.AutoEllipsis = $true
        $card.Controls.Add($l)
        return $l
    }

    [void](Add-L ([string][char]0x25CF) 10 8 22 13 $false $dotColor)
    [void](Add-L "$(Get-RowValue $r 'Name')" 32 9 230 11 $true $null)
    [void](Add-L "#$num" 280 11 40 9 $false $cDim)
    [void](Add-L "$(Get-RowValue $r 'Email')" 32 34 290 9 $false $cDim)
    [void](Add-L $stateText 32 54 290 9 $true $dotColor)

    $pct = 0
    $done = Get-RowValue $r "Done%"
    if ($done -match '([\d\.]+)') { $pct = [int][math]::Min(100, [double]$Matches[1]) }
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(14, 84)
    $bar.Size = New-Object System.Drawing.Size(250, 14)
    $bar.Value = $pct
    $card.Controls.Add($bar)
    [void](Add-L "$(if ($done -and $done -ne '-') { $done } else { '0%' })" 270 81 55 9 $true $cText)

    $online = Get-RowValue $r "Online Total"; if (-not $online -or $online -eq "-") { $online = "?" }
    $local  = Get-RowValue $r "Local"; if (-not $local -or $local -eq "-") { $local = "0" }
    [void](Add-L "On Drive:  $online" 14 108 160 9 $false $null)
    [void](Add-L "Downloaded:  $local" 170 108 155 9 $false $null)
    $files = Get-RowValue $r "Files"; if (-not $files -or $files -eq "-") { $files = "0" }
    $when = Get-RowValue $r "Downloaded"; if (-not $when) { $when = "never" }
    [void](Add-L "Files: $files" 14 130 160 9 $false $cDim)
    [void](Add-L "Last: $when" 170 130 155 9 $false $cDim)
    $own = Get-RowValue $r "My: own"; $oth = Get-RowValue $r "My: others"; $sc = Get-RowValue $r "My: shortcuts"
    if ($own -and $own -ne "-") {
        $short = { param($v) if ($v -and $v -ne "-") { ($v -replace '\s*\(.*$', '') } else { "-" } }
        $mixLbl = Add-L "Own $(& $short $own)  |  Others $(& $short $oth)  |  Shortcuts $(& $short $sc)" 14 152 310 8.5 $false $cDim
        $tips.SetToolTip($mixLbl, "My Drive`nOwn files: $own`nFiles of other owners: $oth`nShortcuts: $sc")
    }
    $ver = Get-RowValue $r "Verified"
    $verOk = $ver -like "VERIFIED*"
    $verText = $(if ($verOk) { "VERIFIED - safe" } elseif ($ver -and $ver -ne "-") { "Check: $ver" } else { "Not verified yet" })
    [void](Add-L $verText 14 180 300 9 $true $(if ($verOk) { $icGreen } elseif ($ver -and $ver -ne "-") { $icRed } else { $cDim }))

    # Click anywhere on the card to select it
    $frame.Tag = $num
    $card.Tag = $num
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.add_Click({ param($sender, $e) Switch-CardSelection ([int]$sender.Tag) })
    foreach ($c in $card.Controls) {
        $c.Tag = $num
        $c.Cursor = [System.Windows.Forms.Cursors]::Hand
        $c.add_Click({ param($sender, $e) Switch-CardSelection ([int]$sender.Tag) })
    }
    $tips.SetToolTip($card, "Signed in as: $(Get-RowValue $r 'Signed in as')`nGoogle says: $(Get-RowValue $r 'Drive used (Google)')`nFolders: $(Get-RowValue $r 'Folders')   Duplicates: $(Get-RowValue $r 'Dupes')")
    $script:cardFrames[$num] = $frame
    Set-CardLook $num
    return $frame
}

function Update-Dashboard($rows) {
    $cards.SuspendLayout()
    $cards.Controls.Clear()
    $tiles.Controls.Clear()
    $script:rowsByNum = @{}
    $script:cardFrames = @{}
    $accountRows = @($rows | Where-Object { "$(Get-RowValue $_ '#')" -match '^\d+$' })
    $total = @($rows | Where-Object { (Get-RowValue $_ "Auth") -eq "TOTAL ALL" }) | Select-Object -First 1
    foreach ($r in $accountRows) { $script:rowsByNum[[int](Get-RowValue $r "#")] = $r }
    # Forget selections of accounts that are gone
    foreach ($n in @($script:selected)) { if (-not $script:rowsByNum.ContainsKey($n)) { [void]$script:selected.Remove($n) } }

    $signed = @($accountRows | Where-Object { (Get-RowValue $_ "Auth") -like "AUTHENTICATED*" }).Count
    $verified = @($accountRows | Where-Object { (Get-RowValue $_ "Verified") -like "VERIFIED*" }).Count
    $tiles.Controls.Add((New-Tile "Accounts signed in" "$signed / $($accountRows.Count)" $icBlue))
    if ($total) {
        $tiles.Controls.Add((New-Tile "On Drive (ticked parts)" "$(Get-RowValue $total 'Online Total')" $icGold))
        $tiles.Controls.Add((New-Tile "Downloaded" "$(Get-RowValue $total 'Local')" $icGreen))
        $tiles.Controls.Add((New-Tile "Done" "$(Get-RowValue $total 'Done%')" $icGreen))
    }
    $tiles.Controls.Add((New-Tile "Verified (safe to delete)" "$verified / $($accountRows.Count)" $icPurple))

    foreach ($r in $accountRows) { $cards.Controls.Add((New-Card $r)) }
    $cards.ResumeLayout()
    $n = $script:selected.Count
    $lblSel.Text = $(if ($n -eq 0) { "Click a card to select it" } else { "$n selected" })
}

# ---------- window ----------
$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form = New-Object System.Windows.Forms.Form
$form.Text = "G Nexus - Google Drive Migration Engine"
$form.Size = New-Object System.Drawing.Size(1400, 820)
$form.StartPosition = "CenterScreen"
$form.Font = $font
$form.WindowState = "Maximized"
$form.BackColor = $cBack
$form.ForeColor = $cText
$script:appIcon = $null
try { $script:appIcon = New-Object System.Drawing.Icon (Join-Path $here "gnexus.ico") } catch { }
if ($script:appIcon) { $form.Icon = $script:appIcon }

# Header: logo, name, slogan, backup location
$cCyan = New-Color 90 200 230
$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 96
$header.BackColor = $cBack
function New-Picture($file, $x, $y) {
    $path = Join-Path $here $file
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $pb = New-Object System.Windows.Forms.PictureBox
    # Read into memory so the image file is not locked while G Nexus is open
    $ms = New-Object System.IO.MemoryStream(,[System.IO.File]::ReadAllBytes($path))
    $pb.Image = [System.Drawing.Image]::FromStream($ms)
    $pb.SizeMode = "AutoSize"
    $pb.Location = New-Object System.Drawing.Point($x, $y)
    $pb.BackColor = $cBack
    $header.Controls.Add($pb)
    return $pb
}
$picMark = New-Picture "gnexus-symbol.png" 24 16
$picWord = New-Picture "gnexus-wordmark.png" 160 20
if (-not $picWord) {
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "G NEXUS"
    $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblName.AutoSize = $true
    $lblName.ForeColor = $cText
    $lblName.Location = New-Object System.Drawing.Point(160, 14)
    $header.Controls.Add($lblName)
}
$lblSlogan = New-Object System.Windows.Forms.Label
$lblSlogan.Text = "Google Drive Migration Engine"
$lblSlogan.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$lblSlogan.AutoSize = $true
$lblSlogan.ForeColor = $cCyan
$lblSlogan.Location = New-Object System.Drawing.Point(162, 56)
$header.Controls.Add($lblSlogan)
$title = New-Object System.Windows.Forms.Label
$title.Text = "Saving to:  $($script:dest)"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$title.Dock = "Right"
$title.Width = 700
$title.TextAlign = "MiddleRight"
$title.Padding = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
$title.ForeColor = $cDim
$title.BackColor = $cBack
$header.Controls.Add($title)
# Gold line under the header, like the ring in the logo
$headerLine = New-Object System.Windows.Forms.Panel
$headerLine.Dock = "Top"
$headerLine.Height = 2
$headerLine.BackColor = $cGold

$dash = New-Object System.Windows.Forms.Panel
$dash.Dock = "Fill"
$dash.BackColor = $cGrid
$cards = New-Object System.Windows.Forms.FlowLayoutPanel
$cards.Dock = "Fill"
$cards.AutoScroll = $true
$cards.BackColor = $cGrid
$cards.Padding = New-Object System.Windows.Forms.Padding(8)
$tiles = New-Object System.Windows.Forms.FlowLayoutPanel
$tiles.Dock = "Top"
$tiles.AutoSize = $true
$tiles.BackColor = $cBack
$tiles.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 2)
$selBar = New-Object System.Windows.Forms.FlowLayoutPanel
$selBar.Dock = "Top"
$selBar.Height = 30
$selBar.BackColor = $cGrid
$selBar.Padding = New-Object System.Windows.Forms.Padding(14, 4, 0, 0)
$lblSel = New-Object System.Windows.Forms.Label
$lblSel.Text = "Click a card to select it"
$lblSel.AutoSize = $true
$lblSel.ForeColor = $cDim
$lblSel.Margin = New-Object System.Windows.Forms.Padding(0, 4, 12, 0)
$lnkAll = New-Object System.Windows.Forms.LinkLabel
$lnkAll.Text = "Select all"
$lnkAll.AutoSize = $true
$lnkAll.LinkColor = $cText
$lnkAll.Margin = New-Object System.Windows.Forms.Padding(0, 4, 12, 0)
$lnkAll.add_LinkClicked({ foreach ($k in @($script:rowsByNum.Keys)) { [void]$script:selected.Add([int]$k); Set-CardLook ([int]$k) }; $lblSel.Text = "$($script:selected.Count) selected" })
$lnkNone = New-Object System.Windows.Forms.LinkLabel
$lnkNone.Text = "Select none"
$lnkNone.AutoSize = $true
$lnkNone.LinkColor = $cText
$lnkNone.Margin = New-Object System.Windows.Forms.Padding(0, 4, 12, 0)
$lnkNone.add_LinkClicked({ $old = @($script:selected); $script:selected.Clear(); foreach ($k in $old) { Set-CardLook $k }; $lblSel.Text = "Click a card to select it" })
$selBar.Controls.AddRange(@($lblSel, $lnkAll, $lnkNone))
$dash.Controls.Add($cards)
$dash.Controls.Add($selBar)
$dash.Controls.Add($tiles)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = "AllCells"
$grid.BackgroundColor = $cGrid
$grid.BorderStyle = "None"
$grid.GridColor = $cLine
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = $cPanel
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $cText
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersBorderStyle = "Single"
$grid.DefaultCellStyle.BackColor = $cGrid
$grid.DefaultCellStyle.ForeColor = $cText
$grid.DefaultCellStyle.SelectionBackColor = New-Color 0 110 200
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$grid.RowsDefaultCellStyle.SelectionBackColor = New-Color 0 110 200
$grid.add_DataBindingComplete({ Set-RowColors })

# ---------- settings window ----------
function Show-SettingsForm {
    $cfg = Get-Settings
    if (-not $cfg) { $cfg = [pscustomobject]@{} }
    function Get-Val($name, $default) { if ($cfg.PSObject.Properties[$name] -and "$($cfg.$name)" -ne "") { return $cfg.$name } return $default }

    $sf = New-Object System.Windows.Forms.Form
    $sf.Text = "G Nexus - Settings"
    $sf.Size = New-Object System.Drawing.Size(700, 720)
    $sf.StartPosition = "CenterParent"
    if ($script:appIcon) { $sf.Icon = $script:appIcon }
    $sf.FormBorderStyle = "FixedDialog"
    $sf.MaximizeBox = $false
    $sf.MinimizeBox = $false
    $sf.BackColor = $cPanel
    $sf.ForeColor = $cText
    $sf.Font = $font

    $script:sfY = 14
    function Add-Section($text) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $l.ForeColor = $icGold
        $l.Location = New-Object System.Drawing.Point(16, $script:sfY)
        $l.Size = New-Object System.Drawing.Size(650, 22)
        $sf.Controls.Add($l)
        $script:sfY += 26
    }
    function Add-Row($label, $control, $hint) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $label
        $l.Location = New-Object System.Drawing.Point(28, ($script:sfY + 3))
        $l.Size = New-Object System.Drawing.Size(170, 20)
        $sf.Controls.Add($l)
        $control.Location = New-Object System.Drawing.Point(200, $script:sfY)
        if ($control -is [System.Windows.Forms.TextBox] -or $control -is [System.Windows.Forms.ComboBox] -or $control -is [System.Windows.Forms.NumericUpDown]) {
            $control.BackColor = $cGrid
            $control.ForeColor = $cText
        }
        $sf.Controls.Add($control)
        if ($hint) {
            $h = New-Object System.Windows.Forms.Label
            $h.Text = $hint
            $h.ForeColor = $cDim
            $h.Font = New-Object System.Drawing.Font("Segoe UI", 8)
            $h.Location = New-Object System.Drawing.Point(200, ($script:sfY + 25))
            $h.Size = New-Object System.Drawing.Size(460, 16)
            $sf.Controls.Add($h)
            $script:sfY += 18
        }
        $script:sfY += 32
    }
    function New-Text($value, $width) {
        $t = New-Object System.Windows.Forms.TextBox
        $t.Text = "$value"
        $t.Width = $width
        $t.BorderStyle = "FixedSingle"
        return $t
    }
    function New-Number($value, $min, $max) {
        $n = New-Object System.Windows.Forms.NumericUpDown
        $n.Minimum = $min
        $n.Maximum = $max
        $n.Value = [math]::Max($min, [math]::Min($max, [int]$value))
        $n.Width = 70
        return $n
    }
    function Add-Link($text, $url, $x, $y) {
        $ll = New-Object System.Windows.Forms.LinkLabel
        $ll.Text = $text
        $ll.AutoSize = $true
        $ll.LinkColor = $icBlue
        $ll.ActiveLinkColor = $icGold
        $ll.Location = New-Object System.Drawing.Point($x, $y)
        $ll.Tag = $url
        $ll.add_LinkClicked({ param($sender, $e) Start-Process $sender.Tag })
        $sf.Controls.Add($ll)
    }

    Add-Section "Download location"
    $tDest = New-Text $script:dataDir 360
    Add-Row "Save backups in" $tDest "Each account gets its own folder. Logs: _logs. Settings and accounts: _gnexus."
    $bBrowse = New-Object System.Windows.Forms.Button
    $bBrowse.Text = "Browse..."
    $bBrowse.FlatStyle = "Flat"
    $bBrowse.BackColor = $cBtn
    $bBrowse.Location = New-Object System.Drawing.Point(570, ($script:sfY - 51))
    $bBrowse.Size = New-Object System.Drawing.Size(90, 25)
    $bBrowse.add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.SelectedPath = $tDest.Text
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog($sf) -eq "OK") { $tDest.Text = $dlg.SelectedPath }
    })
    $sf.Controls.Add($bBrowse)
    $tMore = New-Object System.Windows.Forms.TextBox
    $tMore.Multiline = $true
    $tMore.ScrollBars = "Vertical"
    $tMore.BorderStyle = "FixedSingle"
    $tMore.Size = New-Object System.Drawing.Size(360, 50)
    $tMore.Text = (@(Get-Val "MoreLocations" @()) | ForEach-Object { "$_" }) -join "`r`n"
    Add-Row "More places" $tMore ""
    $script:sfY += 26
    $hMore = New-Object System.Windows.Forms.Label
    $hMore.Text = "One folder per line, e.g. E:\Backups. When the first disk is full, My Drive folders go to the next place with room."
    $hMore.ForeColor = $cDim
    $hMore.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $hMore.Location = New-Object System.Drawing.Point(200, ($script:sfY - 6))
    $hMore.Size = New-Object System.Drawing.Size(470, 30)
    $sf.Controls.Add($hMore)
    $script:sfY += 26

    Add-Section "Google app (your own, avoids 'Quota exceeded')"
    $tId = New-Text (Get-Val "ClientId" "") 460
    Add-Row "Client ID" $tId ""
    $tSecret = New-Text (Get-Val "ClientSecret" "") 460
    $tSecret.UseSystemPasswordChar = $true
    Add-Row "Client secret" $tSecret "After changing these, every account must sign in again once."

    Add-Section "Speed"
    $nTransfers = New-Number (Get-Val "Transfers" 8) 1 32
    Add-Row "Files at the same time" $nTransfers "8 is safe. Lower it to 4 if Google says 'Quota exceeded'."
    $cBw = New-Object System.Windows.Forms.ComboBox
    $cBw.Width = 160
    foreach ($o in @("No limit", "2M", "5M", "10M", "20M", "50M")) { [void]$cBw.Items.Add($o) }
    $bw = "$(Get-Val 'BwLimit' '')"
    $cBw.Text = $(if ($bw) { $bw } else { "No limit" })
    Add-Row "Speed limit" $cBw "MB per second (10M = 10 MB/s), so the internet stays usable for others."
    $nRetries = New-Number (Get-Val "Retries" 2) 0 10
    Add-Row "Retries per part" $nRetries "How many times a failed part is tried again."

    Add-Section "Google Docs, Sheets and Slides are saved as"
    $cFmt = New-Object System.Windows.Forms.ComboBox
    $cFmt.DropDownStyle = "DropDownList"
    $cFmt.Width = 460
    $presets = [ordered]@{
        "Microsoft Office (docx, xlsx, pptx)" = "docx,xlsx,pptx,pdf,odt,ods,odp,csv,tsv,rtf,txt,html,epub,md,json,svg"
        "OpenDocument (odt, ods, odp)"        = "odt,ods,odp,pdf,docx,xlsx,pptx,csv,tsv,rtf,txt,html,epub,md,json,svg"
        "PDF"                                 = "pdf,xlsx,pptx,docx,odt,ods,odp,csv,tsv,rtf,txt,html,epub,md,json,svg"
    }
    foreach ($k in $presets.Keys) { [void]$cFmt.Items.Add($k) }
    $curFmt = "$(Get-Val 'ExportFormats' $presets[0])"
    $cFmt.SelectedIndex = 0
    $i = 0
    foreach ($k in $presets.Keys) { if ($presets[$k] -eq $curFmt) { $cFmt.SelectedIndex = $i }; $i++ }
    Add-Row "Format" $cFmt "Other files (Photoshop, Illustrator, PDF, video...) are always saved as they are."

    Add-Section "Defaults"
    $chkTrashDef = New-Object System.Windows.Forms.CheckBox
    $chkTrashDef.Text = "Tick 'Trash' when G Nexus opens"
    $chkTrashDef.AutoSize = $true
    $chkTrashDef.Checked = [bool](Get-Val "IncludeTrash" $false)
    Add-Row "Trash" $chkTrashDef ""
    $chkJunkDef = New-Object System.Windows.Forms.CheckBox
    $chkJunkDef.Text = "Skip Mac junk files (._name, .DS_Store, __MACOSX)"
    $chkJunkDef.AutoSize = $true
    $chkJunkDef.Checked = [bool](Get-Val "SkipJunk" $true)
    Add-Row "Mac files" $chkJunkDef "Files a Mac leaves next to real files. They hold no content."

    Add-Section "Useful links"
    $ly = $script:sfY
    Add-Link "Install rclone" "https://rclone.org/downloads/" 28 $ly
    Add-Link "Create the Google app" "https://console.cloud.google.com/auth/clients" 140 $ly
    Add-Link "Turn on Google Drive API" "https://console.cloud.google.com/apis/library/drive.googleapis.com" 300 $ly
    Add-Link "Google Admin console" "https://admin.google.com/ac/users" 480 $ly
    $script:sfY += 34

    $bSave = New-Object System.Windows.Forms.Button
    $bSave.Text = "Save"
    $bSave.FlatStyle = "Flat"
    $bSave.BackColor = $cGreen
    $bSave.ForeColor = [System.Drawing.Color]::White
    $bSave.Size = New-Object System.Drawing.Size(100, 30)
    $bSave.Location = New-Object System.Drawing.Point(450, $script:sfY)
    $bCancel = New-Object System.Windows.Forms.Button
    $bCancel.Text = "Cancel"
    $bCancel.FlatStyle = "Flat"
    $bCancel.BackColor = $cBtn
    $bCancel.Size = New-Object System.Drawing.Size(100, 30)
    $bCancel.Location = New-Object System.Drawing.Point(560, $script:sfY)
    $bCancel.DialogResult = "Cancel"
    $bAdv = New-Object System.Windows.Forms.LinkLabel
    $bAdv.Text = "Advanced: open settings.json"
    $bAdv.AutoSize = $true
    $bAdv.LinkColor = $cDim
    $bAdv.Location = New-Object System.Drawing.Point(28, ($script:sfY + 7))
    $bAdv.add_LinkClicked({ Start-Process notepad.exe "`"$($script:settingsPath)`"" })
    $sf.Controls.AddRange(@($bSave, $bCancel, $bAdv))
    $sf.CancelButton = $bCancel
    $sf.ClientSize = New-Object System.Drawing.Size(690, ($script:sfY + 46))

    $bSave.add_Click({
        $destText = $tDest.Text.Trim()
        $script:newDataDir = $destText
        $bwText = $cBw.Text.Trim()
        if ($bwText -eq "No limit") { $bwText = "" }
        $new = [ordered]@{
            Destination   = $destText
            ClientId      = $tId.Text.Trim()
            ClientSecret  = $tSecret.Text.Trim()
            IncludeTrash  = $chkTrashDef.Checked
            SkipJunk      = $chkJunkDef.Checked
            Transfers     = [int]$nTransfers.Value
            Checkers      = [int]$nTransfers.Value
            Retries       = [int]$nRetries.Value
            BwLimit       = $bwText
            ExportFormats = $presets[$cFmt.SelectedIndex]
            MoreLocations = @($tMore.Lines | ForEach-Object { "$_".Trim().TrimEnd('\') } | Where-Object { $_ })
        }
        try {
            $new | ConvertTo-Json | Set-Content -LiteralPath $script:settingsPath -Encoding UTF8
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not save settings: $($_.Exception.Message)", "Settings") | Out-Null
            return
        }
        $sf.DialogResult = "OK"
        $sf.Close()
    })

    if ($sf.ShowDialog($form) -eq "OK") {
        $chkJunk.Checked = $chkJunkDef.Checked
        if ($script:newDataDir) { Switch-DataFolder $script:newDataDir }
        Set-Paths
        $txtDest.Text = $script:dest
        $title.Text = "Saving to:  $($script:dest)"
        Update-Table
        $lblStatus.Text = "Settings saved."
    }
}

# ---------- ribbon ----------
# Groups of big icon buttons (icons are glyphs of the Windows "Segoe MDL2 Assets" font)
# Tabs like Office: each tab is its own row of groups; only the chosen one is visible
$script:tabPages = [ordered]@{}
$script:tabButtons = @{}
function New-TabPage($name) {
    $pg = New-Object System.Windows.Forms.FlowLayoutPanel
    $pg.Dock = "Top"
    $pg.AutoSize = $true
    $pg.WrapContents = $true
    $pg.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 2)
    $pg.BackColor = $cPanel
    $pg.Visible = $false
    $script:tabPages[$name] = $pg
    $script:ribbon = $pg
    return $pg
}
function Show-Tab($name) {
    foreach ($k in @($script:tabPages.Keys)) {
        $on = ($k -eq $name)
        $script:tabPages[$k].Visible = $on
        $tb = $script:tabButtons[$k]
        $tb.BackColor = $(if ($on) { $cPanel } else { $cBack })
        $tb.ForeColor = $(if ($on) { $icGold } else { $cDim })
    }
}
$tabStrip = New-Object System.Windows.Forms.FlowLayoutPanel
$tabStrip.Dock = "Top"
$tabStrip.Height = 32
$tabStrip.BackColor = $cBack
$tabStrip.Padding = New-Object System.Windows.Forms.Padding(8, 2, 0, 0)
function Add-TabButton($name) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $name
    $b.AutoSize = $true
    $b.Height = 30
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 2, 0)
    $b.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = $cBtn
    $b.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $b.BackColor = $cBack
    $b.ForeColor = $cDim
    $b.Tag = $name
    $b.add_Click({ param($sender, $e) Show-Tab $sender.Tag })
    $tabStrip.Controls.Add($b)
    $script:tabButtons[$name] = $b
}
$script:actionButtons = @()
$tips = New-Object System.Windows.Forms.ToolTip

function New-Icon($glyph, $color) {
    $bmp = New-Object System.Drawing.Bitmap(30, 30)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $f = New-Object System.Drawing.Font("Segoe MDL2 Assets", 17)
    $br = New-Object System.Drawing.SolidBrush($color)
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = "Center"
    $fmt.LineAlignment = "Center"
    $g.DrawString([string][char]$glyph, $f, $br, (New-Object System.Drawing.RectangleF(0, 0, 30, 30)), $fmt)
    $g.Dispose()
    return $bmp
}

function New-Group($title) {
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.AutoSize = $true
    $t.ColumnCount = 1
    $t.RowCount = 2
    $t.Margin = New-Object System.Windows.Forms.Padding(2, 0, 2, 0)
    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.AutoSize = $true
    $flow.WrapContents = $false
    $flow.Margin = New-Object System.Windows.Forms.Padding(0)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $title
    $lbl.Dock = "Fill"
    $lbl.Height = 18
    $lbl.TextAlign = "MiddleCenter"
    $lbl.ForeColor = $cDim
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $t.Controls.Add($flow, 0, 0)
    $t.Controls.Add($lbl, 0, 1)
    $ribbon.Controls.Add($t)
    $sep = New-Object System.Windows.Forms.Panel
    $sep.Width = 1
    $sep.Height = 92
    $sep.BackColor = $cLine
    $sep.Margin = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
    $ribbon.Controls.Add($sep)
    return $flow
}

function Add-RibbonButton($group, $text, $glyph, $accent, $tip, $onClick, [switch]$always) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Image = New-Icon $glyph $accent
    $b.TextImageRelation = "ImageAboveText"
    $b.ImageAlign = "TopCenter"
    $b.TextAlign = "BottomCenter"
    $b.Size = New-Object System.Drawing.Size(86, 72)
    $b.Margin = New-Object System.Windows.Forms.Padding(1)
    $b.Font = New-Object System.Drawing.Font("Segoe UI", 8.25)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = $cBtn
    $b.FlatAppearance.MouseDownBackColor = $cLine
    $b.BackColor = $cPanel
    $b.ForeColor = $cText
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($tip) { $tips.SetToolTip($b, $tip) }
    $b.add_Click($onClick)
    $group.Controls.Add($b)
    if (-not $always) { $script:actionButtons += $b }
    return $b
}

function Invoke-ForSelected($action, $what) {
    $nums = Get-SelectedNumbers
    if ($nums.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select one or more accounts in the table first (Ctrl+click for several).", "No account selected") | Out-Null
        return
    }
    Start-Engine $action $nums "$what - $(Get-SelectedEmails)"
}

$icGreen  = New-Color 90 210 130
$icBlue   = New-Color 100 170 255
$icPurple = New-Color 190 140 255
$icGold   = New-Color 240 190 70
$icRed    = New-Color 255 105 105
$icGray   = New-Color 200 200 210

[void](New-TabPage "Home")
$gSel = New-Group "Selected accounts"
[void](Add-RibbonButton $gSel "Download" 0xE896 $icGreen "Download the selected accounts (only the ticked parts)" { Invoke-ForSelected "Download" "Download" })
[void](Add-RibbonButton $gSel "Sign in" 0xE77B $icBlue "Sign in the selected accounts in the browser" { Invoke-ForSelected "SignIn" "Sign in" })
[void](Add-RibbonButton $gSel "Sizes" 0xE8EF $icBlue "Measure the selected accounts on Drive" { Invoke-ForSelected "Sizes" "Sizes" })
[void](Add-RibbonButton $gSel "Verify" 0xE73E $icPurple "Byte check: downloaded files vs Drive" { Invoke-ForSelected "Verify" "Verify" })
[void](Add-RibbonButton $gSel "SHA-256" 0xE72E $icPurple "Write a SHA-256 checksum list of the local backup" { Invoke-ForSelected "Sha" "SHA-256" })
[void](Add-RibbonButton $gSel "Clean Mac`nfiles" 0xE74D $icRed "Delete the Mac junk files (._name, .DS_Store, __MACOSX) already downloaded. They hold no content." {
    $r = [System.Windows.Forms.MessageBox]::Show("Delete the Mac junk files (._name, .DS_Store, __MACOSX folders) from the downloaded backup of the selected accounts?`n`nThey only hold Mac Finder info, no file content. Nothing on Google Drive is touched.", "Clean Mac files", "YesNo", "Question")
    if ($r -eq "Yes") { Invoke-ForSelected "Clean" "Clean Mac files" }
})
[void](Add-RibbonButton $gSel "Contacts`nCalendar" 0xE787 $icGray "Open Google's Contacts and Calendar export pages" { Invoke-ForSelected "Contacts" "Contacts + Calendar" })

$gAll = New-Group "All accounts"
[void](Add-RibbonButton $gAll "Check`nsign-in" 0xE895 $icBlue "Check the sign-in of every account (fast)" { Start-Engine "T" @() "Check sign-in of all accounts" })
[void](Add-RibbonButton $gAll "Sign in all" 0xE716 $icBlue "Sign in, one after another, every account that needs it" { Start-Engine "N" @() "Sign in all accounts that need it" })
[void](Add-RibbonButton $gAll "Sizes of all" 0xE9D2 $icBlue "Measure every signed-in account on Drive" { Start-Engine "S" @() "Sizes of all accounts" })
[void](Add-RibbonButton $gAll "Cross-check" 0xE72D $icGray "Find 'Shared with me' files that belong to these accounts" { Start-Engine "X" @() "Shared with me cross-check" })
[void](Add-RibbonButton $gAll "Download`nALL" 0xE896 $icGreen "Download every signed-in account, one after another" {
    $r = [System.Windows.Forms.MessageBox]::Show("Download ALL signed-in accounts? This can take many hours.", "Download ALL", "YesNo")
    if ($r -eq "Yes") { Start-Engine "A" @() "Download ALL accounts" }
})

$gRun = New-Group "Running"
$btnSkip = Add-RibbonButton $gRun "Skip step" 0xE893 $icGold "Jump over the step that is stuck" {
    Set-Content -LiteralPath $script:skipFile -Value "skip"
    $lblStatus.Text = "Skipping the current step..."
} -always
$btnPause = Add-RibbonButton $gRun "Pause" 0xE769 $icGold "Pause the download. Continue it later, also after closing G Nexus or turning off the PC." {
    if ($script:proc -and -not $script:proc.HasExited) {
        $script:paused = $true
        & taskkill.exe /PID $script:proc.Id /T /F | Out-Null
    }
} -always
$btnContinue = Add-RibbonButton $gRun "Continue" 0xE768 $icGreen "Continue the paused download where it stopped" { Resume-Download } -always
$btnStop = Add-RibbonButton $gRun "Stop" 0xE71A $icRed "Cancel what is running (downloaded files are kept)" {
    if ($script:proc -and -not $script:proc.HasExited) {
        $r = [System.Windows.Forms.MessageBox]::Show("Stop $($script:running)?`nFiles already downloaded are kept.`n(To continue it later, use Pause instead.)", "Stop", "YesNo")
        if ($r -eq "Yes") { Clear-Resume; & taskkill.exe /PID $script:proc.Id /T /F | Out-Null }
    }
} -always
$btnStopAll = Add-RibbonButton $gRun "Stop all" 0xE7BA $icRed "Stop EVERY backup running on this PC, also ones left in the background" {
    $n = @(Get-EngineProcesses).Count
    $r = [System.Windows.Forms.MessageBox]::Show("Stop all $n backup action(s) running on this PC?`nFiles already downloaded are kept.", "Stop all", "YesNo")
    if ($r -ne "Yes") { return }
    foreach ($lp in @(Get-EngineProcesses)) { Stop-EngineTree $lp.ProcessId }
    foreach ($rp in @(Get-Process rclone -ErrorAction SilentlyContinue)) { try { $rp.Kill() } catch { } }
    $lblStatus.Text = "Stopped all backup actions."
} -always
$btnSkip.Enabled = $false
$btnStop.Enabled = $false
$btnPause.Enabled = $false
$btnContinue.Enabled = $false

[void](New-TabPage "Accounts")
$gAcc = New-Group "Accounts"
[void](Add-RibbonButton $gAcc "Add" 0xE8FA $icGreen "Add an account (Name, Email, Password, Folder)" { Show-AddAccount } -always)
[void](Add-RibbonButton $gAcc "Edit list" 0xE70F $icGray "Open accounts.csv to change or remove accounts" { Start-Process notepad.exe "`"$($script:accountsPath)`"" } -always)
[void](Add-RibbonButton $gAcc "Refresh" 0xE72C $icGray "Reload the table" { Update-Table; Update-Log } -always)

$gSign = New-Group "Sign-in"
[void](Add-RibbonButton $gSign "Check`nsign-in" 0xE895 $icBlue "Check the sign-in of every account (fast)" { Start-Engine "T" @() "Check sign-in of all accounts"; Show-Tab "Home" })
[void](Add-RibbonButton $gSign "Sign in all" 0xE716 $icBlue "Sign in, one after another, every account that needs it" { Start-Engine "N" @() "Sign in all accounts that need it"; Show-Tab "Home" })
[void](Add-RibbonButton $gSign "Sign in`nselected" 0xE77B $icBlue "Sign in the selected accounts" { Invoke-ForSelected "SignIn" "Sign in"; Show-Tab "Home" })

[void](New-TabPage "Files & Settings")
$gFiles = New-Group "Files"
[void](Add-RibbonButton $gFiles "Backup`nfolder" 0xE8B7 $icGold "Open the download location" { Start-Process explorer.exe $script:dest } -always)
[void](Add-RibbonButton $gFiles "Clean Mac`nfiles" 0xE74D $icRed "Delete the Mac junk files (._name, .DS_Store, __MACOSX) already downloaded for the selected accounts" {
    $r = [System.Windows.Forms.MessageBox]::Show("Delete the Mac junk files (._name, .DS_Store, __MACOSX folders) from the downloaded backup of the selected accounts?`n`nThey only hold Mac Finder info, no file content. Nothing on Google Drive is touched.", "Clean Mac files", "YesNo", "Question")
    if ($r -eq "Yes") { Invoke-ForSelected "Clean" "Clean Mac files" }
})
[void](Add-RibbonButton $gFiles "Logs" 0xE8A5 $icGray "Open the logs folder" { Start-Process explorer.exe $script:logDir } -always)
[void](Add-RibbonButton $gFiles "Settings" 0xE713 $icGray "Download location, Google app, speed, speed limit, document formats, links" { Show-SettingsForm })
[void](Add-RibbonButton $gFiles "Old menu" 0xE756 $icGray "The same tool as a text menu" { Start-Process (Join-Path $here "Backup-console.bat") } -always)

[void](New-TabPage "Help")
$gHelp = New-Group "Help"
[void](Add-RibbonButton $gHelp "How to use" 0xE897 $icGold "How to use G Nexus and connect your Google app" { Show-Instructions } -always)
[void](Add-RibbonButton $gHelp "Report for`nsupport" 0xEBE8 $icPurple "Make a .zip with logs and errors (no passwords) to send for support" { New-SupportReport } -always)

# ---------- options bar ----------
$opts = New-Object System.Windows.Forms.FlowLayoutPanel
$opts.Dock = "Top"
$opts.AutoSize = $true
$opts.WrapContents = $true
$opts.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 6)
$opts.BackColor = $cBack

function Add-OptLabel($text) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.AutoSize = $true
    $l.ForeColor = $cDim
    $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $l.Margin = New-Object System.Windows.Forms.Padding(14, 7, 4, 0)
    $opts.Controls.Add($l)
}
function New-OptBox($text, $checked) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $text
    $c.Checked = $checked
    $c.AutoSize = $true
    $c.ForeColor = $cText
    $c.Margin = New-Object System.Windows.Forms.Padding(4, 6, 4, 0)
    $opts.Controls.Add($c)
    return $c
}

Add-OptLabel "Save to"
$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Text = $script:dest
$txtDest.ReadOnly = $true
$txtDest.Width = 340
$txtDest.BackColor = $cGrid
$txtDest.ForeColor = $cText
$txtDest.BorderStyle = "FixedSingle"
$txtDest.Margin = New-Object System.Windows.Forms.Padding(2, 4, 2, 0)
$opts.Controls.Add($txtDest)
$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Change..."
$btnDest.AutoSize = $true
$btnDest.FlatStyle = "Flat"
$btnDest.FlatAppearance.BorderColor = $cLine
$btnDest.BackColor = $cBtn
$btnDest.ForeColor = $cText
$btnDest.Margin = New-Object System.Windows.Forms.Padding(2, 2, 2, 0)
$btnDest.add_Click({ Set-Destination })
$opts.Controls.Add($btnDest)
$script:actionButtons += $btnDest

$trashDefault = $false
try { $trashDefault = [bool]((Get-Settings).IncludeTrash) } catch { }
Add-OptLabel "Parts"
$chkPartMy     = New-OptBox "My Drive" $true
$chkPartShared = New-OptBox "Shared with me" $false
$chkPartDrives = New-OptBox "Shared drives" $false
$chkPartTrash  = New-OptBox "Trash" $trashDefault

Add-OptLabel "Shared with me"
$cmbShared = New-Object System.Windows.Forms.ComboBox
$cmbShared.DropDownStyle = "DropDownList"
[void]$cmbShared.Items.Add("Only from outside these accounts")
[void]$cmbShared.Items.Add("All")
$cmbShared.SelectedIndex = 0
$cmbShared.Width = 210
$cmbShared.FlatStyle = "Flat"
$cmbShared.BackColor = $cBtn
$cmbShared.ForeColor = $cText
$cmbShared.Margin = New-Object System.Windows.Forms.Padding(2, 4, 4, 0)
$opts.Controls.Add($cmbShared)
$cmbShared.Enabled = $false
$chkPartShared.add_CheckedChanged({ $cmbShared.Enabled = $chkPartShared.Checked })
# The table shows only the ticked parts: refresh it when a box changes (not while something runs)
foreach ($pb in @($chkPartMy, $chkPartShared, $chkPartDrives, $chkPartTrash)) {
    $pb.add_CheckedChanged({ if (-not $script:loadingSession -and -not ($script:proc -and -not $script:proc.HasExited)) { Update-Table } })
}

Add-OptLabel "Download"
$chkVerify = New-OptBox "Verify after download" $false
$chkOverwrite = New-OptBox "Overwrite files already downloaded" $false
$junkDefault = $true
try { $jd = (Get-Content -LiteralPath $script:settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json).SkipJunk; if ($null -ne $jd) { $junkDefault = [bool]$jd } } catch { }
$chkJunk = New-OptBox "Skip Mac junk files" $junkDefault
$chkNoShortcuts = New-OptBox "Skip shortcuts" $true
$tips.SetToolTip($chkNoShortcuts, "Drive shortcuts point to files or folders that live somewhere else. Ticked: they are not downloaded again here. Untick if a shortcut points to something outside these accounts.")
$chkOwned = New-OptBox "Only files owned by the account" $true
# These change the sizes in the table too
foreach ($pb in @($chkNoShortcuts, $chkOwned)) {
    $pb.add_CheckedChanged({ if (-not $script:loadingSession -and -not ($script:proc -and -not $script:proc.HasExited)) { Update-Table } })
}
$tips.SetToolTip($chkOwned, "My Drive: download only the files this account owns. Files owned by other people are left out (they belong to the owner's backup). Careful with owners outside these accounts.")
$tips.SetToolTip($chkJunk, "Skips ._name files, .DS_Store and __MACOSX folders that a Mac leaves next to real files. They only hold Finder info, no content.")
Add-OptLabel "View"
$chkTable = New-OptBox "Table view" $false
$chkTable.add_CheckedChanged({ $grid.Visible = $chkTable.Checked; $dash.Visible = -not $chkTable.Checked })
$chkConsole = New-OptBox "Show console window" $false
$tips.SetToolTip($chkConsole, "Also show the black console window of the next action (for checking details)")

# Status + log
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Dock = "Bottom"
$lblStatus.Height = 26
$lblStatus.Padding = New-Object System.Windows.Forms.Padding(8, 5, 0, 0)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblStatus.BackColor = $cPanel
$lblStatus.ForeColor = $cText

$progPanel = New-Object System.Windows.Forms.Panel
$progPanel.Dock = "Bottom"
$progPanel.Height = 30
$progPanel.BackColor = $cBack
$progBar = New-Object System.Windows.Forms.ProgressBar
$progBar.Dock = "Left"
$progBar.Width = 360
$progBar.Minimum = 0
$progBar.Maximum = 100
$lblProg = New-Object System.Windows.Forms.Label
$lblProg.Dock = "Fill"
$lblProg.Padding = New-Object System.Windows.Forms.Padding(8, 7, 0, 0)
$lblProg.Font = New-Object System.Drawing.Font("Consolas", 9)
$lblProg.ForeColor = New-Color 120 220 140
$progPanel.Controls.Add($lblProg)
$progPanel.Controls.Add($progBar)

function Update-Progress {
    $running = $script:proc -and -not $script:proc.HasExited
    if (-not $running -or -not (Test-Path -LiteralPath $progressFile)) {
        $progBar.Style = "Blocks"; $progBar.Value = 0
        $lblProg.Text = $(if ($running) { "Starting..." } else { "" })
        return
    }
    try { $pr = Get-Content -LiteralPath $progressFile -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return }
    if ([int]$pr.Percent -lt 0) {
        $progBar.Style = "Marquee"
    } else {
        $progBar.Style = "Blocks"
        $progBar.Value = [math]::Max(0, [math]::Min(100, [int]$pr.Percent))
    }
    $lblProg.Text = "$($pr.Activity)   $($pr.Status)"
}

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = "Bottom"
$txtLog.Height = 170
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = New-Color 16 16 20
$txtLog.BorderStyle = "None"
$txtLog.ForeColor = [System.Drawing.Color]::Gainsboro

$help = New-Object System.Windows.Forms.Label
$help.Dock = "Bottom"
$help.Height = 22
$help.Padding = New-Object System.Windows.Forms.Padding(8, 3, 0, 0)
$help.ForeColor = $cDim
$help.BackColor = $cBack
$help.Text = "Click an account (Ctrl+click for several), then a button. Sign-in opens your browser; everything else shows below. Green = signed in, red = needs sign-in, gold = not checked. Output:"

# Order matters for docking: Fill first, then edges
$form.Controls.Add($grid)
$form.Controls.Add($dash)
$grid.Visible = $false
$form.Controls.Add($opts)
foreach ($k in @($script:tabPages.Keys)) { Add-TabButton $k }
# Added in reverse so they stack in tab order (all hidden but one)
$pagesRev = @($script:tabPages.Values); [array]::Reverse($pagesRev)
foreach ($pg in $pagesRev) { $form.Controls.Add($pg) }
$form.Controls.Add($tabStrip)
Show-Tab "Home"
$form.Controls.Add($headerLine)
$form.Controls.Add($header)
$form.Controls.Add($help)
$form.Controls.Add($txtLog)
$form.Controls.Add($progPanel)
$form.Controls.Add($lblStatus)

# Watch the running action and the log
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$script:tick = 0
$timer.add_Tick({
    try {
    # Every 5 s: how many backup actions run on this PC (also from closed windows)
    $script:tick++
    if ($script:tick % 5 -eq 1) {
        $n = @(Get-EngineProcesses).Count
        $btnStopAll.Text = $(if ($n -gt 0) { "Stop all ($n)" } else { "Stop all" })
        $btnStopAll.BackColor = $(if ($n -gt 1) { New-Color 110 30 30 } else { $cPanel })
        $mine = $script:proc -and -not $script:proc.HasExited
        if ($n -gt 0 -and -not $mine) { $lblStatus.Text = "$n backup action(s) running in the background (from a closed window). 'Stop all' stops them." }
        elseif ($n -gt 1) { $lblStatus.Text = "WARNING: $n backup actions are running at the same time. Use 'Stop all' and start again." }
    }
    Update-Log
    Update-Progress
    if ($script:proc -and $script:proc.HasExited) {
        $code = $script:proc.ExitCode
        $script:proc = $null
        # Ended normally: nothing to continue. Paused, stopped by a full disk, or killed: keep it.
        if ($script:isDownload -and -not $script:paused -and $code -eq 0) { Clear-Resume }
        Set-Busy $false
        Update-Table
        $lblStatus.Text = "Finished at $((Get-Date).ToString('HH:mm')): $($script:running). Table updated."
        if ($script:paused) { $lblStatus.Text = "Paused: $($script:running). Press Continue when you want (also after closing G Nexus)." }
        elseif ($code -eq 3) { $lblStatus.Text = "Stopped at $((Get-Date).ToString('HH:mm')): the disk is full. Free up space, then press Continue." }
        $script:paused = $false
        if (Test-Path -LiteralPath (Join-Path $logDir "errors.log")) {
            $ef = Get-Item -LiteralPath (Join-Path $logDir "errors.log")
            if ($ef.LastWriteTime -gt (Get-Date).AddMinutes(-1)) { $lblStatus.Text += "  Some errors were logged: use 'Report for support' if something looks wrong." }
        }
    }
    } catch { Save-AppErrors }
})

# Backup actions still running from an earlier G Nexus window (they keep going after it closes)
function Get-EngineProcesses {
    try {
        return @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
            Where-Object { $_.CommandLine -match 'backup-drive\.ps1' -and $_.CommandLine -match '-Action' -and $_.CommandLine -notmatch 'Export' -and $_.ProcessId -ne $PID })
    } catch { return @() }
}
function Stop-EngineTree($procId) { & taskkill.exe /PID $procId /T /F | Out-Null }

$form.add_FormClosing({
    param($sender, $e)
    if ($script:proc -and -not $script:proc.HasExited) {
        $r = [System.Windows.Forms.MessageBox]::Show("$($script:running) is still running.`n`nYes = pause it and close$(if ($script:isDownload) { ' (Continue next time)' })`nNo = close and let it keep running in the background`nCancel = don't close", "G Nexus", "YesNoCancel")
        if ($r -eq "Cancel") { $e.Cancel = $true; return }
        if ($r -eq "Yes") { Stop-EngineTree $script:proc.Id }
    }
    Save-Session
})

# ---------- session: the ticked options and selected accounts, remembered between runs ----------
# (Overwrite is never remembered: it always starts unticked)
$sessionFile = Join-Path $script:cfgDir "session.json"
$sessionBoxes = [ordered]@{
    PartMy = $chkPartMy; PartShared = $chkPartShared; PartDrives = $chkPartDrives; PartTrash = $chkPartTrash
    Verify = $chkVerify; SkipJunk = $chkJunk; NoShortcuts = $chkNoShortcuts; OwnedOnly = $chkOwned
    TableView = $chkTable; Console = $chkConsole
}
function Save-Session {
    try {
        $o = [ordered]@{}
        foreach ($k in $sessionBoxes.Keys) { $o[$k] = [bool]$sessionBoxes[$k].Checked }
        $o.SharedMode = $cmbShared.SelectedIndex
        $o.Selected = @(foreach ($n in @($script:selected)) { if ($script:rowsByNum.ContainsKey($n)) { "$(Get-RowValue $script:rowsByNum[$n] 'Email')" } })
        $o | ConvertTo-Json | Set-Content -LiteralPath $sessionFile -Encoding UTF8
    } catch { Save-AppErrors }
}
$script:sessionSelected = @()
function Restore-Session {
    if (-not (Test-Path -LiteralPath $sessionFile)) { return }
    $script:loadingSession = $true
    try {
        $o = Get-Content -LiteralPath $sessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in $sessionBoxes.Keys) { if ($null -ne $o.$k) { $sessionBoxes[$k].Checked = [bool]$o.$k } }
        if ($null -ne $o.SharedMode -and [int]$o.SharedMode -ge 0 -and [int]$o.SharedMode -lt $cmbShared.Items.Count) { $cmbShared.SelectedIndex = [int]$o.SharedMode }
        $script:sessionSelected = @($o.Selected | ForEach-Object { "$_" })
    } catch { Save-AppErrors }
    $script:loadingSession = $false
}
# Selects the remembered accounts once the cards exist
function Select-SessionAccounts {
    foreach ($k in @($script:rowsByNum.Keys)) {
        if ($script:sessionSelected -contains "$(Get-RowValue $script:rowsByNum[$k] 'Email')") { [void]$script:selected.Add([int]$k); Set-CardLook ([int]$k) }
    }
    $n = $script:selected.Count
    $lblSel.Text = $(if ($n -eq 0) { "Click a card to select it" } else { "$n selected" })
    $script:sessionSelected = @()
}
Restore-Session

$form.add_Shown({
    $left = @(Get-EngineProcesses)
    if ($left.Count -gt 0) {
        $r = [System.Windows.Forms.MessageBox]::Show("$($left.Count) backup action(s) from an earlier G Nexus window are still running in the background.`nRunning two at once makes both slow.`n`nStop them now? (files already downloaded are kept)", "G Nexus", "YesNo")
        if ($r -eq "Yes") { foreach ($lp in $left) { Stop-EngineTree $lp.ProcessId } }
    }
    $sf = $script:settingsPath
    $noClient = $true
    try { if ((Get-Content -LiteralPath $sf -Raw -Encoding UTF8 | ConvertFrom-Json).ClientId) { $noClient = $false } } catch { }
    if ($noClient) { Show-Instructions }
    Update-Table
    Select-SessionAccounts
    Update-Log
    # A download that was paused (or cut off) last time
    $rs = Read-Resume
    if ($null -ne $rs -and @(Get-EngineProcesses).Count -eq 0) {
        $r = [System.Windows.Forms.MessageBox]::Show("A download was paused on $($rs.Time):`n$($rs.Text)`n`nContinue it now?`n(No = later, with the Continue button)", "Continue download", "YesNo", "Question")
        if ($r -eq "Yes") { Resume-Download }
    }
    Update-ResumeButton
    $timer.Start()
})

# Shortcut with the G Nexus icon next to the scripts (copy it to the desktop if you like)
try {
    $lnk = Join-Path $here "G Nexus.lnk"
    $ico = Join-Path $here "gnexus.ico"
    if ((Test-Path -LiteralPath $ico) -and -not (Test-Path -LiteralPath $lnk)) {
        $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
        $sc.TargetPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
        $sc.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$(Join-Path $here 'backup-app.ps1')`""
        $sc.WorkingDirectory = $here
        $sc.IconLocation = "$ico,0"
        $sc.WindowStyle = 7
        $sc.Description = "G Nexus - Google Drive Migration Engine"
        $sc.Save()
    }
} catch { Save-AppErrors }

[void]$form.ShowDialog()
Save-AppErrors
$timer.Stop()
