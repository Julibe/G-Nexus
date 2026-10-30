# G Nexus - stop every backup action running on this PC (also the ones in the background)
$procs = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -match 'backup-drive\.ps1' -and $_.ProcessId -ne $PID })
foreach ($p in $procs) { & taskkill.exe /PID $p.ProcessId /T /F | Out-Null }
$rc = @(Get-Process rclone -ErrorAction SilentlyContinue)
foreach ($r in $rc) { try { $r.Kill() } catch { } }
Write-Host "Stopped $($procs.Count) backup action(s) and $($rc.Count) rclone process(es). Files already downloaded are kept."
