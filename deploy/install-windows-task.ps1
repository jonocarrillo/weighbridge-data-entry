# Registers the weighbridge server to start automatically at boot (Windows).
# Run once from an ADMIN PowerShell:
#   powershell -ExecutionPolicy Bypass -File deploy\install-windows-task.ps1
# Remove with:
#   Unregister-ScheduledTask -TaskName "Weighbridge" -Confirm:$false

$bat = Join-Path $PSScriptRoot 'run-windows.bat'

$action    = New-ScheduledTaskAction -Execute $bat
$trigger   = New-ScheduledTaskTrigger -AtStartup
$settings  = New-ScheduledTaskSettingsSet `
               -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
               -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
               -ExecutionTimeLimit (New-TimeSpan -Days 0)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

Register-ScheduledTask -TaskName "Weighbridge" -Action $action -Trigger $trigger `
  -Settings $settings -Principal $principal -Force

Write-Host "Installed. The server will start on every boot (no login needed)."
Write-Host "Start it now with:  Start-ScheduledTask -TaskName Weighbridge"
