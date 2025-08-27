# MoonFlower Simple Automation Setup

Write-Host "Setting up MoonFlower Simple Automation..." -ForegroundColor Green
Write-Host "=============================================="

$ScriptPath = "C:\Users\Lenovo\Documents\Automate2\Automata2\moonflower_simple.bat"
$WorkingDirectory = "C:\Users\Lenovo\Documents\Automate2\Automata2"

# Remove existing tasks
$ExistingTasks = Get-ScheduledTask -TaskName "MoonFlower*" -ErrorAction SilentlyContinue
foreach ($Task in $ExistingTasks) {
    Unregister-ScheduledTask -TaskName $Task.TaskName -Confirm:$false
    Write-Host "Removed: $($Task.TaskName)" -ForegroundColor Yellow
}

# Task configurations
$Tasks = @(
    @{ Name = "MoonFlower_Morning_CSV"; Time = "09:30"; Description = "Morning CSV Download" },
    @{ Name = "MoonFlower_Midday_CSV"; Time = "12:30"; Description = "Midday CSV Download + Validation" },
    @{ Name = "MoonFlower_Excel"; Time = "12:35"; Description = "Excel Generation + Validation" },
    @{ Name = "MoonFlower_Phase1"; Time = "13:00"; Description = "VBS Phase 1 Login" },
    @{ Name = "MoonFlower_Phase4"; Time = "17:30"; Description = "VBS Phase 4 Report" },
    @{ Name = "MoonFlower_Email"; Time = "08:00"; Description = "Morning Email to GM" }
)

$SuccessCount = 0
foreach ($Task in $Tasks) {
    try {
        $Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$ScriptPath`"" -WorkingDirectory $WorkingDirectory
        $Trigger = New-ScheduledTaskTrigger -Daily -At $Task.Time
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd
        $Principal = New-ScheduledTaskPrincipal -UserId "$env:COMPUTERNAME\$env:USERNAME" -LogonType S4U -RunLevel Highest
        
        Register-ScheduledTask -TaskName $Task.Name -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $Task.Description -Force | Out-Null
        
        Write-Host "Created: $($Task.Name) at $($Task.Time)" -ForegroundColor Green
        $SuccessCount++
        
    } catch {
        Write-Host "Failed: $($Task.Name)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "SETUP COMPLETE!" -ForegroundColor Green
Write-Host "Successfully created: $SuccessCount/$($Tasks.Count) tasks"
Write-Host ""
Write-Host "SCHEDULE:"
Write-Host "09:30 - Morning CSV Download"
Write-Host "12:30 - Midday CSV Download + Validation"
Write-Host "12:35 - Excel Generation + Validation"
Write-Host "13:00 - VBS Phase 1 Login"
Write-Host "17:30 - VBS Phase 4 Report"
Write-Host "08:00 - Morning Email to GM"
Write-Host ""
Write-Host "NOTIFICATIONS: faseenm@gmail.com"
Write-Host "GM EMAILS: ramon.logan@absons.ae"
Write-Host ""
Write-Host "Test command: Start-ScheduledTask -TaskName MoonFlower_Phase1" 