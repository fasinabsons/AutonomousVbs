# MoonFlower Complete Automation Setup
# Sets up all Windows Task Scheduler tasks for the complete workflow

Write-Host "Setting up MoonFlower Complete Automation System..." -ForegroundColor Green
Write-Host "==========================================================="

# Define paths
$ScriptPath = "C:\Users\Lenovo\Documents\Automate2\Automata2\moonflower_complete.bat"
$WorkingDirectory = "C:\Users\Lenovo\Documents\Automate2\Automata2"

# Task configurations
$Tasks = @(
    @{
        Name = "MoonFlower_Morning_CSV"
        Description = "MoonFlower Morning CSV Download (9:30 AM)"
        Time = "09:30"
        Arguments = ""
    },
    @{
        Name = "MoonFlower_Complete_Workflow" 
        Description = "MoonFlower Complete Workflow - CSV + Excel + VBS (12:30 PM)"
        Time = "12:30"
        Arguments = ""
    },
    @{
        Name = "MoonFlower_Evening_Report"
        Description = "MoonFlower Evening Report Generation (5:30 PM)"
        Time = "17:30"
        Arguments = ""
    },
    @{
        Name = "MoonFlower_Morning_Email"
        Description = "MoonFlower Morning Email Delivery (8:00 AM)"
        Time = "08:00"
        Arguments = ""
    }
)

# Remove existing tasks
Write-Host "Removing existing MoonFlower tasks..." -ForegroundColor Yellow
foreach ($Task in $Tasks) {
    $ExistingTask = Get-ScheduledTask -TaskName $Task.Name -ErrorAction SilentlyContinue
    if ($ExistingTask) {
        Unregister-ScheduledTask -TaskName $Task.Name -Confirm:$false
        Write-Host "Removed existing task: $($Task.Name)" -ForegroundColor Yellow
    }
}

# Create new tasks
Write-Host "Creating new MoonFlower automation tasks..." -ForegroundColor Cyan

foreach ($Task in $Tasks) {
    try {
        # Create task action
        $Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$ScriptPath`"" -WorkingDirectory $WorkingDirectory
        
        # Create task trigger (daily at specified time)
        $Trigger = New-ScheduledTaskTrigger -Daily -At $Task.Time
        
        # Create task settings for locked PC execution
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd -RunOnlyIfNetworkAvailable
        
        # Create task principal (run with highest privileges and whether user is logged on or not)
        $Principal = New-ScheduledTaskPrincipal -UserId "$env:COMPUTERNAME\$env:USERNAME" -LogonType S4U -RunLevel Highest
        
        # Register the scheduled task
        Register-ScheduledTask -TaskName $Task.Name -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $Task.Description -Force
        
        Write-Host "✅ Created task: $($Task.Name) at $($Task.Time)" -ForegroundColor Green
        
    } catch {
        Write-Host "❌ Failed to create task: $($Task.Name)" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "LOCKED PC CONFIGURATION SUMMARY:" -ForegroundColor Cyan
Write-Host "================================="
Write-Host "✅ Run whether user is logged on or not: YES"
Write-Host "✅ Run with highest privileges: YES"
Write-Host "✅ Allow start if on batteries: YES"
Write-Host "✅ Don't stop if going on batteries: YES"
Write-Host "✅ Start when available: YES"
Write-Host "✅ Network connectivity required: YES"

Write-Host ""
Write-Host "SCHEDULE SUMMARY:" -ForegroundColor Yellow
Write-Host "=================="
foreach ($Task in $Tasks) {
    Write-Host "$($Task.Time) - $($Task.Description)"
}

Write-Host ""
Write-Host "MANUAL TESTING COMMANDS:" -ForegroundColor Magenta
Write-Host "========================"
Write-Host "To test any task manually (even on locked PC):"
foreach ($Task in $Tasks) {
    Write-Host "Start-ScheduledTask -TaskName '$($Task.Name)'"
}

Write-Host ""
Write-Host "MONITORING COMMANDS:" -ForegroundColor Cyan
Write-Host "==================="
Write-Host "Get-ScheduledTask -TaskName 'MoonFlower*' | Format-Table Name, State, NextRunTime"
Write-Host "Get-ScheduledTaskInfo -TaskName 'MoonFlower*' | Format-Table TaskName, LastRunTime, LastTaskResult"

Write-Host ""
Write-Host "LOG LOCATIONS:" -ForegroundColor White
Write-Host "=============="
Write-Host "Main Logs: C:\Users\Lenovo\Documents\Automate2\Automata2\EHC_Logs\[date]\"
Write-Host "Task Scheduler Logs: Event Viewer > Applications and Services Logs > Microsoft > Windows > TaskScheduler"

Write-Host ""
Write-Host "==========================================================="
Write-Host "MoonFlower Complete Automation Setup COMPLETED!" -ForegroundColor Green
Write-Host "==========================================================="
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Test individual components manually first"
Write-Host "2. Run a test task: Start-ScheduledTask -TaskName 'MoonFlower_Morning_CSV'"
Write-Host "3. Monitor logs in EHC_Logs folder"
Write-Host "4. Verify email delivery functionality"
Write-Host "5. Test locked PC execution" 