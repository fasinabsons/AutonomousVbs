# MoonFlower Final Automation Setup
# Complete Task Scheduler configuration for production deployment

Write-Host "🌙 Setting up MoonFlower Final Automation System..." -ForegroundColor Green
Write-Host "=================================================================="

# Define paths
$ScriptPath = "C:\Users\Lenovo\Documents\Automate2\Automata2\moonflower_final.bat"
$WorkingDirectory = "C:\Users\Lenovo\Documents\Automate2\Automata2"

# Verify BAT file exists
if (-not (Test-Path $ScriptPath)) {
    Write-Host "❌ ERROR: BAT file not found at $ScriptPath" -ForegroundColor Red
    exit 1
}

Write-Host "✅ BAT file found: $ScriptPath" -ForegroundColor Green

# Remove any existing MoonFlower tasks
Write-Host "🧹 Removing existing MoonFlower tasks..." -ForegroundColor Yellow
$ExistingTasks = Get-ScheduledTask -TaskName "MoonFlower*" -ErrorAction SilentlyContinue
foreach ($Task in $ExistingTasks) {
    Unregister-ScheduledTask -TaskName $Task.TaskName -Confirm:$false
    Write-Host "   Removed: $($Task.TaskName)" -ForegroundColor Yellow
}

# Task configurations
$Tasks = @(
    @{
        Name = "MoonFlower_Morning_CSV"
        Description = "MoonFlower Morning CSV Download (9:30 AM)"
        Time = "09:30"
    },
    @{
        Name = "MoonFlower_Complete_Workflow"
        Description = "MoonFlower Complete Workflow - All Phases (12:30 PM)"
        Time = "12:30"
    },
    @{
        Name = "MoonFlower_Evening_Report"
        Description = "MoonFlower Evening Report Generation (5:30 PM)"
        Time = "17:30"
    },
    @{
        Name = "MoonFlower_Morning_Email"
        Description = "MoonFlower Morning Email to GM (8:00 AM)"
        Time = "08:00"
    }
)

Write-Host "🔧 Creating MoonFlower automation tasks..." -ForegroundColor Cyan

$SuccessCount = 0
foreach ($Task in $Tasks) {
    try {
        # Create task action
        $Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$ScriptPath`"" -WorkingDirectory $WorkingDirectory
        
        # Create task trigger (daily at specified time)
        $Trigger = New-ScheduledTaskTrigger -Daily -At $Task.Time
        
        # Create task settings for LOCKED PC execution
        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -DontStopOnIdleEnd `
            -RunOnlyIfNetworkAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Hours 6)
        
        # Create task principal for LOCKED PC (run whether user is logged on or not)
        $Principal = New-ScheduledTaskPrincipal -UserId "$env:COMPUTERNAME\$env:USERNAME" -LogonType S4U -RunLevel Highest
        
        # Register the scheduled task
        Register-ScheduledTask -TaskName $Task.Name -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $Task.Description -Force | Out-Null
        
        Write-Host "✅ Created: $($Task.Name) at $($Task.Time)" -ForegroundColor Green
        $SuccessCount++
        
    } catch {
        Write-Host "❌ Failed: $($Task.Name) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "📊 SETUP SUMMARY:" -ForegroundColor Cyan
Write-Host "==================="
Write-Host "✅ Successfully created: $SuccessCount/$($Tasks.Count) tasks"

if ($SuccessCount -eq $Tasks.Count) {
    Write-Host "🎉 ALL TASKS CREATED SUCCESSFULLY!" -ForegroundColor Green
} else {
    Write-Host "⚠️  Some tasks failed to create. Check permissions." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🔐 LOCKED PC CONFIGURATION:" -ForegroundColor Cyan
Write-Host "============================="
Write-Host "✅ Run whether user is logged on or not: YES"
Write-Host "✅ Run with highest privileges: YES"
Write-Host "✅ Allow start if on batteries: YES"
Write-Host "✅ Don't stop if going on batteries: YES"
Write-Host "✅ Start when available: YES"
Write-Host "✅ Network connectivity required: YES"
Write-Host "✅ Execution time limit: 6 hours"

Write-Host ""
Write-Host "📅 DAILY SCHEDULE:" -ForegroundColor Yellow
Write-Host "=================="
Write-Host "09:30 - Morning CSV Download"
Write-Host "12:30 - Complete Workflow (CSV + Excel + VBS Upload)"
Write-Host "17:30 - Evening Report Generation"
Write-Host "08:00 - Morning Email to GM"

Write-Host ""
Write-Host "📧 NOTIFICATIONS:" -ForegroundColor Magenta
Write-Host "=================="
Write-Host "📩 Milestone updates sent to: faseenm@gmail.com"
Write-Host "📄 Daily PDF reports sent to: ramon.logan@absons.ae"
Write-Host "📧 From: mohamed.fasin@absons.ae"

Write-Host ""
Write-Host "🧪 TESTING COMMANDS:" -ForegroundColor White
Write-Host "===================="
Write-Host "# Test individual tasks:"
foreach ($Task in $Tasks) {
    Write-Host "Start-ScheduledTask -TaskName '$($Task.Name)'" -ForegroundColor Gray
}

Write-Host ""
Write-Host "# Monitor task status:"
Write-Host "Get-ScheduledTask -TaskName MoonFlower* | Format-Table Name, State, NextRunTime" -ForegroundColor Gray
Write-Host "Get-ScheduledTaskInfo -TaskName MoonFlower* | Format-Table TaskName, LastRunTime, LastTaskResult" -ForegroundColor Gray

Write-Host ""
Write-Host "📁 LOG LOCATIONS:" -ForegroundColor White
Write-Host "=================="
Write-Host "Main Logs: C:\Users\Lenovo\Documents\Automate2\Automata2\EHC_Logs\[date]\"
Write-Host "Task Logs: Event Viewer > Task Scheduler"

Write-Host ""
Write-Host "=================================================================="
Write-Host "🎉 MoonFlower Final Automation Setup COMPLETED!" -ForegroundColor Green
Write-Host "=================================================================="

Write-Host ""
Write-Host "🚀 NEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Lock your PC (Windows + L) to test"
Write-Host "2. Tasks will run automatically at scheduled times"
Write-Host "3. Check faseenm@gmail.com for milestone notifications"
Write-Host "4. Check ramon.logan@absons.ae for daily PDF reports"
Write-Host "5. Monitor EHC_Logs folder for execution details"

Write-Host ""
Write-Host "✨ Your MoonFlower automation is ready for production! ✨" -ForegroundColor Green 