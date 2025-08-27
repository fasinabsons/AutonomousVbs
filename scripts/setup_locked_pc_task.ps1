# Setup Windows Task Scheduler for VBS Phase 3 - Locked PC Execution
# This script creates a scheduled task that can run even when PC is locked

Write-Host "Setting up VBS Phase 3 for Locked PC execution..." -ForegroundColor Green
Write-Host "============================================================"

# Define task parameters
$TaskName = "VBS_Phase3_LockedPC"
$Description = "VBS Phase 3 Data Upload - Runs even when PC is locked"
$ScriptPath = "C:\Users\Lenovo\Documents\Automate2\Automata2\scripts\run_phase3_locked_pc.bat"
$WorkingDirectory = "C:\Users\Lenovo\Documents\Automate2\Automata2"

# Check if task already exists
$ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($ExistingTask) {
    Write-Host "Task '$TaskName' already exists. Removing old task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Create task action
$Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$ScriptPath`"" -WorkingDirectory $WorkingDirectory

# Create task trigger (manual trigger - can be scheduled later)
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)

# Create task settings for locked PC execution
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd

# Create task principal (run with highest privileges and whether user is logged on or not)
$Principal = New-ScheduledTaskPrincipal -UserId "$env:COMPUTERNAME\$env:USERNAME" -LogonType S4U -RunLevel Highest

# Register the scheduled task
try {
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $Description -Force
    
    Write-Host "✅ SUCCESS: Task '$TaskName' created successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "TASK CONFIGURATION:" -ForegroundColor Cyan
    Write-Host "- Task Name: $TaskName"
    Write-Host "- Description: $Description"
    Write-Host "- Script Path: $ScriptPath"
    Write-Host "- Working Directory: $WorkingDirectory"
    Write-Host "- Run Level: Highest"
    Write-Host "- Run Whether User Is Logged On Or Not: Yes"
    Write-Host "- Allow Start If On Batteries: Yes"
    Write-Host "- Don't Stop If Going On Batteries: Yes"
    Write-Host ""
    Write-Host "USAGE INSTRUCTIONS:" -ForegroundColor Yellow
    Write-Host "1. To run Phase 3 immediately (even on locked PC):"
    Write-Host "   Start-ScheduledTask -TaskName '$TaskName'"
    Write-Host ""
    Write-Host "2. To run from Command Prompt:"
    Write-Host "   schtasks /run /tn `"$TaskName`""
    Write-Host ""
    Write-Host "3. To schedule for specific time:"
    Write-Host "   Open Task Scheduler and modify the trigger"
    Write-Host ""
    Write-Host "LOCKED PC BENEFITS:" -ForegroundColor Green
    Write-Host "✅ Runs even when PC is locked"
    Write-Host "✅ Runs even when user is not logged in"
    Write-Host "✅ Has highest privileges for system access"
    Write-Host "✅ Works with battery power"
    Write-Host "✅ Comprehensive logging to EHC_Logs folder"
    
} catch {
    Write-Host "❌ ERROR: Failed to create scheduled task" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "============================================================"
Write-Host "Setup completed successfully!" -ForegroundColor Green 