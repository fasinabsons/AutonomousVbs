# PowerShell Script for 365-Day MoonFlower Automation Schedule
# Handles locked/unlocked states, plugged/unplugged states, and PC restarts
# Ensures all tasks run regardless of system state

param(
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Status
)

$ErrorActionPreference = "Stop"

# Project configuration
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogPath = Join-Path $ProjectRoot "EHC_Logs"

# Ensure log directory exists
if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$LogFile = Join-Path $LogPath "scheduler_setup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param($Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogFile -Value $LogMessage
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-MoonFlowerSchedule {
    Write-Log "🚀 Installing 365-Day MoonFlower Automation Schedule"
    
    if (!(Test-Administrator)) {
        Write-Log "❌ ERROR: Administrator privileges required"
        Write-Log "Please run as Administrator"
        exit 1
    }
    
    # Task definitions with enhanced settings for locked/unlocked states
    $Tasks = @(
        @{
            Name = "MoonFlower_Download_Morning"
            Description = "Download CSV files at 9:30 AM daily"
            Time = "09:30"
            Script = "2_Download_Files.bat"
            RunLevel = "Highest"
            StartWhenAvailable = $true
            WakeToRun = $true
        },
        @{
            Name = "MoonFlower_Download_Afternoon" 
            Description = "Download CSV files and merge Excel at 12:30 PM daily"
            Time = "12:30"
            Script = "2_Download_Files.bat"
            RunLevel = "Highest"
            StartWhenAvailable = $true
            WakeToRun = $true
        },
        @{
            Name = "MoonFlower_VBS_Upload"
            Description = "VBS upload process at 1:00 PM daily"
            Time = "13:00"
            Script = "3_VBS_Upload.bat"
            RunLevel = "Highest"
            StartWhenAvailable = $true
            WakeToRun = $true
        },
        @{
            Name = "MoonFlower_Startup_Recovery"
            Description = "Check for missed schedules on startup and recover"
            Time = $null
            Script = "scripts\startup_recovery.ps1 -CheckAndRun"
            RunLevel = "Highest"
            StartWhenAvailable = $true
            WakeToRun = $true
            PowerShellTask = $true
            StartupTrigger = $true
        },
        @{
            Name = "MoonFlower_VBS_Report"
            Description = "Generate VBS PDF report at 4:00 PM daily"
            Time = "16:00"
            Script = "4_VBS_Report.bat"
            RunLevel = "Highest"
            StartWhenAvailable = $true
            WakeToRun = $true
        },
        @{
            Name = "MoonFlower_Email_Morning"
            Description = "Send morning email with PDF report at 8:00 AM daily"
            Time = "08:00"
            Script = "1_Email_Morning.bat"
            RunLevel = "Highest"
            StartWhenAvailable = $true
            WakeToRun = $true
        }
    )
    
    Write-Log "📋 Creating 6 scheduled tasks for 365-day operation..."
    
    foreach ($Task in $Tasks) {
        try {
            Write-Log "Creating task: $($Task.Name)"
            
            # Remove existing task if it exists
            try {
                Unregister-ScheduledTask -TaskName $Task.Name -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "Removed existing task: $($Task.Name)"
            } catch {
                # Task doesn't exist, continue
            }
            
            # Create action
            if ($Task.RestartTask) {
                $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command `"Restart-Computer -Force`""
            } elseif ($Task.PowerShellTask) {
                $ScriptPath = Join-Path $ProjectRoot $Task.Script.Split(' ')[0]
                $Arguments = $Task.Script.Split(' ', 2)[1]
                $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$ScriptPath`" $Arguments" -WorkingDirectory $ProjectRoot
            } else {
                $ScriptPath = Join-Path $ProjectRoot $Task.Script
                $Action = New-ScheduledTaskAction -Execute $ScriptPath -WorkingDirectory $ProjectRoot
            }
            
            # Create trigger for daily execution or startup
            if ($Task.StartupTrigger) {
                $Trigger = New-ScheduledTaskTrigger -AtStartup
            } else {
                $Trigger = New-ScheduledTaskTrigger -Daily -At $Task.Time
            }
            
            # Enhanced settings for locked/unlocked states and power management
            $Settings = New-ScheduledTaskSettingsSet `
                -AllowStartIfOnBatteries `
                -DontStopIfGoingOnBatteries `
                -StartWhenAvailable:$Task.StartWhenAvailable `
                -WakeToRun:$Task.WakeToRun `
                -DontStopOnIdleEnd `
                -RestartOnIdle `
                -RestartCount 3 `
                -RestartInterval (New-TimeSpan -Minutes 5) `
                -ExecutionTimeLimit (New-TimeSpan -Hours 4)
            
            # Principal with highest privileges and run regardless of user logon
            $Principal = New-ScheduledTaskPrincipal `
                -UserId "SYSTEM" `
                -LogonType ServiceAccount `
                -RunLevel $Task.RunLevel
            
            # Register the task
            $ScheduledTask = Register-ScheduledTask `
                -TaskName $Task.Name `
                -Description $Task.Description `
                -Action $Action `
                -Trigger $Trigger `
                -Settings $Settings `
                -Principal $Principal `
                -Force
            
            # Additional configuration for locked state execution
            $TaskPath = "\Microsoft\Windows\TaskScheduler\$($Task.Name)"
            $TaskXml = Export-ScheduledTask -TaskName $Task.Name
            
            # Modify XML to ensure it runs when locked
            $TaskXml = $TaskXml -replace '<LogonType>Password</LogonType>', '<LogonType>ServiceAccount</LogonType>'
            $TaskXml = $TaskXml -replace '<LogonType>Interactive</LogonType>', '<LogonType>ServiceAccount</LogonType>'
            
            # Re-register with modified XML
            Register-ScheduledTask -TaskName $Task.Name -Xml $TaskXml -Force | Out-Null
            
            Write-Log "✅ Successfully created: $($Task.Name) at $($Task.Time)"
            
        } catch {
            Write-Log "❌ Failed to create task $($Task.Name): $($_.Exception.Message)"
        }
    }
    
    Write-Log "🎯 365-Day Schedule Installation Summary:"
    Write-Log "   📧 8:00 AM  - Morning email with PDF report"
    Write-Log "   📥 9:30 AM  - Download CSV files (Morning)"
    Write-Log "   📊 12:30 PM - Download CSV + Excel merge (Afternoon)" 
    Write-Log "   ⬆️ 1:00 PM  - VBS Upload (3-hour process)"
    Write-Log "   📋 4:00 PM  - VBS Report generation"
    Write-Log "   🚀 Startup  - Recovery check for missed schedules"
    Write-Log ""
    Write-Log "🔐 All tasks configured to run:"
    Write-Log "   ✅ When PC is locked or unlocked"
    Write-Log "   ✅ When PC is plugged in or on battery"
    Write-Log "   ✅ Will wake PC from sleep if needed"
    Write-Log "   ✅ Run with highest privileges"
    Write-Log "   ✅ 365 days without interruption"
    
    Write-Log "🎉 365-Day MoonFlower Automation Schedule installed successfully!"
}

function Uninstall-MoonFlowerSchedule {
    Write-Log "🗑️ Uninstalling MoonFlower Automation Schedule"
    
    if (!(Test-Administrator)) {
        Write-Log "❌ ERROR: Administrator privileges required"
        exit 1
    }
    
    $TaskNames = @(
        "MoonFlower_Download_Morning",
        "MoonFlower_Download_Afternoon", 
        "MoonFlower_VBS_Upload",
        "MoonFlower_Startup_Recovery",
        "MoonFlower_VBS_Report",
        "MoonFlower_Email_Morning"
    )
    
    foreach ($TaskName in $TaskNames) {
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log "✅ Removed: $TaskName"
        } catch {
            Write-Log "⚠️ Task not found: $TaskName"
        }
    }
    
    Write-Log "🎉 MoonFlower schedule uninstalled successfully!"
}

function Show-ScheduleStatus {
    Write-Log "📊 MoonFlower 365-Day Schedule Status"
    Write-Log "======================================"
    
    $TaskNames = @(
        "MoonFlower_Download_Morning",
        "MoonFlower_Download_Afternoon",
        "MoonFlower_VBS_Upload", 
        "MoonFlower_Startup_Recovery",
        "MoonFlower_VBS_Report",
        "MoonFlower_Email_Morning"
    )
    
    foreach ($TaskName in $TaskNames) {
        try {
            $Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
            if ($Task) {
                $TaskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
                $NextRun = $TaskInfo.NextRunTime
                $LastRun = $TaskInfo.LastRunTime
                $LastResult = $TaskInfo.LastTaskResult
                
                Write-Log "✅ $TaskName"
                Write-Log "   📅 Next Run: $NextRun"
                Write-Log "   🕒 Last Run: $LastRun"
                Write-Log "   📊 Result: $LastResult"
                Write-Log ""
            } else {
                Write-Log "❌ $TaskName - NOT INSTALLED"
            }
        } catch {
            Write-Log "❌ $TaskName - ERROR: $($_.Exception.Message)"
        }
    }
}

# Main execution
Write-Log "==================================================="
Write-Log "MoonFlower 365-Day Automation Scheduler"
Write-Log "==================================================="
Write-Log "Project Root: $ProjectRoot"
Write-Log "Log File: $LogFile"
Write-Log ""

if ($Install) {
    Install-MoonFlowerSchedule
} elseif ($Uninstall) {
    Uninstall-MoonFlowerSchedule  
} elseif ($Status) {
    Show-ScheduleStatus
} else {
    Write-Log "❓ Usage Instructions:"
    Write-Log "   Install:   .\Setup_365Day_Complete_Schedule.ps1 -Install"
    Write-Log "   Uninstall: .\Setup_365Day_Complete_Schedule.ps1 -Uninstall"
    Write-Log "   Status:    .\Setup_365Day_Complete_Schedule.ps1 -Status"
    Write-Log ""
    Write-Log "🔐 Must run as Administrator"
    Write-Log "🎯 Configures 365-day automation with lock/unlock support"
}

Write-Log "Script execution completed."
