# MoonFlower WiFi Automation - Universal PowerShell Startup Script
# Handles all execution policy and path issues automatically

param(
    [string]$Action = "menu",
    [switch]$Service,
    [switch]$Install,
    [switch]$Manual,
    [switch]$Status
)

# ============================================================================
# EXECUTION POLICY SETUP
# ============================================================================
try {
    # Set execution policy for current process only
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
    
    # Also try to set for current user if we have permissions
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    } catch {
        # Ignore if we can't set user policy
    }
} catch {
    Write-Warning "Could not modify execution policy: $($_.Exception.Message)"
}

# ============================================================================
# CONFIGURATION
# ============================================================================
$Global:Config = @{
    ServiceName = "MoonFlowerAutomation"
    Version = "5.0"
    ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    BasePath = ""
    LogPath = ""
    DateFolder = ""
}

# Calculate base path (parent of scripts directory)
$Global:Config.BasePath = Split-Path -Parent $Global:Config.ScriptPath
$Global:Config.LogPath = Join-Path $Global:Config.BasePath "EHC_Logs"
$Global:Config.DateFolder = (Get-Date).ToString("ddMMM").ToLower()

# ============================================================================
# LOGGING SYSTEM
# ============================================================================
function Initialize-Logging {
    # Create log directories
    $logDir = Join-Path $Global:Config.LogPath $Global:Config.DateFolder
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    $Global:LogFile = Join-Path $logDir "startup.log"
    
    # Create other required directories
    @("EHC_Data", "EHC_Data_Merge", "EHC_Data_Pdf", "templates") | ForEach-Object {
        $dir = Join-Path $Global:Config.BasePath $_
        if (!(Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        
        # Create date subdirectories for data folders
        if ($_ -match "EHC_Data|EHC_Logs") {
            $dateDir = Join-Path $dir $Global:Config.DateFolder
            if (!(Test-Path $dateDir)) {
                New-Item -ItemType Directory -Path $dateDir -Force | Out-Null
            }
        }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$Console
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    if ($Global:LogFile) {
        Add-Content -Path $Global:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
    
    # Write to console
    if ($Console -or $Host.Name -eq "ConsoleHost") {
        switch ($Level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            default { Write-Host $logEntry -ForegroundColor White }
        }
    }
}

# ============================================================================
# PYTHON DETECTION
# ============================================================================
function Find-Python {
    Write-Log "Searching for Python installation..." -Console
    
    # Try standard python command first
    try {
        $pythonVersion = python --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Found Python in PATH: $pythonVersion" -Level "SUCCESS" -Console
            return "python"
        }
    } catch { }
    
    # Try common Python installation paths
    $pythonPaths = @(
        "$env:LOCALAPPDATA\Programs\Python\*\python.exe",
        "$env:PROGRAMFILES\Python*\python.exe",
        "$env:PROGRAMFILES(X86)\Python*\python.exe",
        "C:\Python*\python.exe"
    )
    
    foreach ($pathPattern in $pythonPaths) {
        $pythonExes = Get-ChildItem -Path $pathPattern -ErrorAction SilentlyContinue
        if ($pythonExes) {
            $pythonExe = $pythonExes | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            try {
                $version = & $pythonExe.FullName --version 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "Found Python at: $($pythonExe.FullName) - $version" -Level "SUCCESS" -Console
                    return $pythonExe.FullName
                }
            } catch { }
        }
    }
    
    Write-Log "Python not found" -Level "ERROR" -Console
    return $null
}

# ============================================================================
# AUTOMATION METHOD DETECTION
# ============================================================================
function Get-AutomationMethod {
    Write-Log "Detecting best automation method..." -Console
    
    $pythonCmd = Find-Python
    
    # Method 1: Python orchestrator (preferred)
    $orchestratorPath = Join-Path $Global:Config.BasePath "master_automation_orchestrator.py"
    if ((Test-Path $orchestratorPath) -and $pythonCmd) {
        Write-Log "Using Python orchestrator method" -Level "SUCCESS" -Console
        return @{
            Method = "PYTHON_ORCHESTRATOR"
            Command = $pythonCmd
            Arguments = "`"$orchestratorPath`""
            Description = "Master Python Orchestrator"
        }
    }
    
    # Method 2: VBS Complete Workflow
    $vbsWorkflowPath = Join-Path $Global:Config.BasePath "vbs\vbs_complete_workflow.py"
    if ((Test-Path $vbsWorkflowPath) -and $pythonCmd) {
        Write-Log "Using VBS complete workflow method" -Level "SUCCESS" -Console
        return @{
            Method = "VBS_WORKFLOW"
            Command = $pythonCmd
            Arguments = "`"$vbsWorkflowPath`""
            Description = "VBS Complete Workflow"
        }
    }
    
    # Method 3: PowerShell automation
    $psScriptPath = Join-Path $Global:Config.ScriptPath "MoonFlowerAutomation.ps1"
    if (Test-Path $psScriptPath) {
        Write-Log "Using PowerShell automation method" -Level "SUCCESS" -Console
        return @{
            Method = "POWERSHELL"
            Command = "powershell"
            Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$psScriptPath`""
            Description = "PowerShell Automation"
        }
    }
    
    # Method 4: BAT automation
    $batScriptPath = Join-Path $Global:Config.ScriptPath "moonflower_automation.bat"
    if (Test-Path $batScriptPath) {
        Write-Log "Using BAT automation method" -Level "SUCCESS" -Console
        return @{
            Method = "BAT"
            Command = "cmd"
            Arguments = "/c `"$batScriptPath`""
            Description = "Batch File Automation"
        }
    }
    
    Write-Log "No automation method found" -Level "ERROR" -Console
    return $null
}

# ============================================================================
# SERVICE INSTALLATION
# ============================================================================
function Install-StartupMethods {
    Write-Log "Installing startup methods..." -Console
    
    if (!(Test-IsAdmin)) {
        Write-Log "Administrator privileges required for installation" -Level "ERROR" -Console
        return $false
    }
    
    $success = $true
    
    try {
        # Method 1: Scheduled Task (Primary)
        Write-Log "Creating Scheduled Task..." -Console
        $taskName = $Global:Config.ServiceName
        $scriptPath = $MyInvocation.MyCommand.Path
        
        # Remove existing task
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        
        # Create new task
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -Service"
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "MoonFlower WiFi Automation Service" | Out-Null
        
        Write-Log "Scheduled task created successfully" -Level "SUCCESS" -Console
    }
    catch {
        Write-Log "Failed to create scheduled task: $($_.Exception.Message)" -Level "ERROR" -Console
        $success = $false
    }
    
    try {
        # Method 2: Registry Run Key (Backup)
        Write-Log "Creating Registry startup entry..." -Console
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        $regValue = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MyInvocation.MyCommand.Path)`" -Service"
        Set-ItemProperty -Path $regPath -Name $Global:Config.ServiceName -Value $regValue
        
        Write-Log "Registry startup entry created successfully" -Level "SUCCESS" -Console
    }
    catch {
        Write-Log "Failed to create registry entry: $($_.Exception.Message)" -Level "WARNING" -Console
    }
    
    try {
        # Method 3: Startup Folder (Additional backup)
        Write-Log "Creating startup folder entry..." -Console
        $startupFolder = [Environment]::GetFolderPath("Startup")
        $startupScript = Join-Path $startupFolder "MoonFlowerAutomation.bat"
        
        $startupContent = @"
@echo off
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "$($MyInvocation.MyCommand.Path)" -Service
"@
        Set-Content -Path $startupScript -Value $startupContent
        
        Write-Log "Startup folder entry created successfully" -Level "SUCCESS" -Console
    }
    catch {
        Write-Log "Failed to create startup folder entry: $($_.Exception.Message)" -Level "WARNING" -Console
    }
    
    return $success
}

function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ============================================================================
# HEALTH CHECK
# ============================================================================
function Test-SystemHealth {
    Write-Log "Performing system health check..." -Console
    
    $health = @{
        Python = $false
        DiskSpace = $false
        Memory = $false
        Processes = $false
    }
    
    # Check Python
    $pythonCmd = Find-Python
    $health.Python = $pythonCmd -ne $null
    
    # Check disk space (need at least 1GB free)
    try {
        $drive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq (Split-Path $Global:Config.BasePath -Qualifier) }
        $freeSpaceGB = [math]::Round($drive.FreeSpace / 1GB, 2)
        $health.DiskSpace = $freeSpaceGB -gt 1
        Write-Log "Free disk space: $freeSpaceGB GB" -Console
    } catch {
        Write-Log "Could not check disk space" -Level "WARNING" -Console
    }
    
    # Check available memory (need at least 500MB)
    try {
        $memory = Get-WmiObject -Class Win32_OperatingSystem
        $freeMemoryMB = [math]::Round($memory.FreePhysicalMemory / 1024, 2)
        $health.Memory = $freeMemoryMB -gt 500
        Write-Log "Free memory: $freeMemoryMB MB" -Console
    } catch {
        Write-Log "Could not check memory" -Level "WARNING" -Console
    }
    
    # Check for conflicting processes
    $conflictingProcesses = Get-Process | Where-Object { 
        $_.ProcessName -match "python|powershell" -and 
        $_.CommandLine -match "moonflower|automation" 
    }
    $health.Processes = $conflictingProcesses.Count -eq 0
    
    if ($conflictingProcesses.Count -gt 0) {
        Write-Log "Found $($conflictingProcesses.Count) potentially conflicting processes" -Level "WARNING" -Console
    }
    
    return $health
}

# ============================================================================
# AUTOMATION EXECUTION
# ============================================================================
function Start-Automation {
    param([hashtable]$Method)
    
    Write-Log "Starting automation: $($Method.Description)" -Console
    Write-Log "Command: $($Method.Command) $($Method.Arguments)" -Console
    
    try {
        # Change to base directory
        Set-Location $Global:Config.BasePath
        
        # Start the automation process
        $processArgs = @{
            FilePath = $Method.Command
            ArgumentList = $Method.Arguments
            WindowStyle = "Hidden"
            Wait = $false
            PassThru = $true
        }
        
        $process = Start-Process @processArgs
        
        Write-Log "Automation started successfully (PID: $($process.Id))" -Level "SUCCESS" -Console
        return $true
    }
    catch {
        Write-Log "Failed to start automation: $($_.Exception.Message)" -Level "ERROR" -Console
        return $false
    }
}

# ============================================================================
# SERVICE LOOP
# ============================================================================
function Start-ServiceLoop {
    Write-Log "=========================================" -Console
    Write-Log "MoonFlower WiFi Automation Service Started" -Console
    Write-Log "Version: $($Global:Config.Version)" -Console
    Write-Log "Base Path: $($Global:Config.BasePath)" -Console
    Write-Log "=========================================" -Console
    
    # Get automation method
    $automationMethod = Get-AutomationMethod
    if (!$automationMethod) {
        Write-Log "No automation method available - exiting" -Level "ERROR" -Console
        return
    }
    
    # Perform health check
    $health = Test-SystemHealth
    if (!$health.Python) {
        Write-Log "Python not available - some features may not work" -Level "WARNING" -Console
    }
    
    # Start automation
    if (Start-Automation -Method $automationMethod) {
        Write-Log "Service loop started successfully" -Level "SUCCESS" -Console
    } else {
        Write-Log "Failed to start service loop" -Level "ERROR" -Console
    }
}

# ============================================================================
# STATUS DISPLAY
# ============================================================================
function Show-Status {
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "MoonFlower Automation Status" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Basic Information
    Write-Host "System Information:" -ForegroundColor Yellow
    Write-Host "  Version: $($Global:Config.Version)"
    Write-Host "  Base Path: $($Global:Config.BasePath)"
    Write-Host "  Script Path: $($Global:Config.ScriptPath)"
    Write-Host "  Date Folder: $($Global:Config.DateFolder)"
    Write-Host ""
    
    # Installation Status
    Write-Host "Installation Status:" -ForegroundColor Yellow
    $taskExists = Get-ScheduledTask -TaskName $Global:Config.ServiceName -ErrorAction SilentlyContinue
    if ($taskExists) {
        Write-Host "  Scheduled Task: " -NoNewline
        Write-Host "INSTALLED" -ForegroundColor Green
        Write-Host "  Task Status: $($taskExists.State)"
    } else {
        Write-Host "  Scheduled Task: " -NoNewline
        Write-Host "NOT INSTALLED" -ForegroundColor Red
    }
    
    # Automation Method
    Write-Host "Automation Method:" -ForegroundColor Yellow
    $method = Get-AutomationMethod
    if ($method) {
        Write-Host "  Method: $($method.Method)" -ForegroundColor Green
        Write-Host "  Description: $($method.Description)"
        Write-Host "  Command: $($method.Command) $($method.Arguments)"
    } else {
        Write-Host "  No automation method available" -ForegroundColor Red
    }
    Write-Host ""
    
    # Health Status
    Write-Host "System Health:" -ForegroundColor Yellow
    $health = Test-SystemHealth
    foreach ($key in $health.Keys) {
        Write-Host "  $key`: " -NoNewline
        if ($health[$key]) {
            Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "ISSUE" -ForegroundColor Red
        }
    }
    Write-Host ""
    
    # Recent Logs
    Write-Host "Recent Log Entries:" -ForegroundColor Yellow
    if (Test-Path $Global:LogFile) {
        $recentLogs = Get-Content $Global:LogFile -Tail 5
        foreach ($log in $recentLogs) {
            Write-Host "  $log"
        }
    } else {
        Write-Host "  No log file found"
    }
    Write-Host ""
}

# ============================================================================
# INTERACTIVE MENU
# ============================================================================
function Show-Menu {
    while ($true) {
        Clear-Host
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "MoonFlower WiFi Automation" -ForegroundColor Cyan
        Write-Host "Universal Startup Script v$($Global:Config.Version)" -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Select an option:" -ForegroundColor Yellow
        Write-Host "1. Install Service (Run as Administrator)"
        Write-Host "2. Start Service Now"
        Write-Host "3. Manual Test Run"
        Write-Host "4. Show Status"
        Write-Host "5. View Logs"
        Write-Host "6. Exit"
        Write-Host ""
        
        $choice = Read-Host "Enter your choice (1-6)"
        
        switch ($choice) {
            "1" {
                if (Install-StartupMethods) {
                    Write-Host "Service installed successfully!" -ForegroundColor Green
                } else {
                    Write-Host "Service installation failed!" -ForegroundColor Red
                }
                Read-Host "Press Enter to continue"
            }
            "2" {
                Write-Host "Starting service in background..." -ForegroundColor Yellow
                Start-Job -ScriptBlock { 
                    & $using:MyInvocation.MyCommand.Path -Service 
                } | Out-Null
                Write-Host "Service started!" -ForegroundColor Green
                Read-Host "Press Enter to continue"
            }
            "3" {
                Write-Host "Running manual test..." -ForegroundColor Yellow
                $method = Get-AutomationMethod
                if ($method) {
                    Start-Automation -Method $method
                    Write-Host "Manual test started!" -ForegroundColor Green
                } else {
                    Write-Host "No automation method available!" -ForegroundColor Red
                }
                Read-Host "Press Enter to continue"
            }
            "4" {
                Show-Status
                Read-Host "Press Enter to continue"
            }
            "5" {
                if (Test-Path $Global:LogFile) {
                    Write-Host "Recent log entries:" -ForegroundColor Yellow
                    Get-Content $Global:LogFile -Tail 20 | ForEach-Object { Write-Host $_ }
                } else {
                    Write-Host "No log file found" -ForegroundColor Red
                }
                Read-Host "Press Enter to continue"
            }
            "6" {
                Write-Host "Goodbye!" -ForegroundColor Green
                return
            }
            default {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Initialize logging
Initialize-Logging

# Handle command line parameters
if ($Service) {
    Start-ServiceLoop
}
elseif ($Install) {
    if (Install-StartupMethods) {
        Write-Host "Service installed successfully!" -ForegroundColor Green
    } else {
        Write-Host "Service installation failed!" -ForegroundColor Red
    }
}
elseif ($Manual) {
    Write-Log "Manual execution requested" -Console
    $method = Get-AutomationMethod
    if ($method) {
        Start-Automation -Method $method
    } else {
        Write-Log "No automation method available" -Level "ERROR" -Console
    }
}
elseif ($Status) {
    Show-Status
}
else {
    # Default: Show interactive menu
    Show-Menu
}