# MoonFlower WiFi Automation - PowerShell Service
# Complete automation system with advanced features
# Version: 4.0

param(
    [string]$Action = "service",
    [switch]$Manual,
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Status
)

# Set execution policy for this session
try {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
} catch {
    Write-Warning "Could not set execution policy: $($_.Exception.Message)"
}

# ============================================================================
# CONFIGURATION
# ============================================================================
$Global:Config = @{
    ServiceName = "MoonFlowerAutomation"
    Version = "4.0"
    ScriptPath = if ($MyInvocation.MyCommand.Path) { 
        Split-Path -Parent $MyInvocation.MyCommand.Path 
    } else { 
        Split-Path -Parent (Get-Location) 
    }
    LogRetentionDays = 30
    MaxRestartDays = 3
    CheckIntervalMinutes = 5
    
    # Time slots (24-hour format)
    TimeSlots = @{
        Slot1 = "09:30"  # Morning
        Slot2 = "13:00"  # Afternoon  
        Slot3 = "13:30"  # Backup
    }
    
    # Email settings
    Email = @{
        Enabled = $true
        Subject = "MoonFlower Automation - Excel Ready for Upload"
        Body = "The daily automation cycle has completed successfully. Excel file is ready for upload to VBS system."
    }
}

# ============================================================================
# LOGGING SYSTEM
# ============================================================================
function Initialize-Logging {
    $dateFolder = (Get-Date).ToString("ddMMM").ToLower()
    $logDir = Join-Path $Global:Config.ScriptPath "EHC_Logs\$dateFolder"
    
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    $Global:LogFile = Join-Path $logDir "automation.log"
    
    # Create log directories
    @("EHC_Data", "EHC_Data_Merge", "EHC_Data_Pdf", "EHC_Logs") | ForEach-Object {
        $dir = Join-Path $Global:Config.ScriptPath $_
        if (!(Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
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
    Add-Content -Path $Global:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    
    # Write to console if requested or if running interactively
    if ($Console -or $Host.Name -eq "ConsoleHost") {
        switch ($Level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            default { Write-Host $logEntry }
        }
    }
    
    # Write to Windows Event Log for important events
    if ($Level -eq "ERROR") {
        try {
            Write-EventLog -LogName Application -Source "MoonFlowerAutomation" -EventId 1001 -EntryType Error -Message $Message -ErrorAction SilentlyContinue
        } catch {
            # Event source doesn't exist, create it
            try {
                New-EventLog -LogName Application -Source "MoonFlowerAutomation" -ErrorAction SilentlyContinue
                Write-EventLog -LogName Application -Source "MoonFlowerAutomation" -EventId 1001 -EntryType Error -Message $Message -ErrorAction SilentlyContinue
            } catch { }
        }
    }
}

# ============================================================================
# SERVICE INSTALLATION
# ============================================================================
function Install-AutomationService {
    Write-Log "Installing MoonFlower Automation Service..." -Level "INFO" -Console
    
    if (!(Test-IsAdmin)) {
        Write-Log "Administrator privileges required for installation" -Level "ERROR" -Console
        return $false
    }
    
    try {
        # Method 1: Scheduled Task (Primary)
        $taskName = $Global:Config.ServiceName
        $scriptPath = $MyInvocation.MyCommand.Path
        
        # Remove existing task
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        
        # Create new task
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "MoonFlower WiFi Automation Service" | Out-Null
        
        Write-Log "Scheduled task created successfully" -Level "SUCCESS" -Console
        
        # Method 2: Registry Run Key (Backup)
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        $regValue = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        Set-ItemProperty -Path $regPath -Name $Global:Config.ServiceName -Value $regValue -ErrorAction SilentlyContinue
        
        Write-Log "Registry startup entry created" -Level "SUCCESS" -Console
        
        # Create service configuration file
        $configFile = Join-Path $Global:Config.ScriptPath "service_config.json"
        $serviceConfig = @{
            ServiceName = $Global:Config.ServiceName
            Version = $Global:Config.Version
            InstallDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ScriptPath = $scriptPath
            LastUpdate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $serviceConfig | ConvertTo-Json | Set-Content -Path $configFile
        
        Write-Log "Service installed successfully" -Level "SUCCESS" -Console
        return $true
    }
    catch {
        Write-Log "Service installation failed: $($_.Exception.Message)" -Level "ERROR" -Console
        return $false
    }
}

function Uninstall-AutomationService {
    Write-Log "Uninstalling MoonFlower Automation Service..." -Level "INFO" -Console
    
    if (!(Test-IsAdmin)) {
        Write-Log "Administrator privileges required for uninstallation" -Level "ERROR" -Console
        return $false
    }
    
    try {
        # Remove scheduled task
        Unregister-ScheduledTask -TaskName $Global:Config.ServiceName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Scheduled task removed" -Level "SUCCESS" -Console
        
        # Remove registry entry
        Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name $Global:Config.ServiceName -ErrorAction SilentlyContinue
        Write-Log "Registry entry removed" -Level "SUCCESS" -Console
        
        # Remove configuration file
        $configFile = Join-Path $Global:Config.ScriptPath "service_config.json"
        if (Test-Path $configFile) {
            Remove-Item $configFile -Force -ErrorAction SilentlyContinue
        }
        
        Write-Log "Service uninstalled successfully" -Level "SUCCESS" -Console
        return $true
    }
    catch {
        Write-Log "Service uninstallation failed: $($_.Exception.Message)" -Level "ERROR" -Console
        return $false
    }
}

function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ============================================================================
# SELF-UPDATE SYSTEM
# ============================================================================
function Test-ForUpdates {
    $updateFile = Join-Path $Global:Config.ScriptPath "MoonFlowerAutomation_new.ps1"
    
    if (Test-Path $updateFile) {
        Write-Log "Update detected - applying update..." -Level "INFO" -Console
        
        try {
            # Backup current version
            $backupFile = Join-Path $Global:Config.ScriptPath "MoonFlowerAutomation_backup.ps1"
            Copy-Item $MyInvocation.MyCommand.Path $backupFile -Force
            
            # Apply update
            Copy-Item $updateFile $MyInvocation.MyCommand.Path -Force
            Remove-Item $updateFile -Force
            
            Write-Log "Update applied successfully - restarting service..." -Level "SUCCESS" -Console
            
            # Restart service
            Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MyInvocation.MyCommand.Path)`"" -WindowStyle Hidden
            exit 0
        }
        catch {
            Write-Log "Update failed: $($_.Exception.Message)" -Level "ERROR" -Console
        }
    }
}

# ============================================================================
# PC RESTART MANAGEMENT
# ============================================================================
function Test-RestartNeeded {
    $restartFile = Join-Path $Global:Config.ScriptPath "restart_counter.txt"
    $currentDate = Get-Date -Format "yyyy-MM-dd"
    
    if (!(Test-Path $restartFile)) {
        Set-Content -Path $restartFile -Value "0|$currentDate"
        return $false
    }
    
    $restartData = Get-Content $restartFile
    $parts = $restartData -split '\|'
    $dayCount = [int]$parts[0]
    $lastDate = $parts[1]
    
    # Check if it's a new day
    if ($lastDate -ne $currentDate) {
        $dayCount++
        Set-Content -Path $restartFile -Value "$dayCount|$currentDate"
    }
    
    # Check if restart is needed (every 3 days at 2 AM)
    $currentTime = Get-Date
    if ($dayCount -ge $Global:Config.MaxRestartDays -and $currentTime.Hour -eq 2 -and $currentTime.Minute -lt 5) {
        Write-Log "Initiating PC restart after $dayCount days" -Level "WARNING" -Console
        Set-Content -Path $restartFile -Value "0|$currentDate"
        
        # Schedule restart in 2 minutes
        shutdown /r /t 120 /c "MoonFlower Automation: Scheduled maintenance restart"
        return $true
    }
    
    return $false
}

# ============================================================================
# CSV DOWNLOAD FUNCTIONS
# ============================================================================
function Invoke-CSVDownload {
    param([int]$SlotNumber, [string]$SlotName)
    
    Write-Log "Starting CSV download - $SlotName (Slot $SlotNumber)" -Level "INFO"
    
    try {
        # Check if Python is available
        $pythonCmd = "python"
        $pythonCheck = Get-Command python -ErrorAction SilentlyContinue
        if (!$pythonCheck) {
            # Try alternative Python locations
            $pythonPaths = @(
                "C:\Python*\python.exe",
                "$env:LOCALAPPDATA\Programs\Python\*\python.exe",
                "$env:PROGRAMFILES\Python*\python.exe",
                "$env:PROGRAMFILES(X86)\Python*\python.exe"
            )
            
            $foundPython = $false
            foreach ($path in $pythonPaths) {
                $pythonExe = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($pythonExe) {
                    $pythonCmd = $pythonExe.FullName
                    $foundPython = $true
                    Write-Log "Found Python at: $pythonCmd" -Level "INFO"
                    break
                }
            }
            
            if (!$foundPython) {
                Write-Log "Python not found anywhere" -Level "ERROR"
                return $false
            }
        }
        
        # Check if CSV downloader exists
        $csvScript = Join-Path $Global:Config.ScriptPath "wifi\csv_downloader.py"
        if (!(Test-Path $csvScript)) {
            Write-Log "CSV downloader script not found: $csvScript" -Level "ERROR"
            return $false
        }
        
        # Run CSV download with timeout
        $process = Start-Process -FilePath $pythonCmd -ArgumentList "`"$csvScript`" --slot $SlotNumber" -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "csv_output.log" -RedirectStandardError "csv_error.log"
        
        if ($process.ExitCode -eq 0) {
            Write-Log "$SlotName completed successfully" -Level "SUCCESS"
            return $true
        } else {
            $errorContent = Get-Content "csv_error.log" -ErrorAction SilentlyContinue
            Write-Log "$SlotName failed with exit code $($process.ExitCode): $errorContent" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "$SlotName failed with exception: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
    finally {
        # Clean up log files
        Remove-Item "csv_output.log", "csv_error.log" -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-AllCSVDownloads {
    Write-Log "=== STARTING CSV DOWNLOADS ===" -Level "INFO"
    
    $results = @()
    
    # Slot 1: 9:30 AM
    $results += Invoke-CSVDownload -SlotNumber 1 -SlotName "Morning Slot (9:30 AM)"
    Start-Sleep -Seconds 60
    
    # Slot 2: 1:00 PM
    $results += Invoke-CSVDownload -SlotNumber 2 -SlotName "Afternoon Slot (1:00 PM)"
    Start-Sleep -Seconds 60
    
    # Slot 3: 1:30 PM (Backup)
    $results += Invoke-CSVDownload -SlotNumber 3 -SlotName "Backup Slot (1:30 PM)"
    
    $successCount = ($results | Where-Object { $_ -eq $true }).Count
    Write-Log "CSV Downloads completed: $successCount/3 successful" -Level "INFO"
    
    Write-Log "=== CSV DOWNLOADS COMPLETED ===" -Level "INFO"
    return $successCount -gt 0
}

# ============================================================================
# EXCEL MERGE FUNCTION
# ============================================================================
function Invoke-ExcelMerge {
    Write-Log "=== STARTING EXCEL MERGE ===" -Level "INFO"
    
    try {
        # Check CSV file count
        $dateFolder = (Get-Date).ToString("ddMMM").ToLower()
        $csvDir = Join-Path $Global:Config.ScriptPath "EHC_Data\$dateFolder"
        
        if (!(Test-Path $csvDir)) {
            Write-Log "CSV directory not found: $csvDir" -Level "ERROR"
            return $false
        }
        
        $csvFiles = Get-ChildItem -Path $csvDir -Filter "*.csv"
        $csvCount = $csvFiles.Count
        
        Write-Log "Found $csvCount CSV files in $csvDir" -Level "INFO"
        
        if ($csvCount -lt 2) {
            Write-Log "Need at least 2 CSV files for merging, found $csvCount" -Level "ERROR"
            return $false
        }
        
        # Run Excel generation using Python
        Write-Log "Generating Excel file from $csvCount CSV files..." -Level "INFO"
        
        $pythonCode = @"
from excel.excel_generator import ExcelGenerator
from utils.file_manager import FileManager
import sys

try:
    fm = FileManager()
    eg = ExcelGenerator(fm)
    result = eg.execute_excel_generation()
    
    if result.get('success'):
        print(f'SUCCESS: Excel generated - {result.get("excel_file")}')
        print(f'Rows processed: {result.get("total_rows")}')
        sys.exit(0)
    else:
        print(f'ERROR: Excel generation failed - {result.get("error")}')
        sys.exit(1)
except Exception as e:
    print(f'EXCEPTION: {str(e)}')
    sys.exit(1)
"@
        
        $process = Start-Process -FilePath "python" -ArgumentList "-c", "`"$pythonCode`"" -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "excel_output.log" -RedirectStandardError "excel_error.log"
        
        if ($process.ExitCode -eq 0) {
            $output = Get-Content "excel_output.log" -ErrorAction SilentlyContinue
            Write-Log "Excel merge completed successfully: $output" -Level "SUCCESS"
            return $true
        } else {
            $error = Get-Content "excel_error.log" -ErrorAction SilentlyContinue
            Write-Log "Excel merge failed: $error" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Excel merge failed with exception: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
    finally {
        # Clean up log files
        Remove-Item "excel_output.log", "excel_error.log" -Force -ErrorAction SilentlyContinue
        Write-Log "=== EXCEL MERGE COMPLETED ===" -Level "INFO"
    }
}

# ============================================================================
# VBS AUTOMATION FUNCTION
# ============================================================================
function Invoke-VBSPhase1 {
    Write-Log "=== STARTING VBS PHASE 1 ===" -Level "INFO"
    
    try {
        # Check if VBS Phase 1 script exists
        $vbsScript = Join-Path $Global:Config.ScriptPath "vbs\vbs_phase1_login.py"
        if (!(Test-Path $vbsScript)) {
            Write-Log "VBS Phase 1 script not found: $vbsScript" -Level "ERROR"
            return $false
        }
        
        Write-Log "Running VBS Phase 1 (Login and Import)..." -Level "INFO"
        
        $process = Start-Process -FilePath "python" -ArgumentList "`"$vbsScript`"" -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "vbs_output.log" -RedirectStandardError "vbs_error.log"
        
        if ($process.ExitCode -eq 0) {
            Write-Log "VBS Phase 1 completed successfully" -Level "SUCCESS"
            return $true
        } else {
            $error = Get-Content "vbs_error.log" -ErrorAction SilentlyContinue
            Write-Log "VBS Phase 1 failed: $error" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "VBS Phase 1 failed with exception: $($_.Exception.Message)" -Level "WARNING"
        return $false
    }
    finally {
        # Clean up log files
        Remove-Item "vbs_output.log", "vbs_error.log" -Force -ErrorAction SilentlyContinue
        Write-Log "=== VBS PHASE 1 COMPLETED ===" -Level "INFO"
    }
}

# ============================================================================
# EMAIL NOTIFICATION FUNCTION
# ============================================================================
function Send-EmailNotification {
    Write-Log "=== SENDING EMAIL NOTIFICATION ===" -Level "INFO"
    
    try {
        # Check if email system is available
        $emailScript = Join-Path $Global:Config.ScriptPath "email\email_delivery.py"
        if (!(Test-Path $emailScript)) {
            Write-Log "Email system not available: $emailScript" -Level "WARNING"
            return $false
        }
        
        Write-Log "Sending notification: Time to upload Excel file" -Level "INFO"
        
        $process = Start-Process -FilePath "python" -ArgumentList "`"$emailScript`" --type upload_ready" -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "email_output.log" -RedirectStandardError "email_error.log"
        
        if ($process.ExitCode -eq 0) {
            Write-Log "Email notification sent successfully" -Level "SUCCESS"
            return $true
        } else {
            $error = Get-Content "email_error.log" -ErrorAction SilentlyContinue
            Write-Log "Email notification failed: $error" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Email notification failed with exception: $($_.Exception.Message)" -Level "WARNING"
        return $false
    }
    finally {
        # Clean up log files
        Remove-Item "email_output.log", "email_error.log" -Force -ErrorAction SilentlyContinue
        Write-Log "=== EMAIL NOTIFICATION COMPLETED ===" -Level "INFO"
    }
}

# ============================================================================
# MAIN AUTOMATION CYCLE
# ============================================================================
function Invoke-AutomationCycle {
    Write-Log "=========================================" -Level "INFO" -Console
    Write-Log "STARTING DAILY AUTOMATION CYCLE" -Level "INFO" -Console
    Write-Log "Date: $(Get-Date -Format 'yyyy-MM-dd')" -Level "INFO" -Console
    Write-Log "Time: $(Get-Date -Format 'HH:mm:ss')" -Level "INFO" -Console
    Write-Log "Version: $($Global:Config.Version)" -Level "INFO" -Console
    Write-Log "=========================================" -Level "INFO" -Console
    
    $results = @{
        CSVSuccess = $false
        ExcelSuccess = $false
        VBSSuccess = $false
        EmailSuccess = $false
    }
    
    try {
        # Step 1: Run CSV Downloads
        $results.CSVSuccess = Invoke-AllCSVDownloads
        
        # Step 2: Merge CSV files into Excel (only if CSV was successful)
        if ($results.CSVSuccess) {
            $results.ExcelSuccess = Invoke-ExcelMerge
            
            # Step 3: Run VBS Phase 1 (only if Excel was successful)
            if ($results.ExcelSuccess) {
                $results.VBSSuccess = Invoke-VBSPhase1
                
                # Step 4: Send email notification
                $results.EmailSuccess = Send-EmailNotification
            } else {
                Write-Log "Skipping VBS and email due to Excel merge failure" -Level "WARNING"
            }
        } else {
            Write-Log "Skipping Excel, VBS, and email due to CSV download failure" -Level "WARNING"
        }
        
        # Update last run timestamp
        $lastRunFile = Join-Path $Global:Config.ScriptPath "last_run.json"
        $lastRunData = @{
            Date = Get-Date -Format "yyyy-MM-dd"
            Time = Get-Date -Format "HH:mm:ss"
            Results = $results
        }
        $lastRunData | ConvertTo-Json | Set-Content -Path $lastRunFile
        
        Write-Log "=========================================" -Level "INFO" -Console
        Write-Log "DAILY AUTOMATION CYCLE COMPLETED" -Level "INFO" -Console
        Write-Log "CSV Success: $($results.CSVSuccess)" -Level "INFO" -Console
        Write-Log "Excel Success: $($results.ExcelSuccess)" -Level "INFO" -Console
        Write-Log "VBS Success: $($results.VBSSuccess)" -Level "INFO" -Console
        Write-Log "Email Success: $($results.EmailSuccess)" -Level "INFO" -Console
        Write-Log "=========================================" -Level "INFO" -Console
        
        return $results
    }
    catch {
        Write-Log "Automation cycle failed with exception: $($_.Exception.Message)" -Level "ERROR" -Console
        return $results
    }
}

# ============================================================================
# TIME MANAGEMENT
# ============================================================================
function Test-TimeSlot {
    $currentTime = Get-Date -Format "HH:mm"
    $currentDate = Get-Date -Format "yyyy-MM-dd"
    
    # Check if it's time for automation (9:30 AM)
    if ($currentTime -eq $Global:Config.TimeSlots.Slot1) {
        # Check if already run today
        $lastRunFile = Join-Path $Global:Config.ScriptPath "last_run.json"
        if (Test-Path $lastRunFile) {
            try {
                $lastRun = Get-Content $lastRunFile | ConvertFrom-Json
                if ($lastRun.Date -eq $currentDate) {
                    Write-Log "Automation already run today, skipping" -Level "INFO"
                    return $false
                }
            }
            catch {
                Write-Log "Could not read last run file, proceeding with automation" -Level "WARNING"
            }
        }
        
        Write-Log "Time slot reached: $currentTime - Starting automation cycle" -Level "INFO" -Console
        return $true
    }
    
    return $false
}

# ============================================================================
# SERVICE STATUS
# ============================================================================
function Show-ServiceStatus {
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "MoonFlower Automation Service Status" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Service Information
    Write-Host "Service Information:" -ForegroundColor Yellow
    Write-Host "  Name: $($Global:Config.ServiceName)"
    Write-Host "  Version: $($Global:Config.Version)"
    Write-Host "  Script Path: $($Global:Config.ScriptPath)"
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
    
    $regValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name $Global:Config.ServiceName -ErrorAction SilentlyContinue
    if ($regValue) {
        Write-Host "  Registry Entry: " -NoNewline
        Write-Host "INSTALLED" -ForegroundColor Green
    } else {
        Write-Host "  Registry Entry: " -NoNewline
        Write-Host "NOT INSTALLED" -ForegroundColor Red
    }
    Write-Host ""
    
    # File Structure
    Write-Host "File Structure:" -ForegroundColor Yellow
    $dateFolder = (Get-Date).ToString("ddMMM").ToLower()
    
    @("EHC_Data", "EHC_Data_Merge", "EHC_Data_Pdf", "EHC_Logs") | ForEach-Object {
        $dir = Join-Path $Global:Config.ScriptPath $_
        if (Test-Path $dir) {
            Write-Host "  $_: " -NoNewline
            Write-Host "EXISTS" -ForegroundColor Green
            
            $todayDir = Join-Path $dir $dateFolder
            if (Test-Path $todayDir) {
                $fileCount = (Get-ChildItem $todayDir -File).Count
                Write-Host "    Today's folder: $fileCount files"
            }
        } else {
            Write-Host "  $_: " -NoNewline
            Write-Host "MISSING" -ForegroundColor Red
        }
    }
    Write-Host ""
    
    # Last Run Information
    Write-Host "Last Run Information:" -ForegroundColor Yellow
    $lastRunFile = Join-Path $Global:Config.ScriptPath "last_run.json"
    if (Test-Path $lastRunFile) {
        try {
            $lastRun = Get-Content $lastRunFile | ConvertFrom-Json
            Write-Host "  Date: $($lastRun.Date)"
            Write-Host "  Time: $($lastRun.Time)"
            Write-Host "  CSV Success: $($lastRun.Results.CSVSuccess)"
            Write-Host "  Excel Success: $($lastRun.Results.ExcelSuccess)"
            Write-Host "  VBS Success: $($lastRun.Results.VBSSuccess)"
            Write-Host "  Email Success: $($lastRun.Results.EmailSuccess)"
        }
        catch {
            Write-Host "  Could not read last run information" -ForegroundColor Red
        }
    } else {
        Write-Host "  No previous runs recorded" -ForegroundColor Yellow
    }
    Write-Host ""
    
    # Next Run Time
    Write-Host "Next Run Time:" -ForegroundColor Yellow
    Write-Host "  Scheduled for: $($Global:Config.TimeSlots.Slot1) daily"
    Write-Host "  Current time: $(Get-Date -Format 'HH:mm')"
    Write-Host ""
}

# ============================================================================
# MAIN SERVICE LOOP
# ============================================================================
function Start-ServiceLoop {
    Write-Log "=========================================" -Level "INFO" -Console
    Write-Log "MoonFlower WiFi Automation Service Started" -Level "INFO" -Console
    Write-Log "Version: $($Global:Config.Version)" -Level "INFO" -Console
    Write-Log "Script Path: $($Global:Config.ScriptPath)" -Level "INFO" -Console
    Write-Log "=========================================" -Level "INFO" -Console
    
    # Install service if not already installed
    $taskExists = Get-ScheduledTask -TaskName $Global:Config.ServiceName -ErrorAction SilentlyContinue
    if (!$taskExists) {
        Write-Log "Service not installed, installing now..." -Level "INFO" -Console
        Install-AutomationService | Out-Null
    }
    
    # Main service loop
    while ($true) {
        try {
            # Check for updates
            Test-ForUpdates
            
            # Check if PC restart is needed
            if (Test-RestartNeeded) {
                Write-Log "PC restart initiated, service will resume after reboot" -Level "INFO" -Console
                break
            }
            
            # Check time slots for automation
            if (Test-TimeSlot) {
                Invoke-AutomationCycle
            }
            
            # Wait before next check
            Start-Sleep -Seconds ($Global:Config.CheckIntervalMinutes * 60)
        }
        catch {
            Write-Log "Error in service loop: $($_.Exception.Message)" -Level "ERROR" -Console
            Start-Sleep -Seconds 300  # Wait 5 minutes before retrying
        }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Initialize logging
Initialize-Logging

# Set working directory
Set-Location $Global:Config.ScriptPath

# Handle command line arguments
switch ($Action.ToLower()) {
    "install" {
        if (Install-AutomationService) {
            Write-Host "Service installed successfully!" -ForegroundColor Green
            exit 0
        } else {
            Write-Host "Service installation failed!" -ForegroundColor Red
            exit 1
        }
    }
    
    "uninstall" {
        if (Uninstall-AutomationService) {
            Write-Host "Service uninstalled successfully!" -ForegroundColor Green
            exit 0
        } else {
            Write-Host "Service uninstallation failed!" -ForegroundColor Red
            exit 1
        }
    }
    
    "status" {
        Show-ServiceStatus
        exit 0
    }
    
    "manual" {
        Write-Log "Manual execution requested" -Level "INFO" -Console
        $results = Invoke-AutomationCycle
        Write-Host "Manual execution completed. Check logs for details." -ForegroundColor Green
        exit 0
    }
    
    "csv" {
        Write-Log "CSV download only requested" -Level "INFO" -Console
        $success = Invoke-AllCSVDownloads
        if ($success) {
            Write-Host "CSV downloads completed successfully!" -ForegroundColor Green
        } else {
            Write-Host "CSV downloads failed!" -ForegroundColor Red
        }
        exit 0
    }
    
    "excel" {
        Write-Log "Excel merge only requested" -Level "INFO" -Console
        $success = Invoke-ExcelMerge
        if ($success) {
            Write-Host "Excel merge completed successfully!" -ForegroundColor Green
        } else {
            Write-Host "Excel merge failed!" -ForegroundColor Red
        }
        exit 0
    }
    
    "vbs" {
        Write-Log "VBS Phase 1 only requested" -Level "INFO" -Console
        $success = Invoke-VBSPhase1
        if ($success) {
            Write-Host "VBS Phase 1 completed successfully!" -ForegroundColor Green
        } else {
            Write-Host "VBS Phase 1 failed!" -ForegroundColor Red
        }
        exit 0
    }
    
    default {
        # Start main service
        Start-ServiceLoop
    }
}