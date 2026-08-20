<#
.SYNOPSIS
    Advanced Printer Management Tool for Windows Environments.
.DESCRIPTION
    Automates printer installation, removal, test page dispatch, spooler cleanup,
    and inventorying with interactive TUI, GUI dialogs, logs, and progress reporting.
.VERSION
    3.1.0
#>

# Load Windows Forms for GUI dialogs if available
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
} catch {
    # Non-GUI environments (e.g. Server Core)
}

# Configuration
$script:LogFile = Join-Path -Path $(if ($PSScriptRoot) { $PSScriptRoot } else { "." }) -ChildPath "PrinterManagement.log"

# Logging function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "SUCCESS", "WARN", "ERROR")]
        [string]$Level = "INFO",

        [string]$CustomLogPath = $script:LogFile
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "[$timestamp] [$Level] $Message"

    $colorMap = @{
        "INFO"    = "Cyan"
        "SUCCESS" = "Green"
        "WARN"    = "Yellow"
        "ERROR"   = "Red"
    }

    $iconMap = @{
        "INFO"    = "[i]"
        "SUCCESS" = "[OK]"
        "WARN"    = "[!]"
        "ERROR"   = "[FAIL]"
    }

    Write-Host "$($iconMap[$Level]) $Message" -ForegroundColor $colorMap[$Level]

    # Append to log file
    try {
        Add-Content -Path $CustomLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Silent continue if log file write fails
    }
}

# Function to display banner
function Show-Banner {
    Clear-Host
    Write-Host @'
 +============================================================================+
 |   ___      _      __          __  ___                                      |
 |  / _ \____(_)__  / /____ ____/  |/  /__ ____  ___ ____ ____ __ _  ___ ___  |
 | / ___/ __/ / _ \/ __/ -_) __/ /|_/ / _ `/ _ \/ _ `/ _ `/ -_)  ' \/ -_) _ \ |
 |/_/  /_/ /_/_//_/\__/\__/_/ /_/  /_/\_,_/_//_/\_,_/\_, /\__/_/_/_/\__/_//_/ |
 |                                                  /___/                     |
 |                      Windows Printer Management Suite                      |
 |                                Version 3.1.0                               |
 +============================================================================+
'@ -ForegroundColor Cyan
}

# GUI & CLI File Picker
function Get-ValidFilePath {
    param (
        [string]$FilePath = "",
        [string]$Title = "Select Printer CSV File",
        [string]$Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    )

    if (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $clean = $FilePath.Trim('"').Trim("'")
        if (Test-Path -Path $clean -PathType Leaf) {
            return (Resolve-Path -Path $clean).Path
        }
        Write-Log "Specified path '$clean' does not exist." "ERROR"
        return $null
    }

    Write-Host "`nSelect CSV File:" -ForegroundColor Yellow
    Write-Host "  - Press [Enter] to open GUI File Explorer" -ForegroundColor DarkGray
    Write-Host "  - Or type / paste the file path directly" -ForegroundColor DarkGray
    $inputPath = Read-Host -Prompt "File path"

    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        # Attempt to open GUI File Dialog
        try {
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = [System.IO.Directory]::GetCurrentDirectory()
            $openFileDialog.Filter = $Filter
            $openFileDialog.Title = $Title
            $openFileDialog.ShowHelp = $false

            $dialogResult = $openFileDialog.ShowDialog()
            if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Log "Selected file via GUI: $($openFileDialog.FileName)" "INFO"
                return $openFileDialog.FileName
            } else {
                Write-Log "File selection cancelled by user." "WARN"
                return $null
            }
        } catch {
            Write-Log "GUI file picker unavailable in this environment. Please enter path manually." "WARN"
            $manualPath = Read-Host -Prompt "Enter CSV file path"
            if (-not [string]::IsNullOrWhiteSpace($manualPath) -and (Test-Path -Path $manualPath.Trim('"').Trim("'") -PathType Leaf)) {
                return (Resolve-Path -Path $manualPath.Trim('"').Trim("'")).Path
            }
            return $null
        }
    } else {
        $cleanPath = $inputPath.Trim('"').Trim("'")
        if (Test-Path -Path $cleanPath -PathType Leaf) {
            $resolved = (Resolve-Path -Path $cleanPath).Path
            Write-Log "Selected file via CLI: $resolved" "INFO"
            return $resolved
        } else {
            Write-Log "The path '$cleanPath' is invalid or does not exist." "ERROR"
            return $null
        }
    }
}

# Smart CSV Importer
function Import-SmartCsv {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Log "File not found: $Path" "ERROR"
        return $null
    }

    # Detect delimiter safely
    $firstLine = (Get-Content -Path $Path -TotalCount 1)
    $delimiter = ';'
    if ($firstLine -match ';') {
        $delimiter = ';'
    } elseif ($firstLine -match ',') {
        $delimiter = ','
    } elseif ($firstLine -match "`t") {
        $delimiter = "`t"
    }

    Write-Log "Reading CSV with detected delimiter: '$delimiter'" "INFO"

    try {
        $data = Import-Csv -Path $Path -Delimiter $delimiter -Encoding UTF8 -ErrorAction Stop

        # Normalize properties
        $normalizedList = @()
        foreach ($row in $data) {
            $nameVal = if ($row.Name) { $row.Name } elseif ($row.PrinterName) { $row.PrinterName } else { "" }
            $portVal = if ($row.LocalPort) { $row.LocalPort } elseif ($row.Port) { $row.Port } elseif ($row.PortName) { $row.PortName } else { "" }
            $driverVal = if ($row.DriverName) { $row.DriverName } elseif ($row.Driver) { $row.Driver } else { "" }

            if (-not [string]::IsNullOrWhiteSpace($nameVal)) {
                $normalizedList += [PSCustomObject]@{
                    Name       = $nameVal.ToString().Trim()
                    LocalPort  = $portVal.ToString().Trim()
                    DriverName = $driverVal.ToString().Trim()
                }
            }
        }
        return $normalizedList
    } catch {
        Write-Log "Error parsing CSV file: $_" "ERROR"
        return $null
    }
}

# Function to add printers
function Add-Printers {
    param (
        [string]$FilePath = ""
    )

    Write-Host "`n=========================== [ 1. ADD PRINTERS ] ===========================" -ForegroundColor Yellow

    $printersFile = Get-ValidFilePath -FilePath $FilePath -Title "Select CSV File to Add Printers"
    if (-not $printersFile) {
        if (-not $FilePath) { Read-Host -Prompt "`nPress Enter to return to menu..." }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    $printerList = @(Import-SmartCsv -Path $printersFile)
    if (-not $printerList -or $printerList.Count -eq 0) {
        Write-Log "No valid printer records found in CSV." "WARN"
        if (-not $FilePath) { Read-Host -Prompt "`nPress Enter to return to menu..." }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    $total = $printerList.Count
    Write-Log "Starting installation of $total printer(s)..." "INFO"

    $index = 0
    $successCount = 0
    $failCount = 0

    foreach ($printer in $printerList) {
        $index++
        $percent = [int](($index / $total) * 100)
        Write-Progress -Activity "Installing Printers" -Status "[$index/$total] Processing: $($printer.Name)" -PercentComplete $percent

        $pName   = $printer.Name
        $pPort   = $printer.LocalPort
        $pDriver = $printer.DriverName

        # Case 1: Shared Network Printer (UNC Path: \\server\printer)
        if ($pPort -like "\\*" -or $pName -like "\\*") {
            $connectionPath = if ($pPort -like "\\*") { $pPort } else { $pName }
            Write-Log "Connecting to network shared printer: '$connectionPath'..." "INFO"
            try {
                Add-Printer -ConnectionName $connectionPath -ErrorAction Stop
                Write-Log "Successfully connected to shared printer '$connectionPath'." "SUCCESS"
                $successCount++
            } catch {
                Write-Log "Failed to connect to shared printer '$connectionPath': $_" "ERROR"
                $failCount++
            }
            continue
        }

        # Case 2: Standard Printer (Local / TCP-IP)
        if ([string]::IsNullOrWhiteSpace($pPort) -or [string]::IsNullOrWhiteSpace($pDriver)) {
            Write-Log "Skipping '$pName': Missing Port or Driver." "WARN"
            $failCount++
            continue
        }

        # Create Port if needed
        $portExists = Get-PrinterPort -Name $pPort -ErrorAction SilentlyContinue
        if (-not $portExists) {
            try {
                # If IPv4 address or standard TCP hostname
                $isIpAddress = $pPort -match "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
                $isHostname  = ($pPort -match "^[a-zA-Z0-9\.\-_]+$" -and $pPort -notmatch "^(USB|COM|LPT|FILE|PORTPROMPT|NUL)\d*[:]?$")

                if ($isIpAddress -or $isHostname) {
                    Add-PrinterPort -Name $pPort -PrinterHostAddress $pPort -ErrorAction Stop
                } else {
                    Add-PrinterPort -Name $pPort -ErrorAction Stop
                }
                Write-Log "Created printer port '$pPort'." "SUCCESS"
            } catch {
                Write-Log "Error creating port '$pPort': $_" "ERROR"
                $failCount++
                continue
            }
        }

        # Verify Driver
        $driverExists = Get-PrinterDriver -Name $pDriver -ErrorAction SilentlyContinue
        if (-not $driverExists) {
            Write-Log "Driver '$pDriver' is NOT installed on this system. Skipping '$pName'." "ERROR"
            $failCount++
            continue
        }

        # Add Printer
        try {
            Add-Printer -Name $pName -PortName $pPort -DriverName $pDriver -ErrorAction Stop
            Write-Log "Installed printer '$pName' on port '$pPort' with driver '$pDriver'." "SUCCESS"
            $successCount++
        } catch {
            Write-Log "Failed to add printer '$pName': $_" "ERROR"
            $failCount++
        }
    }

    Write-Progress -Activity "Installing Printers" -Completed
    Write-Host "`n---------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log "Finished: $successCount installed successfully, $failCount failed." $(if ($failCount -eq 0) { "SUCCESS" } else { "WARN" })
    if (-not $FilePath) { Read-Host -Prompt "`nPress Enter to return to menu..." }

    return @{ Total = $total; Success = $successCount; Failed = $failCount }
}

# Function to remove printers
function Remove-Printers {
    param (
        [string]$FilePath = "",
        [array]$PrinterList = @(),
        [switch]$Force
    )

    Write-Host "`n========================= [ 2. REMOVE PRINTERS ] =========================" -ForegroundColor Yellow

    $printersToRemove = @()

    if ($PrinterList -and $PrinterList.Count -gt 0) {
        $printersToRemove = $PrinterList
    } elseif (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $printersToRemove = @(Import-SmartCsv -Path $FilePath)
    } else {
        Write-Host " [1] Remove printers specified in CSV" -ForegroundColor Cyan
        Write-Host " [2] Select printers interactively from GUI list" -ForegroundColor Cyan
        Write-Host " [3] Cancel" -ForegroundColor DarkGray
        $subOption = Read-Host "`nSelect an option (1-3)"

        if ($subOption -eq "1") {
            $printersFile = Get-ValidFilePath -Title "Select CSV File to Remove Printers"
            if (-not $printersFile) { return @{ Total = 0; Success = 0; Failed = 0 } }
            $printersToRemove = @(Import-SmartCsv -Path $printersFile)
        } elseif ($subOption -eq "2") {
            try {
                $installed = Get-CimInstance -ClassName Win32_Printer | Select-Object Name, PortName, DriverName
                $selected = $installed | Out-GridView -Title "Select printers to remove (Hold Ctrl/Shift to select multiple)" -PassThru
                if ($selected) {
                    $printersToRemove = $selected | ForEach-Object {
                        [PSCustomObject]@{
                            Name      = $_.Name
                            LocalPort = $_.PortName
                        }
                    }
                } else {
                    Write-Log "No printers selected." "WARN"
                    Read-Host -Prompt "`nPress Enter to return to menu..."
                    return @{ Total = 0; Success = 0; Failed = 0 }
                }
            } catch {
                Write-Log "Error opening selection grid: $_" "ERROR"
                return @{ Total = 0; Success = 0; Failed = 0 }
            }
        } else {
            return @{ Total = 0; Success = 0; Failed = 0 }
        }
    }

    if (-not $printersToRemove -or $printersToRemove.Count -eq 0) {
        Write-Log "No printers specified for removal." "WARN"
        if (-not $Force) { Read-Host -Prompt "`nPress Enter to return to menu..." }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    # Security Confirmation
    if (-not $Force) {
        Write-Host "`n[!] ATTENTION: You are about to remove $($printersToRemove.Count) printer(s)." -ForegroundColor Yellow
        $confirm = Read-Host "Are you sure you want to proceed? (Y/N)"
        if ($confirm -notmatch '^(y|s|yes|si)$') {
            Write-Log "Printer removal aborted by user." "INFO"
            Read-Host -Prompt "`nPress Enter to return to menu..."
            return @{ Total = 0; Success = 0; Failed = 0 }
        }
    }

    $total = $printersToRemove.Count
    $index = 0
    $successCount = 0
    $failCount = 0

    foreach ($item in $printersToRemove) {
        $index++
        $pName = $item.Name
        $pPort = if ($item.LocalPort) { $item.LocalPort } else { "" }

        Write-Progress -Activity "Removing Printers" -Status "[$index/$total] Removing: $pName" -PercentComplete $([int](($index / $total) * 100))

        # Remove Printer
        $printerExists = Get-Printer -Name $pName -ErrorAction SilentlyContinue
        if ($printerExists) {
            try {
                Remove-Printer -Name $pName -ErrorAction Stop
                Write-Log "Printer '$pName' removed." "SUCCESS"
                $successCount++
            } catch {
                Write-Log "Error removing printer '$pName': $_" "ERROR"
                $failCount++
            }
        } else {
            Write-Log "Printer '$pName' does not exist." "WARN"
        }

        # Remove Port if provided and port exists
        if (-not [string]::IsNullOrWhiteSpace($pPort)) {
            $portExists = Get-PrinterPort -Name $pPort -ErrorAction SilentlyContinue
            if ($portExists) {
                try {
                    Remove-PrinterPort -Name $pPort -ErrorAction Stop
                    Write-Log "Port '$pPort' removed." "SUCCESS"
                } catch {
                    Write-Log "Error removing port '$pPort': $_" "WARN"
                }
            }
        }
    }

    Write-Progress -Activity "Removing Printers" -Completed
    Write-Host "`n---------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log "Removal process completed. Processed: $total printer(s)." "SUCCESS"
    if (-not $Force) { Read-Host -Prompt "`nPress Enter to return to menu..." }

    return @{ Total = $total; Success = $successCount; Failed = $failCount }
}

# Function to send test pages
function Send-TestPages {
    param (
        [string]$FilePath = "",
        [array]$PrinterList = @()
    )

    Write-Host "`n======================= [ 3. SEND TEST PAGES ] =======================" -ForegroundColor Yellow

    $printersToTest = @()
    if ($PrinterList -and $PrinterList.Count -gt 0) {
        $printersToTest = $PrinterList
    } elseif (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $printersToTest = @(Import-SmartCsv -Path $FilePath)
    } else {
        $printersFile = Get-ValidFilePath -Title "Select CSV File for Test Pages"
        if (-not $printersFile) {
            Read-Host -Prompt "`nPress Enter to return to menu..."
            return @{ Total = 0; Success = 0; Failed = 0 }
        }
        $printersToTest = @(Import-SmartCsv -Path $printersFile)
    }

    if (-not $printersToTest -or $printersToTest.Count -eq 0) {
        Write-Log "No valid printers found in CSV." "WARN"
        if (-not $FilePath) { Read-Host -Prompt "`nPress Enter to return to menu..." }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    $total = $printersToTest.Count
    $index = 0
    $successCount = 0
    $failCount = 0

    Write-Log "Sending test pages to $total printer(s)..." "INFO"

    foreach ($printer in $printersToTest) {
        $index++
        $pName = $printer.Name
        Write-Progress -Activity "Dispatching Test Pages" -Status "[$index/$total] Testing: $pName" -PercentComplete $([int](($index / $total) * 100))

        $printerObj = Get-Printer -Name $pName -ErrorAction SilentlyContinue
        if (-not $printerObj) {
            Write-Log "Printer '$pName' is not installed locally. Skipping." "WARN"
            $failCount++
            continue
        }

        # Escape single quotes in printer name for WQL filter
        $escapedName = $pName -replace "'", "''"

        try {
            $wmiPrinter = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$escapedName'" -ErrorAction Stop
            $result = Invoke-CimMethod -InputObject $wmiPrinter -MethodName "PrintTestPage" -ErrorAction Stop
            if ($result.ReturnValue -eq 0) {
                Write-Log "Test page dispatched to '$pName'." "SUCCESS"
                $successCount++
            } else {
                Write-Log "PrintTestPage returned code $($result.ReturnValue) for '$pName'." "WARN"
                $failCount++
            }
        } catch {
            try {
                Start-Process rundll32.exe -ArgumentList "printui.dll,PrintUIEntry /k /n `"$pName`"" -NoNewWindow -Wait
                Write-Log "Test page dispatched via printui for '$pName'." "SUCCESS"
                $successCount++
            } catch {
                Write-Log "Failed to send test page to '$pName': $_" "ERROR"
                $failCount++
            }
        }
    }

    Write-Progress -Activity "Dispatching Test Pages" -Completed
    Write-Host "`n---------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log "Test pages dispatch process completed: $successCount dispatched, $failCount failed." "SUCCESS"
    if (-not $FilePath) { Read-Host -Prompt "`nPress Enter to return to menu..." }

    return @{ Total = $total; Success = $successCount; Failed = $failCount }
}

# Function to clear print queues
function Clear-PrintQueues {
    param (
        [switch]$Force
    )

    Write-Host "`n====================== [ 4. CLEAR PRINT QUEUES ] ======================" -ForegroundColor Yellow
    Write-Log "This operation will stop the Spooler service and purge all pending jobs." "WARN"

    if (-not $Force) {
        $confirm = Read-Host "`nAre you sure you want to stop Spooler and clear queues? (Y/N)"
        if ($confirm -notmatch '^(y|s|yes|si)$') {
            Write-Log "Print queue clearing cancelled." "INFO"
            Read-Host -Prompt "`nPress Enter to return to menu..."
            return $false
        }
    }

    try {
        Write-Log "Stopping Spooler service..." "INFO"
        Stop-Service -Name Spooler -Force -ErrorAction Stop

        $spooler = Get-Service -Name Spooler
        $timeout = 10
        while ($spooler.Status -ne 'Stopped' -and $timeout -gt 0) {
            Start-Sleep -Milliseconds 500
            $spooler.Refresh()
            $timeout--
        }

        # Extra 1s wait to ensure file handles release
        Start-Sleep -Seconds 1

        Write-Log "Purging spool files from system directory..." "INFO"
        Remove-Item -Path "$env:windir\System32\spool\PRINTERS\*" -Force -Recurse -ErrorAction SilentlyContinue

        Write-Log "Restarting Spooler service..." "INFO"
        Start-Service -Name Spooler -ErrorAction Stop

        Write-Log "Print queues cleared and Spooler restarted successfully!" "SUCCESS"
        if (-not $Force) { Read-Host -Prompt "`nPress Enter to return to menu..." }
        return $true
    } catch {
        Write-Log "Error while clearing print queues: $_" "ERROR"
        if (-not $Force) { Read-Host -Prompt "`nPress Enter to return to menu..." }
        return $false
    }
}

# Function to inventory printers
function Inventory-Printers {
    param (
        [string]$OutputPath = ".\inventory.csv",
        [switch]$NoGrid,
        [switch]$NoPrompt
    )

    Write-Host "`n===================== [ 5. INVENTORY PRINTERS ] =====================" -ForegroundColor Yellow

    try {
        Write-Log "Gathering system printer information..." "INFO"
        $printers = @(Get-CimInstance -ClassName Win32_Printer | Select-Object Name, DriverName, PortName, ShareName, Published, PrinterStatus, Default, Location, Comment)

        if (-not $printers -or $printers.Count -eq 0) {
            Write-Log "No installed printers found on this system." "WARN"
            if (-not $NoGrid -and -not $NoPrompt) { Read-Host -Prompt "`nPress Enter to return to menu..." }
            return @()
        }

        # Export to CSV
        $printers | Export-Csv -Path $OutputPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "Inventory exported to '$OutputPath' ($($printers.Count) printer(s) found)." "SUCCESS"

        # Format and display clean console table
        Write-Host "`nInstalled Printers on this System:" -ForegroundColor Cyan
        $displayData = $printers | Select-Object @{Name="Printer Name"; Expression={$_.Name}},
                                                 @{Name="Port"; Expression={$_.PortName}},
                                                 @{Name="Driver"; Expression={$_.DriverName}},
                                                 @{Name="Default"; Expression={if ($_.Default) { "YES" } else { "No" }}}
        $displayData | Format-Table -AutoSize | Out-String | Write-Host

        return $printers
    } catch {
        Write-Log "Error generating printer inventory: $_" "ERROR"
        return $null
    } finally {
        if (-not $NoGrid -and -not $NoPrompt) { Read-Host -Prompt "`nPress Enter to return to menu..." }
    }
}

# Function to generate CSV template
function New-PrinterTemplateCsv {
    param (
        [string]$OutputPath = ".\template_printers.csv",
        [switch]$NoOpen
    )

    Write-Host "`n==================== [ 6. GENERATE CSV TEMPLATE ] ====================" -ForegroundColor Yellow

    $sampleData = @"
Name;LocalPort;DriverName
Office_HP_LaserJet;192.168.1.50;HP Universal Printing PCL 6
Finance_Canon_MFP;\\printserver01\Canon_Finance;Canon Generic Plus PCL6
Warehouse_Zebra_Labels;USB001;ZDesigner ZD420-203dpi ZPL
HR_Epson_WorkForce;192.168.1.55;EPSON WF-C5790 Series
"@

    try {
        Set-Content -Path $OutputPath -Value $sampleData -Encoding UTF8 -ErrorAction Stop
        Write-Log "Template CSV created successfully at: $OutputPath" "SUCCESS"
        Write-Host "`nSample content generated:" -ForegroundColor Cyan
        Write-Host $sampleData -ForegroundColor DarkGray

        if (-not $NoOpen) {
            $open = Read-Host "`nDo you want to open the generated CSV file now? (Y/N)"
            if ($open -match '^(y|s|yes|si)$') {
                try { Invoke-Item $OutputPath } catch { }
            }
        }
        return $OutputPath
    } catch {
        Write-Log "Error creating template CSV: $_" "ERROR"
        return $null
    } finally {
        if (-not $NoOpen) { Read-Host -Prompt "`nPress Enter to return to menu..." }
    }
}

# Function to view activity log
function Show-ActivityLog {
    param (
        [string]$LogPath = $script:LogFile,
        [int]$Tail = 25
    )

    Write-Host "`n======================= [ 7. VIEW ACTIVITY LOG ] =======================" -ForegroundColor Yellow

    if (Test-Path -Path $LogPath) {
        Write-Host "Log location: $LogPath`n" -ForegroundColor DarkGray
        Get-Content -Path $LogPath -Tail $Tail | ForEach-Object {
            if ($_ -match "\[SUCCESS\]") { Write-Host $_ -ForegroundColor Green }
            elseif ($_ -match "\[ERROR\]") { Write-Host $_ -ForegroundColor Red }
            elseif ($_ -match "\[WARN\]") { Write-Host $_ -ForegroundColor Yellow }
            else { Write-Host $_ -ForegroundColor Cyan }
        }

        $open = Read-Host "`nOpen full log file in default editor? (Y/N)"
        if ($open -match '^(y|s|yes|si)$') {
            try { Invoke-Item $LogPath } catch { }
        }
    } else {
        Write-Log "No activity log file found yet." "INFO"
        Read-Host -Prompt "`nPress Enter to return to menu..."
    }
}

# Function to start the interactive management console
function Start-PrinterManagement {
    [CmdletBinding()]
    param()

    # Check for Administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "`n[!] ERROR: This script requires administrator privileges." -ForegroundColor Red
        Write-Host "Please right-click PowerShell and select 'Run as Administrator'.`n" -ForegroundColor Yellow
        Read-Host -Prompt "Press Enter to exit..."
        return
    }

    do {
        Show-Banner

        Write-Host " [1] Add Printers (Bulk CSV / TCP-IP / Shared UNC)" -ForegroundColor Cyan
        Write-Host " [2] Remove Printers (CSV or Interactive List Selection)" -ForegroundColor Cyan
        Write-Host " [3] Send Test Pages (Bulk CSV)" -ForegroundColor Cyan
        Write-Host " [4] Clear Print Queue (Purge Spooler)" -ForegroundColor Cyan
        Write-Host " [5] Inventory Printers (Console Table & CSV Export)" -ForegroundColor Cyan
        Write-Host " [6] Generate CSV Template" -ForegroundColor Cyan
        Write-Host " [7] View Activity Log" -ForegroundColor Cyan
        Write-Host " [8] Exit" -ForegroundColor Yellow
        Write-Host ""

        $option = Read-Host "Select an option (1-8)"

        switch ($option) {
            "1" { Add-Printers }
            "2" { Remove-Printers }
            "3" { Send-TestPages }
            "4" { Clear-PrintQueues }
            "5" { Inventory-Printers }
            "6" { New-PrinterTemplateCsv }
            "7" { Show-ActivityLog }
            "8" {
                Write-Host "`nExiting PrinterManagement. Goodbye!`n" -ForegroundColor Cyan
            }
            default {
                Write-Warning "Invalid option. Please choose a number between 1 and 8."
                Start-Sleep -Seconds 1
            }
        }
    } while ($option -ne "8")
}

# Auto-run menu when executed directly (not in test mode)
if (-not $env:PRINTER_MANAGEMENT_TEST_MODE) {
    Start-PrinterManagement
}

