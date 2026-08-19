#Requires -Module Pester

<#
.SYNOPSIS
    Unit and Integration Test Suite for PrinterManagement.ps1
.DESCRIPTION
    Uses Pester (v5+) to validate all functions, CSV parsing, logging,
    port creation logic, driver validation, WMI/CIM methods, and spooler purges.
#>

# Define mockable stub functions with full parameter signatures for PrintManagement module
if (-not (Get-Command -Name Get-Printer -ErrorAction SilentlyContinue)) {
    function global:Get-Printer { [CmdletBinding()] param([Parameter(Position=0)]$Name, $ErrorAction) }
}
if (-not (Get-Command -Name Add-Printer -ErrorAction SilentlyContinue)) {
    function global:Add-Printer { [CmdletBinding()] param($Name, $PortName, $DriverName, $ConnectionName, $ErrorAction) }
}
if (-not (Get-Command -Name Remove-Printer -ErrorAction SilentlyContinue)) {
    function global:Remove-Printer { [CmdletBinding()] param([Parameter(Position=0)]$Name, $ErrorAction) }
}
if (-not (Get-Command -Name Get-PrinterPort -ErrorAction SilentlyContinue)) {
    function global:Get-PrinterPort { [CmdletBinding()] param([Parameter(Position=0)]$Name, $ErrorAction) }
}
if (-not (Get-Command -Name Add-PrinterPort -ErrorAction SilentlyContinue)) {
    function global:Add-PrinterPort { [CmdletBinding()] param([Parameter(Position=0)]$Name, $PrinterHostAddress, $ErrorAction) }
}
if (-not (Get-Command -Name Remove-PrinterPort -ErrorAction SilentlyContinue)) {
    function global:Remove-PrinterPort { [CmdletBinding()] param([Parameter(Position=0)]$Name, $ErrorAction) }
}
if (-not (Get-Command -Name Get-PrinterDriver -ErrorAction SilentlyContinue)) {
    function global:Get-PrinterDriver { [CmdletBinding()] param([Parameter(Position=0)]$Name, $ErrorAction) }
}
if (-not (Get-Command -Name Out-GridView -ErrorAction SilentlyContinue)) {
    function global:Out-GridView { [CmdletBinding()] param([Parameter(ValueFromPipeline)]$InputObject, $Title, [switch]$PassThru) }
}

$env:PRINTER_MANAGEMENT_TEST_MODE = "true"

BeforeAll {
    try {
        $env:PRINTER_MANAGEMENT_TEST_MODE = "true"
        $script:ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\PrinterManagement.ps1"

        # Try to import Windows PrintManagement module if available
        try {
            Import-Module PrintManagement -ErrorAction SilentlyContinue
        } catch { }

        # Dot-source script functions into test scope
        try {
            . $script:ScriptPath
        } catch {
            Write-Host "ERROR DOT-SOURCING PrinterManagement.ps1: $_" -ForegroundColor Red
            throw
        }

        # Create temporary directory for test artifacts
        $script:TestTempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "PM_Tests_$([System.Guid]::NewGuid().ToString('N'))"
        New-Item -ItemType Directory -Path $script:TestTempDir -Force | Out-Null
    } catch {
        Write-Host "CRITICAL ERROR IN BeforeAll: $_" -ForegroundColor Red
        throw
    }
}

AfterAll {
    $env:PRINTER_MANAGEMENT_TEST_MODE = $null
    if ($null -ne $script:TestTempDir -and (Test-Path -Path $script:TestTempDir)) {
        Remove-Item -Path $script:TestTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe "1. System Logging (Write-Log)" {
    It "Formats log message with timestamp and level icon" {
        $testLog = Join-Path -Path $script:TestTempDir -ChildPath "test_log.log"
        Write-Log -Message "Test message" -Level "SUCCESS" -CustomLogPath $testLog

        Test-Path -Path $testLog | Should -BeTrue

        $fileContent = Get-Content -Path $testLog -Raw
        $fileContent | Should -Match "\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] \[SUCCESS\] Test message"
    }

    It "Supports all defined log levels (INFO, SUCCESS, WARN, ERROR)" {
        $testLog = Join-Path -Path $script:TestTempDir -ChildPath "test_levels.log"
        Write-Log -Message "Info msg" -Level "INFO" -CustomLogPath $testLog
        Write-Log -Message "Success msg" -Level "SUCCESS" -CustomLogPath $testLog
        Write-Log -Message "Warn msg" -Level "WARN" -CustomLogPath $testLog
        Write-Log -Message "Error msg" -Level "ERROR" -CustomLogPath $testLog

        $content = Get-Content -Path $testLog
        $content.Count | Should -Be 4
        $content[0] | Should -Match "\[INFO\]"
        $content[1] | Should -Match "\[SUCCESS\]"
        $content[2] | Should -Match "\[WARN\]"
        $content[3] | Should -Match "\[ERROR\]"
    }
}

Describe "2. Smart CSV Parser (Import-SmartCsv)" {
    It "Parses CSV with semicolon (;) delimiter and standard headers" {
        $csvPath = Join-Path -Path $script:TestTempDir -ChildPath "printers_semicolon.csv"
        @"
Name;LocalPort;DriverName
Office_HP;192.168.1.50;HP Universal Printing PCL 6
Finance_Canon;192.168.1.60;Canon Generic Plus PCL6
"@ | Set-Content -Path $csvPath -Encoding UTF8

        $result = Import-SmartCsv -Path $csvPath
        $result | Should -Not -BeNullOrEmpty
        $result.Count | Should -Be 2
        $result[0].Name | Should -Be "Office_HP"
        $result[0].LocalPort | Should -Be "192.168.1.50"
        $result[0].DriverName | Should -Be "HP Universal Printing PCL 6"
    }

    It "Parses CSV with comma (,) delimiter and alternate headers (PrinterName, Port, Driver)" {
        $csvPath = Join-Path -Path $script:TestTempDir -ChildPath "printers_comma.csv"
        @"
PrinterName,Port,Driver
Zebra_Label,USB001,ZDesigner ZD420-203dpi ZPL
Color_Laser,10.0.0.15,Epson Universal
"@ | Set-Content -Path $csvPath -Encoding UTF8

        $result = Import-SmartCsv -Path $csvPath
        $result | Should -Not -BeNullOrEmpty
        $result.Count | Should -Be 2
        $result[0].Name | Should -Be "Zebra_Label"
        $result[0].LocalPort | Should -Be "USB001"
        $result[0].DriverName | Should -Be "ZDesigner ZD420-203dpi ZPL"
    }

    It "Parses CSV with Tab (`t) delimiter" {
        $csvPath = Join-Path -Path $script:TestTempDir -ChildPath "printers_tab.csv"
        "Name`tLocalPort`tDriverName`nPrinter1`t192.168.1.200`tGeneric Driver" | Set-Content -Path $csvPath -Encoding UTF8

        $result = Import-SmartCsv -Path $csvPath
        $result | Should -Not -BeNullOrEmpty
        $result[0].Name | Should -Be "Printer1"
        $result[0].LocalPort | Should -Be "192.168.1.200"
    }

    It "Returns $null when file does not exist" {
        $nonExistent = Join-Path -Path $script:TestTempDir -ChildPath "missing.csv"
        $result = Import-SmartCsv -Path $nonExistent
        ($null -eq $result) | Should -BeTrue
    }
}

Describe "3. File Path Resolution (Get-ValidFilePath)" {
    It "Resolves valid existing file path directly" {
        $testFile = Join-Path -Path $script:TestTempDir -ChildPath "existing.csv"
        Set-Content -Path $testFile -Value "Name;LocalPort;DriverName"

        $resolved = Get-ValidFilePath -FilePath $testFile
        $resolved | Should -Be (Resolve-Path $testFile).Path
    }

    It "Returns $null when direct file path does not exist" {
        $missingFile = Join-Path -Path $script:TestTempDir -ChildPath "does_not_exist.csv"
        $resolved = Get-ValidFilePath -FilePath $missingFile
        ($null -eq $resolved) | Should -BeTrue
    }
}

Describe "4. Template Generator (New-PrinterTemplateCsv)" {
    It "Generates a valid CSV template file with sample records" {
        $templatePath = Join-Path -Path $script:TestTempDir -ChildPath "sample_template.csv"
        $result = New-PrinterTemplateCsv -OutputPath $templatePath -NoOpen

        $result | Should -Be $templatePath
        Test-Path -Path $templatePath | Should -BeTrue

        $imported = Import-SmartCsv -Path $templatePath
        $imported | Should -Not -BeNullOrEmpty
        $imported.Count | Should -BeGreaterThan 2

        # Verify sample entries
        $imported[0].Name | Should -Be "Office_HP_LaserJet"
        $imported[1].LocalPort | Should -Match "^\\\\"
    }
}

Describe "5. Printer Installation (Add-Printers)" {
    Context "When adding a shared network printer (UNC)" {
        It "Calls Add-Printer with -ConnectionName" {
            $csvPath = Join-Path -Path $script:TestTempDir -ChildPath "unc_printers.csv"
            @"
Name;LocalPort;DriverName
Shared_HP;\\printserver01\Shared_HP;HP Universal Driver
"@ | Set-Content -Path $csvPath -Encoding UTF8

            Mock Add-Printer { return } -Verifiable -ParameterFilter { $ConnectionName -eq "\\printserver01\Shared_HP" }

            $result = Add-Printers -FilePath $csvPath

            $result.Total | Should -Be 1
            $result.Success | Should -Be 1
            $result.Failed | Should -Be 0
            Assert-VerifiableMock
        }
    }

    Context "When adding a standard TCP/IP printer" {
        It "Creates port and adds printer if driver exists" {
            $csvPath = Join-Path -Path $script:TestTempDir -ChildPath "tcp_printers.csv"
            @"
Name;LocalPort;DriverName
Office_Epson;192.168.1.100;Epson Driver
"@ | Set-Content -Path $csvPath -Encoding UTF8

            Mock Get-PrinterPort { return $null }
            Mock Add-PrinterPort { return } -Verifiable -ParameterFilter { $Name -eq "192.168.1.100" -and $PrinterHostAddress -eq "192.168.1.100" }
            Mock Get-PrinterDriver { return [PSCustomObject]@{ Name = "Epson Driver" } }
            Mock Add-Printer { return } -Verifiable -ParameterFilter { $Name -eq "Office_Epson" -and $PortName -eq "192.168.1.100" }

            $result = Add-Printers -FilePath $csvPath

            $result.Total | Should -Be 1
            $result.Success | Should -Be 1
            $result.Failed | Should -Be 0
            Assert-VerifiableMock
        }

        It "Skips installation and warns when driver is not installed" {
            $csvPath = Join-Path -Path $script:TestTempDir -ChildPath "missing_driver.csv"
            @"
Name;LocalPort;DriverName
Office_Canon;192.168.1.110;Missing Driver XYZ
"@ | Set-Content -Path $csvPath -Encoding UTF8

            Mock Get-PrinterPort { return [PSCustomObject]@{ Name = "192.168.1.110" } }
            Mock Get-PrinterDriver { return $null }
            Mock Add-Printer { throw "Should not be called" }

            $result = Add-Printers -FilePath $csvPath

            $result.Total | Should -Be 1
            $result.Success | Should -Be 0
            $result.Failed | Should -Be 1
        }
    }
}

Describe "6. Printer Removal (Remove-Printers)" {
    It "Removes existing printer and port" {
        $printerList = @(
            [PSCustomObject]@{ Name = "Printer_ToDelete"; LocalPort = "192.168.1.99" }
        )

        Mock Get-Printer { return [PSCustomObject]@{ Name = "Printer_ToDelete" } }
        Mock Remove-Printer { return } -Verifiable -ParameterFilter { $Name -eq "Printer_ToDelete" }
        Mock Get-PrinterPort { return [PSCustomObject]@{ Name = "192.168.1.99" } }
        Mock Remove-PrinterPort { return } -Verifiable -ParameterFilter { $Name -eq "192.168.1.99" }

        $result = Remove-Printers -PrinterList $printerList -Force

        $result.Total | Should -Be 1
        $result.Success | Should -Be 1
        $result.Failed | Should -Be 0
        Assert-VerifiableMock
    }
}

Describe "7. Print Queue Clearing (Clear-PrintQueues)" {
    It "Stops Spooler service, deletes spool files, and restarts service" {
        $spoolerMock = [PSCustomObject]@{
            Status = 'Stopped'
            Refresh = { }
        }

        Mock Stop-Service { return } -Verifiable -ParameterFilter { $Name -eq "Spooler" -and $Force -eq $true }
        Mock Get-Service { return $spoolerMock }
        Mock Remove-Item { return } -Verifiable
        Mock Start-Service { return } -Verifiable -ParameterFilter { $Name -eq "Spooler" }

        $result = Clear-PrintQueues -Force
        $result | Should -BeTrue
        Assert-VerifiableMock
    }
}

Describe "8. Inventory Generation (Inventory-Printers)" {
    It "Queries CIM instances and exports to CSV" {
        $outputPath = Join-Path -Path $script:TestTempDir -ChildPath "test_inventory.csv"

        $cimPrinters = @(
            [PSCustomObject]@{
                Name = "Printer1"; DriverName = "Driver1"; PortName = "192.168.1.1"
                ShareName = ""; Published = $false; PrinterStatus = 3; Default = $true
                Location = "Floor 1"; Comment = ""
            },
            [PSCustomObject]@{
                Name = "Printer2"; DriverName = "Driver2"; PortName = "192.168.1.2"
                ShareName = "P2"; Published = $true; PrinterStatus = 3; Default = $false
                Location = "Floor 2"; Comment = ""
            }
        )

        Mock Get-CimInstance { return $cimPrinters } -ParameterFilter { $ClassName -eq "Win32_Printer" }

        $result = Inventory-Printers -OutputPath $outputPath -NoGrid

        $result | Should -Not -BeNullOrEmpty
        $result.Count | Should -Be 2
        Test-Path -Path $outputPath | Should -BeTrue

        $exported = Import-Csv -Path $outputPath -Delimiter ';'
        $exported.Count | Should -Be 2
        $exported[0].Name | Should -Be "Printer1"
    }
}
