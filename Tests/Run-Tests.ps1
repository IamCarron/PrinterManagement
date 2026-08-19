<#
.SYNOPSIS
    Automated Test Runner for PrinterManagement.
.DESCRIPTION
    Runs the Pester test suite (or installs Pester if missing) and reports test results.
#>

[CmdletBinding()]
param (
    [switch]$CodeCoverage,
    [switch]$CI
)

$testsPath = Join-Path -Path $PSScriptRoot -ChildPath "PrinterManagement.Tests.ps1"
$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\PrinterManagement.ps1"

Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "       PrinterManagement - Automated Test Suite Runner     " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Check if Pester is installed
$pesterModule = Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1

if (-not $pesterModule -or $pesterModule.Version.Major -lt 5) {
    Write-Host "[!] Pester 5+ not found. Attempting to install Pester from PSGallery..." -ForegroundColor Yellow
    try {
        Install-Module -Name Pester -MinimumVersion 5.3.0 -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
        Import-Module Pester -MinimumVersion 5.3.0
        Write-Host "[✓] Pester installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "[✗] Could not install Pester automatically: $_" -ForegroundColor Red
        Write-Host "Please run: Install-Module -Name Pester -Scope CurrentUser" -ForegroundColor Yellow
        exit 1
    }
} else {
    Import-Module $pesterModule.Name -MinimumVersion 5.0.0
}

# Configure Pester configuration
$pesterConfig = New-PesterConfiguration
$pesterConfig.Run.Path = $testsPath
$pesterConfig.Output.Verbosity = "Detailed"
$pesterConfig.Run.PassThru = $true

if ($CodeCoverage) {
    $pesterConfig.CodeCoverage.Enabled = $true
    $pesterConfig.CodeCoverage.Path = $scriptPath
}

if ($CI) {
    $pesterConfig.TestResult.Enabled = $true
    $pesterConfig.TestResult.OutputPath = ".\TestResults.xml"
    $pesterConfig.TestResult.OutputFormat = "NUnitXml"
}

# Run tests
$result = Invoke-Pester -Configuration $pesterConfig

Write-Host "`n═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                   Test Execution Summary                  " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Total Tests: $($result.TotalCount)"
Write-Host "Passed:      $($result.PassedCount)" -ForegroundColor Green
Write-Host "Failed:      $($result.FailedCount)" -ForegroundColor $(if ($result.FailedCount -gt 0) { "Red" } else { "Green" })
Write-Host "Skipped:     $($result.SkippedCount)" -ForegroundColor Yellow

if ($result.FailedCount -gt 0) {
    exit 1
} else {
    exit 0
}
