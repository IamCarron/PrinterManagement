# Printer Management Suite 🖨️

[![PowerShell Tests & Quality Check](https://github.com/IamCarron/PrinterManagement/actions/workflows/test.yml/badge.svg)](https://github.com/IamCarron/PrinterManagement/actions/workflows/test.yml)
[![PowerShell](https://img.shields.io/badge/Language-PowerShell-5391FE.svg?logo=powershell&logoColor=white)](https://microsoft.com/powershell)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011%20%2F%20Server-0078D6.svg?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Tests](https://img.shields.io/badge/Tests-Pester%20v5-28A745.svg?logo=pester&logoColor=white)](https://pester.dev)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/IamCarron/PrinterManagement/pulls)

Automated, robust, and enterprise-grade PowerShell automation suite for managing printers in Windows environments. Designed for System Administrators, IT Support, and DevOps to streamline bulk deployments, maintenance, diagnostics, and inventorying.

---

## 📑 Table of Contents

- [🌟 Features](#-features)
- [🖥️ Interactive Console](#️-interactive-console)
- [📋 Requirements](#-requirements)
- [🚀 Quick Start](#-quick-start)
- [📊 CSV File Specifications](#-csv-file-specifications)
- [🧪 Automated Unit Tests](#-automated-unit-tests)
- [🛡️ Safety & Parachute Guards](#️-safety--parachute-guards)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)

---

## 🌟 Features

- ⚡ **Multi-Protocol Bulk Installation:** Deploy dozens of printers automatically from CSV files via standard TCP/IP (IPv4 / Hostname), Shared Network Printers (`\\server\share`), or Local Ports (`USB001`, `LPT1:`).
- 🗑️ **Dual-Mode Printer Removal:** Remove printers in bulk via CSV lists or interactively using a native Windows `Out-GridView` selection dialog with multi-select support (`Ctrl` / `Shift`).
- 📄 **Automated Test Page Dispatch:** Send hardware test prints via CIM/WMI methods (`Win32_Printer.PrintTestPage()`) with automatic fallback to native Windows `printui.dll`.
- 🧹 **Safe Spooler Queue Purge:** Stops the `Spooler` service cleanly, verifies process termination and file lock releases, clears corrupted print jobs (`.SHD` / `.SPL`), and restarts the Spooler service.
- 📊 **Real-Time Inventory & Console Table Inspection:** Collects system-wide printer data using modern CIM cmdlets, displays an organized, clean table directly in the terminal, and exports to UTF-8 CSV (`inventory.csv`).
- 🗂️ **Hybrid GUI / CLI File Picker:** Open native Windows Explorer dialogs by pressing `[Enter]` or type/paste file paths directly in the console.
- 🛡️ **Smart Delimiter & Header Parser:** Auto-detects delimiters (Semicolon `;`, Comma `,`, or Tab `\t`) and normalizes varied column names (`Driver`/`DriverName`, `Port`/`LocalPort`, `Name`/`PrinterName`).
- 📈 **Native Progress Bars:** Integrated `Write-Progress` tracking provides visual percentage and status updates during batch installations and test page dispatching.
- 📜 **Centralized Activity Logging:** Every operation, warning, success, and error is recorded with timestamps in `PrinterManagement.log`.
- 🌐 **Cross-PowerShell Compatibility:** 100% pure ASCII user interface and robust encoding protection, compatible with Windows PowerShell 5.1 and modern PowerShell 7+ (pwsh).

---

## 🖥️ Interactive Console

```text
  ===========================================================================
    ___      _      __          __  ___
   / _ \____(_)__  / /____ ____/  |/  /__ ____  ___ ____ ____ __ _  ___ ___
  / ___/ __/ / _ \/ __/ -_) __/ /|_/ / _ `/ _ \/ _ `/ _ `/ -_)  ' \/ -_) _ \
 /_/  /_/ /_/_//_/\__/\__/_/ /_/  /_/\_,_/_//_/\_,_/\_, /\__/_/_/_/\__/_//_/
                                                    /___/
  ===========================================================================
   v3.2.0 | Windows Printer Management Suite
   github.com/IamCarron/PrinterManagement
  ===========================================================================

   >> OPERATIONS -------------------------------------------------
      [1]  Add Printers         Bulk CSV / TCP-IP / Shared UNC
      [2]  Remove Printers      CSV or interactive selection

   >> DIAGNOSTICS ------------------------------------------------
      [3]  Send Test Pages      CIM dispatch with printui fallback
      [4]  Clear Print Queue    Purge Spooler & restart service

   >> SYSTEM -----------------------------------------------------
      [5]  Inventory Printers   GUI table & CSV export
      [6]  View Activity Log    Review recent operations

   ---------------------------------------------------------------
      [0]  Exit

                         Made with <3 by IamCarron
```

---

## 📋 Requirements

- **Operating System:** Windows 10 / 11 or Windows Server 2016 / 2019 / 2022 / 2025.
- **PowerShell Version:** Windows PowerShell 5.1 (Built-in) or PowerShell 7.x (Core).
- **Execution Privileges:** Elevated Administrator rights (`Run as Administrator`).
- **Print Drivers:** Required manufacturer printer drivers must be installed prior to bulk deployment.

---

## 🚀 Quick Start

### 1. Clone the Repository
```powershell
git clone https://github.com/IamCarron/PrinterManagement.git
cd PrinterManagement
```

### 2. Launch with Administrator Privileges
Open PowerShell as **Administrator** and run:
```powershell
.\PrinterManagement.ps1
```

### 3. Prepare Your Printer CSV
Create a CSV file with your printers (see the [CSV File Specifications](#-csv-file-specifications) section below), or copy the sample template directly from this README.

---

## 📊 CSV File Specifications

The parser automatically detects delimiters (`;`, `,`, or `\t`). Semicolon (`;`) is recommended for international Excel compatibility:

| Column | Description | Example Values |
| :--- | :--- | :--- |
| **`Name`** | Display name for the printer in Windows. | `Office_HP_LaserJet`, `Finance_Canon` |
| **`LocalPort`** | Port identifier (IPv4 address, Hostname, UNC path, or USB). | `192.168.1.50`, `\\printserver\share`, `USB001` |
| **`DriverName`** | Exact name of the pre-installed print driver on the host. | `HP Universal Printing PCL 6`, `Canon Generic Plus PCL6` |

### Sample `template_printers.csv`:
```csv
Name;LocalPort;DriverName
Office_HP_LaserJet;192.168.1.50;HP Universal Printing PCL 6
Finance_Canon_MFP;\\printserver01\Canon_Finance;Canon Generic Plus PCL6
Warehouse_Zebra_Labels;USB001;ZDesigner ZD420-203dpi ZPL
HR_Epson_WorkForce;192.168.1.55;EPSON WF-C5790 Series
```

---

## 🧪 Automated Unit Tests

The repository includes a comprehensive unit and integration test suite built with **Pester 5+** located in `Tests/`.

### Running Tests Locally:
```powershell
# Automated test runner (installs Pester automatically if not present):
.\Tests\Run-Tests.ps1

# Or run directly via Pester:
Invoke-Pester .\Tests\PrinterManagement.Tests.ps1 -Output Detailed
```

### Continuous Integration (CI):
Automated testing is configured via **GitHub Actions** (`.github/workflows/test.yml`), verifying code quality and test execution across both **Windows PowerShell 5.1** and **PowerShell 7+ (pwsh)** on every push and pull request.

---

## 🛡️ Safety & Parachute Guards

- 🪂 **Interactive Deletion Safeguards:** Removal operations require explicit confirmation (`Y/N`) before modifying the system unless `-Force` is supplied programmatically.
- 🔍 **Pre-Execution Driver Validation:** Verifies that the required printer driver exists locally before attempting printer creation, avoiding system errors or corrupt configurations.
- 🧹 **Controlled Spooler Restart:** Stops the print spooler safely and ensures all pending file handles are released before purging `.SPL` / `.SHD` spool files.
- 💉 **CIM / WQL Escape Protection:** Automatically escapes single quotes and special characters in printer names to prevent WMI/CIM injection and query errors.
- 📋 **Zero Pipeline Pollution:** All helper routines are isolated to guarantee pure data pipelines across scripts and test runners.

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!

1. Fork the Project.
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`).
3. Commit your Changes (`git commit -m 'feat: add some amazing feature'`).
4. Push to the Branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request.

---

## 📄 License

Distributed under the **Apache License 2.0**. See the [LICENSE](LICENSE) file for more information.



