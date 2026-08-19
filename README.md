# 🖨️ PrinterManagement

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg?logo=powershell)](https://microsoft.com/powershell)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011%20%2F%20Server-0078D6.svg?logo=windows)](https://www.microsoft.com/windows)
[![Tests](https://img.shields.io/badge/Tests-Pester%20v5-28A745.svg?logo=pester)](https://pester.dev)
[![License](https://img.shields.io/badge/License-Apache%202.0-orange.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-3.1.0-brightgreen.svg)]()

A robust, enterprise-grade PowerShell automation suite for managing printers in Windows environments. Designed for System Administrators, IT Support, and DevOps, it streamlines bulk deployments, maintenance, diagnostics, and inventorying with an interactive TUI, native GUI file dialogs, and comprehensive logging.

---

## 📑 Table of Contents
- [✨ Key Features](#-key-features)
- [🖥️ Interactive Menu](#️-interactive-menu)
- [🚀 Quick Start](#-quick-start)
- [📊 CSV File Specifications](#-csv-file-specifications)
- [🧪 Automated Unit Tests](#-automated-unit-tests)
- [⚙️ Requirements](#️-requirements)
- [🔒 Safety & Best Practices](#-safety--best-practices)
- [📄 License](#-license)

---

## ✨ Key Features

- **📥 Multi-Protocol Bulk Installation:**
  - Standard TCP/IP network printers (IPv4 and Hostnames / FQDN).
  - Shared network printers via UNC paths (`\\printserver\printer_share`) using native Windows connection mapping.
  - Local ports (`USB001`, `LPT1:`, `COM1:`).
- **🗑️ Dual-Mode Printer Removal:**
  - Bulk removal from CSV lists.
  - Visual interactive removal using Windows `Out-GridView` (select multiple printers with `Ctrl`/`Shift`).
- **📄 Automated Test Page Dispatch:**
  - Dispatches test pages using CIM/WMI methods (`Invoke-CimMethod -MethodName PrintTestPage`) with fallback to `printui.dll`.
  - Safely handles special characters and quotes in printer names.
- **🧹 Non-Destructive Spooler Queue Purge:**
  - Safely halts the `Spooler` service (`-Force`).
  - Active polling ensures the service is stopped and file locks are released before deleting `.SHD` / `.SPL` files.
  - Restarts the service cleanly and reports status.
- **📊 Real-Time Inventory & GridView Inspection:**
  - Queries `Win32_Printer` via CIM for instant inventorying.
  - Exports to UTF-8 delimited CSV (`inventory.csv`).
  - Displays results in a native, filterable, and searchable graphical table.
- **🗂️ Hybrid CLI / GUI File Picker:**
  - Press `[Enter]` to open the standard Windows Explorer file selection dialog, or paste file paths directly.
- **🛡️ Smart CSV Parser:**
  - Auto-detects delimiters (Semicolon `;`, Comma `,`, or Tab `\t`).
  - Automatically maps column variations (`DriverName`/`Driver`, `LocalPort`/`Port`, `Name`/`PrinterName`).
- **📈 Native Progress Reporting:**
  - Real-time progress bars (`Write-Progress`) show current operation percentage and printer names during bulk operations.
- **📜 Centralized Activity Logging:**
  - All operations, successes, warnings, and errors are saved with timestamps to `PrinterManagement.log`.

---

## 🖥️ Interactive Menu

```text
 ╔════════════════════════════════════════════════════════════════════════════╗
 ║   ___      _      __          __  ___                                      ║
 ║  / _ \____(_)__  / /____ ____/  |/  /__ ____  ___ ____ ____ __ _  ___ ___  ║
 ║ / ___/ __/ / _ \/ __/ -_) __/ /|_/ / _ `/ _ \/ _ `/ _ `/ -_)  ' \/ -_) _ \ ║
 ║/_/  /_/ /_/_//_/\__/\__/_/ /_/  /_/\_,_/_//_/\_,_/\_, /\__/_/_/_/\__/_//_/ ║
 ║                                                  /___/                     ║
 ║                      Windows Printer Management Suite                      ║
 ║                                Version 3.1.0                               ║
 ╚════════════════════════════════════════════════════════════════════════════╝

 1. 📥 Add Printers (Bulk CSV / TCP-IP / Shared UNC)
 2. 🗑️  Remove Printers (CSV or Interactive GUI Selection)
 3. 📄 Send Test Pages (Bulk CSV)
 4. 🧹 Clear Print Queue (Purge Spooler)
 5. 📊 Inventory Printers (CSV Export & GUI GridView)
 6. 📝 Generate CSV Template
 7. 📜 View Activity Log
 8. 🚪 Exit
```

---

## 🚀 Quick Start

### 1. Clone or Download the Repository
```powershell
git clone https://github.com/IamCarron/PrinterManagement.git
cd PrinterManagement
```

### 2. Run with Administrator Privileges
Right-click PowerShell and select **Run as Administrator**, then execute:
```powershell
.\PrinterManagement.ps1
```

### 3. Generate a Ready-to-Use CSV Template
Select option `6` from the menu to create `template_printers.csv`, or create your own custom CSV file.

---

## 📊 CSV File Specifications

The parser automatically detects delimiters (`;`, `,`, or `\t`). Semicolon is recommended for international Excel compatibility:

| Column | Description | Example Values |
| :--- | :--- | :--- |
| **`Name`** | Display name for the printer in Windows. | `Office_HP_LaserJet`, `Finance_Canon` |
| **`LocalPort`** | Port identifier (IPv4, Hostname, UNC path, or USB). | `192.168.1.50`, `\\server\share`, `USB001` |
| **`DriverName`** | Exact name of the pre-installed print driver. | `HP Universal Printing PCL 6`, `Canon Generic Plus PCL6` |

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

The project includes a full unit and integration test suite written with **Pester 5+** under `Tests/`.

### Running Tests Locally:
```powershell
# Run the automated test runner (installs Pester if missing):
.\Tests\Run-Tests.ps1

# Or run directly with Pester:
Invoke-Pester .\Tests\PrinterManagement.Tests.ps1 -Output Detailed
```

### Continuous Integration (CI):
Automated testing is configured via **GitHub Actions** (`.github/workflows/test.yml`), verifying execution on both Windows PowerShell 5.1 and PowerShell 7+ on every commit.

---

## ⚙️ Requirements

- **Operating System:** Windows 10 / 11, Windows Server 2016 / 2019 / 2022 / 2025.
- **PowerShell:** PowerShell 5.1 (Built-in) or PowerShell 7.x (Core).
- **Permissions:** Elevated Administrator privileges (`Run as Administrator`).
- **Drivers:** Manufacturer printer drivers must be installed on the host before running bulk creation.

---

## 🔒 Safety & Best Practices

- **Queue Cleaning Safety:** Option `4` requests explicit confirmation before stopping the Spooler service to prevent unintentional loss of active print jobs.
- **CIM Modernization:** Uses `Get-CimInstance` instead of deprecated `Get-WmiObject`, ensuring full compatibility with modern PowerShell Core and Windows Server Core.
- **Apostrophe & WQL Injection Protection:** Single quotes in printer names (e.g. `User's Printer`) are escaped automatically to prevent CIM/WMI filter exceptions.

---

## 📄 License

This project is licensed under the **Apache License 2.0**. See the [LICENSE](LICENSE) file for details.



