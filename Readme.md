# User Audit Report Generator

<p align="center">
    <strong>Comprehensive Active Directory and Entra ID user audit reporting tool</strong>
</p>

<p align="center">
    <a href="#quick-start">Quick Start</a> • 
    <a href="#features">Features</a> • 
    <a href="#installation">Installation</a> • 
    <a href="#usage">Usage</a> • 
    <a href="#report-analysis">Report Analysis</a>
</p>

---

## 📋 Table of Contents

- [Overview](#overview)
- [✨ Features](#features)
- [📋 Prerequisites](#prerequisites)
- [🚀 Quick Start](#quick-start)
- [📦 Installation](#installation)
- [⚙️ Configuration](#configuration)
- [🔧 Usage](#usage)
- [📊 Report Analysis](#report-analysis)
- [🛠️ Troubleshooting](#troubleshooting)
- [🔒 Security Considerations](#security-considerations)
- [🤝 Contributing](#contributing)

---

## Overview

The **User Audit Report Generator** is a comprehensive PowerShell tool designed for IT administrators to perform thorough audits of user accounts across both **Active Directory (AD)** and **Microsoft Entra ID (Azure AD)** environments. This script generates detailed reports with security recommendations, helping organizations maintain proper user account hygiene and security compliance.

### Key Capabilities

- **🏢 On-Premises Integration**: Full Active Directory user account analysis
- **☁️ Cloud Integration**: Microsoft Entra ID user account auditing
- **🔄 Hybrid Analysis**: Seamless integration of on-premises and cloud user data
- **🛡️ Security Assessment**: Automated security recommendations and risk identification
- **📊 Professional Reporting**: Excel output with conditional formatting and visual indicators
- **🔍 Service Account Detection**: Automatic identification of service accounts and MSAs

---

## ✨ Features

### 🏢 Active Directory Analysis
- **Complete User Enumeration**: Audits all AD user accounts with comprehensive properties
- **Admin Rights Detection**: Identifies Enterprise Admins and Domain Admins
- **Service Account Recognition**: Detects MSAs, gMSAs, and known service accounts
- **Password Policy Compliance**: Analyzes password age, expiration, and policy settings
- **Account Status Monitoring**: Tracks enabled/disabled, locked, and expired accounts

### ☁️ Entra ID Integration
- **Cloud User Analysis**: Audits Entra ID (Azure AD) user accounts
- **Global Admin Detection**: Identifies Global Administrator role assignments
- **Premium License Support**: Utilizes SignInActivity data when available
- **Hybrid User Reconciliation**: Merges on-premises and cloud user data
- **Last Logon Correlation**: Combines AD and Entra ID logon timestamps

### 📊 Advanced Reporting
- **Excel Output**: Professional XLSX reports with automated formatting
- **Conditional Formatting**: Visual indicators for security risks and compliance issues
- **CSV Fallback**: Alternative output format when Excel module unavailable
- **Customizable Layout**: Organized columns with proper data typing and formatting
- **Interactive Features**: Freeze panes, auto-filtering, and sortable columns

### 🛡️ Security Intelligence
- **Risk Assessment**: Automated identification of security risks and compliance issues
- **Actionable Recommendations**: Specific guidance for account remediation
- **Stale Account Detection**: Identifies accounts not used in 90+ days
- **Password Analysis**: Flags accounts with old passwords or never-expiring passwords
- **Admin Account Monitoring**: Special attention to privileged account security

---

## 📋 Prerequisites

### System Requirements

| Component | Requirement | Notes |
|-----------|-------------|--------|
| **PowerShell** | 5.1 or later | PowerShell 5.1 and 7.x supported |
| **Operating System** | Windows Server 2012 R2+ | Domain-joined system preferred |
| **Permissions** | Domain User minimum | Domain Admin for complete analysis |
| **Memory** | 4GB minimum | 8GB+ recommended for large environments |
| **Disk Space** | 100MB minimum | Additional space for report output |

### Required PowerShell Modules

#### ✅ Essential Modules
```powershell
# Active Directory Module (Required)
Import-Module ActiveDirectory
```

#### 🔧 Optional Modules (Auto-Install Available)
```powershell
# Excel Report Generation
Install-Module ImportExcel -Force

# Microsoft Graph API Modules
Install-Module Microsoft.Graph.Authentication -Force
Install-Module Microsoft.Graph.Users -Force  
Install-Module Microsoft.Graph.DirectoryObjects -Force
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Force
```

### Account Permissions

#### Active Directory Requirements
- **Minimum**: Domain User with read access to AD
- **Recommended**: Domain Admin for complete user enumeration
- **Service Accounts**: Read access to service account OUs

#### Entra ID Requirements  
- **Authentication**: Interactive browser-based authentication
- **Required Permissions**:
  - `Directory.Read.All` - Read directory data
  - `User.Read.All` - Read all user profiles
  - `AuditLog.Read.All` - Read audit log data
- **Role Requirements**: Global Reader or Global Administrator

---

## 🚀 Quick Start

### Basic Execution (AD Only)
```powershell
# Run the script with default settings
.\UserAuditReport.ps1
```

### Complete Analysis (AD + Entra ID)
1. **Launch PowerShell as Administrator** (for module installation)
2. **Execute the script**:
   ```powershell
   .\UserAuditReport.ps1
   ```
3. **Follow Interactive Prompts**:
   - Install ImportExcel module? **Y** (recommended)
   - Connect to Entra ID? **Y** (for hybrid analysis)
   - Authenticate to Microsoft Graph when prompted

### Expected Output
```
C:\Temp\yourdomain.com_Users_Report_12302024_1430.xlsx
```

---

## 📦 Installation

### Method 1: Direct Download
1. Download `UserAuditReport.ps1` from the repository
2. Save to your preferred directory
3. Ensure execution policy allows script execution:
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   .\UserAuditReport.ps1
   ```

### Method 2: Clone Repository
```bash
git clone https://github.com/yourusername/user-audit-report.git
cd user-audit-report
```

---

## ⚙️ Configuration

### Service Account Database

The script includes a comprehensive database of known service accounts. You can customize this by editing the `$KnownServiceAccounts` hashtable:

```powershell
$KnownServiceAccounts = @{
    "svc-backup" = "Backup Service Account"
    "svc-monitoring" = "Monitoring Service Account" 
    "svc-custom" = "Custom Application Service Account"
    # Add your organization's service accounts here
}
```

### Supported Service Account Patterns

The script automatically detects service accounts using these patterns:
- **MSA/gMSA**: Managed Service Accounts from AD
- **Prefix Patterns**: `svc-*`, `*svc*`
- **Microsoft Patterns**: `MSOL_*`, `AAD_*` (Entra Connect)
- **Application Specific**: Sophos, Veeam, SQL, etc.

### Report Customization

#### Output Location
```powershell
# Default: C:\Temp\
# Modify in script or ensure C:\Temp\ exists
```

#### Excel Formatting Options
- **Conditional Formatting**: Automatically applied based on risk levels
- **Column Sizing**: Auto-sized for optimal viewing
- **Date Formatting**: Standardized MM/dd/yyyy hh:mm AM/PM format
- **Color Coding**: 
  - 🔴 **Red**: Critical issues (180+ days inactive, never-expiring passwords)
  - 🟡 **Yellow**: Warnings (90+ days inactive, disabled accounts)
  - 🟢 **Green**: Good status (recent activity, proper configuration)

---

## 🔧 Usage

### Interactive Execution

When you run the script, it will prompt for configuration options:

#### 1. ImportExcel Module Installation
```
WARNING: ImportExcel module is not installed. Without it the report will output in CSV and you will have to format it manually.
If authorized to install modules on this system, would you like to install it for this script? (Y/N)
```
- **Y**: Installs ImportExcel module and generates formatted XLSX report
- **N**: Falls back to CSV output format

#### 2. Entra ID Connection
```
Would you like to connect to Entra ID? (Y/N)
```
- **Y**: Enables hybrid analysis with cloud user data
- **N**: AD-only analysis

#### 3. Graph Module Installation (if needed)
```
WARNING: Graph API modules required for this report are not installed. The report will display on-premises AD Users only.
If authorized to install modules on this system, would you like to install the required Graph API modules for this script? (Y/N)
```

### Command Line Examples

```powershell
# Basic AD-only report
.\UserAuditReport.ps1

# Run from specific location
PowerShell.exe -ExecutionPolicy Bypass -File "C:\Scripts\UserAuditReport.ps1"

# Scheduled task execution
PowerShell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "UserAuditReport.ps1"
```

---

## 📊 Report Analysis

### Report Structure

The generated report contains the following columns:

#### User Identification
| Column | Description | Data Type |
|--------|-------------|-----------|
| **Name** | Display name of the user | Text |
| **SamAccountName** | AD logon name | Text |
| **On-Prem UserPrincipalName** | AD UPN | Text |
| **Cloud UserPrincipalName** | Entra ID UPN | Text |
| **Email Address** | Primary email address | Text |

#### Account Classification
| Column | Description | Values |
|--------|-------------|--------|
| **User Type** | Account classification | On-Prem, Cloud, Hybrid |
| **Known Service Account** | Service account detection | True/False |
| **EnterpriseAdmin** | Enterprise Admin membership | True/False |
| **DomainAdmin** | Domain Admin membership | True/False |
| **AzGlobalAdmin** | Global Admin membership | True/False |

#### Account Status
| Column | Description | Data Type |
|--------|-------------|-----------|
| **Enabled** | Account enabled status | True/False |
| **AccountExpiredDate** | Account expiration date | Date/Time |
| **Account Locked** | Lockout status | True/False |
| **PasswordExpired** | Password expiration status | True/False |

#### Security Analysis
| Column | Description | Data Type |
|--------|-------------|-----------|
| **PasswordLastSet** | Last password change | Date/Time |
| **LastLogonDate** | Most recent logon | Date/Time |
| **PasswordNeverExpires** | Password policy exception | True/False |
| **CannotChangePassword** | Password change restriction | True/False |

#### Audit Information
| Column | Description | Purpose |
|--------|-------------|---------|
| **Date Created** | Account creation date | Age analysis |
| **Recommended Actions** | Security recommendations | Remediation guidance |
| **Notes** | Manual notes field | Custom annotations |
| **Action** | Planned actions | Change tracking |
| **Follow Up** | Follow-up requirements | Task management |
| **Resolution** | Resolution status | Completion tracking |

### Visual Indicators

#### Color Coding System
- **🔴 Red Background**: 
  - Accounts inactive for 180+ days
  - Passwords not changed in 90+ days
  - Never-expiring passwords (non-service accounts)
  
- **🟡 Yellow Background**:
  - Disabled user accounts
  - Accounts inactive for 90-180 days
  - Expired accounts
  - Locked accounts
  - Password change restrictions

- **🟢 Green Background**:
  - Accounts active within 90 days
  - Administrative accounts (with bold formatting)
  - Service accounts (properly identified)

#### Special Formatting
- **Bold Text**: Administrative roles and service accounts
- **Date Formatting**: Consistent MM/dd/yyyy hh:mm AM/PM format
- **Auto-sizing**: Columns automatically sized for content
- **Freeze Panes**: Top row and first column frozen for easy navigation

---

## 🛠️ Troubleshooting

### Common Issues and Solutions

#### ❌ Active Directory Module Not Found
**Error**: `Import-Module : The specified module 'ActiveDirectory' was not loaded`

**Solutions**:
1. **Install RSAT Tools**:
   ```powershell
   # Windows 10/11
   Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
   
   # Windows Server
   Install-WindowsFeature -Name RSAT-AD-PowerShell
   ```

2. **Alternative Installation**:
   ```powershell
   # Using Windows Features
   Enable-WindowsOptionalFeature -Online -FeatureName RSATClient-Roles-AD-Powershell
   ```

#### ❌ Execution Policy Restrictions
**Error**: `Execution of scripts is disabled on this system`

**Solutions**:
```powershell
# Temporary bypass (current session only)
PowerShell.exe -ExecutionPolicy Bypass -File "UserAuditReport.ps1"

# Permanent change (requires admin)
Set-ExecutionPolicy RemoteSigned -Scope LocalMachine

# User-specific change
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### ❌ Graph API Authentication Failures
**Error**: `Connect-MgGraph : Authentication failed` or `Insufficient privileges`

**Solutions**:
1. **Clear Authentication Cache**:
   ```powershell
   Disconnect-MgGraph
   Clear-MgContext
   ```

2. **Verify Required Permissions**:
   - Directory.Read.All
   - User.Read.All  
   - AuditLog.Read.All

3. **Check Admin Consent**:
   - Ensure admin consent granted for application permissions
   - Use Global Administrator account for initial setup

#### ❌ Memory or Performance Issues
**Symptoms**: Script hangs or runs slowly with large user bases

**Solutions**:
```powershell
# Increase PowerShell memory limits
$PSDefaultParameterValues = @{
    '*:MaximumReceivedDataSizePerCommand' = 500MB
    '*:MaximumReceivedObjectSize' = 200MB
}

# Process users in batches for very large environments
# (Modify script to process in chunks of 1000 users)
```

#### ❌ Excel Export Issues
**Error**: `Export-Excel : The term 'Export-Excel' is not recognized`

**Solutions**:
1. **Manual Module Installation**:
   ```powershell
   Install-Module ImportExcel -Force -AllowClobber
   Import-Module ImportExcel
   ```

2. **Troubleshoot Installation**:
   ```powershell
   # Check module availability
   Get-Module -ListAvailable ImportExcel
   
   # Update PowerShellGet if needed
   Install-Module PowerShellGet -Force
   ```

3. **CSV Fallback**:
   - Script automatically falls back to CSV if ImportExcel unavailable
   - Manual formatting required for CSV output

### Debug Information Collection

#### Enable Detailed Logging
```powershell
# Add to beginning of script for debugging
$VerbosePreference = "Continue"
$DebugPreference = "Continue"

# Run with verbose output
.\UserAuditReport.ps1 -Verbose
```

#### PowerShell Module Diagnostics
```powershell
# Check PowerShell version
$PSVersionTable

# List installed modules
Get-Module -ListAvailable | Where-Object {$_.Name -like "*Graph*" -or $_.Name -like "*ImportExcel*" -or $_.Name -like "*ActiveDirectory*"}

# Check execution policy
Get-ExecutionPolicy -List

# Test AD connectivity
Test-ComputerSecureChannel
```

---

## 🔒 Security Considerations

### Data Protection
- **Sensitive Information**: Report contains privileged account information
- **Access Control**: Restrict report access to authorized personnel only
- **Storage Security**: Store reports in secure, encrypted locations
- **Retention Policy**: Implement appropriate data retention policies

### Account Security
- **Least Privilege**: Run script with minimum required permissions
- **Service Account**: Consider dedicated service account for automated runs
- **Audit Trail**: Log script executions and report access
- **Regular Updates**: Keep PowerShell modules updated for security patches

### Network Security
- **Secure Channels**: All communications use encrypted channels (HTTPS/TLS)
- **Authentication**: Multi-factor authentication recommended for Graph API access
- **Firewall Rules**: Ensure PowerShell can reach required endpoints
- **Proxy Configuration**: Configure proxy settings if required

---

## 🤝 Contributing

We welcome contributions to improve this User Audit Report tool!

### How to Contribute

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature-enhancement`
3. **Make your changes** with clear documentation
4. **Test thoroughly** in different environments
5. **Submit a pull request** with detailed description

### Contribution Guidelines

- **Code Style**: Follow PowerShell best practices and existing style
- **Documentation**: Update README and inline comments
- **Testing**: Test with various AD/Entra ID configurations
- **Error Handling**: Include appropriate try/catch blocks
- **Backwards Compatibility**: Maintain compatibility with existing features

### Areas for Enhancement

- 🏗️ **Additional Modules**: Support for other identity providers
- 📊 **Enhanced Reporting**: Additional report formats and visualizations
- 🔧 **Performance**: Optimization for very large environments
- 🌐 **Localization**: Multi-language support
- 📱 **Integration**: API endpoints for integration with other tools
- 🛡️ **Security**: Additional security analysis features

### Reporting Issues

When reporting issues, please include:
- **PowerShell Version**: `$PSVersionTable.PSVersion`
- **Module Versions**: `Get-Module -ListAvailable | Select Name, Version`
- **Environment**: AD size, Entra ID configuration, hybrid setup
- **Error Messages**: Complete error text and stack traces
- **Steps to Reproduce**: Detailed reproduction steps
- **Expected Behavior**: What should happen vs. what actually happens

---

<p align="center">
    <strong>🚀 Ready to audit your user accounts?</strong><br>
    <a href="#quick-start">Get Started Now</a>
</p>
