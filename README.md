# 365AdminTool - Microsoft 365 Admin PowerShell Module

This PowerShell module provides cmdlets for managing and retrieving information from Microsoft 365 (formerly Office 365) environments. It utilizes Microsoft Graph API and Exchange Online Management PowerShell Module for connectivity and operations.

## Features

- Retrieve information about Microsoft 365 domains, users, licenses, and admin roles.
- Query DNS records and check DKIM settings for domains.
- Manage mailbox settings, including email forwarding and message copies.
- Check MFA status and methods for users.
- Connect to Microsoft Graph (MgGraph) and Exchange Online.
- Detailed error handling and connection management.

## Prerequisites

1. **PowerShell**: Ensure PowerShell 5.1 or later is installed.
2. **Microsoft Graph Module**: Install using `Install-Module Microsoft.Graph -Scope AllUsers` if not already installed.
3. **Exchange Online Management Module**: Install using `Install-Module ExchangeOnlineManagement` if needed for Exchange operations.
4. **Microsoft 365 Account**: You need admin access to a Microsoft 365 tenant to perform most operations.

## Installation

1. Clone the repository or download the module files.
   ```bash
   git clone https://github.com/RapidScripter/365AdminTool.git

2. Open PowerShell with Administrator privileges.

3. Either load the script directly using:
   ```powershell
   <path_to_script>\365AdminTool.ps1
   
4. OR rename the script to <b>365AdminTool.psm1</b> and save it as a module within your PowerShell modules in a new folder also called <b>365AdminTool</b>. Then the functions within the script will work as cmdlets. After renaming and saving the script/module:
   ```powershell
   Import-Module <path_to_module>\365AdminTool.psm1

5. View the list of Commands: `<b>Get-365Command</b>`

## Command List

### 1. Get-365DNSInfo

- **Summary:** Returns information about each mail-configured domain in Microsoft 365.
  
- **Usage:** `Get-365DNSInfo [-Domain <string>]`
  
- **Example:**
  ```powershell
  # Retrieve information for all domains
  $domainsInfo = Get-365DNSInfo
  $domainsInfo | Format-List

  # Export the information to a CSV file
  $domainsInfo | Export-Csv -NoTypeInformation -Path M365MailSetup.csv

  # Retrieve information for a specific domain
  Get-365DNSInfo -Domain "example.com"

### 2. Resolve-DNSSummary

- **Summary:** Query DNS for a specific Domain and return a Summary.
  
- **Usage:** `Resolve-DNSSummary -Domain <string>`
- **Usage:** `Resolve-DNSSummary -Name <string>`
  
- **Example:**
  ```powershell
  Resolve-DNSSummary -Domain example.com
  Resolve-DNSSummary -Name example.com

### 3. Get-365licenses

- **Summary:** Returns a summary of all Microsoft subscriptions/licenses that are configured.

- **Usage:** `Get-365licenses`

- **Example:**
  ```powershell
  Get-365licenses

### 4. Get-365user

- **Summary:** Gets details about users within a Microsoft 365 account.

- **Usage:** `Get-365user [-userPrincipalName <string>]`

- **Example:**
  ```powershell
  # Retrieve information for all users
  Get-365user
  
  # Retrieve information for a specific user
  Get-365user -userPrincipalName info@contoso.com
  
  # Export user information to a CSV file
  Get-365user | Export-Csv -NoTypeInformation listOfUsers.csv

### 5. Get-365Whoami

- **Summary:** Retrieves information about the currently signed-in user(s) for various Microsoft 365 services.

- **Usage:** `Get-365Whoami [-DontElaborate] [-checkIfSignedInTo <string>]`

- **Example:**
  ```powershell
  # Retrieve a summary of signed-in users
  Get-365Whoami -DontElaborate
  
  # Check if signed in to a specific service
  Get-365Whoami -checkIfSignedInTo MgGraph

### 6. Get-365Domains

- **Summary:** Gets a summarized list of domains from Microsoft 365.

- **Usage:** `Get-365Domains [-EmailEnabled]`

- **Example:**
  ```powershell
  # Retrieve a list of all domains
  Get-365Domains
  
  # Retrieve only email-enabled domains
  Get-365Domains -EmailEnabled

### 7. Connect-365

- **Summary:** Connects to Microsoft Graph (MgGraph) using the MS prompt.

- **Usage:** `Connect-365 [-SilentifAlreadyConnected]`

- **Example:**
  ```powershell
  Connect-365

### 8. Disconnect-365

- **Summary:** Disconnects from Microsoft Graph (MgGraph) and Exchange Online.

- **Usage:** `Disconnect-365`

- **Example:**
  ```powershell
  Disconnect-365

### 9. Get-365Admins

- **Summary:** Gets details showing admin roles assigned to users in Microsoft 365.

- **Usage:** `Get-365Admins`

- **Example:**
  ```powershell
  Get-365Admins

### 10. Get-365UserMFAMethods

- **Summary:** Gets the MFA status and methods for the specified user(s).

- **Usage:** `Get-365UserMFAMethods -userId <string>`

- **Example:**
  ```powershell
  # Retrieve MFA methods for a specific user by UPN
  Get-365UserMFAMethods -userId info@contoso.com
  
  # Retrieve MFA methods for a specific user by ID
  Get-365UserMFAMethods -userId fe636523-5608-438d-83f5-41b5c9a7fe95

### 11. Connect-JustToExchange

- **Summary:** Connects to Exchange Online, installing the required module if necessary.

- **Usage:** `Connect-JustToExchange [-Identity <string>]`

- **Example:**
  ```powershell
  # Connect to Exchange Online using the currently signed-in identity
  Connect-JustToExchange
  
  # Connect to Exchange Online using a specified identity
  Connect-JustToExchange -Identity info@contoso.com

### 12. Set-MailBoxMessageSentAsCopy

- **Summary:** Sets a mailbox to keep a copy of emails sent on behalf of or SendAs another mailbox.

- **Usage:** `Set-MailBoxMessageSentAsCopy [-UserPrincipalName <string>] [-AllMailboxes] [-KeepSentCopy <boolean>]`

- **Example:**
  ```powershell
  # Set a mailbox to keep a copy of sent items
  Set-MailBoxMessageSentAsCopy -UserPrincipalName "info@contoso.com"
  
  # Set all mailboxes to keep a copy of sent items
  Set-MailBoxMessageSentAsCopy -AllMailboxes -KeepSentCopy $true

### 13. Get-365Command

- **Summary:** Generates a list of cmdlets within this script/module.

- **Usage:** `Get-365Command`

- **Example:**
  ```powershell
  Get-365Command
