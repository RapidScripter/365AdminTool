<#
.SYNOPSIS
  Returns information about each mail-configured domain in M365.
.DESCRIPTION
  Ensure you Connect-MgGraph -Scopes "Domain.Read.All" first.
  MgGraph can be installed with Install-Module Microsoft.Graph -Scope AllUsers.
  It takes a while, so make sure it is not already installed before you try to install again.
  
  [Optional] To retrieve DKIM settings from 365:
  Connect-ExchangeOnline.
  Exchange-online module can be installed with Install-Module ExchangeOnlineManagement.

  Requires at least:
  Connect-MgGraph -Scopes "AuditLog.Read.All","Mail.Read","Domain.Read.All".
  
  Or try:
  Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","AuditLog.Read.All","Mail.Read","Domain.Read.All","RoleManagement.Read.All","Policy.Read.All","Directory.Read.All","Organization.Read.All".
.PARAMETER Domain
  Optional: If used, will return information ONLY on the specific domain name (and only if it is also within the 365 Account).
  The FQN of the domain (e.g. imatec.co.nz).
.EXAMPLE
  # Retrieve information for all domains
  $domainsInfo = Get-365DNSInfo
  $domainsInfo | Format-List

  # Export the information to a CSV file
  $domainsInfo | Export-Csv -NoTypeInformation -Path M365MailSetup.csv

  # Retrieve information for a specific domain
  Get-365DNSInfo -Domain "example.com"
#>
function Get-365DNSInfo {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName)]
    [Alias("Name", "id")]
    [string[]]$Domain 
  )
  begin {
    Connect-365 -SilentIfAlreadyConnected
    Connect-JustToExchange
  }
  process {
    Write-Verbose "Retrieving data from MgGraph..."
    if ($Domain) {
      $domains = Get-MgDomain | Where-Object Id -in $Domain
    } else {
      $domains = Get-MgDomain | Where-Object Id -NotLike "*.onmicrosoft.com"
    }

    foreach ($adomain in $domains) {
      $domainid = $adomain.id
      $configuredForMail = $adomain.supportedServices -contains "Email"
      $DNSrecs = Get-MgDomainServiceConfigurationRecord -DomainId $domainid
      $spfs = ($DNSrecs | Where-Object recordType -eq "Txt" | Select-Object -ExpandProperty AdditionalProperties -ErrorAction SilentlyContinue).text -join ", "
      $MXrecs = ($DNSrecs | Where-Object recordType -eq "Mx").AdditionalProperties.mailExchange -join ", "
      $Autodiscover = ($DNSrecs | Where-Object { ($_.recordType -eq "CNAME") -and ($_.AdditionalProperties.canonicalName -like "autodiscover.*") }).AdditionalProperties.canonicalName
      [string]$M365DKIM = (Get-DkimSigningConfig -Identity $domainid -ErrorAction SilentlyContinue).Enabled
      $resolvedDNS = Resolve-DNSSummary -Domain $domainid
      $AutoDiscover365 = [bool]($Autodiscover -and $resolvedDNS.AutoDiscover)
      if (!$M365DKIM) { $M365DKIM = "Not yet configured: $domainid is not configured for DKIM" }

      $arec = [PSCustomObject]@{
        Name                 = $domainid
        M365_MailEnabled     = $configuredForMail
        Autodiscover365      = $AutoDiscover365
        M365_DKIM_Configured = $M365DKIM
        SOA                  = $resolvedDNS.Provider
        M365_spf             = $spfs
        DNS_spf              = $resolvedDNS.SPF
        M365_mx              = $MXrecs
        DNS_mx               = $resolvedDNS.MX
        DNS_DKIM_SMX         = $resolvedDNS.DKIM_SMX
        DNS_DKIM_M365        = $resolvedDNS.DKIM_365
      }
      $arec
    }
  }
  end {}
}


<#
.SYNOPSIS
Query DNS for a specific Domain - return a Summary

.DESCRIPTION
Query DNS for a specific Domain - return a Summary
provides summary of MX, Home IP (usually also WWW), www, SPF and identifies if DKIM is configured for our commonly used systems

.PARAMETER Domain
has an alias of Name
has an alias of id
the FQN (Domain) that needs to be resolved
This MUST be the DOMAIN suffix only, do not include the hostname
i.e use example.com , and not www.example.com

.EXAMPLE
Resolve-DNSSummary -Domain example.com   
Resolve-DNSSummary -Name example.com       
#>
function Resolve-DNSSummary {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName)]
    [Alias("Name")]
    [Alias("id")]
    [string[]] $Domain
  )
  begin {}
  Process {
    foreach ($adomain in $Domain) {
      $SOA = (Resolve-DnsName -Name $adomain -Type SOA -ErrorAction SilentlyContinue).PrimaryServer
      $spfDNS = (Resolve-DnsName -Name $adomain -Type TXT -ErrorAction SilentlyContinue | Where-Object { $_.Strings -Like "*v=spf1*" }).strings -join ", "
      $MxinDNS = (Resolve-DnsName -Name $adomain -Type MX -ErrorAction SilentlyContinue | Where-Object Name -eq $adomain).NameExchange -join ", " 
      $dnsroot = (Resolve-DnsName -Name $adomain -ErrorAction SilentlyContinue | Where-Object Name -eq $adomain).IP4Address -join ", " 
      $www = (Resolve-DnsName -Name www.$adomain -ErrorAction SilentlyContinue | Where-Object Name -eq $adomain).IP4Address -join ", " 
      $Autodiscover = (Resolve-DnsName -Name autodiscover.$adomain -Type CNAME -ErrorAction SilentlyContinue).NameHost

      $arec = [PSCustomObject]@{
        Name         = $adomain
        Home         = $dnsroot
        www          = $www
        Provider     = $SOA
        AutoDiscover = $Autodiscover 
        MX           = $MxinDNS
        DKIM_SMX     = ""
        DKIM_365     = ""
        SPF_SMX      = ""
        SPF_365      = ""
        SPF          = $spfDNS
      }

      switch ($SOA) {
        { $SOA -Like "*1stDomains*" } { $arec.Provider = "1stDomains" }
        { $SOA -Like "*cms-tool*" } { $arec.Provider = "WebsiteWorld" }
        { $SOA -Like "*cloudflare*" } { $arec.Provider = "CloudFlare" }
        { $SOA -Like "*crazydomains*" } { $arec.Provider = "CrazyDomains" }
        { $SOA -Like "*domaincontrol*" } { $arec.Provider = "Bluehost.com (domaincontrol.com)" }
        { $SOA -Like "*cpanel.com*" } { $arec.Provider = "Domainz.co.nz (server-cpanel.com)" }
        { $SOA -Like "*onlydomains.com*" } { $arec.Provider = "OnlyDomains" }
        { $SOA -Like "*omninet.co.nz*" } { $arec.Provider = "OmniNet" }
        { $SOA -Like "*wix.com" } { $arec.Provider = "wix.com" }
        { $SOA -Like "*wixdns.net*" } { $arec.Provider = "wix.com (wixdns.net)" }
      }

      if ($spfDNS -Like "*include:spf.nz.smxemail.com*all") { $arec.SPF_SMX = $true }
      if ($spfDNS -Like "*include:spf.protection.outlook.com*all") { $arec.SPF_365 = $true }

      $DKIMsmxinDNS1 = (Resolve-DnsName -Name smx1._domainkey.$adomain -Type CNAME -ErrorAction SilentlyContinue) | Select-Object NameHost
      $DKIMsmxinDNS2 = (Resolve-DnsName -Name smx2._domainkey.$adomain -Type CNAME -ErrorAction SilentlyContinue) | Select-Object NameHost
      if ($DKIMsmxinDNS1 -or $DKIMsmxinDNS2) {
        $arec.DKIM_SMX = "$($DKIMsmxinDNS1.NameHost), $($DKIMsmxinDNS2.NameHost)"
      }
      $DKIMM365inDNS1 = (Resolve-DnsName -Name selector1._domainkey.$adomain -Type CNAME -ErrorAction SilentlyContinue) | Select-Object NameHost
      $DKIMM365inDNS2 = (Resolve-DnsName -Name selector2._domainkey.$adomain -Type CNAME -ErrorAction SilentlyContinue) | Select-Object NameHost
      if ($DKIMM365inDNS1 -or $DKIMM365inDNS2) {
        $arec.DKIM_365 = "$($DKIMM365inDNS1.NameHost), $($DKIMM365inDNS2.NameHost)"
      }
      if ($arec.Home) { $arec }
      else {
        Write-Host "Resolve-DNSSummary: Did not find records for domain: $adomain" -ForegroundColor Red
      }
    }
  }
  end {}
}


<#
.SYNOPSIS
Returns a summary of all Microsoft subscriptions/licenses that are configured.

.DESCRIPTION
Returns a description of subscriptions used within the account and shows the amount of available licenses left in each subscription.

Requires at least:
Connect-MgGraph -Scopes "Organization.Read.All"

.EXAMPLE
Get-365licenses
#>
function Get-365licenses {
  [CmdletBinding()]
  param ()
  
  Connect-365 -SilentIfAlreadyConnected
  
  $licensedetails = Get-MgSubscribedSku | Where-Object { $_.AppliesTo -eq "User" -and $_.CapabilityStatus -eq "Enabled" } | Select-Object SkuPartNumber, @{Name = "Prepaid"; Expression = { $_.PrepaidUnits.Enabled } }, ConsumedUnits, SkuId
  
  foreach ($license in $licensedetails) {
    switch ($license.SkuPartNumber) {
      "O365_BUSINESS_ESSENTIALS" { $license.SkuPartNumber = 'Microsoft 365 Business Basic' }
      "O365_BUSINESS_PREMIUM" { $license.SkuPartNumber = 'Microsoft 365 Business Standard' }
      "EXCHANGESTANDARD" { $license.SkuPartNumber = 'Exchange Online (Plan 1)' }
      "STANDARDPACK" { $license.SkuPartNumber = 'Office 365 E1' }
    }
    
    $availableUnits = [int]($license.Prepaid) - [int]($license.ConsumedUnits)
    $license | Add-Member -NotePropertyName "AvailableUnits" -NotePropertyValue $availableUnits -Force
  }
  
  $licensedetails
}


<#
.SYNOPSIS
Gets details about users within a Microsoft 365 account.

.DESCRIPTION
Retrieves information about all users in a Microsoft 365 account. Provides a list of licenses used by each user. Collections such as Licenses, email-aliases, signInActivity are in JSON format (suitable for exporting to CSV).

Requires at least the following rights:
Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All"
Or alternatively:
Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","AuditLog.Read.All","Mail.Read","Domain.Read.All","RoleManagement.Read.All","Policy.Read.All","Directory.Read.All","Organization.Read.All"

If you want to see mail statistics, also connect to Exchange Online:
Connect-ExchangeOnline

.PARAMETER userPrincipalName
Allows you to retrieve data about a specific user by their userPrincipalName.

.EXAMPLE
get-365user
get-365user -userPrincipalName info@contoso.com
get-365user | Export-Csv -NoTypeInformation listOfUsers.csv
$variable = get-365user
#>
function Get-365user {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName)]
    [Alias('Name')]
    [string[]]$userPrincipalName,

    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName)]
    [Alias('id')]
    [string]$userid,

    [switch]$basicInfoOnly,
    [switch]$ShowMFA,
    [switch]$showMailBox,
    [switch]$EnablebUsersOnly
  )

  begin {
    Connect-365 -SilentIfAlreadyConnected
    if ($showMailBox) {
      Connect-JustToExchange
    }
  }

  process {
    $filter = ""
    if ($userPrincipalName) {
      $filter = "&`$filter=userPrincipalName in ('$($userPrincipalName -join "','")')"
      Write-Debug "Get-365User: Filter Name $filter"
    }
    if ($userid) {
      $filter = "&`$filter=id in ('$($userid -join "','")')"
    }

    if ($basicInfoOnly) {
      $basicEndpoint = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,Mail,accountEnabled,onPremisesSamAccountName,userType'
      $result = Invoke-MgGraphRequest -Method GET "$basicEndpoint$filter" -OutputType PSObject
      if ($result) {
        if ($EnablebUsersOnly -and $result.value) { $result.value = $result.value | Where-Object { $_.accountEnabled -eq $true } }
        $result.value
      }
      return
    }

    $endpoint = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,Mail,proxyAddresses,licenseAssignmentStates,accountEnabled,lastPasswordChangeDateTime,onPremisesSyncEnabled,onPremisesDomainName,onPremisesDistinguishedName,onPremisesSamAccountName,userType'

    try {
      $result = Invoke-MgGraphRequest -Method GET "$endpoint,signInActivity$filter" -OutputType PSObject
      $result.value | Add-Member -NotePropertyName get_errors -NotePropertyValue "No errors getting this information from Microsoft 365"
    }
    catch {
      $result = Invoke-MgGraphRequest -Method GET "$endpoint$filter" -OutputType PSObject
      $result.value | Add-Member -NotePropertyName signInActivity -NotePropertyValue ""
      $result.value | Add-Member -NotePropertyName get_errors -NotePropertyValue "Cannot get signInActivity, tenant is neither B2C nor has premium license"
    }

    if ($result) {
      $users = $result.value
      if ($EnablebUsersOnly -and $users) { $users = $users | Where-Object { $_.accountEnabled -eq $true } }
      if ($showMailBox) {
        $users | Add-Member -NotePropertyName "MailSize" -NotePropertyValue ""
        $users | Add-Member -NotePropertyName "MailSizeLimit" -NotePropertyValue ""
        $users | Add-Member -NotePropertyName "MailBoxType" -NotePropertyValue ""
        $users | Add-Member -NotePropertyName "LastUserMailAction" -NotePropertyValue ""
      }
      if ($ShowMFA) {
        $users | Add-Member -NotePropertyName "MFAInfo" -NotePropertyValue ""
      }

      $lic = Get-365licenses
      foreach ($user in $users) {
        $userskus = @()
        $user.proxyAddresses = ($user.proxyAddresses | Where-Object { $_ -like "SMTP*" }) -replace "SMTP:", "" | ConvertTo-Json -Compress

        foreach ($userlic in $user.licenseAssignmentStates) {
          $alic = ($lic | Where-Object { $_.SkuId -eq $userlic.skuid }).SkuPartNumber
          if (!$alic) { $alic = $userlic.skuid }
          if ($userlic.state -ne "Active") { $alic = "$alic <Inactive>" }
          $userskus += $alic
        }
        $user.licenseAssignmentStates = $userskus | ConvertTo-Json -Compress

        if ($user.signInActivity) { $user.signInActivity = $user.signInActivity | Select-Object lastSignInDateTime, lastNonInteractiveSignInDateTime | ConvertTo-Json -Compress }

        if ($showMailBox -and $user.mail) {
          $maildetail = Get-exomailboxStatistics -UserPrincipalName $user.mail -Properties MailboxTypeDetail, SystemMessageSizeShutoffQuota, LastUserActionTime -ErrorAction SilentlyContinue
          if ($maildetail.MailboxTypeDetail) {
            $user.MailSize = $maildetail.TotalItemSize
            $user.MailSizeLimit = $maildetail.SystemMessageSizeShutoffQuota
            $user.MailBoxType = $maildetail.MailboxTypeDetail
            $user.LastUserMailAction = $maildetail.LastUserActionTime
          }
          else {
            $user.mail = ""
            $user.proxyAddresses = ""
          }
        }

        if ($ShowMFA) {
          $user.MFAInfo = Get-365UserMFAMethods -userId $user.id | ConvertTo-Json -Compress
        }

        $user
      }
    }
    else {
      Write-Host "Get-365user: Did not find any user entries based on $filter"
    }
  }

  end {}
}


<#
.SYNOPSIS
    Retrieves information about the currently signed-in user(s) for various Microsoft 365 services.

.DESCRIPTION
    This function retrieves information about the user(s) currently signed in to Microsoft 365 services:
    - MgGraph (Microsoft Graph)
    - Exchange Online
    - Azure AD

.PARAMETER DontElaborate
    Use this switch to suppress detailed scope information for MgGraph.

.PARAMETER checkIfSignedInTo
    Specifies the service to check for signed-in user. Allowed values: "MgGraph", "Exchange", "AzureAD".
    Returns the UserPrincipalName (UPN) or Account ID if signed in, or $null if not.

.EXAMPLE
    get-365Whoami -DontElaborate
    Returns a summary of signed-in users without detailed scope information.

.EXAMPLE
    get-365Whoami -checkIfSignedInTo MgGraph
    Returns the UPN of the user signed in to MgGraph.

#>
function Get-365Whoami {
    [CmdletBinding()]
    param(
        [switch]
        $DontElaborate,
        [ValidateSet("MgGraph", "Exchange", "AzureAD")]
        [string] $checkIfSignedInTo
    )

    # Initialize result variables
    $results = @{
        MgGraph  = ""
        Exchange = ""
        AzureAD  = ""
    }

    # Function to check Microsoft Graph sign-in
    function Get-MgGraphSignIn {
        try {
            Write-Verbose "Checking login for MgGraph..."
            $graphResult = Invoke-MgGraphRequest -Method GET 'https://graph.microsoft.com/v1.0/me?$select=userPrincipalName' -ErrorAction Stop
            return $graphResult.userPrincipalName
        }
        catch {
            Write-Warning "Failed to retrieve MgGraph user info: $_"
            return $null
        }
    }

    # Function to check Exchange Online sign-in
    function Get-ExchangeSignIn {
        try {
            Write-Verbose "Checking login for Exchange Online..."
            $exchangeResult = Get-ConnectionInformation -ErrorAction Stop
            if ($exchangeResult) {
                return $exchangeResult.UserPrincipalName
            }
        }
        catch {
            Write-Warning "Failed to retrieve Exchange Online user info: $_"
            return $null
        }
    }

    # Function to check Azure AD sign-in
    function Get-AzureADSignIn {
        try {
            Write-Verbose "Checking login for Azure AD..."
            $azureResult = Get-AzureADCurrentSessionInfo -ErrorAction Stop
            if ($azureResult) {
                return $azureResult.Account.ID
            }
        }
        catch {
            Write-Warning "Failed to retrieve Azure AD user info: $_"
            return $null
        }
    }

    # Perform sign-in checks based on parameters
    if ($checkIfSignedInTo -in "MgGraph", $null) {
        $results.MgGraph = Get-MgGraphSignIn
    }

    if ($checkIfSignedInTo -in "Exchange", $null) {
        $results.Exchange = Get-ExchangeSignIn
    }

    if ($checkIfSignedInTo -in "AzureAD", $null) {
        $results.AzureAD = Get-AzureADSignIn
    }

    # Output based on parameters
    if ($checkIfSignedInTo) {
        return $results.$checkIfSignedInTo
    }
    else {
        [PSCustomObject]@{
            MgGraph  = $results.MgGraph
            Exchange = $results.Exchange
            AzureAD  = $results.AzureAD
            MSoline  = "Not checked"  # Placeholder for additional services
        }
    }

    # Output scopes if not suppressed and signed in to MgGraph
    if ($results.MgGraph -and -not $DontElaborate) {
        try {
            Write-Verbose "Getting MgGraph scopes..."
            $mgContext = Get-MgContext
            Write-Verbose "MgGraph Scopes are:"
            $mgContext.Scopes | ConvertTo-Json -Compress
        }
        catch {
            Write-Warning "Failed to retrieve MgGraph scopes: $_"
        }
    }
}


<#
.SYNOPSIS
Gets a summarized list of domains from Microsoft 365.

.DESCRIPTION
Retrieves a summarized list of domains from Microsoft 365. Requires prior authentication with appropriate scopes.

.PARAMETER EmailEnabled
Filters domains that support email services.

.EXAMPLE
get-365Domains

.NOTES
This function utilizes 'Connect-365 -SilentifAlreadyConnected' to ensure a connection before retrieving domain information.
#>
function Get-365Domains {
    [CmdletBinding()]
    param (
        [switch]$EmailEnabled
    )

    # Ensure connection to Microsoft Graph with required scopes
    Connect-365 -SilentifAlreadyConnected

    # Retrieve domains from Microsoft 365
    $domains = Get-MgDomain | Select-Object Id, IsDefault, IsVerified, SupportedServices

    # Filter domains if EmailEnabled switch is specified
    if ($EmailEnabled) {
        $domains = $domains | Where-Object { $_.SupportedServices -contains "Email" }
    }

    # Output the list of domains
    return $domains
}


<#
.SYNOPSIS
Connects to Microsoft Graph (MgGraph) using the MS prompt.

.DESCRIPTION
Connects to Microsoft Graph (MgGraph). Depending on your workstation setup, it may auto-connect with prior credentials without prompting for new ones. If you need to log in with different credentials, use Disconnect-365 first. Some scripts may also need to connect to ExchangeOnline, in which case the script will prompt when required.

.EXAMPLE
Disconnect-365
Connect-365

.PARAMETER SilentifAlreadyConnected
Use this switch to suppress the output message when already connected to MgGraph.

.NOTES
Ensure the Microsoft Graph module is installed before running this function.
#>
function Connect-365 {
    [CmdletBinding()]
    param (
        [Switch]$SilentifAlreadyConnected
    )

    # Check if Microsoft Graph module is installed
    if (-not (Get-InstalledModule Microsoft.Graph)) {
        Write-Host "Microsoft Graph module not found." -ForegroundColor Black -BackgroundColor Yellow
        $install = Read-Host "Do you want to install the Microsoft Graph Module? (Y/N)"

        if ($install -match "[yY]") {
            Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        }
        else {
            Write-Host "Microsoft Graph module is required." -ForegroundColor Black -BackgroundColor Yellow
            exit
        }
    }

    # Check current connections to MgGraph
    $connections = (Get-365Whoami -DontElaborate).MgGraph

    # If already connected and SilentifAlreadyConnected switch is not used, display message
    if ($connections -and !$SilentifAlreadyConnected) {
        Write-Host "Already connected to MgGraph with UserPrincipalName: $connections" -ForegroundColor Cyan
        return
    }

    # If not connected, prompt for connection to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.Read.All,Group.Read.All,AuditLog.Read.All,Mail.Read,Domain.Read.All,RoleManagement.Read.All,Policy.Read.All,Directory.Read.All,Organization.Read.All,UserAuthenticationMethod.Read.All,AuthenticationContext.Read.All" -NoWelcome

    # Get updated connections after connection attempt
    $connections = (Get-365Whoami -DontElaborate).MgGraph
    if ($connections) {
        Write-Host "Successfully connected to MgGraph with UserPrincipalName: $connections" -ForegroundColor Green
    }
    else {
        Write-Host "Failed to connect to MgGraph. Please check your credentials." -ForegroundColor Red
    }
}


<#
.SYNOPSIS
Disconnects from Microsoft Graph (MgGraph) and Exchange Online.

.DESCRIPTION
Disconnects from Microsoft Graph (MgGraph) and Exchange Online if connected.

.EXAMPLE
Disconnect-365
Disconnects from both MgGraph and Exchange Online if connected.

.NOTES
Ensure the respective modules (Microsoft.Graph, ExchangeOnlineManagement) are imported before running this function.
#>
function Disconnect-365 {
    [CmdletBinding()]
    param (
    )

    # Disconnect from MgGraph
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    # Disconnect from Exchange Online if signed in
    if (Get-365Whoami -checkIfSignedInTo Exchange) {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }
}


<#
.SYNOPSIS
Gets details showing admin roles assigned to users in Microsoft 365.

.DESCRIPTION
Retrieves details showing admin roles assigned to users in Microsoft 365.
Requires connection to Microsoft Graph.

.EXAMPLE
Get-365Admins
Gets details of admin roles assigned to users in Microsoft 365.

.NOTES
Ensure the Connect-365 function is defined and properly connects to Microsoft Graph.
#>
function Get-365Admins {
    [CmdletBinding()]
    param (
    )

    # Connect to Microsoft Graph
    Connect-365 -SilentifAlreadyConnected

    # Retrieve admin roles
    $adminRoles = Get-MgDirectoryRole | Select-Object DisplayName, Id, Description

    foreach ($role in $adminRoles) {
        # Get members of each admin role
        $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | Where-Object { $_.AdditionalProperties."@odata.type" -eq "#microsoft.graph.user" }

        foreach ($member in $roleMembers) {
            # Get detailed user information
            $user = Get-365User -UserPrincipalName $member.UserPrincipalName -BasicInfoOnly -EnableUsersOnly

            if ($user) {
                [PSCustomObject]@{
                    Role              = $role.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    UserDisplayName   = $user.DisplayName
                    Description       = $role.Description
                }
            }
        }
    }
}


<#
.SYNOPSIS
Gets the MFA status and methods for the specified user(s).

.DESCRIPTION
Retrieves the MFA status and methods for the specified user(s) in Microsoft 365 using Microsoft Graph API.
Requires prior connection to Microsoft Graph.

.PARAMETER userId
Specifies the UserPrincipalName or ID of the user(s) to retrieve MFA information for.

.EXAMPLE
Get-365UserMFAMethods -userId info@contoso.com -Verbose
Get-365UserMFAMethods -userId fe636523-5608-438d-83f5-41b5c9a7fe95
#>
Function Get-365UserMFAMethods {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName)]
    [Alias('Name')]
    [string[]] $userId
  )

  begin {
    # Connect to Office 365 or Microsoft 365
    function Connect-365 {
      # Replace with your actual connection logic
      Write-Verbose "Connecting to Office 365 or Microsoft 365..."
      # Example: Connect-MicrosoftTeams or Connect-MicrosoftGraph
    }

    # Ensure connection is established
    Connect-365
  }
  
  process {
    foreach ($auser in $userId) {
      Write-Verbose "Getting MFA methods for user: $auser"
      
      try {
        [array]$mfaData = Get-MgUserAuthenticationMethod -UserId $auser -ErrorAction Stop
      }
      catch {
        Write-Warning "Failed to retrieve MFA methods for user $auser $($_.Exception.Message)"
        continue
      }
    
      if (!$mfaData) { return }
  
      # Initialize MFA details object
      $mfaMethods = [PSCustomObject]@{
        Name                  = $auser
        status                = ""
        authApp               = ""
        phoneAuth             = ""
        fido                  = ""
        helloForBusiness      = ""
        helloForBusinessCount = 0
        emailAuth             = ""
        tempPass              = ""
        passwordLess          = ""
        softwareAuth          = ""
        authDevice            = ""
        authPhoneNr           = ""
        SSPREmail             = ""
        OtherInfo             = ""
      }
  
      # Process each authentication method
      foreach ($method in $mfaData) {
        Switch ($method.AdditionalProperties["@odata.type"]) {
          "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { 
            $mfaMethods.authApp = $true
            $mfaMethods.authDevice += "$($method.AdditionalProperties["displayName"]),"
            $mfaMethods.status = "enabled"
          } 
          "#microsoft.graph.phoneAuthenticationMethod" { 
            $mfaMethods.phoneAuth = $true
            $mfaMethods.authPhoneNr = $method.AdditionalProperties["phoneType", "phoneNumber"] -join ' '
            $mfaMethods.status = "enabled"
          } 
          "#microsoft.graph.fido2AuthenticationMethod" { 
            $mfaMethods.fido = $true
            $mfaMethods.OtherInfo += "Fido-Model:$($method.AdditionalProperties["model"]),"
            $mfaMethods.status = "enabled"
          } 
          "#microsoft.graph.passwordAuthenticationMethod" { 
            if ($mfaMethods.status -ne "enabled") { $mfaMethods.status = "disabled" }
          }
          "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
            $mfaMethods.helloForBusiness = $true
            $mfaMethods.OtherInfo += "Hello-Device:$($method.AdditionalProperties["displayName"]),"
            $mfaMethods.status = "enabled"
            $mfaMethods.helloForBusinessCount++
          } 
          "#microsoft.graph.emailAuthenticationMethod" { 
            $mfaMethods.emailAuth = $true
            $mfaMethods.SSPREmail = $method.AdditionalProperties["emailAddress"] 
            $mfaMethods.status = "enabled"
          }               
          "microsoft.graph.temporaryAccessPassAuthenticationMethod" { 
            $mfaMethods.tempPass = $true
            $mfaMethods.OtherInfo += "TempPass-LifeTime:$($method.AdditionalProperties["lifetimeInMinutes"]),"
            $mfaMethods.status = "enabled"
          }
          "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { 
            $mfaMethods.passwordLess = $true
            $mfaMethods.OtherInfo += "passwordless-devicve:$($method.AdditionalProperties["displayName"]),"
            $mfaMethods.status = "enabled"
          }
          "#microsoft.graph.softwareOathAuthenticationMethod" { 
            $mfaMethods.softwareAuth = $true
            $mfaMethods.status = "enabled"
          }
        }
      }
      
      # Trim trailing commas from authDevice and OtherInfo
      $mfaMethods.authDevice = $mfaMethods.authDevice.TrimEnd(",")
      $mfaMethods.OtherInfo = $mfaMethods.OtherInfo.TrimEnd(",")
  
      # Output the MFA details object
      Write-Output $mfaMethods
    }
  }
}


<#
.SYNOPSIS
Connects to Exchange Online, installing the required module if necessary.

.DESCRIPTION
Connects to Exchange Online, ensuring the ExchangeOnlineManagement module is installed if not already present. It verifies if already connected to Exchange and disconnects if connected with a different identity.

.PARAMETER Identity
Specifies the UserPrincipalName to connect to Exchange Online. If not provided, it uses the identity signed in to MgGraph.

.EXAMPLE
Connect-JustToExchange
Connects to Exchange Online using the identity signed in to MgGraph.

Connect-JustToExchange -Identity info@contoso.com
Connects to Exchange Online using the specified identity.
#>
function Connect-JustToExchange {
    [CmdletBinding()]
    param(
        [string]$Identity
    )

    # Check if already signed in to Exchange
    $isSignedInToExchange = Get-365Whoami -checkIfSignedInTo Exchange

    if (!$Identity) {
        # If no specific Identity provided, use the identity signed in to MgGraph
        $Identity = Get-365Whoami -checkIfSignedInTo MgGraph
    }

    # Disconnect if signed in with a different identity than required
    if ($isSignedInToExchange -ne $Identity -and $isSignedInToExchange) {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        $isSignedInToExchange = $null
    }

    if (!$isSignedInToExchange) {
        Write-Host "Connecting to Exchange Online ($Identity)..." -ForegroundColor Cyan

        # Check if ExchangeOnlineManagement module is installed
        if (-not (Get-InstalledModule ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
            Write-Host "Microsoft ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Black -BackgroundColor Yellow
            $install = Read-Host "Do you want to install the Microsoft ExchangeOnlineManagement module? (Y/N)"

            if ($install -match "[yY]") {
                Install-Module ExchangeOnlineManagement -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
            }
            else {
                Write-Host "ExchangeOnlineManagement module is required to connect to Exchange Online." -ForegroundColor Black -BackgroundColor Yellow
                return
            }
        }

        # Connect to Exchange Online
        Connect-ExchangeOnline -UserPrincipalName $Identity -ShowBanner:$false
    }
    else {
        Write-Host "Already connected to Exchange Online with user $Identity" -ForegroundColor Green
    }
}


<#
.SYNOPSIS
Sets a mailbox to keep a copy of emails sent on behalf of or SendAs another mailbox.

.DESCRIPTION
By default, Office 365 mailboxes do not keep a duplicate copy of an email sent on behalf of or SendAs another mailbox. This command changes that behavior by setting a mailbox to keep a copy of the sent item in its own sent items folder.

.PARAMETER UserPrincipalName
The email address of the mailbox to apply the settings to. Not applicable if the AllMailboxes parameter is used.

.PARAMETER AllMailboxes
Ensures that all user or shared mailboxes have their settings changed.

.PARAMETER KeepSentCopy
If set to $true, ensures the target mailbox keeps a copy of the sent item. If $false, only the mailbox that originates the SendAs or SendOnBehalf will keep a copy of the email in its sent items.

.EXAMPLE
Set-MailBoxMessageSentAsCopy -UserPrincipalName "info@contoso.com"

Set-MailBoxMessageSentAsCopy -AllMailboxes -KeepSentCopy $true
#>
function Set-MailBoxMessageSentAsCopy {
    [CmdletBinding(DefaultParameterSetName = "UserPrincipalName")]
    param (
        [Parameter(ParameterSetName = "UserPrincipalName", ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias("Identity")]
        [string[]]$UserPrincipalName,

        [Parameter(ParameterSetName = "AllMailboxes")]
        [switch]$AllMailboxes,

        [bool]$KeepSentCopy = $true
    )

    begin {
        Connect-JustToExchange
        if ($AllMailboxes) {
            Write-Host "Setting all user/shared mailboxes to keep SendAs/SendOnBehalf sent item copy = $KeepSentCopy"
            $mailboxes = Get-EXOMailbox -Filter "(RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'UserMailbox')"
            foreach ($mailbox in $mailboxes) {
                Set-Mailbox -Identity $mailbox.UserPrincipalName -MessageCopyForSendOnBehalfEnabled $KeepSentCopy -MessageCopyForSentAsEnabled $KeepSentCopy
                Write-Host "Set mailbox $($mailbox.UserPrincipalName) to keep SendAs/SendOnBehalf sent item copy = $KeepSentCopy"
            }
            return
        }
    }

    process {
        if ($UserPrincipalName) {
            foreach ($upn in $UserPrincipalName) {
                Set-Mailbox -Identity $upn -MessageCopyForSendOnBehalfEnabled $KeepSentCopy -MessageCopyForSentAsEnabled $KeepSentCopy
                Write-Host "Set mailbox $upn to keep SendAs/SendOnBehalf sent item copy = $KeepSentCopy"
            }
        }
    }
}


<#
.SYNOPSIS
Generates a list of cmdlets within this script/module.

.DESCRIPTION
Generates a list of cmdlets within this script/module.
FYI: If this script is renamed as a *.psm1 (instead of a *.ps1) and installed within a folder "365AdminTool" under the PowerShell\modules directory, you can call these commands without manually importing the script.

.EXAMPLE
Get-365Command
#>
function Get-365Command {
    [CmdletBinding()]
    param ()
    
    $module = "365AdminTool"
    $commands = Get-Command -Module $module
    
    if ($commands) {
        $commands
        return
    }
    
    Write-Host "Get-365Command: Will only show you the full list of commands when 365AdminTool is installed as a module (*.psm1)." -ForegroundColor Yellow
    Write-Host "Since you ran this script as . ./365AdminTool.ps1, the list below is manual and may be inaccurate." -ForegroundColor Yellow
    Write-Host @"
CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Get-365Command                                     1.0        $module
Function        Connect-365                                        1.0        $module
Function        Disconnect-365                                     1.0        $module
Function        Connect-JustToExchange                             1.0        $module
Function        Get-365Admins                                      1.0        $module
Function        Get-365DNSInfo                                     1.0        $module
Function        Get-365Domains                                     1.0        $module
Function        Get-365licenses                                    1.0        $module
Function        Get-365user                                        1.0        $module
Function        Get-365UserMFAMethods                              1.0        $module
Function        Get-365Whoami                                      1.0        $module
Function        Resolve-DNSSummary                                 1.0        $module
Function        Set-MailBoxMessageSentAsCopy                       1.0        $module
"@
}

Get-365Command