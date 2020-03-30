<#
.SYNOPSIS
    Gets the Office 365 Health Center messages and incidents
    Displays on several tabs with the summary providing a dashbord overview

.DESCRIPTION
    This script is designed to run on a scheduled basis.
    It requires and Azure AD Application for your tenant.
    The script will build an HTML page displaying information and links to tenant messages and articles.
    Several tenants configurations can be held.
    Office 365 service health incidents generate email alerts as well as logging to the event log.
    Alternative links can be added (ie to an external dashboard) should data retrieval fail

    Requires Azure AD powershell module (Install-Module AzureAD)

.INPUTS
    None

.OUTPUTS
    None

.EXAMPLE
    PS C:\> O365SH.ps1 -configXML ..\config\production.xml

.EXAMPLE
	Build all linked incident and advisory documents as local HTML files

    PS C:\> O365SH.ps1 -configXML ..\config\production.xml -RebuildDocs


.NOTES
    Author:  Jonathan Christie
    Email:   jonathan.christie (at) boilerhouseit.com
    Date:    02 Feb 2019
    PSVer:   2.0/3.0/4.0/5.0
    Version: 2.0.1
    Updated:
    UpdNote:

    Wishlist:

    Completed:

    Outstanding:

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [String]$configXML = "..\config\profile-bhitprod.xml",
    [Parameter(Mandatory = $false)] [Switch]$RebuildDocs = $false
)

$swScript = [system.diagnostics.stopwatch]::StartNew()
Write-Verbose "Changing Directory to $PSScriptRoot"
Set-Location $PSScriptRoot
Import-Module "..\common\O365ServiceHealth.psm1"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)] [string]$info
    )
    # verify the Log is setup and if not create the file
    if ($script:loginitialized -eq $false) {
        $script:FileHeader >> $script:logfile
        $script:loginitialized = $True
    }
    $info = $(Get-Date).ToString() + ": " + $info
    $info >> $script:logfile
}

if ([system.IO.path]::IsPathRooted($configXML) -eq $false) {
    #its not an absolute path. Find the absolute path
    $configXML = Resolve-Path $configXML
}
$config = LoadConfig $configXML

#Configure local event log
[string]$evtLogname = $config.EventLog
#[string]$evtSource = $config.MonitorEvtSource
[string]$evtSource = "Toolbox"
if ($config.UseEventlog -like 'true') {
    [boolean]$UseEventLog = $true
    #check source and log exists
    $CheckLog = [System.Diagnostics.EventLog]::Exists("$($evtLogname)")
    $CheckSource = [System.Diagnostics.EventLog]::SourceExists("$($evtSource)")
    if ((! $CheckLog) -or (! $CheckSource)) {
        New-EventLog -LogName $evtLogname -Source $evtSource
    }
}
else { [boolean]$UseEventLog = $false }

#Declare variables
[string]$tenantID = $config.TenantID
[string]$tenantMSName = $config.TenantMSName
[string]$appID = $config.AppID
[string]$clientSecret = $config.AppSecret
[string]$emailEnabled = $config.EmailEnabled
[string]$SMTPUser = $config.EmailUser
[string]$SMTPPassword = $config.EmailPassword
[string]$SMTPKey = $config.EmailKey

[string]$HTMLFile = $config.ToolboxHTML
[string]$rptName = $config.ToolboxName
[int]$pageRefresh = $config.ToolboxRefresh
[string]$toolboxNotes = $config.ToolboxNotes

[string]$rptProfile = $config.TenantShortName
[string]$rptTenantName = $config.TenantName

[string]$pathLogs = $config.LogPath
[string]$pathHTML = $config.HTMLPath
[string]$pathIPURLs = $config.IPURLsPath
[string]$pathWorking = $config.WorkingPath

[string]$cnameFilename = $config.CNAMEFilename
[string[]]$cnameURLs = $config.CNAMEUrls.split(",")
$cnameURLs = $cnameURLs.Replace('"', '')
$cnameURLs = $cnameURLs.Trim()
[string[]]$cnameResolvers = $config.CNAMEResolvers.split(",")
$cnameResolvers = $cnameResolvers.Replace('"', '')
$cnameResolvers = $cnameResolvers.Trim()

[string[]]$cnameResolverDesc = $config.CNAMEResolverDesc.split(",")
$cnameResolverDesc = $cnameResolverDesc.Replace('"', '')
$cnameResolverDesc = $cnameResolverDesc.Trim()
[string]$cnameNotes = $config.CnameNotes

if ($cnameresolvers[0] -eq "") {
    $cnameResolvers = @(Get-DnsClientServerAddress | Sort-Object interfaceindex | Select-Object -ExpandProperty serveraddresses | Where-Object { $_ -like '*.*' } | Select-Object -First 1)
    $cnameResolverDesc = @("Default")
}

[string]$pacEnabled = $config.PACEnabled
[string]$pacProxy = $config.PACProxy
[string]$pacType1Filename = $config.PACType1Filename
[string]$pacType2Filename = $config.PACType2Filename



[string[]]$emailIPURLAlerts = $config.IPURLsAlertsTo
[string]$fileIPURLsNotes = "$($config.IPURLsNotesFilename)-$($rptProfile).csv"
[string]$fileIPURLsNotesAll = "$($config.IPURLsNotesFilename)All-$($rptProfile).csv"
[string]$fileCustomNotes = "$($config.CustomNotesFilename)-$($rptProfile).csv"
[int]$IPURLHistory = $config.IPURLHistory
[string]$proxyHost = $config.ProxyHost
[array]$customURLs = @()

if ($IPURLHistory -le 1) { $IPURLHistory = 6 }
#Check diagnostics and save as boolean
if ($config.DiagnosticsEnabled -like 'true') { [boolean]$diagEnabled = $true } else { [boolean]$diagEnabled = $false }
if ($config.DiagnosticsURLs -like 'true') { [boolean]$diagURLs = $true } else { [boolean]$diagURLs = $false }
if ($config.DiagnosticsVerbose -like 'true') { [boolean]$diagVerbose = $true } else { [boolean]$diagVerbose = $false }
if ($config.MiscDiagsWeb -like 'true') { [boolean]$diagWeb = $true } else { [boolean]$diagWeb = $false }
if ($config.MiscDiagsPorts -like 'true') { [boolean]$diagPorts = $true } else { [boolean]$diagPorts = $false }
if ($config.MiscDiagsEnabled -like 'true') { [boolean]$miscDiagsEnabled = $true } else { [boolean]$miscDiagsEnabled = $false }
if ($config.EmailEnabled -like 'true') { [boolean]$emailEnabled = $true } else { [boolean]$emailEnabled = $false }
if ($config.PACEnabled -like 'true') { [boolean]$pacEnabled = $true } else { [boolean]$pacEnabled = $false }
if ($config.CNAMEEnabled -like 'true') { [boolean]$cnameEnabled = $true } else { [boolean]$cnameEnabled = $false }

[boolean]$rptOutage = $false
[boolean]$fileMissing = $false

[string]$cssfile = ".\O365Health.css"

# Get Email credentials
# Check for a username. No username, no need for credentials (internal mail host?)
[PSCredential]$emailCreds = $null
if ($emailEnabled -and $smtpuser -notlike '') {
    #Email credentials have been specified, so build the credentials.
    #See readme on how to build credentials files
    $EmailCreds = getCreds $SMTPUser $SMTPPassword $SMTPKey
}


#Check the various file paths, set default, create and make absolute reference if necessary
$pathLogs = CheckDirectory $pathLogs
$pathHTML = CheckDirectory $pathHTML
$pathIPURLs = CheckDirectory $pathIPURLs
$pathWorking = CheckDirectory $pathWorking

# setup the logfile
# If logfile exists, the set flag to keep logfile
$script:DailyLogFile = "$($pathLogs)\O365Toolbox-$($rptprofile)-$(Get-Date -format yyMMdd).log"
$script:LogFile = "$($pathLogs)\tmpO365Toolbox-$($rptprofile)-$(Get-Date -format yyMMddHHmmss).log"
$script:LogInitialized = $false
$script:FileHeader = "*** Toolbox Information ***"

$evtMessage = "Config File: $($configXML)"
Write-Log $evtMessage
$evtMessage = "Log Path: $($pathLogs)"
Write-Log $evtMessage
$evtMessage = "HTML Output: $($pathHTML)"
Write-Log $evtMessage

if ($config.CNAMEEnabled -like 'true') { [boolean]$cnameEnabled = $true } else { [boolean]$cnameEnabled = $false }

#Create event logs if set
if ($UseEventLog) {
    $evtCheck = Get-EventLog -List -ErrorAction SilentlyContinue | Where-Object { $_.LogDisplayName -eq $evtLogname }
    if (!($evtCheck)) {
        New-EventLog -LogName $evtLogname -Source $evtSource
        Write-ELog -LogName $evtLogname -Source $evtSource -Message "Event log created." -EventId 1 -EntryType Information
    }
}

#Proxy Configuration
if ($config.ProxyEnabled -like 'true') {
    [boolean]$ProxyServer = $true
    $evtMessage = "Using proxy server $($proxyHost) for connectivity"
    Write-Log $evtMessage
}
else {
    [boolean]$ProxyServer = $false
    $evtMessage = "No proxy to be used."
    Write-Log $evtMessage
}


#Connect to Azure app and grab the service status
ConnectAzureAD
[string]$urlResource = "https://manage.office.com/.default"
[uri]$authority = "https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token"
$reqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = $urlResource
    Client_ID     = $appID
    Client_Secret = $clientSecret
}

if ($proxyServer) {
    $bearerToken = invoke-RestMethod -uri $authority -Method Post -Body $reqTokenBody -Proxy $proxyHost -ProxyUseDefaultCredentials
}
else {
    $bearerToken = Invoke-RestMethod -uri $authority -Method Post -Body $reqTokenBody
}

if ($null -eq $bearerToken) {
    $evtMessage = "ERROR - No authentication result for Azure AD App"
    Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 10 -EntryType Error
    Write-Log $evtMessage
}

#Now remove any system default proxy in order to test no-proxy paths.
$defaultproxy = [System.Net.WebProxy]::GetDefaultProxy()
[System.Net.GlobalProxySelection]::Select = [System.Net.GlobalProxySelection]::GetEmptyWebProxy()

function EnsureAzureADModule() {
    # Query for installed Azure AD modules
    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
    if ($null -eq $AadModule) {
        Write-Output "AzureAD PowerShell module not found, looking for AzureADPreview"
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
    }

    if ($null -eq $AadModule) {
        Write-Output
        Write-Output "AzureAD Powershell module not installed..." -f Red
        Write-Output "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        Write-Output "Script can't continue..." -f Red
        Write-Output
        exit
    }

    if ($AadModule.count -gt 1) {
        $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
        $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }
        # Checking if there are multiple versions of the same module found
        if ($AadModule.count -gt 1) {
            $aadModule = $AadModule | Select-Object -Unique
        }
        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }
    else {
        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }

    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
}

function BuildHTML {
    Param (
        [Parameter(Mandatory = $true)] $Title,
        [Parameter(Mandatory = $true)] $contentOne,
        [Parameter(Mandatory = $true)] $contentTwo,
        [Parameter(Mandatory = $true)] $contentThree,
        [Parameter(Mandatory = $true)] $contentFour,
        [Parameter(Mandatory = $true)] $contentFive,
        [Parameter(Mandatory = $true)] $contentLast,
        [Parameter(Mandatory = $true)] $swStopWatch,
        [Parameter(Mandatory = $true)] $HTMLOutput
    )
    [array]$htmlHeader = @()
    [array]$htmlBody = @()
    [array]$htmlFooter = @()

    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<link rel="stylesheet" href="O365Health.css">
<style>
</style>
<title>$($rptTitle)</title>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
<meta http-equiv="Pragma" content="no-cache" />
<meta http-equiv="Expires" content="0" />
</head>
"@

    $htmlBody = @"
<body>
    <p>Page refreshed: <span id="datetime"></span><span>&nbsp;&nbsp;Data refresh: $(Get-Date -f 'dd-MMM-yyyy HH:mm:ss')</span></p>
	<div class="tab">
    <button class="tablinks" onclick="openTab(event,'Diagnostics')" id="defaultOpen">Diagnostics</button>
    <button class="tablinks" onclick="openTab(event,'Licences')">Licences</button>
    <button class="tablinks" onclick="openTab(event,'IPsandURLs')">IPs and URLs</button>
    <button class="tablinks" onclick="openTab(event,'URLs')">URLs</button>
"@

    if ($cnameenabled) {
        $htmlBody += @"
    <button class="tablinks" onclick="openTab(event,'CNAMEs')">CNAMEs</button>
"@
    }

    $htmlBody += @"
    <button class="tablinks" onclick="openTab(event,'Logs')">Logs</button>
</div>

<!-- Tab content -->
<div id="Diagnostics" class="tabcontent">
    $($contentOne)
</div>

<div id="Licences" class="tabcontent">
    $($contentTwo)
</div>

<div id="IPsandURLs" class="tabcontent">
    $($contentThree)
</div>

<div id="URLs" class="tabcontent">
    $($contentFour)
</div>
"@

    if ($cnameenabled) {
        $htmlBody += @"
<div id="CNAMEs" class="tabcontent">
    $($contentFive)
</div>
"@
    }
    $htmlBody += @"
<div id="Logs" class="tabcontent">
    $($contentLast)
</div>
"@

    $htmlFooter = @"
<script>
var dt = new Date();
document.getElementById("datetime").innerHTML = (("0"+dt.getDate()).slice(-2)) +"-"+ (("0"+(dt.getMonth()+1)).slice(-2)) +"-"+ (dt.getFullYear()) +" "+ (("0"+dt.getHours()).slice(-2)) +":"+ (("0"+dt.getMinutes()).slice(-2)) +":"+ (("0"+dt.getSeconds()).slice(-2));
</script>

<script>
function openTab(evt, tabName) {
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }

    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active","");
    }
    document.getElementById(tabName).style.display = "block";
    evt.currentTarget.className += " active";
}

document.getElementById("defaultOpen").click();
</script>
$($scriptRuntime)
</body>
</html>
"@

    #Add in code to refresh page
    #Editing after file is generated increases the file size drastically
    $addJava = "<script language=""JavaScript"" type=""text/javascript"">"
    $addJava += "setTimeout(""location.href='$($HTMLOutput)'"",$($pageRefresh*60*1000));"
    $addjava += "</script>"

    $htmlReport = $htmlHeader + $addJava + $htmlBody + $htmlFooter
    $htmlReport | Out-File "$($pathHTML)\$($HTMLOutput)"
}

#Diagnostics
#Get the CRL endpoints and check they are valid
#shout out to Aaron at undocumentedfeatures.com for his AAD Connect test tool
# Test Online Networking Only
#For testing there are options: Full tests, include client script (download and run from client)
function OnlineEndPoints {
    Param(
        [Parameter(Mandatory = $true)] [boolean]$diagWeb,
        [Parameter(Mandatory = $true)] [boolean]$diagPorts,
        [Parameter(Mandatory = $true)] [boolean]$diagURLs

    )
    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting Online Endpoints tests.</p><br/>"
    #See https://support.office.com/en-us/article/office-365-urls-and-ip-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2
    $CRL = @(
        "http://ocsp.msocsp.com",
        "http://crl.microsoft.com/pki/crl/products/microsoftrootcert.crl",
        "http://mscrl.microsoft.com/pki/mscorp/crl/msitwww2.crl",
        "http://ocsp.verisign.com",
        "http://ocsp.entrust.net"
    )
    $RequiredResources = @(
        "adminwebservice.microsoftonline.com",
        "adminwebservice-s1-co2.microsoftonline.com",
        "login.microsoftonline.com",
        "provisioningapi.microsoftonline.com",
        "login.windows.net",
        "secure.aadcdn.microsoftonline-p.com",
        "management.core.windows.net",
        "bba800-anchor.microsoftonline.com",
        "graph.windows.net",
        "aadcdn.msauth.net",
        "aadcdn.msftauth.net",
        "ccscdn.msauth.net",
        "ccscdn.msftauth.net"
    )
    $RequiredResourcesEndpoints = @(
        "https://adminwebservice.microsoftonline.com/provisioningservice.svc",
        "https://adminwebservice-s1-co2.microsoftonline.com/provisioningservice.svc",
        "https://login.microsoftonline.com",
        "https://provisioningapi.microsoftonline.com/provisioningwebservice.svc",
        "https://login.windows.net",
        "https://secure.aadcdn.microsoftonline-p.com/ests/2.1.5975.9/content/cdnbundles/jquery.1.11.min.js"
    )
    $OptionalResources = @(
        "management.azure.com",
        "policykeyservice.dc.ad.msft.net",
        "s1.adhybridhealth.azure.com",
        "autoupdate.msappproxy.net",
        "adds.aadconnecthealth.azure.com",
        "enterpriseregistration.windows.net" # device registration
    )
    $OptionalResourcesEndpoints = @(
        "https://policykeyservice.dc.ad.msft.net/clientregistrationmanager.svc",
        "https://device.login.microsoftonline.com" # Hybrid device registration
    )
    $SeamlessSSOEndpoints = @(
        "autologon.microsoftazuread-sso.com",
        "aadg.windows.net.nsatc.net",
        "0.register.msappproxy.net",
        "0.registration.msappproxy.net",
        "proxy.cloudwebappproxy.net"
    )
    # Use the AdditionalResources array to specify items that need a port test on a port other
    # than 80 or 443.
    $AdditionalResources = @(
        "watchdog.servicebus.windows.net:5671")

    if ($diagWeb) {
        # CRL Endpoint tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting CRL Endpoint Tests (Invoke-WebRequest)</p><br/>"
        foreach ($url in $CRL) {
            $rptTests += checkURL $url $diagVerbose $proxyServer $proxyHost $false
        } # End Foreach CRL

        # Required Resources Endpoints tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Required Resources Endpoints (Invoke-WebRequest).</p><br/>"
        foreach ($url in $RequiredResourcesEndpoints) {
            $rptTests += checkURL $url $diagVerbose $proxyServer $proxyHost $false
        } # End Foreach RequiredResourcesEndpoints
	
        # Optional Resources Endpoints tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Optional Resources Endpoints (Invoke-WebRequest).</p><br/>"
        foreach ($url in $OptionalResourcesEndpoints) {
            $rptTests += checkURL $url $diagVerbose $proxyServer $proxyHost $false
        } # End Foreach RequiredResourcesEndpoints
    } #End web tests

    # Required Resource tests
    if ($diagPorts) {
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Required Resources (TCP:443) DNS Resolution may fail from clients.</p><br/>"
        foreach ($url in $RequiredResources) {
            try { [array]$ResourceAddresses = (Resolve-DnsName $url -ErrorAction stop -QuickTimeout).IP4Address }
            catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to resolve host URL $($url).</p><br/>"; Continue }
            foreach ($ip4 in $ResourceAddresses) {
                try {
                    $Result = Test-NetConnection $ip4 -Port 443 -ea stop -wa silentlycontinue -InformationLevel Quiet
                    switch ($Result) {
                        true { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>TCP connection to $($url) [$($ip4)]:443 success.</p><br/>" }
                        false { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>TCP connection to $($url) [$($ip4)]:443 failed.</p><br/>" }
                    }
                }
                catch {
                    $ErrorMessage = $_
                    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Error resolving or connecting to $($url) [$($ip4)]:443</p><br/>"
                    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>$($ErrorMessage)</p><br/>"
                }
            } 
        } # End Foreach Resources

        # Option Resources tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Optional Resources (TCP:443) DNS Resolution may fail from clients.</p><br/>"
        foreach ($url in $OptionalResources) {
            try { [array]$ResourceAddresses = (Resolve-DnsName $url -ErrorAction stop -QuickTimeout).IP4Address }
            catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Unable to resolve host URL $($url).</p><br/>"; Continue }
		
            foreach ($ip4 in $ResourceAddresses) {
                try {
                    $Result = Test-NetConnection $ip4 -Port 443 -ea stop -wa silentlycontinue -InformationLevel Quiet
                    switch ($Result) {
                        true { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>TCP connection to $($url) [$($ip4)]:443 success.</p><br/>" }
                        false {
                            $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>TCP connection to $($url) [$($ip4)]:443 failed.</p><br/>"
                        }
                    }
                }
                catch {
                    $ErrorMessage = $_
                    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Error resolving or connecting to $($url) [$($ip4)]:443</p><br/>"
                    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>$($ErrorMessage)</p><br/>"
                }
            }
        } # End Foreach OptionalResources


        # Seamless SSO Endpoints
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Seamless SSO Endpoints (TCP:443) DNS Resolution may fail from clients.</p><br/>"
        foreach ($url in $SeamlessSSOEndpoints) {
            try { [array]$ResourceAddresses = (Resolve-DnsName $url -ErrorAction stop -QuickTimeout).IP4Address }
            catch { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to resolve host URL $($url).</p><br/>"; Continue }
		
            foreach ($ip4 in $ResourceAddresses) {
                try {
                    $Result = Test-NetConnection $ip4 -Port 443 -ea stop -wa silentlycontinue -InformationLevel Quiet
                    switch ($Result) {
                        true { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>TCP connection to $($url) [$($ip4)]:443 success.</p><br/>" }
                        false { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>TCP connection to $($url) [$($ip4)]:443 failed.</p><br/>" }
                    }
                }
                catch {
                    $ErrorMessage = $_
                    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Error resolving or connecting to $($url) [$($ip4)]:443</p><br/>"
                    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage)</p><br/>"
				
                }
            }
        } # End Foreach Resources

        # Additional Resources tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Additional Resources Endpoints (Resolve-DNS; Test-NetConnection).</p><br/>"
        If ($AdditionalResources) {
            foreach ($url in $AdditionalResources) {
                if ($url -match "\:") {
                    $Name = $url.Split(":")[0]
                    try { [array]$Resources = (Resolve-DnsName $Name -ErrorAction stop -QuickTimeout).IP4Address }
                    catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Unable to resolve host $($Name).</p><br/>"; Continue }
				
                    #[array]$Resources = (Resolve-DnsName $Name).Ip4Address
                    $ResourcesPort = $url.Split(":")[1]
                }
                Else {
                    $Name = $url
                    try { [array]$Resources = (Resolve-DnsName $Name -ErrorAction stop -QuickTimeout).IP4Address }
                    catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Unable to resolve host URL $($url).</p><br/>"; Continue }
				
                    #[array]$Resources = (Resolve-DnsName $Name).IP4Address
                    $ResourcesPort = "443"
                }
                foreach ($ip4 in $Resources) {
                    try {
                        $Result = Test-NetConnection $ip4 -Port $ResourcesPort -ea stop -wa silentlycontinue -InformationLevel Quiet
                        switch ($Result) {
                            true { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>TCP connection to $($Name) [$($ip4)]:$($ResourcesPort) success.</p><br/>" }
                            false {
                                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>TCP connection to $($Name) [$($ip4)]:$($ResourcesPort) failed.</p><br>"
							
                                If ($DebugLogging) { }
                            }
                        }
                    }
                    catch {
                        $ErrorMessage = $_
                        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Error resolving or connecting to $($Name) [$($ip4)]:$($ResourcesPort)</p><br/>"
                        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>$($ErrorMessage)</p><br/>"
                    }
                } # End ForEach ip4
            } # End ForEach AdditionalResources
        } # End IF AdditionalResources
    }
    $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Finished Online Endpoints tests.</p><br/>"
    return $rptTests
} # End Function OnlineEndPoints

#https://docs.microsoft.com/en-gb/azure/active-directory/users-groups-roles/licensing-service-plan-reference

$SkuNames = @{
    "AAD_BASIC"                                   = "Azure Active Directory Basic"
    "AAD_PREMIUM"                                 = "Azure Active Directory Premium P1"
    "AAD_PREMIUM_P2"                              = "Azure Active Directory Premium P2"
    "AAD_SMB"                                     = ""
    "ADALLOM_STANDALONE"                          = "Microsoft Cloud App Security"
    "ADALLOM_O365"                                = "Office 365 Cloud App Security"
    "ADALLOM_S_DISCOVERY"                         = "Enterprise Mobility + Security E3"
    "ADALLOM_S_O365"                              = "Office 365 Advanced Security Management"
    "ADALLOM_S_STANDALONE"                        = "Microsoft Cloud App Security"
    "ATA"                                         = "Advanced Threat Analytics"
    "ATP_ENTERPRISE"                              = "Exchange Online Advanced Threat Protection"
    "ATP_ENTERPRISE_FACULTY"                      = "Exchange Online Advanced Threat Protection"
    "AX_ENTERPRISE_USER"                          = "AX Enterprise User"
    "AX_SANDBOX_INSTANCE_TIER2"                   = "AX SandBox Tier 2"
    "AX_SELF-SERVE_USER"                          = "AX Self-Serve User"
    "AX_TASK_USER"                                = "AX Task User"
    "AX7_USER_TRIAL"                              = "Microsoft Dynamics AX7 User Trial"
    "BI_AZURE_P0"                                 = "Power BI (free)"
    "BI_AZURE_P1"                                 = "Power BI Reporting and Analytics"
    "BI_AZURE_P2"                                 = "Power BI Pro"
    "BPOS_S_TODO_1"                               = "Microsoft To Do (Plan 1)"
    "BPOS_S_TODO_2"                               = "Microsoft To-Do (Plan 2)"
    "BPOS_S_TODO_3"                               = "Microsoft To-Do (Plan 3)"
    "CRM_ONLINE_PORTAL"                           = "Dynamics CRM Online Portal"
    "CRMINSTANCE"                                 = "Dynamics CRM Production Instance"
    "CRMIUR"                                      = "CRM for Partners"
    "CRMPLAN1"                                    = "Dynamics CRM Online (Plan 1)"
    "CRMPLAN2"                                    = "Dynamics CRM Online (Plan 2)"
    "CRMSTANDARD"                                 = "Dynamics CRM Online Professional"
    "CRMSTORAGE"                                  = "Dynamics CRM Online Additional Storage"
    "CRMTESTINSTANCE"                             = "Dyamics CRM Additional Production Instance"
    "D365_CSI_EMBED_CE"                           = "Dynamics 365 Customer Service Insights for CE Plan"
    "Deskless"                                    = "Microsoft StaffHub"
    "DESKLESSPACK"                                = "Office 365 (Plan K1)"
    "DESKLESSPACK_GOV"                            = "Office 365 (Plan K1) for Gov"
    "DESKLESSPACK_YAMMER"                         = "Office 365 (Plan K1) with Yammer"
    "DESKLESSWOFFPACK"                            = "Office 365 (Plan K2)"
    "DESKLESSWOFFPACK_GOV"                        = "Office 365 (Plan K2) for Gov"
    "DEVELOPERPACK"                               = "Office 365 Enterprise E3 Developer"
    "DEVELOPERPACK_E5"                            = "Office 365 Enterprise E5 Developer"
    "DYN365_AI_SERVICE_INSIGHTS"                  = "Dynamics 365 Customer Service Insights"
    "DYN365_AI_SERVICE_INSIGHTS_VIRAL"            = "Dynamics 365 Customer Service Insights Viral"
    "DYN365_BUSINESS_Marketing"                   = "Dynamics 365 for Marketing"
    "DYN365_CDS_DYN_APPS"                         = "Dynamics 365 for Talent"
    "DYN365_CDS_FORMS_PRO"                        = "Dynamics 365 Forms Pro"
    "DYN365_CDS_PROJECT"                          = "Dynamics 365 Project"
    "DYN365_CDS_VIRAL"                            = "Dynamics 365 Common Data Source"
    "DYN365_CUSTOMER_INSIGHTS_VIRAL"              = "Dynamics 365 Customer Insights "
    "DYN365_ENTERPRISE_CUSTOMER_SERVICE"          = "Dynamics 365 for Customer Service Enterprise Edition"
    "DYN365_ENTERPRISE_P1"                        = "Dynamics 365 Customer Engagement Plan"
    "DYN365_ENTERPRISE_P1_IW"                     = "Dynamics 365 P1 Trial for Information Workers"
    "DYN365_ENTERPRISE_PLAN1"                     = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
    "DYN365_ENTERPRISE_SALES"                     = "Dynamics Office 365 Enterprise Sales"
    "DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE"     = "Dynamics Office 365 Enterprise Sales and Customer Service"
    "DYN365_Enterprise_Talent_Attract_TeamMember" = "Dynamics 365 for Talent - Attract Experience Team member"
    "DYN365_Enterprise_Talent_Onboard_TeamMember" = "Dynamics 365 for Talent - Onboard Experience"
    "DYN365_ENTERPRISE_TEAM_MEMBERS"              = "Dynamics 365 For Team Members Enterprise Edition"
    "DYN365_FINANCIALS_BUSINESS_SKU"              = "Dynamics 365 for Financials Business Edition"
    "DYN365_FINANCIALS_TEAM_MEMBERS_SKU"          = "Dynamics 365 for Team Members Business Edition"
    "DYN365_MARKETING_APPLICATION_ADDON"          = "Dynamics 365 for Marketing Additional Application"
    "DYN365_MARKETING_CONTACT_ADDON"              = "Dynamics 365 for Marketing Addnl Contacts"
    "DYN365_MARKETING_CONTACT_CE_PLAN_ADDON"      = "Dynamics 365 for Marketing Addnl Contacts for CE"
    "DYN365_MARKETING_SANDBOX_APPLICATION_ADDON"  = "Dynamics 365 for Marketing Additional Non-Prod Application"
    "DYN365BC_MS_INVOICING"                       = "Microsoft Invoicing"
    "Dynamics_365_for_HCM_Trial"                  = "Dynamics 365 for Talent"
    "Dynamics_365_for_Operations"                 = "Dynamics 365 for Operations"
    "Dynamics_365_for_Operations_Team_Members"    = "Dynamics 365 for Operations Team Member"
    "Dynamics_365_for_Retail_Team_members"        = "Dynamics 365 for Retail Team Members"
    "Dynamics_365_for_Talent_Team_members"        = "Dynamics 365 for Talent Team Members"
    "Dynamics_365_Hiring_Free_PLAN"               = "Dynamics 365 for Talent"
    "Dynamics_365_Onboarding_Free_PLAN"           = "Dynamics 365 for Talent"
    "DYNAMICS_365_ONBOARDING_SKU"                 = "Dynamics 365 for Talent: Onboard"
    "DYNAMICS_365_TALENT_ONBOARD"                 = "Dynamics 365 for Talent: Onboard"
    "ECAL_SERVICES"                               = "Enterprise Client Access Services"
    "EMS"                                         = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                                  = "Enterprise Mobility + Security E5"
    "ENTERPRISEPACK"                              = "Office 365 Enterprise E3"
    "ENTERPRISEPACK_B_PILOT"                      = "Office 365 (Enterprise Preview)"
    "ENTERPRISEPACK_FACULTY"                      = "Office 365 (Plan A3) for Faculty"
    "ENTERPRISEPACK_GOV"                          = "Office 365 (Plan G3) for Gov"
    "ENTERPRISEPACK_STUDENT"                      = "Office 365 (Plan A3) for Students"
    "ENTERPRISEPACKLRG"                           = "Office 365 (Plan E3)"
    "ENTERPRISEPACKWITHOUTPROPLUS"                = "Office 365 (Plan E3) without ProPlus Add-on"
    "ENTERPRISEPREMIUM"                           = "Office 365 (Plan E5)"
    "ENTERPRISEPREMIUM_NOPSTNCONF"                = "Office 365 (Plan E5) without Audio Conferencing"
    "ENTERPRISEWITHSCAL"                          = "Office 365 (Plan E4)"
    "ENTERPRISEWITHSCAL_FACULTY"                  = "Office 365 (Plan A4) for Faculty"
    "ENTERPRISEWITHSCAL_GOV"                      = "Office 365 (Plan G4) for Gov"
    "ENTERPRISEWITHSCAL_STUDENT"                  = "Office 365 (Plan A4) for Students"
    "EOP_ENTERPRISE"                              = "Exchange Online Protection"
    "EOP_ENTERPRISE_FACULTY"                      = "Exchange Online Protection for Faculty"
    "EQUIVIO_ANALYTICS"                           = "Office 365 Advanced eDiscovery"
    "EQUIVIO_ANALYTICS_FACULTY"                   = "Office 365 Advanced Compliance for faculty"
    "ERP_TRIAL_INSTANCE"                          = "ERP Trial Instance"
    "ESKLESSWOFFPACK_GOV"                         = "Office 365 (Plan K2) for Gov"
    "EXCHANGE_ANALYTICS"                          = "Microsoft MyAnalytics"
    "EXCHANGE_L_STANDARD"                         = "Exchange Online (Plan 1)"
    "EXCHANGE_S_ARCHIVE_ADDON"                    = "Exchange Online Archiving"
    "EXCHANGE_S_ARCHIVE_ADDON_GOV"                = "Exchange Online Archiving Govt"
    "EXCHANGE_S_DESKLESS"                         = "Exchange Online Kiosk"
    "EXCHANGE_S_DESKLESS_GOV"                     = "Exchange Kiosk"
    "EXCHANGE_S_ENTERPRISE"                       = "Exchange Online (Plan 2) Ent"
    "EXCHANGE_S_ENTERPRISE_GOV"                   = "Exchange Online (Plan 2) Gov"
    "EXCHANGE_S_ESSENTIALS"                       = "Exchange Online Essentials"
    "EXCHANGE_S_FOUNDATION"                       = "Exchange Foundation"
    "EXCHANGE_S_STANDARD"                         = "Exchange Online (Plan 2)"
    "EXCHANGE_S_STANDARD_MIDMARKET"               = "Exchange Online (Plan 1)"
    "EXCHANGEARCHIVE"                             = "Exchange Online Archiving"
    "EXCHANGEARCHIVE_ADDON"                       = "Exchange Online Archiving Add-On"
    "EXCHANGEDESKLESS"                            = "Exchange Online Kiosk"
    "EXCHANGEENTERPRISE"                          = "Exchange Online Plan 2"
    "EXCHANGEENTERPRISE_FACULTY"                  = "Exchange Online (Plan 2) for Faculty"
    "EXCHANGEENTERPRISE_GOV"                      = "Exchange Online (Plan 2) for Gov"
    "EXCHANGEESSENTIALS"                          = "Exchange Online Essentials"
    "EXCHANGEONLINE_MULTIGEO"                     = "Exchange Online MultiGeo"
    "EXCHANGESTANDARD"                            = "Exchange Online"
    "EXCHANGESTANDARD_GOV"                        = "Exchange Online (Plan 1) for Gov"
    "EXCHANGESTANDARD_STUDENT"                    = "Exchange Online (Plan 1) for Students"
    "EXCHANGETELCO"                               = "Exchange Online Telco"
    "FLOW_DYN_APPS"                               = "Microsoft Flow [Dynamics]"
    "FLOW_DYN_P2"                                 = "Flow for Dynamics 365"
    "FLOW_DYN_TEAM"                               = "Flow for Dynamics 365"
    "FLOW_FOR_PROJECT"                            = "Flow for Project"
    "FLOW_FORMS_PRO"                              = "Flow for Forms Pro"
    "FLOW_FREE"                                   = "Microsoft Flow Free"
    "FLOW_O365_P1"                                = "Flow for Office 365 (Plan 1)"
    "FLOW_O365_P2"                                = "Flow for Office 365 (Plan 2)"
    "FLOW_O365_P3"                                = "Flow for Office 365 (Plan 3)"
    "FLOW_P2_VIRAL"                               = "Flow Free (Plan 2)"
    "FLOW_P2_VIRAL_REAL"                          = "Flow (Plan 2) [Trial]"
    "FORMS_PLAN_E1"                               = "Microsoft Forms (Plan E1)"
    "FORMS_PLAN_E3"                               = "Microsoft Forms (Plan E3)"
    "FORMS_PLAN_E5"                               = "Microsoft Forms (Plan E5)"
    "FORMS_PRO"                                   = "Forms Pro"
    "FORMS_PRO_CE"                                = "Forms Pro for Customer Engagement Plan"
    "INFOPROTECTION_P2"                           = "Azure Information Protection Premium P2"
    "INFORMATION_BARRIERS"                        = "Information Barriers"
    "INTUNE_A"                                    = "Intune (Plan A) for Office 365"
    "INTUNE_A_VL"                                 = "Intune (Volume License)"
    "INTUNE_O365"                                 = "Intune"
    "INTUNE_SMBIZ"                                = "Mobile Device Management Small Business"
    "IT_ACADEMY_AD"                               = "Microsoft Imagine Academy"
    "KAIZALA_O365_P2"                             = "Microsoft Kaizala Pro (Plan 2)"
    "KAIZALA_O365_P3"                             = "Microsoft Kaizala Pro (Plan 3)"
    "KAIZALA_STANDALONE"                          = "Microsoft Kaizala"
    "LITEPACK"                                    = "Office 365 (Plan P1)"
    "LITEPACK_P2"                                 = "Office 365 Small Business Premium"
    "LOCKBOX"                                     = "Customer Lockbox"
    "LOCKBOX_ENTERPRISE"                          = "Customer Lockbox"
    "M365_ADVANCED_AUDITING"                      = "Microsoft 365 Advanced Auditing"
    "MCO_TEAMS_IW"                                = "Microsoft Teams"
    "MCOEV"                                       = "Skype for Business Cloud PBX"
    "MCOIMP"                                      = "Skype for Business Online (Plan 1)"
    "MCOLITE"                                     = "Lync Online (Plan 1)"
    "MCOMEETACPEA"                                = "Audio Conferencing Pay-Per-Minute"
    "MCOMEETADV"                                  = "Skype for Business PSTN Conferencing"
    "MCOPSTN1"                                    = "Skype for Business PSTN Domestic Calling"
    "MCOPSTN2"                                    = "Skype for Business PSTN Domestic and International Calling"
    "MCOPSTNC"                                    = "Communication Credits"
    "MCOSTANDARD"                                 = "Skype for Business Online (Plan 2)"
    "MCOSTANDARD_GOV"                             = "Lync Online (Plan 2) for Gov"
    "MCOSTANDARD_MIDMARKET"                       = "Lync Online (Plan 1)"
    "MCVOICECONF"                                 = "Lync Online (Plan 3)"
    "MDM_SALES_COLLABORATION"                     = "Microsoft Dynamics Marketing Sales Collaboration"
    "MEE_FACULTY"                                 = "Minecraft Education Edition Faculty"
    "MEE_STUDENT"                                 = "Minecraft Education Edition Student"
    "MFA_PREMIUM"                                 = "Azure Multi-Factor Authentication"
    "MICROSOFT_BUSINESS_CENTER"                   = "Microsoft Business Center"
    "MICROSOFT_SEARCH"                            = "Microsoft Search"
    "MICROSOFTBOOKINGS"                           = "Microsoft Bookings"
    "MIDSIZEPACK"                                 = "Office 365 Midsize Business"
    "MINECRAFT_Edu_EDITION"                       = "Minecraft Education Edition Faculty"
    "MIP_S_CLP1"                                  = "Information Protection for Office 365 Standard"
    "MIP_S_CLP2"                                  = "Information Protection for Office 365 Premium"
    "MS_TEAMS_IW"                                 = "Microsoft Teams Trial"
    "MYANALYTICS_P2"                              = "Insights by MyAnalytics"
    "NBENTERPRISE"                                = "Microsoft Social Engagement"
    "NBPROFESSIONALFORCRM"                        = "Microsoft Social Listening Professional"
    "O365_BUSINESS"                               = "Office 365 Business"
    "O365_BUSINESS_ESSENTIALS"                    = "Office 365 Business Essentials"
    "O365_BUSINESS_PREMIUM"                       = "Office 365 Business Premium"
    "OFFICE_FORMS_PLAN_2"                         = "Microsoft Forms (Plan 2)"
    "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"          = "Office ProPlus"
    "OFFICE365_MULTIGEO"                          = "Multi-Geo Capabilities in Office 365"
    "OFFICEMOBILE_SUBSCRIPTION"                   = "Office Mobile Apps for Office 365"
    "OFFICESUBSCRIPTION"                          = "Office 365 ProPlus"
    "OFFICESUBSCRIPTION_FACULTY"                  = "Office 365 ProPlus for Faculty"
    "OFFICESUBSCRIPTION_GOV"                      = "Office 365 ProPlus for Gov"
    "OFFICESUBSCRIPTION_STUDENT"                  = "Office 365 ProPlus for Student"
    "ONEDRIVE_BASIC"                              = "OneDrive Basic"
    "ONEDRIVESTANDARD"                            = "OneDrive Standard"
    "PAM_ENTERPRISE"                              = "Office 365 Privileged Access Management"
    "PLANNERSTANDALONE"                           = "Planner Standalone"
    "POWER_BI_ADDON"                              = "Power BI for Office 365 Add-On"
    "POWER_BI_INDIVIDUAL_USER"                    = "Power BI for Office 365 Individual"
    "POWER_BI_PRO"                                = "Power BI Pro"
    "POWER_BI_STANDALONE"                         = "Power BI for Office 365 Standalone"
    "POWER_BI_STANDARD"                           = "Power BI for Office 365 Standard"
    "POWERAPPS_DYN_APPS"                          = "Microsoft PowerApps and Logic flows"
    "POWERAPPS_DYN_P2"                            = "PowerApps for Dynamics 365"
    "POWERAPPS_DYN_TEAM"                          = "PowerApps for Dynamics 365"
    "POWERAPPS_INDIVIDUAL_USER"                   = "Microsoft PowerApps and Logic flows"
    "POWERAPPS_O365_P1"                           = "PowerApps for Office 365"
    "POWERAPPS_O365_P2"                           = "PowerApps"
    "POWERAPPS_O365_P3"                           = "PowerApps for Office 365"
    "POWERAPPS_P2_VIRAL"                          = "Microsoft PowerApps (Plan 2) [Trial]"
    "POWERAPPS_VIRAL"                             = "Microsoft PowerApps [Trial]"
    "POWERAPPSFREE"                               = "Microsoft PowerApps and Logic Flows"
    "POWERFLOWSFREE"                              = "Microsoft PowerApps and Logic Flows"
    "POWERVIDEOSFREE"                             = "Microsoft PowerApps and Logic Flows"
    "PREMIUM_ENCRYPTION"                          = "Premium Encryption in Office 365"
    "PROJECT_CLIENT_SUBSCRIPTION"                 = "Project Pro for Office 365"
    "PROJECT_ESSENTIALS"                          = "Project Lite"
    "PROJECT_MADEIRA_PREVIEW_IW"                  = "Dynamics 365 for Financials for IWs"
    "PROJECT_MADEIRA_PREVIEW_IW_SKU"              = "Dynamics 365 for Financials for IWs"
    "PROJECT_PROFESSIONAL"                        = "Project Online Professional"
    "PROJECTCLIENT"                               = "Project Pro for Office 365"
    "PROJECTESSENTIALS"                           = "Project Lite"
    "PROJECTONLINE_PLAN_1"                        = "Project Online (Plan 1)"
    "PROJECTONLINE_PLAN_1_FACULTY"                = "Project Online (Plan 1) for Faculty"
    "PROJECTONLINE_PLAN_1_STUDENT"                = "Project Online (Plan 1) for Student"
    "PROJECTONLINE_PLAN_2"                        = "Project Online (Plan 2)"
    "PROJECTONLINE_PLAN_2_FACULTY"                = "Project Online (Plan 2) for Faculty"
    "PROJECTONLINE_PLAN_2_STUDENT"                = "Project Online (Plan 2) for Student"
    "ProjectPremium"                              = "Project Online Premium"
    "PROJECTPROFESSIONAL"                         = "Project Online Professional"
    "PROJECTWORKMANAGEMENT"                       = "Office 365 Planner"
    "RIGHTSMANAGEMENT"                            = "Azure Information protection (Plan 1)"
    "RIGHTSMANAGEMENT_ADHOC"                      = "Azure Rights Management Services Ad-hoc"
    "RIGHTSMANAGEMENT_STANDARD_FACULTY"           = "Information Rights Management for Faculty"
    "RIGHTSMANAGEMENT_STANDARD_STUDENT"           = "Information Rights Management for Students"
    "RMS_S_ADHOC"                                 = "Azure Rights Management Services Ad-hoc"
    "RMS_S_ENTERPRISE"                            = "Azure Active Directory Rights Management"
    "RMS_S_ENTERPRISE_GOV"                        = "Windows Azure Active Directory Rights Management"
    "RMS_S_PREMIUM"                               = "Azure Information Protection (Plan 1)"
    "RMS_S_PREMIUM2"                              = "Azure Information Protection Premium (Plan 2)"
    "SCHOOL_DATA_SYNC_P1"                         = "School Data Sync (Plan 1)"
    "SAFEDOCS"                                    = "Office 365 SafeDocs"
    "SHAREPOINT_PROJECT"                          = "Project Online (Plan 2)"
    "SHAREPOINT_PROJECT_EDU"                      = "SharePoint Project Online Service for Edu"
    "SHAREPOINTDESKLESS"                          = "SharePoint Online Kiosk"
    "SHAREPOINTDESKLESS_GOV"                      = "SharePoint Online Kiosk Gov"
    "SHAREPOINTENTERPRISE"                        = "SharePoint Online (Plan 2)"
    "SHAREPOINTENTERPRISE_EDU"                    = "SharePoint Online (Plan 2) for Edu"
    "SHAREPOINTENTERPRISE_GOV"                    = "SharePoint Online (Plan 2) for Gov"
    "SHAREPOINTENTERPRISE_MIDMARKET"              = "SharePoint Online (Plan 1) MidMarket"
    "SHAREPOINTLITE"                              = "SharePoint Online (Plan 1) Lite"
    "SHAREPOINTONLINE_MULTIGEO"                   = "SharePoint Online Multi-Geo"
    "SHAREPOINTPARTNER"                           = "SharePoint Online Partner Access"
    "SHAREPOINTSTANDARD"                          = "SharePoint Online (Plan 1)"
    "SHAREPOINTSTANDARD_EDU"                      = "SharePoint Online (Plan 1) for Edu"
    "SHAREPOINTSTORAGE"                           = "SharePoint Online Storage"
    "SHAREPOINTWAC"                               = "Office Online"
    "SHAREPOINTWAC_EDU"                           = "Office Online for Edu"
    "SHAREPOINTWAC_GOV"                           = "Office Online for Gov"
    "SKU_Dynamics_365_for_HCM_Trial"              = "Dynamics 365 for HCM [Trial]"
    "SMB_APPS"                                    = "Business Apps (free)"
    "SMB_BUSINESS"                                = "Office 365 Business"
    "SMB_BUSINESS_ESSENTIALS"                     = "Office 365 Business Essentials"
    "SMB_BUSINESS_PREMIUM"                        = "Office 365 Business Premium"
    "SPB"                                         = "Microsoft 365 Business"
    "SPE_E3"                                      = "Microsoft 365 E3"
    "SPE_E5"                                      = "Microsoft 365 E5"
    "SPZA"                                        = "Microsoft AppConnect"
    "SPZA_IW"                                     = "App Connect"
    "SQL_IS_SSIM"                                 = "Power BI Information Services"
    "STANDARD_B_PILOT"                            = "Office 365 (Small Business Preview)"
    "STANDARDPACK"                                = "Office 365 (Plan E1)"
    "STANDARDPACK_FACULTY"                        = "Office 365 (Plan A1) for Faculty"
    "STANDARDPACK_GOV"                            = "Office 365 (Plan G1) for Gov"
    "STANDARDPACK_STUDENT"                        = "Office 365 (Plan A1) for Students"
    "STANDARDWOFFPACK"                            = "Office 365 (Plan E2)"
    "STANDARDWOFFPACK_FACULTY"                    = "Office 365 (Plan E1) for Faculty"
    "STANDARDWOFFPACK_GOV"                        = "Office 365 (Plan G2) for Gov"
    "STANDARDWOFFPACK_IW_FACULTY"                 = "Office 365 Edu for Faculty"
    "STANDARDWOFFPACK_IW_STUDENT"                 = "Office 365 Edu for Students"
    "STANDARDWOFFPACK_STUDENT"                    = "Office 365 (Plan E1) for Student"
    "STANDARDWOFFPACKPACK_FACULTY"                = "Office 365 (Plan A2) for Faculty"
    "STANDARDWOFFPACKPACK_STUDENT"                = "Office 365 (Plan A2) for Students"
    "STREAM"                                      = "Microsoft Stream"
    "STREAM_O365_E1"                              = "Microsoft Stream for O365 E1 SKU"
    "STREAM_O365_E3"                              = "Microsoft Stream for Office 365 (Plan E3)"
    "STREAM_O365_E5"                              = "Microsoft Stream for Office 365 (Plan E5)"
    "SWAY"                                        = "Sway"
    "TEAMS_COMMERCIAL_TRIAL"                      = "Microsoft Teams Commercial Cloud (User Initiated)"
    "TEAMS1"                                      = "Microsoft Teams"
    "THREAT_INTELLIGENCE"                         = "Office 365 Threat Intelligence"
    "VISIO_CLIENT_SUBSCRIPTION"                   = "Visio Pro for Office 365 Subscription"
    "VISIOCLIENT"                                 = "Visio Pro for Office 365"
    "VISIOONLINE"                                 = "Visio Online"
    "VISIOONLINE_PLAN1"                           = "Visio Online (Plan 1)"
    "WACONEDRIVEENTERPRISE"                       = "OneDrive for Business (Plan 2)"
    "WACONEDRIVESTANDARD"                         = "OneDrive for Business (Plan 1)"
    "WACSHAREPOINTSTD"                            = "Office Online Standard"
    "WHITEBOARD_PLAN1"                            = "Whiteboard (Plan 1)"
    "WHITEBOARD_PLAN2"                            = "Whiteboard (Plan 2)"
    "WHITEBOARD_PLAN3"                            = "Whiteboard (Plan 3)"
    "WIN10_PRO_ENT_SUB"                           = "Windows 10 Enterprise E3"
    "WINBIZ"                                      = "Windows 10 Business"
    "WINDEFATP"                                   = "Windows Defender ATP"
    "WINDOWS_STORE"                               = "Windows Store for Business"
    "YAMMER_EDU"                                  = "Yammer for Academic"
    "YAMMER_ENTERPRISE"                           = "Yammer Enterprise"
    "YAMMER_MIDSIZE"                              = "Yammer Midsize"
    "COMMUNICATIONS_COMPLIANCE"                   = "Microsoft Communications Compliance"
    "COMMUNICATIONS_DLP"                          = "Microsoft Communications DLP"
    "CUSTOMER_KEY"                                = "Microsoft Customer Key"
    "DATA_INVESTIGATIONS"                         = "Microsoft Data Investigations"
    "INFO_GOVERNANCE"                             = "Microsoft Information Governance"
    "INSIDER_RISK_MANAGEMENT"                     = "Microsoft Insider Risk Management"
    "ML_CLASSIFICATION"                           = "Microsoft ML-based Classification"
    "RECORDS_MANAGEMENT"                          = "Microsoft Records Management"
}


#Connect to Microsoft graph and grab the licence information
# Construct URI
[uri]$uriToken = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
# Construct Body
$body = @{
    client_id     = $appID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
# Get OAuth 2.0 Token
if ($proxyServer) { $tokenRequest = Invoke-WebRequest -Method Post -Uri $uriToken -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -Proxy $proxyHost -ProxyUseDefaultCredentials }
else { $tokenRequest = Invoke-WebRequest -Method Post -Uri $uriToken -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing }
# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
#	Returns the tenant licence information

[uri]$uriLicences = "https://graph.microsoft.com/v1.0/subscribedskus"


#Fetch the information from Office 365 Service Health API
#Get Services: Get the list of subscribed services
$uriError = ""

try {
    if ($proxyServer) {
        [array]$allLicences = @((Invoke-RestMethod -Uri $uriLicences -Headers @{Authorization = "Bearer $Token" } -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).value)
    }
    else {
        [array]$allLicences = @((Invoke-RestMethod -Uri $uriLicences -Headers @{Authorization = "Bearer $Token" } -Method Get).value)
    }
    if ($null -eq $allLicences -or $allLicences.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No licence information retrieved - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>$($allLicences.count) licences retrieved.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No licence information retrieved - verify proxy and network connectivity</p><br/>"
}

if ($uriError -and $emaiLEnabled) {
    $emailSubject = "Error(s) retrieving URL(s)"
    SendEmail $uriError $EmailCreds $config "High" $emailSubject $emailDashAlertsTo
}

$rptO365Info += "<br/>You can add some general information in here if needed.<br />"
$rptO365Info += "ie updates or links to external (ie cloud only) activity to verify Azure AD App is working (ie Flow to Teams Channel)"
if ($altLink) { $rptO365Info += "<a href='$($altLink)' target=_blank> here </a></li></ul><br>" }

#Check office 365 IPs and URLS
#Check before building the page as these will be used for diagnostics checks

#From docs.microsoft.com : https://docs.microsoft.com/en-us/Office365/Enterprise/office-365-ip-web-service
[uri]$ws = "https://endpoints.office.com"
$versionpath = $pathIPURLs + "\O365_endpoints_latestversion-$($rptProfile).txt"
$pathIP4 = $pathIPURLs + "\O365_endpoints_ip4-$($rptProfile).csv"
$pathIP6 = $pathIPURLs + "\O365_endpoints_ip6-$($rptProfile).csv"
$pathIPurl = $pathIPURLs + "\O365_endpoints_urls-$($rptProfile).csv"
$pathIPChanges = $pathIPURLs + "\O365_IPChanges-$($rptProfile).csv"
$pathIPChangeIDX = $pathIPURLs + "\O365_IPChangeIDX-$($rptProfile).csv"
$pathEndpointSetsIDX = $pathIPURLs + "\O365_EndpointSetsIDX-$($rptProfile).csv"
$pacFile1 = "$($pathHTML)\$($pacType1Filename)"
$pacFile2 = "$($pathHTML)\$($pacType2Filename)"


$fileData = "O365_endpoints_data-$($rptProfile).txt"
$pathData = $pathIPURLs + "\" + $fileData
$currentData = $null

if (Test-Path $pathdata) { $currentData = Get-Content $pathData } else { $fileMissing = $true }
if (Test-Path $pathIPurl) { $flatUrls = Import-Csv $pathIPurl } else { $fileMissing = $true }
if (Test-Path $pathIP4) { $flatIp4s = Import-Csv $pathIP4 } else { $fileMissing = $true }
if (Test-Path $pathIP6) { $flatIp6s = Import-Csv $pathIP6 } else { $fileMissing = $true }
if (Test-Path $pathIPChanges) { $flatChanges = Import-Csv $pathIPChanges } else { $fileMissing = $true }
if (Test-Path $pathIPChangeIDX) { $flatChangesIDX = Import-Csv $pathIPChangeIDX } else { $fileMissing = $true }
if (Test-Path $pathEndpointSetsIDX) { $EndPointSetsIDX = Import-Csv $pathEndpointSetsIDX } else { $fileMissing = $true }
if ($PACEnabled) {
    if (Test-Path $pacFile1) { } else { $fileMissing = $true }
    if (Test-Path $pacFile2) { } else { $fileMissing = $true }
}




# fetch client ID and version if version file exists; otherwise create new file and client ID
if (Test-Path $versionpath) {
    $content = Get-Content $versionpath
    $clientRequestId = $content[0]
    $lastVersion = $content[1]
}
else {
    $clientRequestId = [GUID]::NewGuid().Guid
    $lastVersion = "0000000000"
    @($clientRequestId, $lastVersion) | Out-File $versionpath
}

# call version method to check the latest version, and pull new data if version number is different
[uri]$ipurlVersion = "$($ws)/version/Worldwide?clientRequestId=$($clientRequestId)"
if ($proxyServer) { $version = Invoke-RestMethod -Uri ($ipurlVersion) -Proxy $proxyhost -ProxyUseDefaultCredentials }
else { $version = Invoke-RestMethod -Uri ($ipurlVersion) }
$evtMessage = "Downloading IP/URL Versions - this is not rate limited"
Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 701 -EntryType Information
Write-Log $evtMessage

if (($version.latest -gt $lastVersion) -or ($null -like $currentData) -or $fileMissing) {
    $ipurlOutput += "New version of Office 365 worldwide commercial service instance endpoints detected<br />`r`n"
    if ($pacEnabled) {
        #Get Proxy PAC file using MS Script (why re-invent the wheel)
        $pacCreate = "$PSScriptRoot\Get-PacFile.ps1"
        if ($PacProxy) { $paramsPac = @("-Type 1 -clientRequestID $($clientRequestId) -Instance Worldwide -TenantName $($tenantMSName) -DefaultProxySettings $($PACProxy) -FilePath $($pathHTML)\$($pacType1Filename)") }
        else { $paramsPac = @("-Type 1 -clientRequestID $($clientRequestId) -Instance Worldwide -TenantName $($tenantMSName) -FilePath $($pathHTML)\$($pacType1Filename)") }
        $callMe = "'$($pacCreate)' $($paramsPac)"
        Invoke-Expression "& $($callme)"
        if ($PacProxy) { $paramsPac = @("-Type 2 -clientRequestID $($clientRequestId) -Instance Worldwide -TenantName $($tenantMSName) -DefaultProxySettings $($PACProxy) -FilePath $($pathHTML)\$($pacType2Filename)") }
        else { $paramsPac = @("-Type 2 -clientRequestID $($clientRequestId) -Instance Worldwide -TenantName $($tenantMSName) -FilePath $($pathHTML)\$($pacType2Filename)") }
        $callMe = "'$($pacCreate)' $($paramsPac)"
        Invoke-Expression "& $($callme)"
    }
    #Build changes
    [uri]$ipurlChanges = "$($ws)/changes/Worldwide/0000000000?ClientRequestId=$($clientRequestId)"
    if ($proxyServer) { $changes = Invoke-RestMethod -Uri ($ipurlChanges) -Proxy $proxyhost -ProxyUseDefaultCredentials }
    else { $changes = Invoke-RestMethod -Uri ($ipurlChanges) }
    $evtMessage = "Downloading IP/URL changes - this is rate limited to 30 per hour"
    Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 702 -EntryType Information
    Write-Log $evtMessage

    #Flatten the IP/URL Changes
    [array]$allChanges = @()
    $changes | ForEach-Object {
        $change = New-Object PSCustomObject
        $change | Add-Member -MemberType NoteProperty -Name ID -Value $_.ID
        $change | Add-Member -MemberType NoteProperty -Name endpointSetId -Value $_.endpointSetId
        $change | Add-Member -MemberType NoteProperty -Name disposition -Value $_.disposition
        $change | Add-Member -MemberType NoteProperty -Name version -Value $_.version
        $change | Add-Member -MemberType NoteProperty -Name impact -Value $_.impact
        $change | Add-Member -MemberType NoteProperty -Name current -Value $_.current
        $change | Add-Member -MemberType NoteProperty -Name previous -Value $_.previous
        $change | Add-Member -MemberType NoteProperty -Name add -Value $_.add
        $change | Add-Member -MemberType NoteProperty -Name remove -Value $_.remove
        $allChanges += $change
    }
    #Index of changes
    $flatChangesIDX = $changes | ForEach-Object {
        $changeSet = $_
        $idxCustomObjects = [PSCustomObject]@{
            id            = $changeSet.id;
            endpointSetId = $changeSet.endpointSetId;
            Disposition   = $changeSet.disposition;
            version       = $changeSet.version;
            Impact        = $changeSet.Impact;
            current       = $changeSet.current
            previous      = $changeSet.previous
        }
        $idxCustomObjects
    }
    $flatChangesIDX | Export-Csv $pathIPChangeIDX -Encoding UTF8 -NoTypeInformation

    #Adds
    $flatAddChanges = $changes | Where-Object { $_.add -ne $null } | ForEach-Object {
        $changeSet = $_
        $addCustomObjects = [PSCustomObject]@{
            id            = $changeSet.id;
            action        = "Add";
            effectiveDate = [datetime]::parseexact($($changeSet.add.effectiveDate), 'yyyyMMdd', $null).tostring('dd MMM yyyy');
            ips           = $changeSet.add.ips -join ", ";
            urls          = $changeSet.add.urls -join ", ";
        }
        $addCustomObjects
    }
    $flatAddChanges | Export-Csv $pathIPChanges -Encoding UTF8 -NoTypeInformation

    #Removes
    $flatRemoveChanges = $changes | Where-Object { $_.remove -ne $null } | ForEach-Object {
        $changeSet = $_
        $addCustomObjects = [PSCustomObject]@{
            id            = $changeSet.id;
            action        = "Remove";
            effectiveDate = $null
            ips           = $changeSet.remove.ips -join ", ";
            urls          = $changeSet.remove.urls -join ", ";
        }
        $addCustomObjects
    }
    $flatRemoveChanges | Export-Csv $pathIPChanges -Encoding UTF8 -NoTypeInformation -Append
    $flatChanges = [array]$flatAddChanges + $flatRemoveChanges

    # write the new version number to the version file
    @($clientRequestId, $version.latest) | Out-File $versionpath
    # invoke endpoints method to get the new data
    [uri]$ipurlEndpoint = "$($ws)/endpoints/Worldwide?clientRequestId=$($clientRequestId)"
    if ($proxyserver) { $endpointSets = Invoke-RestMethod -Uri ($ipurlEndpoint) -Proxy $proxyHost -ProxyUseDefaultCredentials }
    else { $endpointSets = Invoke-RestMethod -Uri ($ipurlEndpoint) }
    $evtMessage = "Downloading IP/URL endpoints - this is rate limited to 30 per hour"
    Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 703 -EntryType Information
    Write-Log $evtMessage

    # filter results for Allow and Optimize endpoints, and transform these into custom objects with port and category
    # URL results
    $flatUrls = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $urls = $(if ($endpointSet.urls.Count -gt 0) { $endpointSet.urls } else { @() })
        $urlCustomObjects = @()
        $urlCustomObjects = $urls | ForEach-Object {
            [PSCustomObject]@{
                id                     = $endpointSet.id;
                serviceArea            = $endpointSet.ServiceArea;
                serviceAreaDisplayName = $endpointSet.serviceAreaDisplayName;
                category               = $endpointSet.category;
                url                    = $_;
                tcpPorts               = $endpointSet.tcpPorts;
                udpPorts               = $endpointSet.udpPorts;
                notes                  = $endpointSet.notes;
                expressRoute           = $endpointSet.expressRoute;
                required               = $endpointSet.required;
            }
        }
        $urlCustomObjects
    }
    $flatUrls | Export-Csv $pathIPurl -Encoding UTF8 -NoTypeInformation

    # IPv4 results
    $flatIp4s = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $ips = $(if ($endpointSet.ips.Count -gt 0) { $endpointSet.ips } else { @() })
        # IPv4 strings contain dots
        $ip4s = $ips | Where-Object { $_ -like '*.*' }
        $ip4CustomObjects = @()
        if ($endpointSet.category -in ("Allow", "Optimize")) {
            $ip4CustomObjects = $ip4s | ForEach-Object {
                [PSCustomObject]@{
                    category = $endpointSet.category;
                    ip       = $_;
                    tcpPorts = $endpointSet.tcpPorts;
                    udpPorts = $endpointSet.udpPorts;
                }
            }
        }
        $ip4CustomObjects
    }
    $flatIp4s | Export-Csv $pathIP4 -Encoding UTF8 -NoTypeInformation
    # IPv6 results
    $flatIp6s = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $ips = $(if ($endpointSet.ips.Count -gt 0) { $endpointSet.ips } else { @() })
        # IPv6 strings contain colons
        $ip6s = $ips | Where-Object { $_ -like '*:*' }
        $ip6CustomObjects = @()
        if ($endpointSet.category -in ("Optimize")) {
            $ip6CustomObjects = $ip6s | ForEach-Object {
                [PSCustomObject]@{
                    category = $endpointSet.category;
                    ip       = $_;
                    tcpPorts = $endpointSet.tcpPorts;
                    udpPorts = $endpointSet.udpPorts;
                }
            }
        }
        $ip6CustomObjects
    }
    $flatIp6s | Export-Csv $pathIP6 -Encoding UTF8 -NoTypeInformation
    $endpointSets | Select-Object id, servicearea, serviceareadisplayname, category, required | Export-Csv $pathEndpointSetsIDX -Encoding UTF8 -NoTypeInformation
    $EndPointSetsIDX = $endpointSets | Select-Object id, servicearea, serviceareadisplayname, category, required
}
else {
    $ipurlSummary += "Office 365 worldwide commercial service instance endpoints are up-to-date. <br />`r`n"
    $ipurlSummary += "Importing previous results. <br />`r`n"
    $ipurlSummary += "Data available from <a href='https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges' target=_blank>https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges</a><br/>`r`n"
}

$watchCats = $endpointSetsIDX | where-object { $_.category -in ('Allow', 'Optimize') } | Select-Object id, category -Unique

if (($version.latest -gt $lastVersion) -and $emailEnabled) {
    $changesLast1 = $flatchangesIDX | Select-Object version, @{label = "VersionDate"; Expression = { [datetime]::parseexact($_.version, "yyyyMMddHH", $null) } } -Unique | Sort-Object versiondate -Descending | Select-Object -First 1
    #Send email to users on IP/URL change
    $emailSubject = "IPs and URLs: New version $($version.latest)"
    $emailMessage = "new version of Office 365 Worldwide Commercial service instance endpoints"
    $emailMessage += BuildIPURLChanges $changesLast1 $flatchangesIDX $EndPointSetsIDX $flatChanges 1 $watchCats -email
    SendEmail $emailMessage $EmailCreds $config "Normal" $emailSubject $emailIPURLAlerts
}


$flatAddressIPv4 = @($flatIp4s | Where-Object { $_.category -like 'optimize' })
$flatAddressIPv6 = @($flatIp6s | Where-Object { $_.category -like 'optimize' })
$flatAddressURLs = @($flatUrls | Where-Object { $_.category -like 'optimize' })


$changesHTML = $null
$changesHTML = "<p>Full history available <a href='IPURLChangeHistory.html' target=_blank> here </a></p>"
$changesLast5 = $flatchangesIDX | Select-Object version, @{label = "VersionDate"; Expression = { [datetime]::parseexact($_.version, "yyyyMMddHH", $null) } } -Unique | Sort-Object versiondate -Descending | Select-Object -First $($IPURLHistory)
$changesHTML += BuildIPURLChanges $changesLast5 $flatchangesIDX $EndPointSetsIDX $flatChanges 2 $watchCats
$changesHTML += "`r`n</div>"

#Build history list of all changes
$changesAllHTML = $null
$changesAll = $flatchangesIDX | Select-Object version, @{label = "VersionDate"; Expression = { [datetime]::parseexact($_.version, "yyyyMMddHH", $null) } } -Unique | Sort-Object versiondate -Descending
$changesAllHTML = BuildIPURLChanges $changesAll $flatchangesIDX $EndPointSetsIDX $flatChanges 100 $watchCats
ConvertTo-Html -CssUri o365health.css -Body $changesAllHTML -Title "IP and URL Change History" | Out-File -FilePath "$($pathHTML)\IPURLChangeHistory.html" -Encoding UTF8 -Force

if (Test-Path $fileIPURLsNotes) {
    $notesCustom = Import-Csv $fileIPURLsNotes
    #match custom notes data to ID and url
    foreach ($url in $flaturls) {
        $notes = @($notesCustom | Where-Object { $_.ID -like $url.ID -and $_.url -like $url.url })
        $url | Add-Member -MemberType NoteProperty -Name "AmendedURL" -Value $notes.AmendedURL
        $url | Add-Member -MemberType NoteProperty -Name "DirectInternet" -Value $notes.DirectInternet
        $url | Add-Member -MemberType NoteProperty -Name "ProxyAuth" -Value $notes.ProxyAuth
        $url | Add-Member -MemberType NoteProperty -Name "SSLInspection" -Value $notes.SSLInspection
        $url | Add-Member -MemberType NoteProperty -Name "DLP" -Value $notes.DLP
        $url | Add-Member -MemberType NoteProperty -Name "Antivirus" -Value $notes.Antivirus
        $url | Add-Member -MemberType NoteProperty -Name "OurNotes" -Value $notes.Notes
    }
    #Export a full list with additional information
    $flaturls | Export-Csv $fileIPURLsNotesAll -NoTypeInformation -Encoding UTF8
}


# write output to screen
# Clients arent going to want to view this, are they?
$ipurlSummary += "<b>Client Request ID: " + $clientRequestId + "</b><br />`r`n"
$ipurlSummary += "<b>Last Version: " + $lastVersion + "</b><br />`r`n"
$ipurlSummary += "<b>New Version: " + $version.latest + "</b><br />`r`n<br />`r`n"

#Links to PAC Files (if enabled)
if ($PACEnabled) {
    $ipurlOutput += "<b>Example Proxy PAC Files with Optimize and Allow URLS</b><br />`r`n"
    $ipurlOutput += "<b>Optimize URLs Only:</b> <a href='./$($PACType1Filename)' target=_blank>Optimize URLS go Direct</a><br />`r`n"
    $ipurlOutput += "<b>Optimize and Allow URLs:</b> <a href='./$($PACType2Filename)' target=_blank>Optimize and Allow URLS go Direct</a><br />`r`n<br />`r`n"
}

#IPv4
$ipurlOutput += "<b>IPv4 Firewall IP Address Ranges</b><br />`r`n"
$ipurlOutput += "<b>Optimize (Direct connection):</b><br />`r`n"

$ipurlOutput += "$(($flatAddressIPv4.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "<b>Allow:</b><br />`r`n"
$flatAddressIPv4 = @($flatIp4s | Where-Object { $_.category -notlike 'Optimize' })
$ipurlOutput += "$(($flatAddressIPv4.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "All IPv4 networks, TCP/UDP Ports and classifications available to <a href='$(Split-Path $pathIP4 -leaf)' target=_blank>download here</a><br /><br />`r`n"

#IPv6
$ipurlOutput += "<b>IPv6 Firewall IP Address Ranges</b><br />`r`n"
$ipurlOutput += "<b>Optimize (Direct connection):</b><br />`r`n"

$ipurlOutput += "$(($flatAddressIPv6.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "<b>Allow:</b><br />`r`n"
$flatAddressIPv6 = @($flatIp6s | Where-Object { $_.category -notlike 'Optimize' })
$ipurlOutput += "$(($flatAddressIPv6.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "All IPv6 networks, TCP/UDP Ports and classifications available to <a href='$(Split-Path $pathIP6 -leaf)' target=_blank>download here</a><br /><br />`r`n"

#URLs
$ipurlOutput += "<b>URLs</b><br />`r`n"
$ipurlOutput += "<b>Optimize (Direct connection):</b><br />`r`n"
$ipurlOutput += "$(($flatAddressURLs.url | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "<b>Allow:</b><br />`r`n"
$flatAddressURLs = @($flatUrls | Where-Object { $_.category -like 'Allow' })
$ipurlOutput += "$(($flatAddressURLs.url | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "All URLs, TCP/UDP Ports and classifications available to <a href='$(Split-Path $pathIPurl -leaf)' target=_blank>download here</a><br /><br />`r`n"
$ipurlOutput += "Summary information available to <a href='$(Split-Path $pathdata -leaf)' target=_blank>download here</a><br /><br />`r`n"

# write output to data file
Write-Output "Office 365 IP and UL Web Service data" | Out-File $pathData
Write-Output "Worldwide instance" | Out-File $pathData -Append
Write-Output "" | Out-File $pathData -Append
Write-Output ("Version: " + $version.latest) | Out-File $pathData -Append
Write-Output "" | Out-File $pathData -Append
Write-Output "IPv4 Firewall IP Address Ranges" | Out-File $pathData -Append
($flatIp4s.ip | Sort-Object -Unique) -join ", " | Out-File $pathData -Append
Write-Output "" | Out-File $pathData -Append
Write-Output "IPv6 Firewall IP Address Ranges" | Out-File $pathData -Append
($flatIp6s.ip | Sort-Object -Unique) -join ", " | Out-File $pathData -Append
Write-Output "" | Out-File $pathData -Append
Write-Output "URLs for Proxy Server" | Out-File $pathData -Append
($flatUrls.url | Sort-Object -Unique) -join ", " | Out-File $pathData -Append

if (Test-Path $pathdata) { Copy-Item $pathdata -Destination $pathHTML }
if (Test-Path $pathIPurl) { Copy-Item $pathIPurl -Destination $pathHTML }
if (Test-Path $pathIP4) { Copy-Item $pathIP4 -Destination $pathHTML }
if (Test-Path $pathIP6) { Copy-Item $pathIP6 -Destination $pathHTML }
if (Test-Path $fileIPURLsNotes) { Copy-Item $fileIPURLsNotes -Destination $pathHTML }
if (Test-Path $fileIPURLsNotesAll) { Copy-Item $fileIPURLsNotesAll -Destination $pathHTML }
if (Test-Path $fileCustomNotes) { Copy-Item $fileCustomNotes -Destination $pathHTML }

$checkOptHTTP = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*80*' -and $_.category -match 'Optimize') }
$checkOptHTTPs = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*443*' -and $_.category -match 'Optimize') }
$checkAllowHTTP = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*80*' -and $_.category -match 'Allow') }
$checkAllowHTTPs = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*443*' -and $_.category -match 'Allow') }

function checkURL {
    Param(
        [Parameter(Mandatory = $true)] [string]$url,
        [Parameter(Mandatory = $true)] [boolean]$diagVerbose,
        [Parameter(Mandatory = $true)] [boolean]$proxyServer,
        [Parameter(Mandatory = $true)] [string]$proxyHost,
        [Parameter(Mandatory = $true)] [boolean]$urlOptimized
    )
    $intAttempts = 0
    $stopLoop = $false
    do {
        try {
            $intAttempts++
            #Proxy servers should not be used for optimized paths
            if ($ProxyServer) {
                $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -Proxy $proxyHost -ProxyUseDefaultCredentials -UseBasicParsing -TimeoutSec 3
            }
            else {
                $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -UseBasicParsing -TimeoutSec 3
            }
            Switch ($Result.StatusCode) {
                200 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully contacted site $($url).</p><br/>"; break}
                400 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Bad request.</p><br/>"; break}
                401 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Unauthorized.</p><br/>"; break}
                403 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Forbidden.</p><br/>"; break}
                404 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Not found.</p><br/>"; break}
                407 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Proxy authentication required.</p><br/>"; break}
                502 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Bad gateway (likely proxy).</p><br/>"; break}
                503 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Service unavailable (transient, try again).</p><br/>"; break}
                504 { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Gateway timeout (likely proxy).</p><br/>"; break}
                default {
                    $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to contact site  $($url).</p><br/>"
                }
            }
            $stopLoop = $true
        }
        catch {
            if ($intAttempts -ge 3) { $stopLoop = $true }
            else {
                $fault=$error[0].exception
                Switch -Wildcard ($fault) {
    				"*(400)*" { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Connection to site $($url): Bad request [400].</p><br/>"; $stopLoop=$true; break}
                    "*(401)*" { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Connection to site $($url): Unauthorised [401].</p><br/>"; $stopLoop=$true; break}
                    "*(403)*" { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Connection to site $($url): No permissions [403].</p><br/>"; $stopLoop=$true; break}
                    "*(404)*" { $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Connection to site $($url): Not found [404].</p><br/>"; $stopLoop=$true; break}
                    default {
                        $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Attempt $($intAttempts) - Exception: Unable to contact site  $($url).</p><br/>"
                        if ($diagVerbose) {
                            [string]$ErrorMessage = $_
                            $errorMessage = $ErrorMessage.substring(0, [system.math]::Min(250, $ErrorMessage.length))
                            $resultURL += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage).</p><br/>"
                        }
                    }
                }
            }
        }
    }
    while ($stopLoop -eq $false)
    return $resultURL
}
if ($diagURLs -and $diagEnabled) {
    # Microsoft Office 365 URL tests - check the Optimize HTTP connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTP checks for 'Optimize' Sites (Invoke-WebRequest)</p><br/>"
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Direct (Optimized route)</p><br/>"
    foreach ($entry in $checkOptHTTP) {
        $url = "http://$($entry.url)"
        $rptIPURLs += checkURL $url $diagVerbose $false $proxyHost $true
    } # End Foreach URL List
    if ($proxyServer) {
        $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>via Proxy (un-optimized route)</p><br/>"
        foreach ($entry in $checkOptHTTP) {
            $url = "http://$($entry.url)"
            $rptIPURLs += checkURL $url $diagVerbose $proxyServer $proxyHost $true
        } # End Foreach URL List
    }

    # Microsoft Office 365 URL tests - check the Optimize HTTPs connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTPs checks for 'Optimize' Sites (Invoke-WebRequest)</p><br/>"
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Direct (Optimized route)</p><br/>"
    foreach ($entry in $checkOptHTTPs) {
        $url = "https://$($entry.url)"
        $rptIPURLs += checkURL $url $diagVerbose $false $proxyHost $true
    } # End Foreach URL List
    if ($proxyserver) {
        $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>via Proxy (un-optimized route)</p><br/>"
        foreach ($entry in $checkOptHTTPs) {
            $url = "https://$($entry.url)"
            $rptIPURLs += checkURL $url $diagVerbose $proxyServer $proxyHost $true
        } # End Foreach URL List
    }

    # Microsoft Office 365 URL tests - check the Allow HTTP connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTP checks for 'Allow' Sites (Invoke-WebRequest)</p><br/>"
    foreach ($entry in $checkAllowHTTP) {
        $url = "http://$($entry.url)"
        $rptIPURLs += checkURL $url $diagVerbose $proxyServer $proxyHost $false
    } # End Foreach URL List

    # Microsoft Office 365 URL tests - check the Allow HTTPs connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTPs checks for 'Allow' Sites (Invoke-WebRequest)</p><br/>"
    foreach ($entry in $checkAllowHTTPs) {
        $url = "https://$($entry.url)"
        $rptIPURLs += checkURL $url $diagVerbose $proxyServer $proxyHost $false
    } # End Foreach URL List

}

#Start Building the Pages
#Build Div1
if ($miscDiagsEnabled) { $rptWebTests = OnlineEndPoints $diagWeb $diagPorts $diagURLs }
#$rptWebTests = ""

$rptSectionOneOne = "<div class='section'><div class='header'>Office 365 Message Data</div>`n"
$rptSectionOneOne += "<div class='content'>`n"
$rptSectionOneOne += "$($rptO365Info)"
$rptSectionOneOne += "$($toolboxNotes)"
$rptSectionOneOne += "</div></div>`n"
$divOne = $rptSectionOneOne

$rptSectionOneTwo = "<div class='section'><div class='header'>Diagnostics - Microsoft URLs</div>`n"
$rptSectionOneTwo += "<div class='content'>`n"
if ($diagEnabled) {
    $rptSectionOneTwo += "Check connectivity to the current Microsoft 'Optimize' and 'Allow' URLs<br />"
    $rptSectionOneTwo += "Optimize URLs are tested directly and through proxy<br />"
    $rptSectionOneTwo += "$($rptIPURLs)"
}
else {
    $rptSectionOneTwo += "Diagnostics disabled. To show, enable in the configuration file $($configXML): Diagnostics.Enabled = true" 
}
$rptSectionOneTwo += "</div></div>`n"

$divOne += $rptSectionOneTwo

$rptSectionOneThree = "<div class='section'><div class='header'>Diagnostics - Misc URL/Ports</div>`n"
$rptSectionOneThree += "<div class='content'>`n"
if ($miscDiagsEnabled) {
    $rptSectionOneThree += "$($rptWebTests)" 
}
else { 
    $rptSectionOneThree += "Diagnostics disabled. To show, enable in the configuration file $($configXML): MiscDiagnostics.Enabled = true" 
}
$rptSectionOneThree += "</div></div>`n"

$divOne += $rptSectionOneThree

#Build Div2
$rptSectionTwoOne = "<div class='section'><div class='header'>Licences</div>`n"
$rptSectionTwoOne += "<div class='content'>`n"
$rptLicenceDash = "<div class='container'>`n"
foreach ($sku in $allLicences) {
    [string]$cardDetail = ""
    [string]$cardClass = ""
    [string]$NicePartNumber = $null
    $NicePartNumber = ($skunames.GetEnumerator() | Where-Object { $_.name -like "$($sku.skupartnumber)" }).Value
    if ($NicePartNumber -eq "") { $NicePartNumber = $($sku.SkuPartNumber) }
    $NicePartNumber += "<br/><span class=tooltiptext'> $($sku.consumedUnits) / $(($sku.prepaidunits).enabled +($sku.prepaidunits).warning) assigned"
    if (($sku.prepaidunits).warning -gt 0) { $NicePartNumber += "<br/>$(($sku.prepaidunits).warning) in warning state" }
    if (($sku.prepaidunits).suspended -gt 0) { $NicePartNumber += "<br/>$(($sku.prepaidunits).suspended) in suspended state" }
    $NicePartNumber += "</span>"
    [int]$intPlanCount = 0
    $sps = $sku.ServicePlans | Sort-Object servicePlanName
    foreach ($serviceplan in $sps) {
        [string]$NiceServiceName = $null
        $NiceServiceName = $SkuNames.item($serviceplan.serviceplanname)
        if ($NiceServiceName -eq "") { $NiceServiceName = $($serviceplan.serviceplanname) }
        $cardClass = get-statusdisplay $($serviceplan.provisioningStatus) 'ServicePlanStatus'
        $cardDetail += "<div class='sku-item-$($cardClass)'>$($NiceServiceName)<span class='tooltiptext'>$($serviceplan.provisioningStatus)</span></div>`r`n"
        if (($serviceplan.serviceplanname).length -gt 32) { $intPlanCount += 2 } else { $intPlanCount++ }
    }
    $cardClass = Get-StatusDisplay $($sku.CapabilityStatus) 'SkuCapabilityStatus'
    $cardText = SkuCardBuilder $($NicePartNumber) $cardDetail $cardClass $intPlanCount
    $rptLicenceDash += "$cardText`n"
}
$rptLicenceDash += "</div>`n<br/><br/>"
$rptSectionTwoOne += $rptLicenceDash
$rptSectionTwoOne += "</div></div>`r`n"
$divTwo = $rptSectionTwoOne

#Build Div3
#Retrieve latest Office 365 service instances

$rptSectionThreeOne = "<div class='section'><div class='header'>Versions Information</div>`n"
$rptSectionThreeOne += "<div class='content'>`n"
[string]$ipurlVersion = "<b>IP and URL Version information</b><br />"
$rptSectionThreeOne += $ipurlVersion
$rptSectionThreeOne += $ipurlSummary
$rptSectionThreeOne += "</div></div>`n"

$divThree = $rptSectionThreeOne

$rptSectionThreeTwo = "<div class='section'><div class='header'>Current IP and URL Information</div>`n"
$rptSectionThreeTwo += "<div class='content'>`n"
[string]$ipurlCurrent = "<b>Current IP and URL information</b><br />"
$rptSectionThreeTwo += $ipurlCurrent
$rptSectionThreeTwo += $ipurlOutput
$rptSectionThreeTwo += "</div></div>`n"

$divThree += $rptSectionThreeTwo

$rptSectionThreeThree = "<div class='section'><div class='header'>IP and URL History</div>`n"
$rptSectionThreeThree += "<div class='content'>`n"
$rptSectionThreeThree += $changesHTML
$rptSectionThreeThree += "</div></div>`n"

$divThree += $rptSectionThreeThree

#Build Div4
$rptURLTable = ""
if ($null -ne $fileIPURLsNotes) { if (Test-Path $($fileIPURLsNotes)) { $rptURLTable += "Download URL notes from <a href='$($fileIPURLsNotes)' target=_blank>here</a><br />" } }
if ($null -ne $fileIPURLsNotesAll) { if (Test-Path $($fileIPURLsNotesAll)) { $rptURLTable += "Download combined URLs and notes from <a href='$($fileIPURLsNotesAll)' target=_blank>here</a><br />" } }
if ($null -ne $fileCustomNotes) {
    if (Test-Path $($fileCustomNotes)) {
        $rptURLTable += "Download custom URLs and notes from <a href='$($fileCustomNotes)' target=_blank>here</a><br />";
        $customURLs = Import-Csv $fileCustomNotes
    }
}


$rptSectionFourOne = $rptURLTable
$rptSectionFourOne += "<div class='section'><div class='header'>Optimize URLs</div>`n"
$rptSectionFourOne += "<div class='content'>`n"
$rptURLTable = "<div class='tableInc'>"
$rptURLTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-l'>ID</div>`n`t<div class='tableInc-header-l'>serviceArea</div>`n`t<div class='tableInc-header-l'>url</div>`n`t<div class='tableInc-header-l'>tcpPorts</div>`n`t<div class='tableInc-header-l'>udpPorts</div>`n`t<div class='tableInc-header-l'>notes</div>`n`t<div class='tableInc-header-c'>required</div>`n"
if (Test-Path $($fileIPURLsNotes)) { $rptURLTable += "<div class='tableInc-header-l'>AmendedURL</div>`n`t<div class='tableInc-header-c'>DirectInternet</div>`n`t<div class='tableInc-header-c'>ProxyAuth</div>`n`t<div class='tableInc-header-c'>SSLInspection</div>`n`t<div class='tableInc-header-c'>DLP</div>`n`t<div class='tableInc-header-c'>Antivirus</div>`n`t<div class='tableInc-header-l'>ourNotes</div>`n" }
$rptURLTable += "</div>`n"

[array]$urlList = @()
$urlList = $flatUrls | Where-Object { $_.category -like 'optimize' }
foreach ($entry in $urlList) {
    $rptURLTable += "<div class='tableInc-row'>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.id)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.serviceArea)</div>`n`t"
    #$rptURLTable+="<td>$($entry.serviceAreaDisplayName)</td>"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.url)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.tcpPorts)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.udpPorts)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.notes)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-c'>$($entry.required)</div>`n`t"
    if (Test-Path $($fileIPURLsNotes)) {
        $e1, $e2, $e3, $e4, $e5 = $null
        if ($entry.DirectInternet -in 'yes', 'true') { $e1 = "True" } elseif ($entry.DirectInternet -in 'no', 'false') { $e1 = "False" }
        if ($entry.ProxyAuth -in 'yes', 'true') { $e2 = "True" } elseif ($entry.ProxyAuth -in 'no', 'false') { $e2 = "False" }
        if ($entry.SSLInspection -in 'yes', 'true') { $e3 = "True" } elseif ($entry.SSLInspection -in 'no', 'false') { $e3 = "False" }
        if ($entry.DLP -in 'yes', 'true') { $e4 = "True" } elseif ($entry.DLP -in 'no', 'false') { $e4 = "False" }
        if ($entry.Antivirus -in 'yes', 'true') { $e5 = "True" } elseif ($entry.Antivirus -in 'no', 'false') { $e5 = "False" }
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.AmendedURL)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e1)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e2)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e3)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e4)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e5)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.ourNotes)</div>`n"
    }
    $rptURLTable += "</div>`n"
}
$rptURLTable += "</div>`n"
$rptSectionFourOne += $rptURLTable
$rptSectionFourOne += "</div></div>`n"

$divFour = $rptSectionFourOne

#Custom URLS - our URLs particular to our deployment, and documented here
if ($customURLs) {
    $rptSectionFourTwo += "<div class='section'><div class='header'>Additional URLs associated with O365 deployment</div>`n"
    $rptSectionFourTwo += "<div class='content'>`n"
    $rptURLTable = "<div class='tableInc'>"
    $rptURLTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-l'>ID</div>`n`t<div class='tableInc-header-l'>serviceArea</div>`n`t<div class='tableInc-header-l'>url</div>`n`t<div class='tableInc-header-l'>tcpPorts</div>`n`t<div class='tableInc-header-l'>udpPorts</div>`n`t"
    $rptURLTable += "<div class='tableInc-header-c'>DirectInternet</div>`n`t<div class='tableInc-header-c'>ProxyAuth</div>`n`t<div class='tableInc-header-c'>SSLInspection</div>`n`t<div class='tableInc-header-c'>DLP</div>`n`t<div class='tableInc-header-c'>Antivirus</div>`n`t<div class='tableInc-header-l'>Notes</div>`n"
    $rptURLTable += "</div>`n"
    [array]$urlList = @()
    $urlList = $customURLs | Where-Object { $_.ID -gt 0 }
    foreach ($entry in $urlList) {
        $e1, $e2, $e3, $e4, $e5 = $null
        $rptURLTable += "<div class='tableInc-row'>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.id)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.serviceArea)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.url)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.tcpPorts)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.udpPorts)</div>`n`t"
        if ($entry.DirectInternet -in 'yes', 'true') { $e1 = "True" } elseif ($entry.DirectInternet -in 'no', 'false') { $e1 = "False" }
        if ($entry.ProxyAuth -in 'yes', 'true') { $e2 = "True" } elseif ($entry.ProxyAuth -in 'no', 'false') { $e2 = "False" }
        if ($entry.SSLInspection -in 'yes', 'true') { $e3 = "True" } elseif ($entry.SSLInspection -in 'no', 'false') { $e3 = "False" }
        if ($entry.DLP -in 'yes', 'true') { $e4 = "True" } elseif ($entry.DLP -in 'no', 'false') { $e4 = "False" }
        if ($entry.Antivirus -in 'yes', 'true') { $e5 = "True" } elseif ($entry.Antivirus -in 'no', 'false') { $e5 = "False" }
        $rptURLTable += "<div class='tableInc-cell-c'>$($e1)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e2)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e3)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e4)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e5)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.Notes)</div>`n"
        $rptURLTable += "</div>`n"
    }
    $rptURLTable += "</div>`n"
    $rptSectionFourTwo += $rptURLTable
    $rptSectionFourTwo += "</div></div>`n"
}
$divFour += $rptSectionFourTwo

#Microsoft specified Allow URLs
$rptSectionFourThree += "<div class='section'><div class='header'>Allow URLs</div>`n"
$rptSectionFourThree += "<div class='content'>`n"
$rptURLTable = "<div class='tableInc'>"
$rptURLTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-l'>ID</div>`n`t<div class='tableInc-header-l'>serviceArea</div>`n`t<div class='tableInc-header-l'>url</div>`n`t<div class='tableInc-header-l'>tcpPorts</div>`n`t<div class='tableInc-header-l'>udpPorts</div>`n`t<div class='tableInc-header-l'>notes</div>`n`t<div class='tableInc-header-c'>required</div>`n"
if (Test-Path $($fileIPURLsNotes)) { $rptURLTable += "<div class='tableInc-header-l'>AmendedURL</div>`n`t<div class='tableInc-header-c'>DirectInternet</div>`n`t<div class='tableInc-header-c'>ProxyAuth</div>`n`t<div class='tableInc-header-c'>SSLInspection</div>`n`t<div class='tableInc-header-c'>DLP</div>`n`t<div class='tableInc-header-c'>Antivirus</div>`n`t<div class='tableInc-header-l'>ourNotes</div>`n" }
$rptURLTable += "</div>`n"

[array]$urlList = @()
$urlList = $flatUrls | Where-Object { $_.category -like 'allow' }
foreach ($entry in $urlList) {
    $rptURLTable += "<div class='tableInc-row'>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.id)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.serviceArea)</div>`n`t"
    #$rptURLTable+="<td>$($entry.serviceAreaDisplayName)</td>"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.url)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.tcpPorts)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.udpPorts)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.notes)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-c'>$($entry.required)</div>`n`t"
    if (Test-Path $($fileIPURLsNotes)) {
        $e1, $e2, $e3, $e4, $e5 = $null
        if ($entry.DirectInternet -in 'yes', 'true') { $e1 = "True" } elseif ($entry.DirectInternet -in 'no', 'false') { $e1 = "False" }
        if ($entry.ProxyAuth -in 'yes', 'true') { $e2 = "True" } elseif ($entry.ProxyAuth -in 'no', 'false') { $e2 = "False" }
        if ($entry.SSLInspection -in 'yes', 'true') { $e3 = "True" } elseif ($entry.SSLInspection -in 'no', 'false') { $e3 = "False" }
        if ($entry.DLP -in 'yes', 'true') { $e4 = "True" } elseif ($entry.DLP -in 'no', 'false') { $e4 = "False" }
        if ($entry.Antivirus -in 'yes', 'true') { $e5 = "True" } elseif ($entry.Antivirus -in 'no', 'false') { $e5 = "False" }
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.AmendedURL)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e1)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e2)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e3)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e4)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e5)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.ourNotes)</div>`n"
    }
    $rptURLTable += "</div>`n"
}
$rptURLTable += "</div>`n"
$rptSectionFourThree += $rptURLTable
$rptSectionFourThree += "</div></div>`n"

$divFour += $rptSectionFourThree

#Microsoft specified Default URLs
$rptSectionFourFour += "<div class='section'><div class='header'>Default URLs</div>`n"
$rptSectionFourFour += "<div class='content'>`n"
$rptURLTable = "<div class='tableInc'>"
$rptURLTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-l'>ID</div>`n`t<div class='tableInc-header-l'>serviceArea</div>`n`t<div class='tableInc-header-l'>url</div>`n`t<div class='tableInc-header-l'>tcpPorts</div>`n`t<div class='tableInc-header-l'>udpPorts</div>`n`t<div class='tableInc-header-l'>notes</div>`n`t<div class='tableInc-header-c'>required</div>`n"
if (Test-Path $($fileIPURLsNotes)) { $rptURLTable += "<div class='tableInc-header-l'>AmendedURL</div>`n`t<div class='tableInc-header-c'>DirectInternet</div>`n`t<div class='tableInc-header-c'>ProxyAuth</div>`n`t<div class='tableInc-header-c'>SSLInspection</div>`n`t<div class='tableInc-header-c'>DLP</div>`n`t<div class='tableInc-header-c'>Antivirus</div>`n`t<div class='tableInc-header-l'>ourNotes</div>`n" }
$rptURLTable += "</div>`n"

[array]$urlList = @()
$urlList = $flatUrls | Where-Object { $_.category -like 'default' }
foreach ($entry in $urlList) {
    $rptURLTable += "<div class='tableInc-row'>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.id)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.serviceArea)</div>`n`t"
    #$rptURLTable+="<td>$($entry.serviceAreaDisplayName)</td>"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.url)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.tcpPorts)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.udpPorts)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-l'>$($entry.notes)</div>`n`t"
    $rptURLTable += "<div class='tableInc-cell-c'>$($entry.required)</div>`n`t"
    if (Test-Path $($fileIPURLsNotes)) {
        $e1, $e2, $e3, $e4, $e5 = $null
        if ($entry.DirectInternet -in 'yes', 'true') { $e1 = "True" } elseif ($entry.DirectInternet -in 'no', 'false') { $e1 = "False" }
        if ($entry.ProxyAuth -in 'yes', 'true') { $e2 = "True" } elseif ($entry.ProxyAuth -in 'no', 'false') { $e2 = "False" }
        if ($entry.SSLInspection -in 'yes', 'true') { $e3 = "True" } elseif ($entry.SSLInspection -in 'no', 'false') { $e3 = "False" }
        if ($entry.DLP -in 'yes', 'true') { $e4 = "True" } elseif ($entry.DLP -in 'no', 'false') { $e4 = "False" }
        if ($entry.Antivirus -in 'yes', 'true') { $e5 = "True" } elseif ($entry.Antivirus -in 'no', 'false') { $e5 = "False" }
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.AmendedURL)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e1)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e2)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e3)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e4)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-c'>$($e5)</div>`n`t"
        $rptURLTable += "<div class='tableInc-cell-l'>$($entry.ourNotes)</div>`n"
    }
    $rptURLTable += "</div>`n"
}
$rptURLTable += "</div>`n"
$rptSectionFourFour += $rptURLTable
$rptSectionFourFour += "</div></div>`n"

$divFour += $rptSectionFourFour


#CNAME checks where possible
$divFive = "<!- Start of section five ->"
$rptSectionFiveOne = ""
if ($cnameenabled) {
    [array]$CNAMEResults = $null
    [array]$cnames = $null

    #Import the DNS results being output by the monitor script
    foreach ($dns in $cnameresolvers) {
        try {$dnsResults = Import-Csv "$($pathWorking)\$cnameFilename-$($DNS)-$($rptProfile).csv"}
        catch {
        $evtMessage = "[CNAMES] Cannot import DNS results file $($pathWorking)\$cnameFilename-$($DNS)-$($rptProfile).csv`r`n Check file has been created by monitoring script."
        Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 601 -EntryType Error
        Write-Log $evtMessage
}
        $CNAMEResults += $dnsResults
    }

    $cnames = $CNameResults

    #Build Div5
    $rptSectionFiveOne = "<div class='section'><div class='header'>Information</div>`n"
    $rptSectionFiveOne += "<div class='content'>`n"
    $rptSectionFiveOne += "$($cnameNotes)"
    $rptSectionFiveOne += "</div></div>`n"
    $divFive += $rptSectionFiveOne

    $rptSectionFiveTwo = "<div class='section'><div class='header'>CNAMEs</div>`r`n"
    $rptSectionFiveTwo += "<div class='content'>`r`n"
    #Get CNAMEs
    #For each unique monitor
    foreach ($url in $cnameURLs) {
        $rptCNAMEInfo += "<div class='section'>`r`n"
        $rptCNAMEInfo += "<div class='header'>$($url)</div>`r`n"
        $rptCNAMEInfo += "<div class='tableCname-header'>`r`n"
        $rptCNAMEInfo += "<div class='tableCname-header-hn'>CNAME Host</div>`r`n"
        $rptCNAMEInfo += "<div class='tableCname-header-dom'>Domain</div>`r`n"
        #Now build headers for each of the resolving servers
        foreach ($dns in $cnameresolvers) {
            $dnsServerDesc = $cnameresolverdesc[[array]::indexof($cnameResolvers, $DNS)]
            $rptCNAMEInfo += "<div class='tableCname-header-dtf'>$($dnsServerDesc)<br/>$($dns)<br/>First Seen</div>`r`n"
            $rptCNAMEInfo += "<div class='tableCname-header-dtl'><br/><br/>Last Seen</div>"
        }
        $rptCNAMEInfo += "</div>`r`n"
        $cnameslist = $cnames | Where-Object { $_.monitor -like $url } | Select-Object -Unique namehost, domain
        foreach ($cname in  $cnameslist) {
            $rptCNAMEInfo += "<div class='tableCname-row'>`r`n"
            $rptCNAMEInfo += "<div class='tableCname-cell-hn'>$($cname.namehost)</div>"
            $rptCNAMEInfo += "<div class='tableCname-cell-dom'>$($cname.domain)</div>`r`n"
            foreach ($dns in $cnameresolvers) {
                $spotted = $cnames | Where-Object { $_.resolver -like $dns -and $_.monitor -like $url -and $_.namehost -like $cname.namehost }
                $addedDate = ""
                $lastDate = ""
                if ($spotted.addedDate) {
                    if ((Get-Date $spotted.addedDate) -lt ((Get-Date).addhours(-48))) { $fontcolour = "<p>" }
                    elseif ((Get-Date $spotted.addedDate) -lt ((Get-Date).addhours(-24))) { $fontcolour = "<p class='recentCname'>" }
                    else { $fontcolour = "<p class='newCname'>" }
                    $addedDate = "$($fontcolour)$(Get-Date $spotted.addedDate -Format 'dd-MMM-yy HH:mm')</p>"
                }
                else { $addedDate = "<p class='error'>n/a</p>" }
                $rptCNAMEInfo += "<div class='tableCname-cell-dtf'>$($addedDate)</div>`r`n"
                if ($spotted.lastdate) {
                    if ((Get-Date $spotted.lastdate) -lt ((Get-Date).addhours(-12))) { $fontcolour = "<p class='error'>" }
                    elseif ((Get-Date $spotted.lastdate) -lt ((Get-Date).addhours(-4))) { $fontcolour = "<p class='warning'>" }
                    else { $fontcolour = "<p class='ok'>" }
                    $lastDate = "$($fontcolour)$(Get-Date $spotted.lastDate -Format 'dd-MMM-yy HH:mm')</p>"
                }
                else { $lastDate = "<p class='error'>n/a</p>" }
                $rptCNAMEInfo += "<div class='tableCname-cell-dtl'>$($lastDate)</div>`r`n"
            }
            $rptCNAMEInfo += "</div>`r`n"
        }
        $rptCNAMEInfo += "</div>`r`n"
        #$rptCNAMEInfo += "<div><br/></div>`r`n"
    }
    $rptCNAMEInfo += "</div>`r`n"
    #get name host and added date


    $rptSectionFiveTwo += $rptCNAMEInfo
    $rptSectionFiveTwo += "</div></div>`n"
}
$divFive += $rptSectionFiveTwo

#Build Last
$rptSectionLastOne = "<div class='section'><div class='header'>Logs</div>`n"
$rptSectionLastOne += "<div class='content'>`n"
$rptSectionLastOne += $rptO365Info
$rptSectionLastOne += "</div></div>`n"

$divLast = $rptSectionLastOne

$rptHTMLName = $HTMLFile.Replace(" ", "")
$rptTitle = $rptTenantName + " " + $rptName
if ($rptOutage) { $rptTitle += " Outage detected" }
$evtMessage = "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] Tenant: $($rptProfile) - Generating HTML to '$($pathHTML)\$($rptHTMLName)'`r`n"
$evtLogMessage += $evtMessage
Write-Verbose $evtMessage

BuildHTML $rptTitle $divOne $divTwo $divThree $divFour $divFive $divLast $swScript.Elapsed $rptHTMLName
#Check if .css file exists in HTML file destination
if (!(Test-Path "$($pathHTML)\$($cssfile)")) {
    Write-Log "Copying O365Health.css to directory $($pathHTML)"
    Copy-Item ".\O365Health.css" -Destination "$($pathHTML)"
}

$swScript.Stop()

$evtMessage = "Tenant: $($rptProfile) - Script runtime $($swScript.Elapsed.Minutes)m:$($swScript.Elapsed.Seconds)s:$($swScript.Elapsed.Milliseconds)ms on $env:COMPUTERNAME`r`n"
$evtMessage += "*** Processing finished ***`r`n"
Write-Log $evtMessage
#Re-instate default proxy
[System.Net.GlobalProxySelection]::Select = $defaultProxy

#Append to daily log file.
Get-Content $script:logfile | Add-Content $script:Dailylogfile
Remove-Item $script:logfile
Remove-Module O365ServiceHealth