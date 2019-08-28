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
[string]$evtSource = $config.MonitorEvtSource
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
[string]$appID = $config.AppID
[string]$clientSecret = $config.AppSecret
[string]$SMTPUser = $config.EmailUser
[string]$SMTPPassword = $config.EmailPassword
[string]$SMTPKey = $config.EmailKey

[string]$addLink = $config.DashboardAddLink
[string]$rptName = $config.DashboardName
[int]$pageRefresh = $config.DashboardRefresh
[int]$IncidentDays = $config.DashboardHistory
#Dashboard cards
[string[]]$dashCards = $config.DashboardCards.split(",")
$dashCards = $dashCards.Replace('"', '')
$dashCards = $dashCards.Trim()

#Prefered cards for features
[string[]]$prefDashCards = $config.WallDashCards.split(",")
$prefDashCards = $prefDashCards.Replace('"', '')
$prefDashCards = $prefDashCards.Trim()

[string]$rptProfile = $config.TenantShortName
[string]$rptTenantName = $config.TenantName

[string]$pathLogs = $config.LogPath
[string]$pathHTML = $config.HTMLPath
[string]$HTMLFile = $config.DashboardHTML
[string]$pathIPURLs = $config.IPURLPath

[string[]]$emailIPURLAlerts = $config.IPURLAlertsTo


[string]$diagNotes = $config.DiagnosticsNotes

[int]$maxFeedItems = $config.MaxFeedItems

[boolean]$rptOutage = $false

[string]$cssfile = ".\O365Health.css"

# Get Email credentials
# Check for a username. No username, no need for credentials (internal mail host?)
[PSCredential]$emailCreds = $null
if ($smtpuser -notlike '') {
    #Email credentials have been specified, so build the credentials.
    #See readme on how to build credentials files
    $EmailCreds = getCreds $SMTPUser $SMTPPassword $SMTPKey
}


#If no path has been specified, use the current script location
if (!$pathLogs) {
    $pathLogs = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
}
if (!$pathIPURLs) {
    $pathIPURLs = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
}


#Check and trim the report path
$pathLogs = $pathLogs.TrimEnd("\")
$pathHTML = $pathHTML.TrimEnd("\")
$pathIPURLs = $pathIPURLs.TrimEnd("\")

#Build and Check output directories
#Base for logs
if (!(Test-Path $($pathLogs))) {
    New-Item -ItemType Directory -Path $pathLogs
}
#HTML directory for output main page
if (!(Test-Path $($pathHTML))) {
    New-Item -ItemType Directory -Path "$($pathHTML)"
}
#Document directory for messages and article documents
$pathHTMLDocs = "$($pathHTML)\Docs"
if (!(Test-Path $($pathHTMLDocs))) {
    New-Item -ItemType Directory -Path "$($pathHTMLDocs)"
}
#IP and URL file storage
if (!(Test-Path $($pathIPURLs))) {
    New-Item -ItemType Directory -Path $pathIPURLs
}

if ([system.IO.path]::IsPathRooted($pathLogs) -eq $false) {
    #its not an absolute path. Find the absolute path
    $pathLogs = Resolve-Path $pathLogs
}

# setup the logfile
# If logfile exists, the set flag to keep logfile
$script:DailyLogFile = "$($pathLogs)\O365Dashboard-$($rptprofile)-$(Get-Date -format yyMMdd).log"
$script:LogFile = "$($pathLogs)\tmpO365Dashboard-$($rptprofile)-$(Get-Date -format yyMMddHHmmss).log"
$script:LogInitialized = $false
$script:FileHeader = "*** Application Information ***"

$evtMessage = "Config File: $($configXML)"
Write-Log $evtMessage
$evtMessage = "Log Path: $($pathLogs)"
Write-Log $evtMessage
$evtMessage = "HTML Output: $($pathHTML)"
Write-Log $evtMessage

#Create event logs if set
if ($UseEventLog) {
    $evtCheck = Get-EventLog -List -ErrorAction SilentlyContinue | Where-Object { $_.LogDisplayName -eq $evtLogname }
    if (!($evtCheck)) {
        New-EventLog -LogName $evtLogname -Source $evtSource
        Write-EventLog -LogName $evtLogname -Source $evtSource -Message "Event log created." -EventId 1 -EntryType Information
    }
}

#Proxy Configuration
if ($config.UseProxy -like 'true') {
    [boolean]$ProxyServer = $true
    $evtMessage = "Using proxy server $($proxyHost) for connectivity"
    Write-Log $evtMessage
}
else {
    [boolean]$ProxyServer = $false
    $evtMessage = "No proxy to be used."
    Write-Log $evtMessage
}

#Check diagnostics and save as boolean
if ($config.DiagnosticsWeb -like 'true') { [boolean]$diagWeb = $true } else { [boolean]$diagWeb = $false }
if ($config.DiagnosticsPorts -like 'true') { [boolean]$diagPorts = $true } else { [boolean]$diagPorts = $false }
if ($config.DiagnosticsURLs -like 'true') { [boolean]$diagURLs = $true } else { [boolean]$diagURLs = $false }

#Connect to Azure app and grab the service status
ConnectAzureAD
[uri]$urlOrca = "https://manage.office.com"
[uri]$authority = "https://login.microsoftonline.com/$TenantID"
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
$clientCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $appId, $clientSecret
$authenticationResult = ($authContext.AcquireTokenAsync($urlOrca, $clientCredential)).Result;
$bearerToken = $authenticationResult.AccessToken
if ($null -eq $authenticationResult) {
    $evtMessage = "ERROR - No authentication result for Auzre AD App"
    Write-EventLog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 1 -EntryType Error
    Write-Log $evtMessage
}

$authHeader = @{
    'Content-Type'  = 'application/json'
    'Authorization' = "Bearer " + $bearerToken
}

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
        [Parameter(Mandatory = $true)] $contentSix,
        [Parameter(Mandatory = $true)] $contentSeven,
        [Parameter(Mandatory = $true)] $contentEight,
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
    <button class="tablinks" onclick="openTab(event,'Overview')" id="defaultOpen">Overview</button>
    <button class="tablinks" onclick="openTab(event,'Features')">Features</button>
    <button class="tablinks" onclick="openTab(event,'Incidents')">Incidents</button>
    <button class="tablinks" onclick="openTab(event,'Advisories')">Advisories</button>
    <button class="tablinks" onclick="openTab(event,'Licences')">Licences</button>
    <button class="tablinks" onclick="openTab(event,'Diagnostics')">Diagnostics</button>
    <button class="tablinks" onclick="openTab(event,'IPsandURLs')">IPs and URLs</button>
    <button class="tablinks" onclick="openTab(event,'Roadmap')">Office 365 Roadmap</button>
</div>

<!-- Tab content -->
<div id="Overview" class="tabcontent">
    $($contentOne)
</div>

<div id="Features" class="tabcontent">
    $($contentTwo)
</div>

<div id="Incidents" class="tabcontent">
    $($contentThree)
</div>

<div id="Advisories" class="tabcontent">
    $($contentFour)
</div>

<div id="Licences" class="tabcontent">
    $($contentFive)
</div>

<div id="Diagnostics" class="tabcontent">
    $($contentSix)
</div>

<div id="IPsandURLs" class="tabcontent">
    $($contentSeven)
</div>

<div id="Roadmap" class="tabcontent">
    $($contentEight)
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
            try {
                if ($ProxyServer) {
                    $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -Proxy $proxyHost -ProxyUseDefaultCredentials
                }
                else {
                    $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue
                }
                Switch ($Result.StatusCode) {
                    200 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully obtained CRL from $($url).</p><br/>" }
                    400 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to obtain CRL from $($url): Bad request.</p><br/>" }
                    401 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Unauthorized.</p><br/>" }
                    403 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Forbidden.</p><br/>" }
                    404 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Not found.</p><br/>" }
                    407 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Proxy authentication required.</p><br/>" }
                    502 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Bad gateway (likely proxy).</p><br/>" }
                    503 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Service unavailable (transient, try again).</p><br/>" }
                    504 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to obtain CRL from $($url): Gateway timeout (likely proxy).</p><br/>" }
                    default {
                        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to obtain CRL from $($url).</p><br/>"
                    }
                }
            }
            catch {
                $ErrorMessage = $_
                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Exception: Unable to obtain CRL from $($url).</p><br/>"
                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage).</p><br/>"
            }
        } # End Foreach CRL

        # Required Resources Endpoints tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Required Resources Endpoints (Invoke-WebRequest).</p><br/>"
        foreach ($url in $RequiredResourcesEndpoints) {
            try {
                if ($ProxyServer) {
                    $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -Proxy $proxyHost -ProxyUseDefaultCredentials
                }
                else {
                    $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue
                }
                Switch ($Result.StatusCode) {
                    200 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully connected to $($url).</p><br/>" }
                    400 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact $($url): Bad request.</p><br/>" }
                    401 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Unauthorized.</p><br/>" }
                    403 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Forbidden.</p><br/>" }
                    404 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Not found.</p><br/>" }
                    407 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Proxy authentication required.</p><br/>" }
                    502 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Bad gateway (likely proxy).</p><br/>" }
                    503 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Service unavailable (transient, try again).</p><br/>" }
                    504 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Gateway timeout (likely proxy).</p><br/>" }
                    default {
                        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url).</p><br/>"
                    }
                }
            }
            catch {
                $ErrorMessage = $_
                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Exception: Failed to contact $($url) [$($ip4)]:443</p><br/>"
                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>$($ErrorMessage)</p><br/>"
			
            }
        } # End Foreach RequiredResourcesEndpoints
	
        # Optional Resources Endpoints tests
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Optional Resources Endpoints (Invoke-WebRequest).</p><br/>"
        foreach ($url in $OptionalResourcesEndpoints) {
            try {
                if ($ProxyServer) {
                    $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -Proxy $proxyHost -ProxyUseDefaultCredentials
                }
                else {
                    $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue
                }
                Switch ($Result.StatusCode) {
                    200 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully connected to $($url).</p><br/>" }
                    400 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact $($url): Bad request.</p><br/>" }
                    401 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Unauthorized.</p><br/>" }
                    403 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Forbidden.</p><br/>" }
                    404 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Not found.</p><br/>" }
                    407 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Proxy authentication required.</p><br/>" }
                    502 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Bad gateway (likely proxy).</p><br/>" }
                    503 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Service unavailable (transient, try again).</p><br/>" }
                    504 { $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url): Gateway timeout (likely proxy).</p><br/>" }
                    default {
                        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact $($url).</p><br/>"
                    }
                }
            }
            catch {
                $ErrorMessage = $_
                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Exception: Failed to contact $($url) [$($ip4)]:443</p><br/>"
                $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>$($ErrorMessage)</p><br/>"
			
            }
        } # End Foreach RequiredResourcesEndpoints


    } #End web tests

    # Required Resource tests
    if ($diagPorts) {
        $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Testing Required Resources (TCP:443) DNS Resolution may fail from clients.</p><br/>"
        foreach ($url in $RequiredResources) {
            try { [array]$ResourceAddresses = (Resolve-DnsName $url -ErrorAction stop).IP4Address }
            catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to resolve host URL $($url).</p><br/>"; Continue }
            foreach ($ip4 in $ResourceAddresses) {
                try {
                    $Result = Test-NetConnection $ip4 -Port 443 -ea stop -wa silentlycontinue
                    switch ($Result.TcpTestSucceeded) {
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
            try { [array]$ResourceAddresses = (Resolve-DnsName $url -ErrorAction stop).IP4Address }
            catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Unable to resolve host URL $($url).</p><br/>"; Continue }
		
            foreach ($ip4 in $ResourceAddresses) {
                try {
                    $Result = Test-NetConnection $ip4 -Port 443 -ea stop -wa silentlycontinue
                    switch ($Result.TcpTestSucceeded) {
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
            try { [array]$ResourceAddresses = (Resolve-DnsName $url -ErrorAction stop).IP4Address }
            catch { $ErrorMesage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to resolve host URL $($url).</p><br/>"; Continue }
		
            foreach ($ip4 in $ResourceAddresses) {
                try {
                    $Result = Test-NetConnection $ip4 -Port 443 -ea stop -wa silentlycontinue
                    switch ($Result.TcpTestSucceeded) {
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
                    try { [array]$Resources = (Resolve-DnsName $Name -ErrorAction stop).IP4Address }
                    catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Unable to resolve host $($Name).</p><br/>"; Continue }
				
                    #[array]$Resources = (Resolve-DnsName $Name).Ip4Address
                    $ResourcesPort = $url.Split(":")[1]
                }
                Else {
                    $Name = $url
                    try { [array]$Resources = (Resolve-DnsName $Name -ErrorAction stop).IP4Address }
                    catch { $ErrorMessage = $_; $rptTests += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Unable to resolve host URL $($url).</p><br/>"; Continue }
				
                    #[array]$Resources = (Resolve-DnsName $Name).IP4Address
                    $ResourcesPort = "443"
                }
                foreach ($ip4 in $Resources) {
                    try {
                        $Result = Test-NetConnection $ip4 -Port $ResourcesPort -ea stop -wa silentlycontinue
                        switch ($Result.TcpTestSucceeded) {
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
    "AAD_BASIC"                               = "Azure Active Directory Basic"
    "AAD_PREMIUM"                             = "Azure Active Directory Premium P1"
    "AAD_PREMIUM_P2"                          = "Azure Active Directory Premium P2"
    "AAD_SMB"                                 = ""
    "ADALLOM_O365"                            = "Office 365 Cloud App Security"
    "ADALLOM_S_DISCOVERY"                     = "Unknown"
    "ATP_ENTERPRISE"                          = "Exchange Online Advanced Threat Protection"
    "BI_AZURE_P1"                             = "Power BI Reporting and Analytics"
    "BPOS_S_TODO_1"                           = "Microsoft To Do"
    "CRMIUR"                                  = "CMRIUR"
    "CRMPLAN2"                                = "MICROSOFT DYNAMICS CRM ONLINE BASIC"
    "CRMSTANDARD"                             = "Microsoft Dynamics CRM Online Professional"
    "DESKLESS"                                = ""
    "DESKLESSPACK"                            = "Office 365 (Plan K1)"
    "DESKLESSPACK_GOV"                        = "Microsoft Office 365 (Plan K1) for Government"
    "DESKLESSWOFFPACK"                        = "Office 365 (Plan K2)"
    "DEVELOPERPACK"                           = "OFFICE 365 ENTERPRISE E3 DEVELOPER"
    "DYN365_ENTERPRISE_CUSTOMER_SERVICE"      = "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
    "DYN365_ENTERPRISE_P1_IW"                 = "Dynamics 365 P1 Trial for Information Workers"
    "DYN365_ENTERPRISE_PLAN1"                 = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
    "DYN365_ENTERPRISE_SALES"                 = "Dynamics Office 365 Enterprise Sales"
    "DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE" = "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
    "DYN365_ENTERPRISE_TEAM_MEMBERS"          = "Dynamics 365 For Team Members Enterprise Edition"
    "DYN365_FINANCIALS_BUSINESS_SKU"          = "Dynamics 365 for Financials Business Edition"
    "DYN365_FINANCIALS_TEAM_MEMBERS_SKU"      = "Dynamics 365 for Team Members Business Edition"
    "Dynamics_365_for_Operations"             = "DYNAMICS 365 UNF OPS PLAN ENT EDITION"
    "ECAL_SERVICES"                           = "ECAL"
    "EMS"                                     = "ENTERPRISE MOBILITY + SECURITY E3"
    "EMSPREMIUM"                              = "ENTERPRISE MOBILITY + SECURITY E5"
    "ENTERPRISEPACK"                          = "Enterprise Plan E3"
    "ENTERPRISEPACK_B_PILOT"                  = "Office 365 (Enterprise Preview)"
    "ENTERPRISEPACK_FACULTY"                  = "Office 365 (Plan A3) for Faculty"
    "ENTERPRISEPACK_GOV"                      = "Microsoft Office 365 (Plan G3) for Government"
    "ENTERPRISEPACK_STUDENT"                  = "Office 365 (Plan A3) for Students"
    "ENTERPRISEPACKLRG"                       = "Enterprise Plan E3"
    "ENTERPRISEPREMIUM"                       = "OFFICE 365 ENTERPRISE E5"
    "ENTERPRISEPREMIUM_NOPSTNCONF"            = "OFFICE 365 ENTERPRISE E5 WITHOUT AUDIO CONFERENCING"
    "ENTERPRISEWITHSCAL"                      = "Enterprise Plan E4"
    "ENTERPRISEWITHSCAL_FACULTY"              = "Office 365 (Plan A4) for Faculty"
    "ENTERPRISEWITHSCAL_GOV"                  = "Microsoft Office 365 (Plan G4) for Government"
    "ENTERPRISEWITHSCAL_STUDENT"              = "Office 365 (Plan A4) for Students"
    "EOP_ENTERPRISE_FACULTY"                  = "Exchange Online Protection for Faculty"
    "EQUIVIO_ANALYTICS"                       = "Office 365 Advanced eDiscovery"
    "ESKLESSWOFFPACK_GOV"                     = "Microsoft Office 365 (Plan K2) for Government"
    "EXCHANGE_L_STANDARD"                     = "Exchange Online (Plan 1)"
    "EXCHANGE_S_ARCHIVE_ADDON"                = ""
    "EXCHANGE_S_ARCHIVE_ADDON_GOV"            = "Exchange Online Archiving"
    "EXCHANGE_S_DESKLESS"                     = "Exchange Online Kiosk"
    "EXCHANGE_S_DESKLESS_GOV"                 = "Exchange Kiosk"
    "EXCHANGE_S_ENTERPRISE_GOV"               = "Exchange Plan 2G"
    "EXCHANGE_S_ESSENTIALS"                   = "EXCHANGE ONLINE ESSENTIALS"
    "EXCHANGE_S_STANDARD"                     = ""
    "EXCHANGE_S_STANDARD_MIDMARKET"           = "Exchange Online (Plan 1)"
    "EXCHANGEARCHIVE"                         = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
    "EXCHANGEARCHIVE_ADDON"                   = "Exchange Online Archiving For Exchange Online"
    "EXCHANGEDESKLESS"                        = "Exchange Online Kiosk"
    "EXCHANGEENTERPRISE"                      = "Exchange Online Plan 2"
    "EXCHANGEENTERPRISE_GOV"                  = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
    "EXCHANGEESSENTIALS"                      = "EXCHANGE ONLINE ESSENTIALS"
    "EXCHANGESTANDARD"                        = "Office 365 Exchange Online Only"
    "EXCHANGESTANDARD_GOV"                    = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
    "EXCHANGESTANDARD_STUDENT"                = "Exchange Online (Plan 1) for Students"
    "EXCHANGETELCO"                           = "EXCHANGE ONLINE POP"
    "FLOW_FREE"                               = "Microsoft Flow Free"
    "INTUNE_A"                                = "Windows Intune Plan A"
    "INTUNE_SMBIZ"                            = ""
    "IT_ACADEMY_AD"                           = "MS IMAGINE ACADEMY"
    "LITEPACK"                                = "OFFICE 365 SMALL BUSINESS"
    "LITEPACK_P2"                             = "OFFICE 365 SMALL BUSINESS PREMIUM"
    "MCOEV"                                   = "SKYPE FOR BUSINESS CLOUD PBX"
    "MCOIMP"                                  = "SKYPE FOR BUSINESS ONLINE (PLAN 1)"
    "MCOLITE"                                 = "Lync Online (Plan 1)"
    "MCOMEETACPEA"                            = "Audio Conferencing Pay-Per-Minute"
    "MCOMEETADV"                              = "AUDIO CONFERENCING"
    "MCOPSTN1"                                = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
    "MCOPSTN2"                                = "SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
    "MCOPSTNC"                                = "Communication Credits"
    "MCOSTANDARD"                             = "SKYPE FOR BUSINESS ONLINE (PLAN 2)"
    "MCOSTANDARD_GOV"                         = "Lync Plan 2G"
    "MCOSTANDARD_MIDMARKET"                   = "Lync Online (Plan 1)"
    "MFA_PREMIUM"                             = "Azure Multi-Factor Authentication"
    "MIDSIZEPACK"                             = "OFFICE 365 MIDSIZE BUSINESS"
    "O365_BUSINESS"                           = "Office 365 Business"
    "O365_BUSINESS_ESSENTIALS"                = "Office 365 Business Essentials"
    "O365_BUSINESS_PREMIUM"                   = "Office 365 Business Premium"
    "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"      = "Office ProPlus"
    "OFFICESUBSCRIPTION"                      = "OFFICE 365 PROPLUS"
    "OFFICESUBSCRIPTION_GOV"                  = "Office ProPlus"
    "OFFICESUBSCRIPTION_STUDENT"              = "Office ProPlus Student Benefit"
    "PLANNERSTANDALONE"                       = "Planner Standalone"
    "POWER_BI_ADDON"                          = "POWER BI FOR OFFICE 365 ADD-ON"
    "POWER_BI_INDIVIDUAL_USE"                 = "Power BI Individual User"
    "POWER_BI_PRO"                            = "Power BI Pro"
    "POWER_BI_STANDALONE"                     = "Power BI Stand Alone"
    "POWER_BI_STANDARD"                       = "Power-BI Standard"
    "POWERAPPS_VIRAL"                         = "Microsoft PowerApps Plan 2 Trial"
    "PROJECT_MADEIRA_PREVIEW_IW_SKU"          = "Dynamics 365 for Financials for IWs"
    "PROJECTCLIENT"                           = "PROJECT FOR OFFICE 365"
    "PROJECTESSENTIALS"                       = "PROJECT ONLINE ESSENTIALS"
    "PROJECTONLINE_PLAN_1"                    = "PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
    "PROJECTONLINE_PLAN_2"                    = "PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
    "ProjectPremium"                          = "Project Online Premium"
    "PROJECTPROFESSIONAL"                     = "PROJECT ONLINE PROFESSIONAL"
    "PROJECTWORKMANAGEMENT"                   = "Office 365 Planner Preview"
    "RIGHTSMANAGEMENT"                        = "AZURE INFORMATION PROTECTION PLAN 1"
    "RIGHTSMANAGEMENT_ADHOC"                  = "Windows Azure Rights Management"
    "RMS_S_ENTERPRISE"                        = "Azure Active Directory Rights Management"
    "RMS_S_ENTERPRISE_GOV"                    = "Windows Azure Active Directory Rights Management"
    "SHAREPOINTDESKLESS"                      = "SharePoint Online Kiosk"
    "SHAREPOINTDESKLESS_GOV"                  = "SharePoint Online Kiosk"
    "SHAREPOINTENTERPRISE"                    = "SHAREPOINT ONLINE (PLAN 2)"
    "SHAREPOINTENTERPRISE_GOV"                = "SharePoint Plan 2G"
    "SHAREPOINTENTERPRISE_MIDMARKET"          = "SharePoint Online (Plan 1)"
    "SHAREPOINTLITE"                          = "SharePoint Online (Plan 1)"
    "SHAREPOINTSTANDARD"                      = "SHAREPOINT ONLINE (PLAN 1)"
    "SHAREPOINTSTORAGE"                       = "SharePoint storage"
    "SHAREPOINTWAC"                           = "Office Online"
    "SHAREPOINTWAC_GOV"                       = "Office Online for Government"
    "SMB_BUSINESS"                            = "OFFICE 365 BUSINESS"
    "SMB_BUSINESS_ESSENTIALS"                 = "OFFICE 365 BUSINESS ESSENTIALS"
    "SMB_BUSINESS_PREMIUM"                    = "OFFICE 365 BUSINESS PREMIUM"
    "SPB"                                     = "Microsoft 365 Business"
    "SPE_E3"                                  = "Microsoft 365 E3"
    "SPZA_IW"                                 = "App Connect"
    "STANDARD_B_PILOT"                        = "Office 365 (Small Business Preview)"
    "STANDARDPACK"                            = "OFFICE 365 ENTERPRISE E1"
    "STANDARDPACK_FACULTY"                    = "Office 365 (Plan A1) for Faculty"
    "STANDARDPACK_GOV"                        = "Microsoft Office 365 (Plan G1) for Government"
    "STANDARDPACK_STUDENT"                    = "Office 365 (Plan A1) for Students"
    "STANDARDWOFFPACK"                        = "OFFICE 365 ENTERPRISE E2"
    "STANDARDWOFFPACK_FACULTY"                = "Office 365 Education E1 for Faculty"
    "STANDARDWOFFPACK_GOV"                    = "Microsoft Office 365 (Plan G2) for Government"
    "STANDARDWOFFPACK_IW_FACULTY"             = "Office 365 Education for Faculty"
    "STANDARDWOFFPACK_IW_STUDENT"             = "Office 365 Education for Students"
    "STANDARDWOFFPACK_STUDENT"                = "Microsoft Office 365 (Plan A2) for Students"
    "STANDARDWOFFPACKPACK_FACULTY"            = "Office 365 (Plan A2) for Faculty"
    "STANDARDWOFFPACKPACK_STUDENT"            = "Office 365 (Plan A2) for Students"
    "TEAMS1"                                  = "Microsoft Teams"
    "VISIOCLIENT"                             = "VISIO Online Plan 2"
    "VISIOONLINE_PLAN1"                       = "Visio Online Plan 1"
    "WACONEDRIVEENTERPRISE"                   = "ONEDRIVE FOR BUSINESS (PLAN 2)"
    "WACONEDRIVESTANDARD"                     = "ONEDRIVE FOR BUSINESS (PLAN 1)"
    "WIN10_PRO_ENT_SUB"                       = "WINDOWS 10 ENTERPRISE E3"
    "WINBIZ"                                  = "Windows 10 Business"
    "YAMMER_ENTERPRISE"                       = "Yammer for the Starship Enterprise"
    "YAMMER_MIDSIZE"                          = "Yammer"
}


#	Returns the list of subscribed services
[uri]$uriServices = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/Services"
#	Returns the current status of the service.
[uri]$uriCurrentStatus = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/CurrentStatus"
#	Returns the historical status of the service, by day, over a certain time range.
[uri]$uriHistoricalStatus = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/HistoricalStatus"
#	Returns the messages about the service over a certain time range.
[uri]$uriMessages = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/Messages"
#   Return the messages on the RSS feed for the O365 roadmap
[uri]$uriO365Roadmap = "https://www.microsoft.com/en-gb/microsoft-365/RoadmapFeatureRSS"
#   Return the messages on the RRS feed for Azure Updates
[uri]$uriAzureUpdates = "https://azurecomcdn.azureedge.net/en-gb/updates/feed/"

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
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uriToken -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
#	Returns the tenant licence information

[uri]$uriLicences = "https://graph.microsoft.com/v1.0/subscribedskus"


#Fetch the information from Office 365 Service Health API
#Get Services: Get the list of subscribed services
try {
    if ($proxyServer) {
        [array]$allSubscribedMessages = @((Invoke-RestMethod -Uri $uriServices -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).Value)
    }
    else {
        [array]$allSubscribedMessages = @((Invoke-RestMethod -Uri $uriServices -Headers $authHeader -Method Get).Value)
    }
    if ($null -eq $allSubscribedMessages -or $allSubscribedMessages.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No Subscribed services returned - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>$($allSubscribedMessages.count) subscribed services returned.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No Subscribed services returned - verify proxy and network connectivity</p><br/>"
}

#Get Current Status: Get a real-time view of current and ongoing service incidents and maintenance events
try {
    if ($proxyServer) {
        [array]$allCurrentStatusMessages = @((Invoke-RestMethod -Uri $uriCurrentStatus -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).Value)
    }
    else {
        [array]$allCurrentStatusMessages = @((Invoke-RestMethod -Uri $uriCurrentStatus -Headers $authHeader -Method Get).Value)
    }
    if ($null -eq $allCurrentStatusMessages -or $allCurrentStatusMessages.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Cannot retrieve the current status of services - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>$($allCurrentStatusMessages.count) services and status returned.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Cannot retrieve the current status of services - verify proxy and network connectivity</p><br/>"
}

#Get Historical Status: Get a historical view of service health, including service incidents and maintenance events.
try {
    if ($proxyServer) {
        [array]$allHistoricalStatusMessages = @((Invoke-RestMethod -Uri $uriHistoricalStatus -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).Value)
    }
    else {
        [array]$allHistoricalStatusMessages = @((Invoke-RestMethod -Uri $uriHistoricalStatus -Headers $authHeader -Method Get).Value)
    }
    if ($null -eq $allHistoricalStatusMessages -or $allHistoricalStatusMessages.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No historical service health messages retrieved - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>$($allHistoricalStatusMessages.count) historical service health messages returned.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No historical service health messages retrieved - verify proxy and network connectivity</p><br/>"
}

#Get Messages: Find Incident, Planned Maintenance, and Message Center communications.
try {
    if ($proxyServer) {
        [array]$allMessages = @((Invoke-RestMethod -Uri $uriMessages -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).Value)
    }
    else {
        [array]$allMessages = @((Invoke-RestMethod -Uri $uriMessages -Headers $authHeader -Method Get).Value)
    }
    if ($null -eq $allMessages -or $allMessages.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No message center messages retrieved - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>$($allMessages.count) message center messages retrieved.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No message center messages retrieved - verify proxy and network connectivity</p><br/>"
}

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

try {
    if ($proxyServer) {
        $Roadmap = @((Invoke-WebRequest -Uri $uriO365Roadmap -Proxy $proxyHost -ProxyUseDefaultCredentials).content)
    }
    else {
        $Roadmap = @((Invoke-WebRequest -Uri $uriO365Roadmap).content)
    }
    if ($null -eq $Roadmap -or $Roadmap.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No Office 365 RSS Feed information - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Office 365 Roadmap RSS feed items retrieved.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No Office 365 RSS Feed information - verify proxy and network connectivity</p><br/>"
}

try {
    if ($proxyServer) {
        $AzureUpdates = @((Invoke-WebRequest -Uri $uriAzureUpdates -Proxy $proxyHost -ProxyUseDefaultCredentials).content)
    }
    else {
        $AzureUpdates = @((Invoke-WebRequest -Uri $uriAzureUpdates).content)
    }
    if ($null -eq $AzureUpdates -or $AzureUpdates.Count -eq 0) {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No Azure Updates RSS Feed information - verify proxy and network connectivity</p><br/>"
    }
    else {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Azure Updates RSS feed items retrieved.</p><br/>"
    }
}
catch {
    $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No Azure Updates RSS Feed information - verify proxy and network connectivity</p><br/>"
}


$rptO365Info += "<br/>You can add some general information in here if needed.<br />"
$rptO365Info += "ie updates or links to external (ie cloud only) activity to verify Azure AD App is working (ie Flow to Teams Channel)"
if ($altLink) { $rptO365Info += "<a href='$($altLink)' target=_blank> here </a></li></ul><br>" }

#Check office 365 IPs and URLS
#Check before building the page as these will be used for diagnostics checks

#From docs.microsoft.com : https://docs.microsoft.com/en-us/Office365/Enterprise/office-365-ip-web-service
[uri]$ws = "https://endpoints.office.com"
$versionpath = $pathIPURLs + "\O365_endpoints_latestversion-$($rptProfile).txt"
$pathIP4 = $pathIPURLs + "\O365_endpoints_ip4-$($rptProfile).txt"
$pathIP6 = $pathIPURLs + "\O365_endpoints_ip6-$($rptProfile).txt"
$pathIPurl = $pathIPURLs + "\O365_endpoints_urls-$($rptProfile).txt"
$fileData = "O365_endpoints_data-$($rptProfile).txt"
$pathData = $pathIPURLs + "\" + $fileData
$currentData = $null
$currentData = Get-Content $pathData

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
$version = Invoke-RestMethod -Uri ($ipurlVersion)
if (($version.latest -gt $lastVersion) -or ($null -eq $currentData)) {
    $ipurlOutput += "New version of Office 365 worldwide commercial service instance endpoints detected<br />`r`n"
    #Send email to users on IP/URL change
    $emailSubject = "IPs and URLs: New version $($version.latest)"
    $emailMessage = "new version of Office 365 Worldwide Commercial service instance endpoints"
    SendReport $emailMessage $EmailCreds $config "Normal" $emailSubject $emailIPURLAlerts

    # write the new version number to the version file
    @($clientRequestId, $version.latest) | Out-File $versionpath
    # invoke endpoints method to get the new data
    [uri]$ipurlEndpoint = "$($ws)/endpoints/Worldwide?clientRequestId=$($clientRequestId)"
    $endpointSets = Invoke-RestMethod -Uri ($ipurlEndpoint)
    # filter results for Allow and Optimize endpoints, and transform these into custom objects with port and category
    # URL results
    $flatUrls = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $urls = $(if ($endpointSet.urls.Count -gt 0) { $endpointSet.urls } else { @() })
        $urlCustomObjects = @()
        if ($endpointSet.category -in ("Allow", "Optimize")) {
            $urlCustomObjects = $urls | ForEach-Object {
                [PSCustomObject]@{
                    category = $endpointSet.category;
                    url      = $_;
                    tcpPorts = $endpointSet.tcpPorts;
                    udpPorts = $endpointSet.udpPorts;
                }
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
}
else {
    $ipurlSummary += "Office 365 worldwide commercial service instance endpoints are up-to-date. <br />`r`n"
    $ipurlSummary += "Importing previous results. <br />`r`n"
    $ipurlSummary += "Data available from <a href='https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges' target=_blank>https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges</a><br/>`r`n"
    $flatUrls = Import-Csv $pathIPurl
    $flatIp4s = Import-Csv $pathIP4
    $flatIp6s = Import-Csv $pathIP6

}
# write output to screen
# Clients arent going to want to view this, are they?
$ipurlSummary += "<b>Client Request ID: " + $clientRequestId + "</b><br />`r`n"
$ipurlSummary += "<b>Last Version: " + $lastVersion + "</b><br />`r`n"
$ipurlSummary += "<b>New Version: " + $version.latest + "</b><br />`r`n"

#IPv4
$ipurlOutput += "<b>IPv4 Firewall IP Address Ranges</b><br />`r`n"
$ipurlOutput += "<b>Optimize (Direct connection):</b><br />`r`n"

$flatAddressIPv4 = @($flatIp4s | Where-Object { $_.category -like 'optimize' })
$ipurlOutput += "$(($flatAddressIPv4.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "<b>Allow:</b><br />`r`n"
$flatAddressIPv4 = @($flatIp4s | Where-Object { $_.category -notlike 'Optimize' })
$ipurlOutput += "$(($flatAddressIPv4.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "All IPv4 networks, TCP/UDP Ports and classifications available to <a href='$(Split-Path $pathIP4 -leaf)' target=_blank>download here</a><br /><br />`r`n"

#IPv6
$ipurlOutput += "<b>IPv6 Firewall IP Address Ranges</b><br />`r`n"
$ipurlOutput += "<b>Optimize (Direct connection):</b><br />`r`n"

$flatAddressIPv6 = @($flatIp6s | Where-Object { $_.category -like 'optimize' })
$ipurlOutput += "$(($flatAddressIPv6.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "<b>Allow:</b><br />`r`n"
$flatAddressIPv6 = @($flatIp6s | Where-Object { $_.category -notlike 'Optimize' })
$ipurlOutput += "$(($flatAddressIPv6.ip | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "All IPv6 networks, TCP/UDP Ports and classifications available to <a href='$(Split-Path $pathIP6 -leaf)' target=_blank>download here</a><br /><br />`r`n"

#URLs
$ipurlOutput += "<b>URLs</b><br />`r`n"
$ipurlOutput += "<b>Optimize (Direct connection):</b><br />`r`n"

$flatAddressURLs = @($flatUrls | Where-Object { $_.category -like 'optimize' })
$ipurlOutput += "$(($flatAddressURLs.url | Sort-Object -unique) -join ', ' | Out-String) <br /><br />`r`n"
$ipurlOutput += "<b>Allow:</b><br />`r`n"
$flatAddressURLs = @($flatUrls | Where-Object { $_.category -notlike 'Optimize' })
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
Copy-Item $pathdata -Destination $pathHTML
Copy-Item $pathIPurl -Destination $pathHTML
Copy-Item $pathIP4 -Destination $pathHTML
Copy-Item $pathIP6 -Destination $pathHTML

$checkOptHTTP = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*80*' -and $_.category -match 'Optimize') }
$checkOptHTTPs = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*443*' -and $_.category -match 'Optimize') }
$checkAllowHTTP = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.category -match 'Allow') -and ($_.tcpPorts -like '*443*' -or $_.tcpPorts -like '*80*') }
$checkAllowHTTPs = $flaturls | Where-Object { ($_.url -notmatch '\*' -and $_.tcpPorts -like '*443*' -and $_.category -match 'Allow') }
if ($diagWeb) {
    # Microsoft Office 365 URL tests - check the Optimize HTTP connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTP checks for 'Optimize' Sites (Invoke-WebRequest)</p><br/>"
    foreach ($entry in $checkOptHTTP) {
        $url = "http://$($entry.url)"
        try {
            #Proxy servers should not be used for optimized paths
            $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -UseBasicParsing
            Switch ($Result.StatusCode) {
                200 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully contacted site $($url).</p><br/>" }
                400 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Bad request.</p><br/>" }
                401 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Unauthorized.</p><br/>" }
                403 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Forbidden.</p><br/>" }
                404 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Not found.</p><br/>" }
                407 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Proxy authentication required.</p><br/>" }
                502 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Bad gateway (likely proxy).</p><br/>" }
                503 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Service unavailable (transient, try again).</p><br/>" }
                504 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Gateway timeout (likely proxy).</p><br/>" }
                default {
                    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to contact site  $($url).</p><br/>"
                }
            }
        }
        catch {
            $ErrorMessage = $_
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Exception: Unable to contact site $($url).</p><br/>"
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage).</p><br/>"
        }
    } # End Foreach URL List

    # Microsoft Office 365 URL tests - check the Optimize HTTPs connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTPs checks for 'Optimize' Sites (Invoke-WebRequest)</p><br/>"
    foreach ($entry in $checkOptHTTPs) {
        $url = "https://$($entry.url)"
        try {
            #Proxy servers should not be used for optimized paths
            $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -UseBasicParsing
            Switch ($Result.StatusCode) {
                200 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully contacted site $($url).</p><br/>" }
                400 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Bad request.</p><br/>" }
                401 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Unauthorized.</p><br/>" }
                403 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Forbidden.</p><br/>" }
                404 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Not found.</p><br/>" }
                407 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Proxy authentication required.</p><br/>" }
                502 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Bad gateway (likely proxy).</p><br/>" }
                503 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Service unavailable (transient, try again).</p><br/>" }
                504 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Gateway timeout (likely proxy).</p><br/>" }
                default {
                    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to contact site  $($url).</p><br/>"
                }
            }
        }
        catch {
            $ErrorMessage = $_
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Exception: Unable to contact site $($url).</p><br/>"
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage).</p><br/>"
        }
    } # End Foreach URL List

    # Microsoft Office 365 URL tests - check the Allow HTTP connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTP checks for 'Allow' Sites (Invoke-WebRequest)</p><br/>"
    foreach ($entry in $checkAllowHTTP) {
        $url = "http://$($entry.url)"
        try {
            if ($ProxyServer) {
                $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -Proxy $proxyHost -ProxyUseDefaultCredentials -UseBasicParsing
            }
            else {
                $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -UseBasicParsing
            }
            Switch ($Result.StatusCode) {
                200 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully contacted site $($url).</p><br/>" }
                400 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Bad request.</p><br/>" }
                401 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Unauthorized.</p><br/>" }
                403 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Forbidden.</p><br/>" }
                404 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Not found.</p><br/>" }
                407 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Proxy authentication required.</p><br/>" }
                502 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Bad gateway (likely proxy).</p><br/>" }
                503 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Service unavailable (transient, try again).</p><br/>" }
                504 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Gateway timeout (likely proxy).</p><br/>" }
                default {
                    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to contact site  $($url).</p><br/>"
                }
            }
        }
        catch {
            $ErrorMessage = $_
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Exception: Unable to contact site $($url).</p><br/>"
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage).</p><br/>"
        }
    } # End Foreach URL List

    # Microsoft Office 365 URL tests - check the Allow HTTPs connections
    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='section'>Starting HTTPs checks for 'Allow' Sites (Invoke-WebRequest)</p><br/>"
    foreach ($entry in $checkAllowHTTPs) {
        $url = "https://$($entry.url)"
        try {
            if ($ProxyServer) {
                $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -Proxy $proxyHost -ProxyUseDefaultCredentials -UseBasicParsing
            }
            else {
                $Result = Invoke-WebRequest -Uri $url -ea stop -wa silentlycontinue -UseBasicParsing
            }
            Switch ($Result.StatusCode) {
                200 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>Successfully contacted site $($url).</p><br/>" }
                400 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Bad request.</p><br/>" }
                401 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Unauthorized.</p><br/>" }
                403 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Forbidden.</p><br/>" }
                404 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='warning'>Failed to contact site $($url): Not found.</p><br/>" }
                407 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Proxy authentication required.</p><br/>" }
                502 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Bad gateway (likely proxy).</p><br/>" }
                503 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Service unavailable (transient, try again).</p><br/>" }
                504 { $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Failed to contact site $($url): Gateway timeout (likely proxy).</p><br/>" }
                default {
                    $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Unable to contact site  $($url).</p><br/>"
                }
            }
        }
        catch {
            $ErrorMessage = $_
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>Exception: Unable to contact site $($url).</p><br/>"
            $rptIPURLs += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>$($ErrorMessage).</p><br/>"
        }
    } # End Foreach URL List

}

#Start Building the Pages
#Build Div1
#Build Summary Dashboard
# 6 cards
$HistoryIncidents = $allMessages | Where-Object { ($_.EndTime -ne $null -and $_.messagetype -notlike 'MessageCenter') } | Sort-Object EndTime -Descending
$rptSectionOneOne = "<div class='section'><div class='header'>Office 365 Dashboard Status</div>`n"
$rptSectionOneOne += "<div class='content'>`n"
$rptSectionOneOne += "<div class='dash-outer'><div class='dash-inner'>`n"
foreach ($card in $dashCards) {
    [array]$item = @()
    [array]$hist = @()
    [int]$advisories = 0
    $item = $allCurrentStatusMessages | Where-Object { $_.WorkloadDisplayName -like $card }
    $hist = $HistoryIncidents | Where-Object { $_.WorkloadDisplayName -like $card -and ($_.status -notlike 'False Positive') } | Sort-Object EndTime -Descending
    $advisories = ($allMessages | Where-Object { ($_.messagetype -like 'MessageCenter' -and $_.AffectedWorkloadDisplayNames -like $card) }).count
    if ($hist.count -gt 0) {
        $days = "{0:N0}" -f (New-TimeSpan -Start (Get-Date $hist[0].EndTime) -End $(Get-Date)).TotalDays
    }
    else {
        $days = "&gt;30"
    }
    $cardClass = Get-StatusDisplay $($item.status) "Class"
    $cardText = cardbuilder $($item.workloaddisplayname) $($Days) $($Hist.count) $advisories $cardClass
    $rptSectionOneOne += "$cardText`n"
}
$rptSectionOneOne += "</div></div>`n" #Close inner and outer divs

#Get Current Status and Issues for non operational services
[array]$CurrentStatusBad = $allCurrentStatusMessages | Where-Object { $_.status -notlike 'ServiceOperational' }
[array]$rptSummaryTable = @()
$rptSummaryTable = "<br/><br/><div class='dash-outer'><div class='dash-inner'>`n"
if ($CurrentStatusBad.count -ge 1) {
    $rptSummaryTable += "<div class='tableWrkld'>`r`n"
    $rptSummaryTable += "<div class='tableWrkld-title'>The following services are reporting service issues</div>`r`n"
    $rptSummaryTable += "<div class='tableWrkld-header'>`n`t<div class='tableWrkld-header-r'>Systems</div>`n`t<div class='tableWrkld-header-c'>Status</div>`n`t<div class='tableWrkld-header-l'>Status at $(Get-Date -Format 'HH:mm')</div>`n</div>`n"
    foreach ($item in $CurrentStatusBad) {
        $statusIcon = Get-StatusDisplay $($item.status) "Icon"
        $rptSummaryTable += "<div class='tableWrkld-row'>`n`t<div class='tableWrkld-cell-r'>$($item.WorkloadDisplayName)</div>`n`t<div class='tableWrkld-cell-c'>$StatusIcon</div>`n`t<div class='tableWrkld-cell-l'>$($item.StatusDisplayName)</div>`n</div>"
    }
}
else {
    $rptSummaryTable += "<div class='tableWrkld'>`r`n"
    if ($authErrMsg) { $rptSummaryTable += "<div class='tableWrkld-title'>$authErrMsg</div>`r`n" }
    else { $rptSummaryTable += "<div class='tableWrkld-title'>No current or recent issues to display</div>`r`n" }
}
#Close table Workld div
$rptSummaryTable += "</div>`n"
#Close div and content div
$rptSummaryTable += "</div></div>`n"
$rptSectionOneOne += $rptSummaryTable
$rptSectionOneOne += "</div></div>`n" #Close content and section

$divOne = $rptSectionOneOne

#Get current and recent incidents
$rptSectionOneTwo = "<div class='section'><div class='header'>Active and Recent Incidents</div>`n"
$rptSectionOneTwo += "<div class='content'>`n"

[array]$CurrentMessagesOpen = @()
[array]$rptActiveTable = @()
$CurrentMessagesOpen = $allMessages | Where-Object { ($_.messagetype -notlike 'MessageCenter' -and $_.EndTime -eq $null) } | Sort-Object LastUpdatedTime -Descending
$rptActiveTable += "<div class='tableInc'>`n"
$rptActiveTable += "<div class='tableInc-title'>Active Messages</div>`n"
if ($CurrentMessagesOpen.count -ge 1) {
    foreach ($item in $CurrentMessagesOpen) {
        $rptOutage = $true
        $LastUpdated = $(Get-Date $item.lastupdatedtime -f 'dd-MMM-yyyy HH:mm')
        $severity = $item.severity
        switch ($severity) {
            "SEV0" { $actionStyle = "style=border:none;font-weight:bold;color:red" }
            "SEV1" { $actionStyle = "style=border:none;color:red" }
            "SEV2" { $actionStyle = "style=border:none;color:blue" }
            default { $actionStyle = "style=border:none;font-weight:bold;color:red" }
        }
        #Build
        $link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($link) {
            $ID = "<a href=$($link) target=_blank>$($item.ImpactDescription)</a>"
        }
        else { $ID = "$($item.ImpactDescription)" }
        $rptActiveTable += "<div class='tableInc-row'><div class='tableInc-cell-l'>$($item.WorkloadDisplayname)</div>`r`n<div class='tableInc-cell-r' $($actionStyle)>$($Severity)</div>`r`n<div class='tableInc-cell-r'>Last Update: $($LastUpdated)</div>`r`n<div class='tableInc-cell-l'>$($item.Status)</div>`r`n<div class='tableInc-cell-l'>$($ID)</div>`r`n</div>`r`n"
    }
}
else {
    $rptActiveTable += "<div class='tableInc-header'><span class='tableInc-header-c'>No open incidents to display</span></div>`n"
}
$rptActiveTable += "</div><br/>`n"

#Show recently closed messages
#create a timespan for recently closed messages - 3 to include weekends
$recentDate = (Get-Date).AddDays(-$IncidentDays)
[array]$RecentMessagesOpen = @()
$RecentMessagesOpen = $allMessages | Where-Object { ($_.messagetype -notlike 'MessageCenter' -and $_.EndTime -ne $null -and ((Get-Date $_.EndTime) -ge (Get-Date $recentdate))) } | Sort-Object LastUpdatedTime -Descending
if ($RecentMessagesOpen.count -ge 1) {
    $rptActiveTable += "<div class='tableInc'>`n"
    $rptActiveTable += "<div class='tableInc-title'>Incidents closed in the last $($incidentdays*24) hours (since $(Get-Date $recentDate -f 'dd-MMM-yyyy HH:mm'))</div>`n"
    foreach ($item in $RecentMessagesOpen) {
        $rptOutage = $true
        $EndTime = $(Get-Date $item.EndTime -f 'dd-MMM-yyyy HH:mm')
        $severity = $item.severity
        switch ($severity) {
            "SEV0" { $actionStyle = "style=border:none;font-weight:bold;color:red" }
            "SEV1" { $actionStyle = "style=border:none;color:red" }
            "SEV2" { $actionStyle = "style=border:none;color:blue" }
            default { $actionStyle = "style=border:none;font-weight:bold;color:red" }
        }
        #Build
        $link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($link) {
            $ID = "<a href=$($link) target=_blank>$($item.ImpactDescription)</a>"
        }
        else { $ID = "$($item.ImpactDescription)" }
        $rptActiveTable += "<div class='tableInc-row'><div class='tableInc-cell-l'>$($item.WorkloadDisplayname)</div>`r`n<div class='tableInc-cell-r' $($actionStyle)>$($Severity)</div>`r`n<div class='tableInc-cell-r'>Closed: $($EndTime)</div>`r`n<div class='tableInc-cell-l'>$($item.Status)</div>`r`n<div class='tableInc-cell-l'>$($ID)</div>`r`n</div>`r`n"
    }
    $rptActiveTable += "</div>`n"
}
else {
    $rptActiveTable += "<div class='tableInc'>`n"
    $rptActiveTable += "<div class='tableInc-header'><span class='tableInc-header-c'>No recent incidents to display</span></div>`n"
    $rptActiveTable += "</div>`n"
}
$rptSectionOneTwo += $rptActiveTable
$rptSectionOneTwo += "</div></div>`r`n" #Close content and section
$divOne += $rptSectionOneTwo

#Get All workload status
[array]$rptWorkloadStatusTable = @()
$rptSectionOneThree = "<div class='section'><div class='header'>Workload Status</div>`n"
$rptSectionOneThree += "<div class='content'>`n"
$allCurrentStatusMessages = $allCurrentStatusMessages | Sort-Object WorkloadDisplayname
$rptWorkloadStatusTable = "<br/><div class='dash-outer'><div class='dash-inner'>`n"
if ($allCurrentStatusMessages.count -ge 1) {
    $rptWorkloadStatusTable += "<div class='tableWrkld'>`r`n"
    $rptWorkloadStatusTable += "<div class='tableWrkld-title'>All workload status</div>`r`n"
    $rptWorkloadStatusTable += "<div class='tableWrkld-header'>`n`t<div class='tableWrkld-header-r'>Systems</div>`n`t<div class='tableWrkld-header-c'>Status</div>`n`t<div class='tableWrkld-header-l'>Status at $(Get-Date -Format 'HH:mm')</div>`n</div>`n"

    foreach ($item in $allCurrentStatusMessages) {
        $statusIcon = Get-StatusDisplay $($item.status) "Icon"
        $rptWorkloadStatusTable += "<div class='tableWrkld-row'>`n`t<div class='tableWrkld-cell-r'>$($item.WorkloadDisplayName)</div>`n`t<div class='tableWrkld-cell-c'>$StatusIcon</div>`n`t<div class='tableWrkld-cell-l'>$($item.StatusDisplayName)</div>`n</div>"
    }
}
else {
    $rptWorkloadStatusTable += "<div class='tableWrkld'>`r`n"
    $rptWorkloadStatusTable += "<div class='tableWrkld-title'>No current or recent issues to display</div>`r`n"
}
$rptWorkloadStatusTable += "</div></div></div>`n"
$rptSectionOneThree += $rptWorkloadStatusTable
$rptSectionOneThree += "</div></div>`r`n" #Close content and section
$divOne += $rptSectionOneThree

#Build Div2
[array]$listLineOne = @()
[array]$listTheRest = @()
foreach ($card in $prefDashCards) { $listLineOne += $allCurrentStatusMessages | Where-Object { $_.workloaddisplayname -like $card } }
$listTheRest = $allCurrentStatusMessages | Where-Object { $_.workloaddisplayname -notin $listlineone.workloaddisplayname } | Sort-Object  workloaddisplayname
$DashWorkloads = $listLineOne + $ListTheRest
$rptFeatureDash = "<div class='container'>`n"
foreach ($workload in $DashWorkloads) {
    [array]$feature = @()
    [string]$cardDetail = ""
    [string]$cardClass = ""
    [boolean]$blnUrgent = $false
    [boolean]$blnErr = $false
    [boolean]$blnWarn = $false
    [boolean]$blnOK = $false
    [int]$intFeatureCount = 0
    foreach ($feature in $workload.FeatureStatus) {
        $cardClass = Get-StatusDisplay $($feature.FeatureServiceStatus) 'Class'
        #If any of the substatus values are not ok, log and set the main card value?
        switch ($CardClass) {
            "urgent" { $blnUrgent = $true }
            "err" { $blnErr = $true }
            "warn" { $blnWarn = $true }
            default { $blnOK = $true }
        }
        $cardDetail += "<div class='feature-item-$($cardClass)'>$($feature.featuredisplayname)<span class='tooltiptext'>$($feature.FeatureServiceStatusDisplayName)</span></div>`r`n"
        if (($feature.FeatureServiceStatusDisplayName).length -gt 29) { $intFeatureCount += 2 } else { $intFeatureCount++ }
    }
    if ($blnUrgent) { $cardClass = "err" }
    elseif ($blnErr) { $cardClass = "err" }
    elseif ($blnWarn) { $cardClass = "warn" }
    else { $cardClass = "ok" }
    $cardText = featurebuilder $($workload.workloaddisplayname) $cardDetail $cardClass $intFeatureCount
    $rptFeatureDash += "$cardText`n"
}
$rptFeatureDash += "</div>`r`n<br/><br/>`r`n"
$divTwo = ($rptFeatureDash)

#Build Div3
$rptSectionThreeOne = "<div class='section'><div class='header'>Office 365 Incident History</div>`n"
$rptSectionThreeOne += "<div class='content'>`n"

#Incident History
#Get all closed (end date) messages of incidents
[array]$HistoryIncidents = @()
$rptHistoryTable = @()
$item = $null
$HistoryIncidents = $allMessages | Where-Object { ($_.EndTime -ne $null -and $_.messagetype -notlike 'MessageCenter') } | Sort-Object EndTime -Descending
if ($HistoryIncidents.count -ge 1) {
    $rptHistoryTable += "<div class='tableInc'>`n"
    $rptHistoryTable += "<div class='tableInc-title'>Closed Incidents</div>`n"
    $rptHistoryTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-c'>Feature</div>`n`t<div class='tableInc-header-c'>Status</div>`n`t<div class='tableInc-header-c'>Description</div>`n`t<div class='tableInc-header-c'>Start Time</div>`n`t<div class='tableInc-header-c'>End Time</div>`n`t<div class='tableInc-header-c'>Last Updated</div>`n</div>`n"
    foreach ($item in $HistoryIncidents) {
        if ($item.StartTime) { $StartTime = $(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm') } else { $StartTime = "" }
        if ($item.EndTime) { $EndTime = $(Get-Date $item.EndTime -f 'dd-MMM-yyyy HH:mm') } else { $EndTime = "" }
        if ($item.LastUpdatedTime) { $LastUpdated = $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm') } else { $LastUpdated = "" }
        $link = ""
        #Build link to detailed message
        $link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($link) {
            $ID = "<a href=$($link) target=_blank>$($item.ID) - $($item.ImpactDescription)</a>"
        }
        else { $ID = "$($item.ID) - $($item.ImpactDescription)" }
        $rptHistoryTable += "<div class='tableInc-row'>`n`t"
        $rptHistoryTable += "<div class='tableInc-cell-l'>$($item.WorkloadDisplayname -join '<br>')</div>`n`t"
        $rptHistoryTable += "<div class='tableInc-cell-l'>$($item.Status)</div>`n`t"
        $rptHistoryTable += "<div class='tableInc-cell-l'>$($ID)</div>`n`t"
        $rptHistoryTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($StartTime)</div>`n`t"
        $rptHistoryTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($EndTime)</div>`n`t"
        $rptHistoryTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($LastUpdated)</div>`n"
        $rptHistoryTable += "</div>`n"
    }
}
else {
    $rptHistoryTable = "<div class='tableInc'>`n"
    $rptHistoryTable += "<div class='tableInc-title'>No Closed Incidents</div>`n"
}
$rptHistoryTable += "</div>`n"
$rptSectionThreeOne += $rptHistoryTable
$rptSectionThreeOne += "</div></div>`n"

$divThree = $rptSectionThreeOne

#Build Div4
$rptSectionFourOne = "<div class='section'><div class='header'>Prevent / Fix Issues</div>`n"
$rptSectionFourOne += "<div class='content'>`n"
#Messages
#Prevent or fix issues
[array]$MessagesFix = @()
[array]$rptMessagesFixTable = @()
$MessagesFix = $allMessages | Where-Object { ($_.messagetype -like 'MessageCenter' -and $_.category -like 'Prevent or Fix Issues') } | Sort-Object MilestoneDate -Descending
if ($MessagesFix.count -ge 1) {
    $rptMessagesFixTable += "<div class='tableInc'>`n"
    $rptMessagesFixTable += "<div class='tableInc-title'>Closed Incidents</div>`n"
    $rptMessagesFixTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-dt'>Feature</div>`n`t<div class='tableInc-header-dt'>Severity</div>`n`t<div class='tableInc-header-dt'>Action</div>`n`t<div class='tableInc-header-dt'>ID</div>`n`t<div class='tableInc-header-l'>Title</div>`n`t<div class='tableInc-header-dt'>Milestone</div>`n<div class='tableInc-header-dt'>Action Rqd By</div>`n</div>`n"
    foreach ($item in $MessagesFix) {
        if ($item.LastUpdatedTime) { $LastUpdated = $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm') }
        if ($item.MilestoneDate) { $MilestoneDate = $(Get-Date $item.MilestoneDate -f 'dd-MMM-yyyy HH:mm') }
        $Workloads = @()
        $Workloads = $item | Select-Object -ExpandProperty AffectedWorkloadDisplaynames
        if (!($Workloads)) { $Workloads = "General" }
        $workloads = $workloads -join "</br>"
        $rptMessagesFixTable += "<div class='tableInc-row'>`n`t"
        $rptMessagesFixTable += "<div class='tableInc-cell-dt'>$($Workloads)</div>`n`t"
        $rptMessagesFixTable += "<div class='tableInc-cell-dt'>$($item.Severity)</div>`n`t"
        $rptMessagesFixTable += "<div class='tableInc-cell-dt'>$($item.ActionType)</div>`n`t"
        #Build advisory and get link
        $link = Get-AdvisoryInHTML $item $RebuildDocs $pathHTMLDocs
        $rptMessagesFixTable += "<div class='tableInc-cell-dt'>$($item.ID)</div>`n`t"
        if ($link) { $rptMessagesFixTable += "<div class='tableInc-cell-l'><a href='$($link)' target='_blank'>$($item.title)</a></div>`n`t" }
        else { $rptMessagesFixTable += "<div class='tableInc-cell-l'>$($item.title)</div>`n`t" }
        $rptMessagesFixTable += "<div class='tableInc-cell-dt'>$($MilestoneDate)</div>`n`t"
        if ($item.ActionRequiredByDate) {
            $ActionRequiredByDate = $(Get-Date $item.ActionRequiredByDate -f 'dd-MMM-yyyy HH:mm')
            $action = (New-TimeSpan -Start $(Get-Date) -End (Get-Date $item.ActionRequiredByDate)).TotalDays
            switch ($action) {
                { $_ -ge 0 -and $_ -lt 7 } { $actionStyle = "style=border:none;font-weight:bold;color:red" }
                { $_ -ge 7 -and $_ -lt 14 } { $actionStyle = "style=border:none;color:red" }
                { $_ -ge 14 -and $_ -lt 21 } { $actionStyle = "style=border:none;color:blue" }
                default { $actionStyle = "style=border:none;" }
            }
        }
        $rptMessagesFixTable += "<div class='tableInc-cell-dt'>$($ActionRequiredByDate)</div>`n"
        $rptMessagesFixTable += "</div>`n"
    }
}
else {
    $rptMessagesFixTable = "<div class='tableInc'>`n"
    $rptMessagesFixTable += "<div class='tableInc-title'>No 'Prevent/Fix Issues' Messages</div>`n"
}
$rptMessagesFixTable += "</div>`n"

$rptSectionFourOne += $rptMessagesFixTable
$rptSectionFourOne += "</div></div>`n"
$divFour = $rptSectionFourOne

$rptSectionFourTwo = "<div class='section'><div class='header'>Plan for Change</div>`n"
$rptSectionFourTwo += "<div class='content'>`n"

#Plan for change Message center messages
[array]$MessagesPFC = @()
[array]$rptMessagesPFCTable = @()
[string]$addText = ""

#Some PFC articles have no milestone.
$rptMessagesPFCTable = "<div class='tableInc'>`n"

$MessagesPFC = $allMessages | Where-Object { ($_.messagetype -like 'MessageCenter' -and $_.category -like 'Plan for Change') }
foreach ($item in $MessagesPFC) {
    if (!($item.properties.name -contains 'MileStoneDate')) { $MilestoneDate = $item.EndTime }
}
$MessagesPFC = $MessagesPFC | Sort-Object MilestoneDate -Descending
if ($MessagesPFC.count -ge 1) {
    $rptMessagesPFCTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-dt'>Feature</div>`n`t<div class='tableInc-header-dt'>Severity</div>`n`t<div class='tableInc-header-dt'>Action</div>`n`t<div class='tableInc-header-dt'>ID</div>`n`t<div class='tableInc-header-l'>Title</div>`n`t<div class='tableInc-header-dt'>Milestone</div>`n<div class='tableInc-header-dt'>Action Rqd By</div>`n</div>`n"
    foreach ($item in $MessagesPFC) {
        $ActionRequiredByDate = $null
        $MilestoneDate = $null
        $LastUpdated = $null
        $addText = ""
        $pubWindow = $null
        if ($item.MilestoneDate) { $MilestoneDate = $(Get-Date $item.MilestoneDate -f 'dd-MMM-yyyy HH:mm') }
        else { $MilestoneDate = $(Get-Date $item.EndTime -f 'dd-MMM-yyyy HH:mm') }
        if ($item.ActionRequiredByDate) { $ActionRequiredByDate = $(Get-Date $item.ActionRequiredByDate -f 'dd-MMM-yyyy HH:mm') }
        $LastUpdated = $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')

        #New text to alert that message is new (24hrs) or updated (7 days)
        $pubWindow = (New-TimeSpan -Start (Get-Date $item.LastUpdatedTime) -End $(Get-Date)).TotalDays
        if ($pubWindow -le 7) { $addtext = "*Updated*" }
        $pubWindow = (New-TimeSpan -Start (Get-Date $item.StartTime)  -End $(Get-Date)).TotalHours
        if ($pubWindow -le 48) { $addtext = "**New**" }

        $Workloads = @()
        $Workloads = $item | Select-Object -ExpandProperty AffectedWorkloadDisplaynames
        if (!($Workloads)) { $Workloads = "General" }
        $workloads = $workloads -join "</br>"
        $rptMessagesPFCTable += "<div class='tableInc-row'>`n`t"
        $rptMessagesPFCTable += "<div class='tableInc-cell-dt'>$($Workloads)</div>`n`t"
        $rptMessagesPFCTable += "<div class='tableInc-cell-dt'>$($item.Severity)</div>`n`t"
        $rptMessagesPFCTable += "<div class='tableInc-cell-dt'>$($item.ActionType)</div>`n`t"
        $link = Get-AdvisoryInHTML $item $RebuildDocs $pathHTMLDocs
        $rptMessagesPFCTable += "<div class='tableInc-cell-dt'>$($item.ID)&nbsp$($addText)</div>`n`t"
        if ($link) { $rptMessagesPFCTable += "<div class='tableInc-cell-l'><a href='$($link)' target='_blank'>$($item.title)</a></div>`n`t" }
        else { $rptMessagesPFCTable += "<div class='tableInc-cell-l'>$($item.title)</div>`n`t" }
        $rptMessagesPFCTable += "<div class='tableInc-cell-dt'>$($MilestoneDate)</div>`n`t"
        $rptMessagesPFCTable += "<div class='tableInc-cell-dt'>$($ActionRequiredByDate)</div>`n"
        $rptMessagesPFCTable += "</div>`n"
    }
}
else {
    $rptMessagesFixTable = "<div class='tableInc'>`n"
    $rptMessagesFixTable += "<div class='tableInc-title'>No Plan for Change Messages</div>`n"
}
$rptMessagesPFCTable += "</div>`n"

$rptSectionFourTwo += $rptMessagesPFCTable
$rptSectionFourTwo += "</div></div>`n"
$divFour += $rptSectionFourTwo

$rptSectionFourThree = "<div class='section'><div class='header'>Other Messages</div>`n"
$rptSectionFourThree += "<div class='content'>`n"

#Remaining message center messages
[array]$HistoryMessages = @()
[array]$rptMessagesTable = @()
$HistoryMessages = $allMessages | Where-Object { ($_.messagetype -like 'MessageCenter' -and $_.category -notlike 'Plan for Change' -and $_.category -notlike 'Prevent or Fix Issues') } | Sort-Object MilestoneDate -Descending
$rptMessagesTable = "<div class='tableInc'>`n"
if ($HistoryMessages.count -ge 1) {
    $rptMessagesTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-dt'>Feature</div>`n`t<div class='tableInc-header-dt'>Category</div>`n`t<div class='tableInc-header-dt'>Severity</div>`n`t<div class='tableInc-header-dt'>ID</div>`n`t<div class='tableInc-header-l'>Title</div>`n`t<div class='tableInc-header-dt'>Milestone</div>`n</div>`n"
    foreach ($item in $HistoryMessages) {
        if ($item.LastUpdatedTime) { $LastUpdated = $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm') }
        if ($item.MilestoneDate) { $MilestoneDate = $(Get-Date $item.MilestoneDate -f 'dd-MMM-yyyy HH:mm') }
        $Workloads = @()
        $Workloads = $item | Select-Object -ExpandProperty AffectedWorkloadDisplaynames
        if (!($Workloads)) { $Workloads = "General" }
        $workloads = $workloads -join "</br>"
        $rptMessagesTable += "<div class='tableInc-row'>`n`t"
        $rptMessagesTable += "<div class='tableInc-cell-dt'>$($Workloads)</div>`n`t"
        $rptMessagesTable += "<div class='tableInc-cell-dt'>$($item.Category)</div>`n`t"
        $rptMessagesTable += "<div class='tableInc-cell-dt'>$($item.Severity)</div>`n`t"
        $link = Get-AdvisoryInHTML $item $RebuildDocs $pathHTMLDocs
        $rptMessagesTable += "<div class='tableInc-cell-dt'>$($item.ID)</div>`n`t"
        if ($link) { $rptMessagesTable += "<div class='tableInc-cell-l'><a href='$($link)' target='_blank'>$($item.title)</a></div>`n`t" }
        else { $rptMessagesTable += "<div class='tableInc-cell-l'>$($item.title)</div>`n`t" }
        $rptMessagesTable += "<div class='tableInc-cell-dt'>$($MilestoneDate)</div>`n"
        $rptMessagesTable += "</div>`n"
    }
}
else {
    $rptMessagesTable = "<div class='tableInc'>`n"
    $rptMessagesTable += "<div class='tableInc-title'>No Previous Messages</div>`n"
}
$rptMessagesTable += "</div>`n"

$rptSectionFourThree += $rptMessagesTable
$rptSectionFourThree += "</div></div>`n"
$divFour += $rptSectionFourThree

#Build Div5
$rptSectionFiveOne = "<div class='section'><div class='header'>Licences</div>`n"
$rptSectionFiveOne += "<div class='content'>`n"
$rptLicenceDash = "<div class='container'>`n"
foreach ($sku in $allLicences) {
    [string]$cardDetail = ""
    [string]$cardClass = ""
    [string]$NicePartNumber = $null
    $NicePartNumber = ($skunames.GetEnumerator() | Where-Object { $_.name -like "$($sku.skupartnumber)" }).Value
    if ($NicePartNumber -eq "") { $NicePartNumber = $($sku.SkuPartNumber) }
    $NicePartNumber += "<br/> $($sku.consumedUnits)/$(($sku.prepaidunits).enabled) used"
    [int]$intPlanCount = 0
    foreach ($serviceplan in $sku.serviceplans) {
        [string]$NiceServiceName = $null
        $NiceServiceName = $SkuNames.item($serviceplan.serviceplanname)
        if ($NiceServiceName -eq "") { $NiceServiceName = $($serviceplan.serviceplanname) }
        $cardClass = get-statusdisplay $($serviceplan.provisioningStatus) 'ServicePlanStatus'
        $cardDetail += "<div class='sku-item-$($cardClass)'>$($NiceServiceName)<span class='tooltiptext'>$($serviceplan.provisioningStatus)</span></div>`r`n"
        if (($serviceplan.serviceplanname).length -gt 29) { $intPlanCount += 2 } else { $intPlanCount++ }
    }
    $cardClass = Get-StatusDisplay $($sku.CapabilityStatus) 'SkuCapabilityStatus'
    $cardText = SkuCardBuilder $($NicePartNumber) $cardDetail $cardClass $intPlanCount
    $rptLicenceDash += "$cardText`n"
}
$rptLicenceDash += "</div>`n<br/><br/>"
$rptSectionFiveOne += $rptLicenceDash
$rptSectionFiveOne += "</div></div>`r`n"
$divFive = $rptSectionFiveOne

#Tab 6


$rptWebTests = OnlineEndPoints $diagWeb $diagPorts $diagURLs
#$rptWebTests = ""

$rptSectionSixOne = "<div class='section'><div class='header'>Office 365 Message Data</div>`n"
$rptSectionSixOne += "<div class='content'>`n"
$rptSectionSixOne += "$($rptO365Info)"
$rptSectionSixOne += "$($diagNotes)"
$rptSectionSixOne += "</div></div>`n"
$divSix = $rptSectionSixOne

$rptSectionSixTwo = "<div class='section'><div class='header'>Diagnostics - Microsoft URLs</div>`n"
$rptSectionSixTwo += "<div class='content'>`n"
$rptSectionSixTwo += "$($rptIPURLs)"
$rptSectionSixTwo += "</div></div>`n"

$divSix += $rptSectionSixTwo

$rptSectionSixThree = "<div class='section'><div class='header'>Diagnostics - Misc URL/Ports</div>`n"
$rptSectionSixThree += "<div class='content'>`n"
$rptSectionSixThree += "$($rptWebTests)"
$rptSectionSixThree += "</div></div>`n"

$divSix += $rptSectionSixThree

#Tab 7 - Network Changes
#Retrieve latest Office 365 service instances

$rptSectionSevenOne = "<div class='section'><div class='header'>Versions Information</div>`n"
$rptSectionSevenOne += "<div class='content'>`n"
[string]$ipurlVersion = "<b>IP and URL Version information</b><br />"
$rptSectionSevenOne += $ipurlVersion
$rptSectionSevenOne += $ipurlSummary
$rptSectionSevenOne += "</div></div>`n"

$divSeven = $rptSectionSevenOne

$rptSectionSevenTwo = "<div class='section'><div class='header'>Current IP and URL Information</div>`n"
$rptSectionSevenTwo += "<div class='content'>`n"
[string]$ipurlCurrent = "<b>Current IP and URL information</b><br />"
$rptSectionSevenTwo += $ipurlCurrent
$rptSectionSevenTwo += $ipurlOutput
$rptSectionSevenTwo += "</div></div>`n"

$divSeven += $rptSectionSevenTwo

$rptSectionSevenThree = "<div class='section'><div class='header'>IP and URL History</div>`n"
$rptSectionSevenThree += "<div class='content'>`n"
[string]$ipurlHistory = "<b>IP and URL history of changes</b><br />"
$rptSectionSevenThree += "</div></div>`n"

$divSeven += $rptSectionSevenThree

#Tab 8 - Office 365 RSS Feed
$rptSectionEightOne = "<div class='section'><div class='header'>Microsoft 365 Roadmap</div>`n"
$rptSectionEightOne += "<div class='content'>`n"
$rptSectionEightOne += "Last $($maxFeedItems) items. Full roadmap can be viewed here: <a href='https://www.microsoft.com/en-us/microsoft-365/roadmap' target=_blank>https://www.microsoft.com/en-us/microsoft-365/roadmap</a><br/>`r`n"
$Roadmap = $Roadmap.replace("ï»¿", "")
[xml]$content = $Roadmap
$feed = $content.rss.channel
$feedMessages = @{ }
$feedMessages = foreach ($msg in $feed.Item) {
    $description = $msg.description
    $description = $description -replace ("`n", '<br>')
    $description = $description -replace ([char]226, "'")
    $description = $description -replace ([char]128, "")
    $description = $description -replace ([char]153, "")
    $description = $description -replace ([char]162, "")
    $description = $description -replace ([char]194, "")
    $description = $description -replace ([char]195, "")
    $description = $description -replace ([char]8217, "'")
    $description = $description -replace ([char]8220, '"')
    $description = $description -replace ([char]8221, '"')
    $description = $description -replace ('\[', '<b><i>')
    $description = $description -replace ('\]', '</i></b>')

    [PSCustomObject]@{
        'LastUpdated' = [datetime]$msg.updated
        'Published'   = [datetime]$msg.pubDate
        'Description' = $description
        'Title'       = $msg.Title
        'Category'    = $msg.Category
        'Link'        = $msg.link
    }
}

$feedMessages = $feedmessages | Sort-Object published -Descending | Select-Object -First $maxFeedItems

if ($feedMessages.count -ge 1) {
    $rptFeedTable += "<div class='tableFeed'>`n"
    $rptFeedTable += "<div class='tableFeed-title'>Microsoft 365 RoadMap</div>`n"
    $rptFeedTable += "<div class='tableFeed-header'>`n`t<div class='tableFeed-header-c'>Category</div>`n`t<div class='tableFeed-header-c'>Title</div>`n`t<div class='tableFeed-header-c'>Description</div>`n`t<div class='tableFeed-header-c'>Published</div>`n`t<div class='tableFeed-header-c'>Last Updated</div>`n</div>`n"
    foreach ($item in $feedMessages) {
        if ($item.LastUpdated) { $LastUpdated = $(Get-Date $item.LastUpdated -f 'dd-MMM-yyyy HH:mm') } else { $StartTime = "" }
        if ($item.Published) { $Published = $(Get-Date $item.Published -f 'dd-MMM-yyyy HH:mm') } else { $EndTime = "" }
        $link = $item.Link
        #Build link to detailed message
        #$link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($item.link) {
            $ID = "<a href=$($item.link) target=_blank>$($item.Title)</a>"
        }
        else { $ID = "$($item.Title)" }
        $rptFeedTable += "<div class='tableFeed-row'>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-cat'>$($item.Category -join ' | ')</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-title'>$($ID)</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-desc'>$($item.description)</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-dt' $($tdStyle2)>$($Published)</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-dt' $($tdStyle2)>$($LastUpdated)</div>`n`t"
        $rptFeedTable += "</div>`n"
    }
}
#Close tablefeed
$rptFeedTable += "</div>"
$rptSectionEightOne += $rptFeedTable
$rptSectionEightOne += "</div></div>`n"

$divEight = $rptSectionEightOne

$rptSectionEightTwo = "<div class='section'><div class='header'>Azure Updates</div>`n"
$rptSectionEightTwo += "<div class='content'>`n"

#Azure Updates URI: https://azurecomcdn.azureedge.net/en-gb/updates/feed/
$rptSectionEightTwo += "Last 20 items. Full roadmap can be viewed here: <a href='https://azure.microsoft.com/en-gb/updates/' target=_blank>https://azure.microsoft.com/en-gb/updates/</a><br/>`r`n"
$AzureUpdates = $AzureUpdates.replace("ï»¿", "")
[xml]$content = $AzureUpdates
$feed = $content.rss.channel
$feedMessages = @{ }
$feedMessages = foreach ($msg in $feed.Item) {
    $description = $msg.description
    $description = $description -replace ("`n", '<br>')
    $description = $description -replace ([char]226, "'")
    $description = $description -replace ([char]128, "")
    $description = $description -replace ([char]153, "")
    $description = $description -replace ([char]162, "")
    $description = $description -replace ([char]194, "")
    $description = $description -replace ([char]195, "")
    $description = $description -replace ([char]8217, "'")
    $description = $description -replace ([char]8220, '"')
    $description = $description -replace ([char]8221, '"')
    $description = $description -replace ('\[', '<b><i>')
    $description = $description -replace ('\]', '</i></b>')

    [PSCustomObject]@{
        'Published'   = [datetime]$msg.pubDate
        'Description' = $description
        'Title'       = $msg.Title
        'Category'    = $msg.Category
        'Link'        = $msg.link
    }
}

$feedMessages = $feedmessages | Sort-Object published -Descending | Select-Object -First $maxFeedItems
$rptFeedTable = $null
if ($feedMessages.count -ge 1) {
    $rptFeedTable += "<div class='tableFeed'>`n"
    $rptFeedTable += "<div class='tableFeed-title'>Azure Updates</div>`n"
    $rptFeedTable += "<div class='tableFeed-header'>`n`t<div class='tableFeed-header-c'>Category</div>`n`t<div class='tableFeed-header-c'>Title</div>`n`t<div class='tableFeed-header-c'>Description</div>`n`t<div class='tableFeed-header-c'>Published</div>`n`t</div>`n"
    foreach ($item in $feedMessages) {
        if ($item.Published) { $Published = $(Get-Date $item.Published -f 'dd-MMM-yyyy HH:mm') } else { $EndTime = "" }
        $link = $item.Link
        #Build link to detailed message
        #$link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($item.link) {
            $ID = "<a href=$($item.link) target=_blank>$($item.Title)</a>"
        }
        else { $ID = "$($item.Title)" }
        $rptFeedTable += "<div class='tableFeed-row'>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-cat'>$($item.Category -join ' | ')</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-title'>$($ID)</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-desc'>$($item.description)</div>`n`t"
        $rptFeedTable += "<div class='tableFeed-cell-dt' $($tdStyle2)>$($Published)</div>`n`t"
        $rptFeedTable += "</div>`n"
    }
}

#Close tablefeed
$rptFeedTable += "</div>"
$rptSectionEightTwo += $rptFeedTable
$rptSectionEightTwo += "</div></div>`n"
$divEight += $rptSectionEightTwo

$rptSectionEightThree = "<div class='section'><div class='header'>Information</div>`n"
$rptSectionEightThree += "<div class='content'>`n"
$rptSectionEightThree += "</div></div>`n"

$divEight += $rptSectionEightThree


$rptHTMLName = ($rptName.Replace(" ", "") + ".html")
$rptTitle = $rptTenantName + " " + $rptName
if ($rptOutage) { $rptTitle += " Outage detected" }
$evtMessage = "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] Tenant: $($rptProfile) - Generating HTML to '$($pathHTML)\$($rptHTMLName)'`r`n"
$evtLogMessage += $evtMessage
Write-Verbose $evtMessage

BuildHTML $rptTitle $divOne $divTwo $divThree $divFour $divFive $divSix $divSeven $divEight $swScript.Elapsed $rptHTMLName
#Check if .css file exists in HTML file destination
if (!(Test-Path $cssfile)) {
    Write-Log "Copying O365Health.css to directory $($pathHTML)"
    Copy-Item "O365Health.css" -Destination "$($pathHTML)"
}

$swScript.Stop()

$evtMessage = "Tenant: $($rptProfile) - Script runtime $($swScript.Elapsed.Minutes)m:$($swScript.Elapsed.Seconds)s:$($swScript.Elapsed.Milliseconds)ms on $env:COMPUTERNAME"
$evtMessage += "*** Processing finished ***`r`n"
Write-Log $evtMessage

#Append to daily log file.
Get-Content $script:logfile | Add-Content $script:Dailylogfile
Remove-Item $script:logfile
Remove-Module O365ServiceHealth