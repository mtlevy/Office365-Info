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
[string]$emailEnabled = $config.EmailEnabled
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
[string]$pathWorking = $config.WorkingPath
[string]$HTMLFile = $config.DashboardHTML
[string]$emailDashAlertsTo = $config.DashboardAlertsTo

[string]$htmlWall = $config.WallHTML
[string]$htmlDiagnostics = $config.ToolboxHTML
[string]$htmlClientDiagnostics = "ClientDiags.txt"

[string]$proxyHost = $config.ProxyHost

if ($config.RSS1Enabled -like 'true') {
    [boolean]$rss1Enabled = $true
}
[string]$rss1Name = $config.RSS1Name
[string]$rss1Feed = $config.RSS1Feed
[string]$rss1URL = $config.RSS1URL
[int]$rss1Items = $config.RSS1Items

if ($config.RSS2Enabled -like 'true') {
    [boolean]$rss2Enabled = $true
}
[string]$rss2Name = $config.RSS2Name
[string]$rss2Feed = $config.RSS2Feed
[string]$rss2URL = $config.RSS2URL
[int]$rss2Items = $config.RSS2Items

[string]$Blogs = $config.Blogs

[boolean]$rptOutage = $false

[string]$cssfile = "O365Health.css"

if ($config.EmailEnabled -like 'true') { [boolean]$emailEnabled = $true } else { [boolean]$emailEnabled = $false }

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
$pathWorking = CheckDirectory $pathWorking
$pathHTMLDocs = CheckDirectory "$($pathHTML)\Docs"
$pathHTMLImg = CheckDirectory "$($pathHTML)\images"



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

$authHeader = @{
    'Content-Type' = 'application/json'
    Authorization  = "$($bearerToken.token_type) $($bearerToken.access_token)"
}

if ($null -eq $bearerToken) {
    $evtMessage = "ERROR - No authentication result for Auzre AD App"
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
    <button class="tablinks" onclick="openTab(event,'Messages')">Messages</button>
    <button class="tablinks" onclick="openTab(event,'Roadmap')">Roadmaps</button>
    <button class="tablinks" onclick="openTab(event,'Log')">Log</button>
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

<div id="Messages" class="tabcontent">
    $($contentFour)
</div>

<div id="Roadmap" class="tabcontent">
    $($contentFive)
</div>

<div id="Log" class="tabcontent">
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

#https://docs.microsoft.com/en-gb/azure/active-directory/users-groups-roles/licensing-service-plan-reference


#	Returns the list of subscribed services
[uri]$uriServices = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/Services"
#	Returns the current status of the service.
[uri]$uriCurrentStatus = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/CurrentStatus"
#	Returns the historical status of the service, by day, over a certain time range.
[uri]$uriHistoricalStatus = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/HistoricalStatus"
#	Returns the messages about the service over a certain time range.
# Range was 30 days but has been shorted to 7 days by default. Filters no longer seem to work
[uri]$uriMessages = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/Messages?$filater=StartTime ge $(Get-Date).AddDays(-30)"

#Fetch the information from Office 365 Service Health API
#Get Services: Get the list of subscribed services
$uriError = ""
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
    $uriError += "Error connecting to $($uriservices)<br/><br/>`r`n"
    $uriError += "Error details:<br/> $($Error[0] | Select-Object *)"
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
    $uriError += "Error connecting to $($uriCurrentStatus)<br/><br/>`r`n"
    $uriError += "Error details:<br/> $($Error[0] | Select-Object *)"
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
    $uriError += "Error connecting to $($uriHistoricalStatus)<br/><br/>`r`n"
    $uriError += "Error details:<br/> $($Error[0] | Select-Object *)"
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
    $uriError += "Error connecting to $($uriMessages)<br/><br/>`r`n"
    $uriError += "Error details:<br/> $($Error[0] | Select-Object *)"
}

if ($rss1Enabled) {
    try {
        if ($proxyServer) {
            $rss1Data = @((Invoke-WebRequest -Uri $rss1Feed -Proxy $proxyHost -ProxyUseDefaultCredentials -UseBasicParsing).content)
        }
        else {
            $rss1Data = @((Invoke-WebRequest -Uri $rss1Feed -UseBasicParsing).content)
        }
        if ($null -eq $rss1Data -or $rss1Data.Count -eq 0) {
            $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No $($rss1Name) RSS Feed information - verify proxy and network connectivity</p><br/>"
        }
        else {
            $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>RSS feed items retrieved for $($rss1Name).</p><br/>"
        }
    }
    catch {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No $($rss1Name) RSS Feed information - verify proxy and network connectivity</p><br/>"
        $uriError += "Error connecting to $($rss1URL)<br/><br/>`r`n"
        $uriError += "Error details:<br/> $($Error[0] | Select-Object *)"
    }
}

if ($rss2Enabled) {
    try {
        if ($proxyServer) {
            $rss2Data = @((Invoke-WebRequest -Uri $rss2Feed -Proxy $proxyHost -ProxyUseDefaultCredentials -UseBasicParsing).content)
        }
        else {
            $rss2Data = @((Invoke-WebRequest -Uri $rss2Feed -UseBasicParsing).content)
        }
        if ($null -eq $rss2Data -or $rss2Data.Count -eq 0) {
            $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No $($rss2Name) RSS Feed information - verify proxy and network connectivity</p><br/>"
        }
        else {
            $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='info'>RSS feed items retrieved for $($rss2Name).</p><br/>"
        }
    }
    catch {
        $rptO365Info += "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] <p class='error'>No $($rss2Name) RSS Feed information - verify proxy and network connectivity</p><br/>"
        $uriError += "Error connecting to $($uriAzureUpdates)<br/><br/>`r`n"
        $uriError += "Error details:<br/> $($Error[0] | Select-Object *)"
    }
}

if ($emailEnabled -and $uriError) {
    $emailSubject = "Error(s) retrieving URL(s)"
    SendEmail $uriError $EmailCreds $config "High" $emailSubject $emailDashAlertsTo
}

$rptO365Info += "<br/>"
$rptO365Info += "Information wall can be found here: <a href=$($htmlWall) target=_blank>Information Wall</a><br />"
$rptO365Info += "Diagnostics can be found here: <a href=$($htmlDiagnostics) target=_blank>Diagnostics page</a><br />"
if (test-path "$($pathHTML)\$($htmlClientDiagnostics)") {
    $rptO365Info += "Client Diagnostics can be download from here: <a href=$($htmlClientDiagnostics) target=_blank>Rename from .txt to .ps1</a><br />"
}

if ($addLink) { $rptO365Info += "<a href='$($addLink)' target=_blank> here </a></li></ul><br>" }

#Start Building the Pages
#Build Div1
#Build Summary Dashboard
# 6 cards
$HistoryIncidents = $allMessages | Where-Object { ($_.EndTime -ne $null -and $_.messagetype -notlike 'MessageCenter') } | Sort-Object EndTime -Descending
$rptSectionOneOne = "<div class='section'><div class='header'>Office 365 Dashboard Status</div>`n"
$rptSectionOneOne += "<div class='content'>`n"
$rptSectionOneOne += "<div class='dash-outer'><div class='dash-inner'>`n"
$card = $null
foreach ($card in $dashCards) {
    [array]$item = @()
    [array]$hist = @()
    [int]$advisories = 0
    $cardClass = $null
    $cardText = $null
    $item = $allCurrentStatusMessages | Where-Object { $_.WorkloadDisplayName -like $card }
    $hist = $HistoryIncidents | Where-Object { $_.WorkloadDisplayName -like $card -and ($_.status -notlike 'False Positive') } | Sort-Object EndTime -Descending
    $advisories = ($allMessages | Where-Object { ($_.messagetype -like 'MessageCenter' -and $_.AffectedWorkloadDisplayNames -like $card) }).count
    if ($hist.count -gt 0) {
        $days = "{0:N0}" -f (New-TimeSpan -Start (Get-Date $hist[0].EndTime) -End $(Get-Date)).TotalDays
    }
    else {
        $days = "&gt;30"
    }
    try { $cardClass = Get-StatusDisplay $($item.status) "Class" }
    catch { Write-Log "No status available for $card - $($item.workloaddisplayname)" }
    try { $cardText = cardbuilder $($item.workloaddisplayname) $($Days) $($Hist.count) $advisories $cardClass }
    catch { Write-Log "Cant find $card in workload. Has name changed or workload replaced?" }
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
            $ID = "<a href=$($link) target=_blank>[$($item.ID)] $($item.ImpactDescription)</a>"
        }
        else { $ID = "[$($item.ID)] $($item.ImpactDescription)" }
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
            $ID = "<a href=$($link) target=_blank>[$($item.ID)] $($item.ImpactDescription)</a>"
        }
        else { $ID = "[$($item.ID)] $($item.ImpactDescription)" }
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
$listTheRest = $allCurrentStatusMessages | Where-Object { $_.workloaddisplayname -notin $listlineone.workloaddisplayname } | Sort-Object  status, workloaddisplayname
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
$rptSectionThreeOne = "<div class='section'><div class='header'>Service Health notes</div>`n"
$rptSectionThreeOne += "<div class='content'>`n"
$rptSectionThreeOne += "<b>Microsoft definitions of 'Incident' and 'Advisory'</b></br>`n"
$rptSectionThreeOne += "An incident is a critical service issue, typically involving noticable user impact.</br>`n"
$rptSectionThreeOne += "An advisory is a service issue that is typically limited in scope or impact.</br>`n"
$rptSectionThreeOne += "</br>Microsoft Severity ranges are typically Sev 0 (Critical), Sev 1 (Error) and Sev2 (Warning)</br>`n"

$rptSectionThreeOne += "</div></div>`n"
$divThree = $rptSectionThreeOne

$rptSectionThreeTwo = "<div class='section'><div class='header'>Office 365 Open Incidents</div>`n"
$rptSectionThreeTwo += "<div class='content'>`n"

#Incident History
#Get all open Incidents
$rptIncidentTable = @()
$item = $null

if ($CurrentMessagesOpen.count -ge 1) {
    $rptIncidentTable += "<div class='tableInc'>`n"
    $rptIncidentTable += "<div class='tableInc-title'>Open Incidents</div>`n"
    $rptIncidentTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-c'>Feature</div>`n`t<div class='tableInc-header-c'>Severity</div>`n`t<div class='tableInc-header-c'>Status</div>`n`t<div class='tableInc-header-c'>Description</div>`n`t<div class='tableInc-header-c'>Start Time</div>`n`t<div class='tableInc-header-c'>Last Updated</div>`n</div>`n"
    foreach ($item in $CurrentMessagesOpen) {
        if ($item.StartTime) { $StartTime = $(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm') } else { $StartTime = "" }
        if ($item.LastUpdatedTime) { $LastUpdated = $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm') } else { $LastUpdated = "" }
        $severity = $item.severity
        switch ($severity) {
            "SEV0" { $actionStyle = "style=border:none;text-align:center;font-weight:bold;color:red" }
            "SEV1" { $actionStyle = "style=border:none;text-align:center;color:red" }
            "SEV2" { $actionStyle = "style=border:none;text-align:center;color:blue" }
            default { $actionStyle = "style=border:none;text-align:center;font-weight:bold;color:red" }
        }
        $link = ""
        #Build link to detailed message
        $link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($link) {
            $ID = "<a href=$($link) target=_blank>$($item.ID) - $($item.ImpactDescription)</a>"
        }
        else { $ID = "$($item.ID) - $($item.ImpactDescription)" }
        $rptIncidentTable += "<div class='tableInc-row'>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-l'>$($item.WorkloadDisplayname -join '<br>')</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-r' $($actionStyle)>$($item.classification) - $($Severity)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-l'>$($item.Status)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-l'>$($ID)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($StartTime)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($LastUpdated)</div>`n"
        $rptIncidentTable += "</div>`n"
    }
}
else {
    $rptIncidentTable = "<div class='tableInc'>`n"
    $rptIncidentTable += "<div class='tableInc-title'>No Open Incidents</div>`n"
}
$rptIncidentTable += "</div>`n"
$rptSectionThreeTwo += $rptIncidentTable
$rptSectionThreeTwo += "</div></div>`n"
$divThree += $rptSectionThreeTwo


$rptSectionThreeThree = "<div class='section'><div class='header'>Office 365 Closed Incidents</div>`n"
$rptSectionThreeThree += "<div class='content'>`n"

#Get Incident History
[array]$HistoryIncidents = @()
$rptIncidentTable = @()

$item = $null
$HistoryIncidents = $allMessages | Where-Object { ($_.EndTime -ne $null -and $_.messagetype -notlike 'MessageCenter') } | Sort-Object EndTime -Descending
if ($HistoryIncidents.count -ge 1) {
    $rptIncidentTable += "<div class='tableInc'>`n"
    $rptIncidentTable += "<div class='tableInc-title'>Closed Incidents</div>`n"
    $rptIncidentTable += "<div class='tableInc-header'>`n`t<div class='tableInc-header-c'>Feature</div>`n`t<div class='tableInc-header-c'>Severity</div>`n`t<div class='tableInc-header-c'>Status</div>`n`t<div class='tableInc-header-c'>Description</div>`n`t<div class='tableInc-header-c'>Start Time</div>`n`t<div class='tableInc-header-c'>End Time</div>`n`t<div class='tableInc-header-c'>Last Updated</div>`n</div>`n"
    foreach ($item in $HistoryIncidents) {
        if ($item.StartTime) { $StartTime = $(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm') } else { $StartTime = "" }
        if ($item.EndTime) { $EndTime = $(Get-Date $item.EndTime -f 'dd-MMM-yyyy HH:mm') } else { $EndTime = "" }
        if ($item.LastUpdatedTime) { $LastUpdated = $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm') } else { $LastUpdated = "" }
        $severity = $item.severity
        switch ($severity) {
            "SEV0" { $actionStyle = "style=border:none;text-align:center;font-weight:bold;color:red" }
            "SEV1" { $actionStyle = "style=border:none;text-align:center;color:red" }
            "SEV2" { $actionStyle = "style=border:none;text-align:center;color:blue" }
            default { $actionStyle = "style=border:none;text-align:center;font-weight:bold;color:red" }
        }
        $link = ""
        #Build link to detailed message
        $link = Get-IncidentInHTML $item $RebuildDocs $pathHTMLDocs
        if ($link) {
            $ID = "<a href=$($link) target=_blank>$($item.ID) - $($item.ImpactDescription)</a>"
        }
        else { $ID = "$($item.ID) - $($item.ImpactDescription)" }
        $rptIncidentTable += "<div class='tableInc-row'>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-l'>$($item.WorkloadDisplayname -join '<br>')</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-r' $($actionStyle)>$($item.classification) - $($Severity)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-l'>$($item.Status)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-l'>$($ID)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($StartTime)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($EndTime)</div>`n`t"
        $rptIncidentTable += "<div class='tableInc-cell-dt' $($tdStyle2)>$($LastUpdated)</div>`n"
        $rptIncidentTable += "</div>`n"
    }
}
else {
    $rptIncidentTable += "<div class='tableInc'>`n"
    $rptIncidentTable += "<div class='tableInc-title'>No Closed Incidents</div>`n"
}
$rptIncidentTable += "</div>`n"
$rptSectionThreeThree += $rptIncidentTable
$rptSectionThreeThree += "</div></div>`n"

$divThree += $rptSectionThreeThree

#Build Div4
#Get current messages
[array]$advNew = @()
[string]$advPath = ""
[string]$advFileName = ""
$allAdvisories = @($allMessages | Where-Object { ($_.messagetype -like 'MessageCenter') } | Sort-Object MilestoneDate -Descending)
#Get previously downloaded messages
$advFileName = "Advisories-$($rptProfile).csv"
$advPath = "$($pathWorking)\$($advFileName)"
if (Test-Path $advPath) {
    $advExisting = Import-Csv $advPath
}
$advNew = @($allAdvisories | Where-Object { ($_.id -notin $advExisting.id) })
if ($advNew.count -gt 0) {
    Write-Log "Adding $($advNew.count) new advisories to local file"
    foreach ($message in $advNew) {
        # Build new exportable list
        $dtmStartTime = $null
        $dtmLastUpdatedTime = $null
        $dtmEndTime = $null
        $dtmActionRequiredByDate = $null
        $dtmMilestoneDate = $null

        if ($null -ne $message.StartTime) { $dtmStartTime = $(Get-Date $message.StartTime -f 'dd-MMM-yyyy HH:mm') } 
        if ($null -ne $message.LastUpdatedTime) { $dtmLastUpdatedTime = $(Get-Date $message.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm') } 
        if ($null -ne $message.EndTime) { $dtmEndTime = $(Get-Date $message.EndTime -f 'dd-MMM-yyyy HH:mm') } 
        if ($null -ne $message.ActionRequiredByDate) { $dtmActionRequiredByDate = $(Get-Date $message.ActionRequiredByDate -f 'dd-MMM-yyyy HH:mm') } 
        if ($null -ne $message.MilestoneDate) { $dtmMilestoneDate = $(Get-Date $message.MilestoneDate -f 'dd-MMM-yyyy HH:mm') } 
        $advTemp = New-Object PSObject
        $advTemp | Add-Member -MemberType NoteProperty -Name ID -Value $message.ID
        $advTemp | Add-Member -MemberType NoteProperty -Name ActionType -Value $message.Actiontype
        $advTemp | Add-Member -MemberType NoteProperty -Name Classification -Value $message.Classification
        $advTemp | Add-Member -MemberType NoteProperty -Name StartTime -Value $($dtmStartTime)
        $advTemp | Add-Member -MemberType NoteProperty -Name LastUpdatedTime -Value $($dtmLastUpdatedTime)
        $advTemp | Add-Member -MemberType NoteProperty -Name EndTime -Value $($dtmEndTime)
        $advTemp | Add-Member -MemberType NoteProperty -Name ActionRequiredByDate -Value $($dtmActionRequiredByDate)
        $advTemp | Add-Member -MemberType NoteProperty -Name MessageType -Value $message.MessageType
        $advTemp | Add-Member -MemberType NoteProperty -Name PostIncidentDocumentURL -Value $message.PostIncidentDocumentURL
        $advTemp | Add-Member -MemberType NoteProperty -Name Severity -Value $message.Severity
        $advTemp | Add-Member -MemberType NoteProperty -Name Title -Value $message.Title
        $advTemp | Add-Member -MemberType NoteProperty -Name Category -Value $message.Category
        $advTemp | Add-Member -MemberType NoteProperty -Name ExternalLink -Value $message.ExternalLink
        $advTemp | Add-Member -MemberType NoteProperty -Name IsMajorChange -Value $message.IsMajorChange
        $advTemp | Add-Member -MemberType NoteProperty -Name AppliesTo -Value $message.AppliesTo
        $advTemp | Add-Member -MemberType NoteProperty -Name MilestoneDate -Value $($dtmMilestoneDate)
        $advTemp | Add-Member -MemberType NoteProperty -Name Milestone -Value $message.Milestone
        $advTemp | Add-Member -MemberType NoteProperty -Name BlogLink -Value $message.BlogLink
        $advTemp | Add-Member -MemberType NoteProperty -Name HelpLink -Value $message.HelpLink

        $advTemp | Export-Csv -Append -Path "$($advPath)" -NoTypeInformation -Encoding UTF8
    }
}
Copy-Item "$($advPath)" -Destination "$($pathHTML)"

$rptSectionFourOne = "<div class='section'>"
$rptSectionFourOne += "<div class='tableinc-cell-r'>Download all items <a href='$($advFileName)' target='_blank'>here</a></div>"
$rptSectionFourOne += "<div class='header'>Prevent / Fix Issues</div>`n"
$rptSectionFourOne += "<div class='content'>`n"
#Messages
#Prevent or fix issues
[array]$MessagesFix = @()
[array]$rptMessagesFixTable = @()
$MessagesFix = $allMessages | Where-Object { ($_.messagetype -like 'MessageCenter' -and $_.category -like 'Prevent or Fix Issues') } | Sort-Object MilestoneDate -Descending
if ($MessagesFix.count -ge 1) {
    $rptMessagesFixTable += "<div class='tableInc'>`n"
    #    $rptMessagesFixTable += "<div class='tableInc-title'>Closed Incidents</div>`n"
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

#Tab 5 - First RSS Feed
$rptSectionFiveOne = "<!-- First RSS Feed goes here-->"
if ($rss1Enabled) {
    $rptSectionFiveOne += "<div class='section'><div class='header'>$($rss1Name)</div>`n"
    $rptSectionFiveOne += "<div class='content'>`n"
    $rptSectionFiveOne += "Last $($rss1Items) items. Full roadmap can be viewed here: <a href='$($rss1URL)' target=_blank>$($rss1URL)</a><br/>`r`n"
    $rss1Data = $rss1Data.replace("ï»¿", "")
    [xml]$content = $rss1Data
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

    $feedMessages = $feedmessages | Sort-Object published -Descending | Select-Object -First $rss1Items

    if ($feedMessages.count -ge 1) {
        $rptFeedTable += "<div class='tableFeed'>`n"
        $rptFeedTable += "<div class='tableFeed-title'>$($rss1Name)</div>`n"
        $rptFeedTable += "<div class='tableFeed-header'>`n`t<div class='tableFeed-header-c'>Tags</div>`n`t<div class='tableFeed-header-c'>Title</div>`n`t<div class='tableFeed-header-c'>Published</div>`n`t<div class='tableFeed-header-c'>Last Updated</div>`n</div>`n"
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
            $rptFeedTable += "<div class='tableFeed-cell-cat'>$($item.Category -join ', ')</div>`n`t"
            $rptFeedTable += "<div class='tableFeed-cell-desc'><p><div class='tableFeed-cell-title'>$($ID)</div></p>`n`t"
            $rptFeedTable += "<div class='tableFeed-cell-desc'>$($item.description)</div></div>`n`t"
            $rptFeedTable += "<div class='tableFeed-cell-dt' $($tdStyle2)>$($Published)</div>`n`t"
            $rptFeedTable += "<div class='tableFeed-cell-dt' $($tdStyle2)>$($LastUpdated)</div>`n`t"
            $rptFeedTable += "</div>`n"
        }
        #Close tablefeed
        $rptFeedTable += "</div>"
    }
    $rptSectionFiveOne += $rptFeedTable
    $rptSectionFiveOne += "</div></div>`n"
}

$divFive = $rptSectionFiveOne

$rptSectionFiveTwo = "<!-- Second RSS Feed goes here-->"
if ($rss2Enabled) {
    $rptSectionFiveTwo += "<div class='section'><div class='header'>$($rss2Name)</div>`n"
    $rptSectionFiveTwo += "<div class='content'>`n"

    $rptSectionFiveTwo += "Last $($rss2Items) items. Full roadmap can be viewed here: <a href='$($rss2URL)' target=_blank>$($rss2URL)</a><br/>`r`n"
    $rss2Data = $rss2Data.replace("ï»¿", "")
    [xml]$content = $rss2Data
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

    $feedMessages = $feedmessages | Sort-Object published -Descending | Select-Object -First $rss2Items
    $rptFeedTable = $null
    if ($feedMessages.count -ge 1) {
        $rptFeedTable += "<div class='tableFeed'>`n"
        $rptFeedTable += "<div class='tableFeed-title'>$($rss2Name)</div>`n"
        $rptFeedTable += "<div class='tableFeed-header'>`n`t<div class='tableFeed-header-c'>Tags</div>`n`t<div class='tableFeed-header-c'>Title</div>`n`t<div class='tableFeed-header-c'>Published</div>`n`t</div>`n"
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
            $rptFeedTable += "<div class='tableFeed-cell-desc'><p><div class='tableFeed-cell-title'>$($ID)</div></p>`n`t"
            $rptFeedTable += "<div class='tableFeed-cell-desc'>$($item.description)</div></div>`n`t"
            $rptFeedTable += "<div class='tableFeed-cell-dt' $($tdStyle2)>$($Published)</div>`n`t"
            $rptFeedTable += "</div>`n"
        }
        #Close tablefeed
        $rptFeedTable += "</div>"
    }

    $rptSectionFiveTwo += $rptFeedTable
    $rptSectionFiveTwo += "</div></div>`n"
}
$divFive += $rptSectionFiveTwo

$rptSectionFiveThree = "<div class='section'><div class='header'>Useful Blogs</div>`n"
$rptSectionFiveThree += "<div class='content'>`n"
$rptSectionFiveThree += "$($Blogs)"
$rptSectionFiveThree += "</div></div>`n"

$divFive += $rptSectionFiveThree



#Tab Last - Logs / Additional info
$rptSectionLastOne = "<div class='section'><div class='header'>Information</div>`n"
$rptSectionLastOne += "<div class='content'>`n"
$rptSectionLastOne += $rptO365Info
$rptSectionLastOne += "</div></div>`n"

$divLast = $rptSectionLastOne

$rptSectionLastTwo = "<div class='section'><div class='header'>Script Runtime</div>`n"
$rptSectionLastTwo += "<div class='content'>`n"

[string]$strTime = "$($swScript.Elapsed.Hours)H:$($swScript.Elapsed.Minutes)m:$($swScript.Elapsed.Seconds)s:$($swScript.Elapsed.Milliseconds)ms"
$rptSectionLastTwo += "Elapsed runtime $strTime"
$rptSectionLastTwo += "</div></div>`n"
$divLast += $rptSectionLastTwo

$rptHTMLName = $HTMLFile.Replace(" ", "")
$rptTitle = $rptTenantName + " " + $rptName
if ($rptOutage) { $rptTitle += " Outage detected" }
$evtMessage = "[$(Get-Date -f 'dd-MMM-yy HH:mm:ss')] Tenant: $($rptProfile) - Generating HTML to '$($pathHTML)\$($rptHTMLName)'`r`n"
$evtLogMessage += $evtMessage
Write-Verbose $evtMessage

BuildHTML $rptTitle $divOne $divTwo $divThree $divFour $divFive $divLast $rptHTMLName
#Check if .css file exists in HTML file destination
if (!(Test-Path "$($pathHTML)\$($cssfile)")) {
    Write-Log "Copying O365Health.css to directory $($pathHTML)"
    Copy-Item "..\common\O365Health.css" -Destination "$($pathHTML)"
    #No CSS file so probably no images
    Copy-Item ".\images\*.jpg" -Destination "$($pathHTMLimg)"
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