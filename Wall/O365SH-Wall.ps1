<#
.SYNOPSIS
    This script creates an Office 365 Health Center Dashboard Wall.
    It displays all Office 365 Service Workloads with Features and Health

.DESCRIPTION
    This script is designed to run on a scheduled basis.
    It requires and Azure AD Application for your tenant.
    The script will build an HTML page displaying a dashboard of tiles for each workload and its feature status

    Configuration file is required allowing script to be called for multiple tenants.
    
    Requires Azure AD powershell module (Install-Module AzureAD)

    Requires Azure AD Application - Tenant ID, Application ID and Secret.

    Output best viewed (used CSS Grid) in firefox. I'm not a web designer :D

.EXAMPLE
    PS .\O365SH.ps1

    Gets information from the default tenant and generates the HTML status wall

.EXAMPLE
    PS .\O365SH.ps1 -configXML '..\profile-sample.xml'

    Uses the specific XML settings file to load tenant information. The file is specified relative to the location of this script, or absolute location

.NOTES
    Author:  Jonathan Christie
    PSVer:   2.0/3.0/4.0/5.0
    Version: 2.0.1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [String]$configXML = "..\profile-test.xml"
)

$swScript = [system.diagnostics.stopwatch]::StartNew()
Write-Verbose "Changing Directory to $PSScriptRoot"
Set-Location $PSScriptRoot
Import-Module "..\common\O365ServiceHealth.psm1"

function Write-Log {
    param (
        [Parameter(Mandatory = $true)] [string]$info
    )
    # verify the Log is setup and if not create the file
    if ($script:loginitialized -eq $false) {
        $script:FileHeader >> $script:logfile
        $script:loginitialized = $True
    }
    $info = $(get-date).ToString() + ": " + $info
    $info >> $script:logfile
}

if ([system.IO.path]::IsPathRooted($configXML) -eq $false) {
    #its not an absolute path. Find the absolute path
    $configXML = Resolve-Path $configXML
}
$config = LoadConfig $configXML

#Configure local event log
[string]$evtLogname = $config.EventLog
[string]$evtSource = $config.WallEventSource
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
[string]$pathLogs = $config.LogPath
[string]$pathHTML = $config.HTMLPath
[string]$HTMLFile = $config.WallHTML
[string[]]$prefDashCards = $config.WallDashCards.split(",")
$prefDashCards=$prefDashCards.Replace('"','')
$prefDashCards=$prefDashCards.Trim()

#Page refresh in minutes
[int]$pageRefresh = $config.WallPageRefresh

[array]$evtCheck = @()
[string]$evtMessage = $null

[string]$tenantID = $config.TenantID
[string]$appID = $config.AppID
[string]$clientSecret = $config.AppSecret

[String]$rptName = $config.ReportName
[string]$rptProfile = $config.TenantShortName
[string]$rptTenantName = $config.TenantName

#If no path has been specified, use the current script location
if (!$pathLogs) {
    $pathLogs = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
}

#Check and trim the report path
$pathLogs = $pathLogs.TrimEnd("\")
$pathHTML = $pathHTML.TrimEnd("\")
#Build and Check output directories
#Base for logs
if (!(Test-Path $($pathLogs))) {
    New-Item -ItemType Directory -Path $pathLogs
}
#HTML directory for output main page
if (!(Test-Path $($pathHTML))) {
    New-Item -ItemType Directory -Path "$($pathHTML)"
}

if ([system.IO.path]::IsPathRooted($pathLogs) -eq $false) {
    #its not an absolute path. Find the absolute path
    $pathLogs = Resolve-Path $pathLogs
}

# setup the logfile
# If logfile exists, the set flag to keep logfile
$script:DailyLogFile = "$($pathLogs)\O365Wall-$($rptprofile)-$(get-date -format yyMMdd).log"
$script:LogFile = "$($pathLogs)\tmpO365Wall-$($rptprofile)-$(get-date -format yyMMddHHmmss).log"
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
[string]$proxyHost = $config.ProxyHost

#Connect to Azure app and grab the service status
ConnectAzureAD
$urlOrca = "https://manage.office.com"
$authority = "https://login.microsoftonline.com/$TenantID"
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

function BuildHTML {
    Param (
        [Parameter(Mandatory = $true)] [string]$Title,
        [Parameter(Mandatory = $true)] [string]$contentOne,
        [Parameter(Mandatory = $true)] [string]$HTMLOutput
    )
    [array]$htmlHeader = @()
    [array]$htmlBody = @()
    [array]$htmlFooter = @()
    [string]$minCSS = ""
    $minCSS = "*{margin:.1px;font-family:""Segoe UI"",Tahoma,Geneva,Verdana,sans-serif;font-size:8pt}.workload-card-ok{background-color:#060;grid-row:span auto}.workload-card-warn{background-color:orange;grid-row:span auto}.workload-card-err{background-color:red;grid-row:span auto}[class^=workload-card]{display:inline-table;table-layout:fixed;color:#fff;font-weight:700;box-shadow:0 4px 8px 0 rgba(0,0,0,.2),0 6px 20px 0 rgba(0,0,0,.19);padding:3px;margin:3px;border:4px solid #fff;width:200px}.wkld-name{color:#fff;text-align:center;font-family:""Lucida Sans"",""Lucida Sans Regular"",""Lucida Grande"",""Lucida Sans Unicode"",Geneva,Verdana,sans-serif;font-size:x-large;border-bottom-style:solid;border-bottom-width:2px;padding:2px}.feature-item-warn{background-color:orange}.feature-item-ok{background-color:#060}.feature-item-err{background-color:red}[class^=feature-item] .tooltiptext{visibility:hidden;width:120px;background-color:#000;color:#fff;text-align:center;border-radius:6px;padding:5px 0;position:absolute;z-index:1;top:-5px;left:40%}[class^=feature-item] .tooltiptext::after{content:"" "";position:absolute;top:50%;right:100%;margin-top:-5px;border-width:5px;border-style:solid;border-color:transparent #000 transparent transparent}[class^=feature-item]:hover .tooltiptext{visibility:visible}[class^=feature-item]{position:relative;padding:5px;color:#fff;border-spacing:2px;border-width:1px;border-bottom-style:solid;font-weight:400;font-size:small}.container{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));grid-auto-flow:dense}"
    $htmlHeader = @"
		<!DOCTYPE html>
		<html>
		<head>
		<style>
		$($minCSS)
        </style>
        <title>$($rptTitle)</title>
        </head>
"@
    $htmlBody = @"
        <body>
        <h1>$($rptTitle)</h1>
        <p>Page refreshed: <span id="datetime"></span><span>&nbsp;&nbsp;Data refresh:$(get-date -f 'MMM dd yyyy HH:mm:ss')</span></p>
        $($contentOne)
"@
    $htmlFooter = @"
        <script>
        var dt = new Date();
        document.getElementById("datetime").innerHTML = (("0"+dt.getDate()).slice(-2)) +"-"+ (("0"+(dt.getMonth()+1)).slice(-2)) +"-"+ (dt.getFullYear()) +" "+ (("0"+dt.getHours()).slice(-2)) +":"+ (("0"+dt.getMinutes()).slice(-2));
        </script>
        </body>
        </html>
"@

    #Add in code to refresh page
    # 300000 is 5 mins (5 *60 * 1000)
    #Assumes the code is scheduled and runs at least every 5 mins
    $addJava = "<script language=""JavaScript"" type=""text/javascript"">"
    $addJava += "setTimeout(""location.href='$($HTMLOutput)'"",$($pageRefresh*60*1000));"
    $addjava += "</script>"

    $htmlReport = $htmlHeader + $addJava + $htmlBody + $htmlFooter
    $htmlReport | Out-File "$($pathHTML)\$($HTMLOutput)"
}

#Report info
# Get Messages
#	Returns the current status of the service.
$uriCurrentStatus = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/CurrentStatus"

if ($proxyServer) {
    [array]$allCurrentStatusMessages = (Invoke-RestMethod -Uri $uriCurrentStatus -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).Value
}
else {
    [array]$allCurrentStatusMessages = (Invoke-RestMethod -Uri $uriCurrentStatus -Headers $authHeader -Method Get).Value
}

if (($null -eq $allCurrentStatusMessages) -or ($allCurrentStatusMessages -eq 0)) {
    $evtMessage = "ERROR - Cannot retrieve the current status of services - verify proxy and network connectivity."
    Write-EventLog -LogName $evtLogname -Source $evtSource -Message $evtMessage -EventId 1 -EntryType Error
    Write-Log $evtMessage
}
else {
    $evtMessage = "$($allCurrentStatusMessages.count) Workloads, Features and Status returned."
    Write-Log $evtMessage
}

#Wall Builder
#Preferred line one cards
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
            "defcon" {
                $evtMessage = "$($feature.FeatureServiceStatus) returned value $($CardClass)"
                Write-Log $evtMessage
            }
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
$rptFeatureDash += "</div></div>`n<br/><br/>"

$rptTitle = $rptTenantName + " " + $rptName
$evtMessage = "Generating HTML to '$($pathHTML)\$($HTMLFile)'"
Write-Log $evtMessage
BuildHTML $rptTitle $rptFeatureDash $HTMLFile

$swScript.Stop()

$evtMessage = "Script runtime $($swScript.Elapsed.Minutes)m:$($swScript.Elapsed.Seconds)s:$($swScript.Elapsed.Milliseconds)ms on $env:COMPUTERNAME`r`n"
$evtMessage += "*** Processing finished ***`r`n"
Write-Log $evtMessage

#Append to daily log file.
Get-Content $script:logfile | Add-Content $script:Dailylogfile
Remove-Item $script:logfile
Remove-Module O365ServiceHealth
