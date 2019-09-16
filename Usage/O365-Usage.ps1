<#
.SYNOPSIS
    This script grabs the office 365 Usage reports in CSV format.

.DESCRIPTION
    This script is designed to run on a scheduled basis.
    Configuration file is required allowing script to be called for multiple tenants.

    Requirements:
    Azure AD powershell module (Install-Module AzureAD)
    Azure AD Application - Tenant ID, Application ID and Secret. (store in config file)

.EXAMPLE
    PS .\O365-Usage.ps1 -configXML '..\profile-sample.xml' -ReportTimeSpan D7 -ReportType All

    Uses the specific XML settings file to load tenant information. The file is specified relative to the location of this script, or absolute location
	All reports listed will be downloaded for a 7 day timespan

.NOTES
    Author:  Jonathan Christie
    PSVer:   2.0/3.0/4.0/5.0
    Version: 2.0.1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [String]$configXML = "..\config\profile-test.xml",
    [Parameter(Mandatory = $false)] [ValidateSet("D7","D30","D90","D180")][String]$ReportTimeSpan = "D7",
    [Parameter(Mandatory = $false)] [ValidateSet("All","UserDetail")][String]$ReportType = "UserDetail"
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
    $info = $(Get-Date).ToString() + ": " + $info
    $info >> $script:logfile
}


if ([system.IO.path]::IsPathRooted($configXML) -eq $false) {
    #its not an absolute path. Find the absolute path
    $configXML = Resolve-Path $configXML
}
$config = LoadConfig $configXML

[string]$pathLogs = $config.LogPath
[string]$pathUsageReports = $config.UsageReportsPath

[array]$allResults = @()

#Configure local event log
[string]$evtLogname = $config.EventLog
[string]$evtSource = $config.UsageEventSource
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

[string]$tenantID = $config.TenantID
[string]$appID = $config.AppID
[string]$clientSecret = $config.AppSecret

[string]$rptProfile = $config.TenantShortName
[string]$proxyHost = $config.ProxyHost

#If no path has been specified, use the current script location
if (!$pathLogs) {
    $pathLogs = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
}

#Check and trim the report path
$pathLogs = $pathLogs.TrimEnd("\")
#Build and Check output directories
#Base for logs
if (!(Test-Path $($pathLogs))) {
    New-Item -ItemType Directory -Path $pathLogs
}
if ([system.IO.path]::IsPathRooted($pathLogs) -eq $false) {
    #its not an absolute path. Find the absolute path
    $pathLogs = Resolve-Path $pathLogs
}
#Check reports output folder
if (!(Test-Path $($pathUsageReports))) {
    New-Item -ItemType Directory -Path $pathUsageReports
}
if ([system.IO.path]::IsPathRooted($pathUsageReports) -eq $false) {
    #its not an absolute path. Find the absolute path
    $pathUsageReports = Resolve-Path $pathUsageReports
}

# setup the logfile
# If logfile exists, the set flag to keep logfile
$script:DailyLogFile = "$($pathLogs)\O365Usage-$($rptprofile)-$(Get-Date -format yyMMdd).log"
$script:LogFile = "$($pathLogs)\tmpO365Usage-$($rptprofile)-$(Get-Date -format yyMMddHHmmss).log"
$script:LogInitialized = $false
$script:FileHeader = "*** Application Information ***"

$evtMessage = "Config File: $($configXML)"
Write-Log $evtMessage
$evtMessage = "Log Path: $($pathLogs)"
Write-Log $evtMessage
$evtMessage = "HTML Output: $($pathHTML)"
Write-Log $evtMessage

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


#Create event logs if set
if ($UseEventLog) {
    $evtCheck = Get-EventLog -List -ErrorAction SilentlyContinue | Where-Object { $_.LogDisplayName -eq $evtLogname }
    if (!($evtCheck)) {
        New-EventLog -LogName $evtLogname -Source $evtSource
        Write-EventLog -LogName $evtLogname -Source $evtSource -Message "Event log created." -EventId 1 -EntryType Information
    }
}

#Keep a list of known issues in CSV. This is useful if the event log is cleared, or not used.
[string]$evtLogAll = $null

#Report info
#Connect to Azure app and grab the service status
ConnectAzureAD
#Connect to Graph API
[uri]$urlOrca = "https://graph.microsoft.com"
[uri]$authority = "https://login.microsoftonline.com/$($TenantID)"
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
$clientCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $appId, $clientSecret
$authenticationResult = ($authContext.AcquireTokenAsync($urlOrca, $clientCredential)).Result;
$bearerToken = $authenticationResult.AccessToken
if ($null -eq $authenticationResult) {
    $evtMessage = "ERROR - No authentication result for Auzre AD App"
    Write-EventLog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 1 -EntryType Error
    Write-Log $evtMessage
}

# Get Messages
$authHeader = @{
    'Content-Type'  = 'application/json'
    'Authorization' = "Bearer " + $bearerToken
}

#O365 reports to generate from https://docs.microsoft.com/en-us/graph/api/resources/report?view=graph-rest-1.0
$rptsAllUsage = @(
    'getTeamsDeviceUsageUserDetail';
    'getTeamsDeviceUsageUserCounts';
    'getTeamsDeviceUsageDistributionUserCounts';
    'getTeamsUserActivityUserDetail';
    'getTeamsUserActivityCounts';
    'getTeamsUserActivityUserCounts';
    'getEmailActivityUserDetail';
    'getEmailActivityCounts';
    'getEmailActivityUserCounts';
    'getEmailAppUsageUserDetail';
    'getEmailAppUsageAppsUserCounts';
    'getEmailAppUsageUserCounts';
    'getEmailAppUsageVersionsUserCounts';
    'getMailboxUsageDetail';
    'getMailboxUsageMailboxCounts';
    'getMailboxUsageQuotaStatusMailboxCounts';
    'getMailboxUsageStorage';
    'getOffice365ActivationsUserDetail';
    'getOffice365ActivationCounts';
    'getOffice365ActivationsUserCounts';
    'getOffice365ActiveUserDetail';
    'getOffice365ActiveUserCounts';
    'getOffice365ServicesUserCounts';
    'getOffice365GroupsActivityDetail';
    'getOffice365GroupsActivityCounts';
    'getOffice365GroupsActivityGroupCounts';
    'getOffice365GroupsActivityStorage';
    'getOffice365GroupsActivityFileCounts';
    'getOneDriveActivityUserDetail';
    'getOneDriveActivityUserCounts';
    'getOneDriveActivityFileCounts';
    'getOneDriveUsageAccountDetail';
    'getOneDriveUsageAccountCounts';
    'getOneDriveUsageFileCounts';
    'getOneDriveUsageStorage';
    'getSharePointActivityUserDetail';
    'getSharePointActivityFileCounts';
    'getSharePointActivityUserCounts';
    'getSharePointActivityPages';
    'getSharePointSiteUsageDetail';
    'getSharePointSiteUsageFileCounts';
    'getSharePointSiteUsageSiteCounts';
    'getSharePointSiteUsageStorage';
    'getSharePointSiteUsagePages';
    'getSkypeForBusinessActivityUserDetail';
    'getSkypeForBusinessActivityCounts';
    'getSkypeForBusinessActivityUserCounts';
    'getSkypeForBusinessDeviceUsageUserDetail';
    'getSkypeForBusinessDeviceUsageDistributionUserCounts';
    'getSkypeForBusinessDeviceUsageUserCounts';
    'getSkypeForBusinessOrganizerActivityCounts';
    'getSkypeForBusinessOrganizerActivityUserCounts';
    'getSkypeForBusinessOrganizerActivityMinuteCounts';
    'getSkypeForBusinessParticipantActivityCounts';
    'getSkypeForBusinessParticipantActivityUserCounts';
    'getSkypeForBusinessParticipantActivityMinuteCounts';
    'getSkypeForBusinessPeerToPeerActivityCounts';
    'getSkypeForBusinessPeerToPeerActivityUserCounts';
    'getSkypeForBusinessPeerToPeerActivityMinuteCounts';
    'getYammerActivityUserDetail';
    'getYammerActivityCounts';
    'getYammerActivityUserCounts';
    'getYammerDeviceUsageUserDetail';
    'getYammerDeviceUsageDistributionUserCounts';
    'getYammerDeviceUsageUserCounts';
    'getYammerGroupsActivityDetail';
    'getYammerGroupsActivityGroupCounts';
    'getYammerGroupsActivityCounts'
)
$rptsUserDetails = @(
    'getTeamsDeviceUsageUserDetail';
    'getTeamsUserActivityUserDetail';
    'getEmailActivityUserDetail';
    'getEmailAppUsageUserDetail';
    'getMailboxUsageDetail';
    'getOffice365ActivationsUserDetail';
    'getOffice365ActiveUserDetail';
    'getOffice365GroupsActivityDetail';
    'getOneDriveActivityUserDetail';
    'getOneDriveUsageAccountDetail';
    'getSharePointActivityUserDetail';
    'getSharePointSiteUsageDetail';
    'getSkypeForBusinessActivityUserDetail';
    'getSkypeForBusinessDeviceUsageUserDetail';
    'getYammerActivityUserDetail';
    'getYammerDeviceUsageUserDetail';
    'getYammerGroupsActivityDetail'
)

$Period = 'D7'
$evtMessage = $null
$evtLogAll = $null
$i = 0
switch ($ReportType)
{
	"All" {$UsageReports=$rptsAllUsage}
	"UserDetail" {$UsageReports=$rptsUserDetails}
}


foreach ($O365Report in $UsageReports) {
    $i++
    Write-Progress -Activity "Downloading data for $($O365Report)" -Status "Storing Office 365 Usage Report $i of $($UsageReports.count)" -PercentComplete (($i / $UsageReports.count) * 100)
    $allResults = $null
    $reportURI = $null
    # Activation reports dont take a time period, appending one causes an error. So dont.
    if ($O365Report -like "getOffice365Activation*") {
        $reportPeriod = ""
        $reportName = "$($O365Report)"
    }
    else {
        $reportPeriod = "(period='$($ReportTimeSpan)')"
        $reportName = "$($O365Report)-$($period)"
    }
    [uri]$reportURI = "https://graph.microsoft.com/v1.0/reports/$($O365Report)$($reportPeriod)"
    $evtMessage = "Working on report URI: $($reportURI)"
    Write-Log $evtMessage

    try {
        if ($proxyServer) {
            $allResults = Invoke-RestMethod -Uri $reportURI -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials
        }
        else {
            $allResults = Invoke-RestMethod -Uri $reportURI -Headers $authHeader -Method Get
        }


        if ($null -eq $allResults) {
            $evtMessage = "No result returned for $($O365Report) : $($reportURI)"
            Write-Log $evtMessage
        }
        else {
            #Results are returned as string, comma seperated.
            #theres some characters at the start of the output to remove
            $allResults = $allResults.replace("ï»¿", "") | ConvertFrom-Csv
            if (!($null -eq $allResults)) {
                $allResults | Export-Csv -Path "$($pathUsageReports)$(Get-Date -f 'yyyyMMdd')-$($reportName).csv" -NoTypeInformation -Encoding UTF8
                $evtMessage = "Exporting data for $($O365Report) : $($reportURI)"
                Write-Log $evtMessage
            }
            else {
                $evtMessage = "No data to export for $($O365Report) : $($reportURI)"
                Write-Log $evtMessage
            }
        }
    }
    catch {
        $evtMessage = "Unable to download report for $($O365Report) : $($reportURI)`r`n"
        $evtMessage += "$($error[0].exception)"
        $evtLog = $evtMessage + "`r`n"
        Write-Log $evtMessage
        Write-EventLog -LogName $evtLogname -Source $evtSource -Message $evtLog -EventId 5 -EntryType Error
    }
}

if ($evtLogAll.Length -gt 35000) { $evtLogAll = $evtLogAll.substring(0, 35000) }
$swScript.Stop()
$evtMessage = "Script runtime $($swScript.Elapsed.Minutes)m:$($swScript.Elapsed.Seconds)s:$($swScript.Elapsed.Milliseconds)ms on $env:COMPUTERNAME`r`n"
$evtMessage += "*** Processing finished ***`r`n"
Write-Log $evtMessage

#Append to daily log file.
Get-Content $script:logfile | Add-Content $script:Dailylogfile
Remove-Item $script:logfile
Remove-Module O365ServiceHealth