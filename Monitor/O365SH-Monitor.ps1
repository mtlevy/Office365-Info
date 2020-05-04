<#
.SYNOPSIS
    Gets the Office 365 Health Center messages and incidents.
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
    PS C:\> O365SH.ps1

.EXAMPLE
    PS C:\> O365SH.ps1 -Tenant Production

.NOTES
    Author:  Jonathan Christie
    Email:   jonathan.christie (at) boilerhouseit.com
    Date:    02 Feb 2019
    PSVer:   2.0/3.0/4.0/5.0
    Version: 2.0.6
    Updated: Single page, monitor only for SysOps
    UpdNote:

    Wishlist:

    Completed:

    Outstanding:

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [String]$configXML = "..\config\profile-test.xml"
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

[string]$tenantID = $config.TenantID
[string]$appID = $config.AppID
[string]$clientSecret = $config.AppSecret
[string]$rptProfile = $config.TenantShortName
[string]$proxyHost = $config.ProxyHost
[string]$emailEnabled = $config.EmailEnabled
[string]$SMTPUser = $config.EmailUser
[string]$SMTPPassword = $config.EmailPassword
[string]$SMTPKey = $config.EmailKey
[string]$pathLogs = $config.LogPath
[string]$evtLogname = $config.EventLog
#[string]$evtSource = $config.MonitorEvtSource
[string]$evtSource = "Monitor"
[boolean]$checklog = $false
[boolean]$checksource = $false
[string[]]$MonitorAlertsTo = $config.MonitorAlertsTo
[string]$emailClosedBgd = "WhiteSmoke"
[string]$pathWorking = $config.WorkingPath

[string[]]$cnameAlertsTo = $config.CNAMEAlertsTo
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

if ($cnameresolvers[0] -eq "") {
    $cnameResolvers = @(Get-DnsClientServerAddress | Sort-Object interfaceindex | Select-Object -ExpandProperty serveraddresses | Where-Object { $_ -like '*.*' } | Select-Object -First 1)
    $cnameResolverDesc = @("Default")
}


#Configure local event log
if ($config.UseEventlog -like 'true') {
    [boolean]$UseEventLog = $true
    #check source and log exists
    $checkLog = [System.Diagnostics.EventLog]::Exists("$($evtLogname)")
    $checkSource = [System.Diagnostics.EventLog]::SourceExists("$($evtSource)")
    if ((! $checkLog) -or (! $checkSource)) {
        New-EventLog -LogName $evtLogname -Source $evtSource
    }
}
else { [boolean]$UseEventLog = $false }

if ($config.EmailEnabled -like 'true') { [boolean]$emailEnabled = $true } else { [boolean]$emailEnabled = $false }
if ($config.CNAMEEnabled -like 'true') { [boolean]$cnameEnabled = $true } else { [boolean]$cnameEnabled = $false }

# Get Email credentials
# Check for a username. No username, no need for credentials (internal mail host?)
[PSCredential]$emailCreds = $null
if ($emailEnabled -and $smtpuser -notlike '') {
    #Email credentials have been specified, so build the credentials.
    #See readme on how to build credentials files
    $EmailCreds = getCreds $SMTPUser $SMTPPassword $SMTPKey
}

# Build logging variables
#If no path has been specified, use the current script location
$pathLogs = CheckDirectory $pathLogs
$pathWorking = CheckDirectory $pathWorking


# setup the logfile
# If logfile exists, the set flag to keep logfile
$script:DailyLogFile = "$($pathLogs)\O365Monitor-$($rptprofile)-$(Get-Date -format yyMMdd).log"
$script:LogFile = "$($pathLogs)\tmpO365Monitor-$($rptprofile)-$(Get-Date -format yyMMddHHmmss).log"
$script:LogInitialized = $false
$script:FileHeader = "*** Application Information ***"

$evtMessage = "Config File: $($configXML)"
Write-Log $evtMessage
$evtMessage = "Log Path: $($pathLogs)"
Write-Log $evtMessage
$evtMessage = "Working Path: $($pathWorking)"
Write-Log $evtMessage

#Create event logs if set
[System.Diagnostics.EventLog]$evtCheck = ""
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
    $evtMessage = "ERROR - No authentication result for Azure AD App"
    Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 10 -EntryType Error
    Write-Log $evtMessage
}

#Script specific
[array]$allMessages = @()
[array]$newIncidents = @()
#Keep a list of known issues in CSV. This is useful if the event log is cleared, or not used.
[string]$knownIssues = "$($pathWorking)\knownIssues-$($rptprofile).csv"
[array]$knownIssuesList = @()

if (Test-Path "$($knownIssues)") { $knownIssuesList = Import-Csv "$($knownIssues)" }
else {
    $evtMessage = "ERROR - Known issues list does not exist. Ignore on first run."
    Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 10 -EntryType Error
    Write-Log $evtMessage
}

if ($knownIssuesList -eq 0) {
    $evtMessage = "ERROR - Known issues list is empty. Ignore on first run."
    Write-ELog -LogName $evtLogname -Source $evtSource -Message "$($rptProfile) : $evtMessage" -EventId 10 -EntryType Error
    Write-Log $evtMessage
}

#	Returns the messages about the service over a certain time range.
[uri]$uriMessages = "https://manage.office.com/api/v1.0/$tenantID/ServiceComms/Messages"

if ($ProxyServer) {
    [array]$allMessages = @((Invoke-RestMethod -Uri $uriMessages -Headers $authHeader -Method Get -Proxy $proxyHost -ProxyUseDefaultCredentials).Value)
}
else {
    [array]$allMessages = @((Invoke-RestMethod -Uri $uriMessages -Headers $authHeader -Method Get).Value)
}

if (($null -eq $allMessages) -or ($allMessages.Count -eq 0)) {
    $evtMessage = "ERROR - Cannot retrieve the current status of services - verify proxy and network connectivity."
    Write-ELog -LogName $evtLogname -Source $evtSource -Message $evtMessage -EventId 11 -EntryType Error
    Write-Log $evtMessage

}
else {
    $evtMessage = "$($allMessages.count) messages returned."
    Write-Log $evtMessage
}

$currentIncidents = @($allMessages | Where-Object { ($_.messagetype -notlike 'MessageCenter') })
$newIncidents = @($currentIncidents | Where-Object { ($_.id -notin $knownIssuesList.id) -and ($null -eq $_.endtime) })

#Get all messages with no end date (open)
#Get events and check event log
#If not logged, create an entry and send an email

$evtFind=$null

if ($newIncidents.count -ge 1) {
    Write-Log "New incidents detected: $($newIncidents.count)"
    foreach ($item in $newIncidents) {
        if ($useEventLog) {
            #Check event log. If no entry (ID) then write entry
            $evtFind = Get-EventLog -LogName $evtLogname -Source $evtSource -Message "*: $($item.ID)*: $($rptProfile)*" -ErrorAction SilentlyContinue
        }
        #Check known issues list
        if ($null -eq $evtFind) {
            $emailPriority = ""
            Write-Log "Building and attempting to send email"
            $mailMessage = "<b>ID</b>`t`t: $($item.ID)<br/>"
            $mailMessage += "<b>Tenant</b>`t`t: $($rptProfile)<br/>"
            $mailMessage += "<b>Feature</b>`t`t: $($item.WorkloadDisplayName)<br/>"
            $mailMessage += "<b>Status</b>`t`t: $($item.Status)<br/>"
            $mailMessage += "<b>Severity</b>`t`t: $($item.Severity)<br/>"
            $mailMessage += "<b>Start Date</b>`t: $(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm')<br/>"
            $mailMessage += "<b>Last Updated</b>`t: $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')<br/>"
            $mailMessage += "<b>Incident Title</b>`t: $($item.title)<br/>"
            $mailMessage += "$($item.ImpactDescription)<br/><br/>"
            $emailPriority = Get-Severity "email" $item.severity
            $emailSubject = "New $($item.Severity) issue: $($item.WorkloadDisplayName) - $($item.Status) [$($item.ID)]"
            if ($MonitorAlertsTo -and $emailEnabled) { SendEmail $mailMessage $EmailCreds $config $emailPriority $emailSubject $MonitorAlertsTo }
            $evtMessage = $mailMessage.Replace("<br/>", "`r`n")
            $evtMessage = $evtMessage.Replace("<b>", "")
            $evtMessage = $evtMessage.Replace("</b>", "")
            $evtID = 0
            $evtErr = ''
            switch ($item.severity) {
                'SEV0' { $evtErr = 'Error'; $evtID = 22 }
                'SEV1' { $evtErr = 'Error'; $evtID = 21 }
                'SEV2' { $evtErr = 'Warning'; $evtID = 20 }
            }
            Write-ELog -LogName $evtLogname -Source $evtSource -Message $evtMessage -EventId $evtID -EntryType $evtErr
            Write-Log $evtMessage
        }
    }
}

[array]$reportClosed = @()
[array]$recentlyClosed = @()
[array]$recentIncidents = @()

#Previously known items (saved list) where there was not an end time
$recentIncidents = $knownIssuesList | Where-Object { ($_.endtime -eq '') }
#Items on current scan (online) where there was an end time
$recentlyClosed = $currentIncidents | Where-Object { ($_.endtime -ne $null) }
#Closed items are previously open items that now have an end time
$reportClosed = $recentlyClosed | Where-Object { $_.id -in $recentIncidents.ID }
#Some items may have been newly added AND closed
$reportClosed += $currentIncidents | Where-Object { $_.id -notin $knownissueslist.ID -and $null -ne $_.endtime }

Write-Log "Closed incidents detected: $($reportClosed.count)"
#Check that closed isnt greater than 10. If it is, its likely to be a new install or issue with corrupt file
if ($knownIssuesList.count -ge 1 -and $reportClosed.count -le 10) {
    foreach ($item in $reportClosed) {
        Write-Log "Building and attempting to send closure email"
        $mailMessage = "<b>Incident Closed</b>`t`t: <b>Closed</b><br/>"
        $mailMessage += "<b>Tenant</b>`t`t: $($rptProfile)<br/>"
        $mailMessage += "<b>ID</b>`t`t: $($item.ID)<br/>"
        $mailMessage += "<b>Feature</b>`t`t: $($item.WorkloadDisplayName)<br/>"
        $mailMessage += "<b>Status</b>`t`t: $($item.Status)<br/>"
        $mailMessage += "<b>Severity</b>`t`t: $($item.Severity)<br/>"
        $mailMessage += "<b>Start Time</b>`t: $(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm')<br/>"
        $mailMessage += "<b>Last Updated</b>`t: $(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')<br/>"
        $mailMessage += "<b>End Time</b>`t: <b>$(Get-Date $item.EndTime -f 'dd-MMM-yyyy HH:mm')</b><br/>"
        $mailMessage += "<b>Incident Title</b>`t: $($item.title)<br/>"
        $mailMessage += "$($item.ImpactDescription)<br/><br/>"
        #Add the last action from microsoft to the email only - not to the event log entry (text can be too long)
        $mailWithLastAction = $mailMessage + "<b>Final Update from Microsoft</b>`t:<br/>"
        $lastMessage = Get-htmlMessage ($item.messages.messagetext | Where-Object { $_ -like '*This is the final update*' -or $_ -like '*Final status:*' })
        $lastMessage = "<div style='background-color:$($emailClosedBgd)'>" + $lastMessage.replace("<br><br>", "<br/>") + "</div>"
        $mailWithLastAction += "$($lastMessage)<br/><br/>"
        $emailSubject = "Closed: $($item.WorkloadDisplayName) - $($item.Status) [$($item.ID)]"
        Write-Log "Sending email to $($MonitorAlertsTo)"
        if ($MonitorAlertsTo -and $emailEnabled) { SendEmail $mailWithLastAction $EmailCreds $config "Normal" $emailSubject $MonitorAlertsTo }
        $evtMessage = $mailMessage.Replace("<br/>", "`r`n")
        $evtMessage = $evtMessage.Replace("<b>", "")
        $evtMessage = $evtMessage.Replace("</b>", "")
        Write-ELog -LogName $evtLogname -Source $evtSource -Message $evtMessage -EventId 30 -EntryType Information
        Write-Log $evtMessage
    }
}

#Update the know lists if issues. they might not have increased, but end times may have been added.
#If empty then skip to avoid overwriting
if ($currentIncidents.count -gt 0) {
    $currentIncidents | Export-Csv "$($knownIssues)" -Encoding UTF8 -NoTypeInformation
}

#Check DNS entries while we're here
if ($cnameEnabled) {
    foreach ($DNSServer in $cnameResolvers) {
        $dnsServerDesc = $cnameresolverdesc[[array]::indexof($cnameResolvers, $DNSServer)]
        [array]$exportNH = @()
        #Define the filename and location to store URL cname results
        $cnameKnownCSV = "$($pathWorking)\$cnameFilename-$($DNSServer)-$($rptProfile).csv"
        if (!(Test-Path $cnameKnownCSV)) {
            $headers = "monitor,namehost,domain,addedDate,lastDate"
            $headers | Add-Content $cnameKnownCSV -Encoding UTF8
        }

        #define the URLs which should be monitored
        #Fetch the list of previously known CNAMES
        $cnameKnown = Import-Csv "$($cnameKnownCSV)"
        if ($cnameKnown.count -ge 1) {
            #check if CSV has the lastDate column. if not, add it in
            if (!(($cnameKnown | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty name) -contains 'lastdate')) {
                $cnameKnown | Add-Member -MemberType NoteProperty -Name lastDate -Value $null
            }
            if (!(($cnameKnown | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty name) -contains 'resolver')) {
                $cnameKnown | Add-Member -MemberType NoteProperty -Name resolver -Value $null
            }
        }
        $addDateTime = Get-Date -f "dd-MMM-yy HH:mm"
        $alert = ""
        #for each entry in the monitored urls
        foreach ($entry in $cnameURLs) {
            [array]$newDomains = @()
            try { $cnames = @(Resolve-DnsName $($entry) -DnsOnly -Server $DNSServer | Where-Object querytype -like 'CNAME') }
            catch { $alert += "<b>Cannot resolve CNAMES for $($entry)</b><br/>$($error[0].exception)"; break; }
            #From the list own previous responses, get matching monitor records
            #update the lastDate for names which have been found
            $updateNH = $cnameknown | Where-Object { ($_.namehost -in $cnames.namehost) -and ($_.monitor -like $entry) }
            foreach ($knownNH in $updateNH) {
                $knownNH.lastDate = $addDateTime
            }
            $exportNH += $updateNH
            #identify the new cnames returned
            $unknownNH = $cnames | Where-Object { ($_.namehost -notin $cnameknown.namehost) }
            foreach ($alias in $unknownNH) {
                $nh = @($alias.nameHost.split("."))
                $aliasDomain = $nh[-2] + "." + $nh[-1]
                $addNew = New-Object PSObject
                $addNew | Add-Member -MemberType NoteProperty -Name monitor -Value $entry
                $addNew | Add-Member -MemberType NoteProperty -Name nameHost -Value $alias.NameHost
                $addNew | Add-Member -MemberType NoteProperty -Name domain -Value $aliasDomain
                $addNew | Add-Member -MemberType NoteProperty -Name addedDate -Value $addDateTime
                $addNew | Add-Member -MemberType NoteProperty -Name lastdate -Value $addDateTime
                $addNew | Add-Member -MemberType NoteProperty -Name resolver -Value $DNSServer

                Write-Log "Adding $($alias.nameHost) to CSV"
                $alert += "<b>{0}</b>: New name host found CNAME <b>{1}</b> via Resolver <b>{2} ({3})</b><br/>" -f $entry, $alias.nameHost, $DNSServer, $dnsServerDesc
                $newDomains += @($addNew)
            }
            $exportNH += $newDomains
        }
        #export the lot of them to the original file
        #Add the items originally
        $notfound = $cnameknown | Where-Object { $exportNH -notcontains $_ }
        $exportNH += $notfound
        $exportNH | Sort-Object addedDate, lastDate | Export-Csv -Path "$($cnameKnownCSV)" -NoTypeInformation -Encoding UTF8

        #if alert email, then send
        if ($alert -ne "") {
            $newDomains = $newDomains | Select-Object -Unique
            $alert += "<br/>Check the following domains can be reached for resolution <b>$($newdomains -join(', '))</b><br/>"
            if ($emailEnabled) {
                $emailSubject = "New CNAME records resolved"
                SendEmail $alert $EmailCreds $config "High" $emailSubject $cnameAlertsTo $cnameKnownCSV
            }
        }
    }
}


$swScript.Stop()
$evtMessage = "Script runtime $($swScript.Elapsed.Minutes)m:$($swScript.Elapsed.Seconds)s:$($swScript.Elapsed.Milliseconds)ms on $env:COMPUTERNAME`r`n"
$evtMessage += "*** Processing finished ***`r`n"
Write-Log $evtMessage

#Append to daily log file.
Get-Content $script:logfile | Add-Content $script:Dailylogfile
Remove-Item $script:logfile
Remove-Module O365ServiceHealth