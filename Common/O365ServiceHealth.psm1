# Shared functions for Office 365 Info powershell
function saveCredentials {
    param (
        [Parameter(Mandatory = $true)] [string]$Password,
        [Parameter(Mandatory = $true)] [boolean]$CreateKey,
        [Parameter(Mandatory = $true)] [string]$KeyPath,
        [Parameter(Mandatory = $true)] [string]$CredsPath
    ) 
    #keypath is path and file name ie c:\credentials\user-tenant.key
    #credspath is path and file to password ie c:\credentials\user-tenant.pwd
    $AESKey = $null
    #Build a key if needed
    if ($CreateKey) {
        $AESKey = New-Object byte[] 32
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
        #Store the AESKey into a file. Protect this file with permissions.
        Set-Content $KeyPath $AESKey
    }
    #Import key file if it hasnt been created
    if ($null -eq $AESKey) { $AESKey = Get-Content $KeyPath }
    $SecurePwd = $Password | ConvertTo-SecureString -AsPlainText -Force
    $SecurePwd | ConvertFrom-SecureString -Key $AESKey | Out-File $CredsPath
}

function getCreds {
    param (
        [Parameter(Mandatory = $true)] [string]$Username,
        [Parameter(Mandatory = $true)] [string]$CredsPath,
        [Parameter(Mandatory = $true)] [string]$KeyPath)

    $AESKey = Get-Content $KeyPath
    $Password = Get-Content $CredsPath | ConvertTo-SecureString -Key $AESKey
    $Credentials = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $Username, $Password
    return $Credentials
}

function CheckDirectory {
    Param (
        [parameter(Mandatory = $false)] [string]$folder
    )
    #If no path has been specified, use the current script location
    if (!$folder) {
        if ($PSIse) {
            $folder = Split-Path $PSIse.CurrentFile.FullPath
        }
        else {
            $folder = $Global:PSScriptRoot
        }
    }

    #Check and trim the path
    $folder = $folder.TrimEnd("\")

    #If path doesnt exist then create
    if (!(Test-Path $($folder))) {
        $result = New-Item -ItemType Directory -Path $folder
    }

    #If path is not absolute, then find it.
    if ([system.IO.path]::IsPathRooted($folder) -eq $false) {
        $folder = Resolve-Path $folder
    }
    return $folder
}

function Write-ELog {
    Param (
        [parameter(Mandatory = $false)] [boolean]$useEventLog,
        [parameter(Mandatory = $false)] [string]$LogName,
        [parameter(Mandatory = $false)] [string]$Source,
        [parameter(Mandatory = $false)] [string]$Message,
        [parameter(Mandatory = $false)] [int]$EventId,
        [parameter(Mandatory = $false)] [string]$EntryType
    )
    if ($useEventLog) {
        Write-ELog -LogName $LogName -Source $Source -Message $Message -EventId $EventId -EntryType $EntryType
    }
}

function ConnectAzureAD() {
    $modAzureAD = (Get-Module -Name "AzureAD" -ListAvailable | Sort-Object version -Descending)[-1]
    if ($null -eq $modAzureAD) {
        $evtMessage = "ERROR - Azure AD module not found. Install using command 'Install-Module AzureAD' from an elevated prompt`r`n"
        Write-Log $evtMessage
        Write-Verbose $evtMessage
        #If tracking exit reasons - use 5 for module not found
        Exit 5
    }
    #Load the latest authentication libraries
    $adal = Join-Path $modAzureAD.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    $adalforms = Join-Path $modAzureAD.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

    #Force TLS 1.2 and ignore SSL Warnings
    IgnoreSSLWarnings
}

function LoadConfig {
    Param (
        [Parameter(Mandatory = $true)] [string]$configFile
    )
    $appSettings = @{ }
    if (!(Test-Path $configFile)) {
        Write-Output "Cannot find configuration file: $($configFile)"
        Exit
    }
    [xml]$configFile = Get-Content "$($configFile)"

    $appSettings = [PSCustomObject]@{
        TenantName          = $configFile.Settings.Tenant.Name
        TenantShortName     = $configFile.Settings.Tenant.ShortName
        TenantMSName        = $configFile.Settings.Tenant.MSName
        TenantDescription   = $configFile.Settings.Tenant.Description
    
        TenantID            = $configFile.Settings.Azure.TenantID
        AppID               = $configFile.Settings.Azure.AppID
        AppSecret           = $configFile.Settings.Azure.AppSecret
    
        LogPath             = $configFile.Settings.Output.LogPath
        HTMLPath            = $configFile.Settings.Output.HTMLPath
        WorkingPath         = $configFile.Settings.Output.WorkingPath
        UseEventLog         = $configFile.Settings.Output.UseEventLog
        EventLog            = $configFile.Settings.Output.EventLog
        HostURL             = $configFile.Settings.Output.HostURL

        EmailEnabled        = $configFile.Settings.Email.Enabled
        EmailHost           = $configFile.Settings.Email.SMTPServer
        EmailPort           = $configFile.Settings.Email.Port
        EmailUseSSL         = $configFile.Settings.Email.UseSSL
        EmailFrom           = $configFile.Settings.Email.From
        EmailUser           = $configFile.Settings.Email.Username
        EmailPassword       = $configFile.Settings.Email.PasswordFile
        EmailKey            = $configFile.Settings.Email.AESKeyFile

        MonitorAlertsTo     = [string[]]$configFile.Settings.Monitor.alertsTo
        MonitorEvtSource    = $configFile.Settings.Monitor.EventSource
        MonitorIgnoreSvc    = [string[]]$configFile.Settings.Monitor.IgnoreSvc
        MonitorIgnoreInc    = [string[]]$configFile.Settings.Monitor.IgnoreInc
  
        WallReportName      = $configFile.Settings.WallDashboard.Name
        WallHTML            = $configFile.Settings.WallDashboard.HTMLFilename
        WallDashCards       = $configFile.Settings.WallDashboard.DashCards
        WallPageRefresh     = $configFile.Settings.WallDashboard.Refresh
        WallEventSource     = $configFile.Settings.WallDashboard.EventSource

        DashboardName       = $configFile.Settings.Dashboard.Name
        DashboardHTML       = $configFile.Settings.Dashboard.HTMLFilename
        DashboardCards      = $configFile.Settings.Dashboard.DashCards
        DashboardRefresh    = $configFile.Settings.Dashboard.Refresh
        DashboardAlertsTo   = $configFile.Settings.Dashboard.AlertsTo
        DashboardEvtSource  = $configFile.Settings.Dashboard.EventSource
        DashboardLogo       = $configFile.Settings.Dashboard.Logo
        DashboardAddLink    = $configFile.Settings.Dashboard.AddLink
        DashboardHistory    = $configFile.Settings.Dashboard.History
        DashboardRecMsgs    = $configFile.Settings.Dashboard.RecentMsgs

        UsageReportsPath    = $configFile.Settings.UsageReports.Path
        UsageEventSource    = $configFile.Settings.UsageReports.EventSource

        ToolboxName         = $configFile.Settings.Toolbox.Name
        ToolboxHTML         = $configFile.Settings.Toolbox.HTMLFilename
        ToolboxNotes        = ($configFile.Settings.Toolbox.Notes).InnerXML
        ToolboxRefresh      = $configFile.Settings.Toolbox.Refresh

        DiagnosticsEnabled  = $configFile.Settings.Diagnostics.Enabled
        DiagnosticsURLs     = $configFile.Settings.Diagnostics.URLs
        DiagnosticsVerbose  = $configFile.Settings.Diagnostics.Verbose

        MiscDiagsEnabled    = $configFile.Settings.MiscDiagnostics.Enabled
        MiscDiagsWeb        = $configFile.Settings.MiscDiagnostics.Web
        MiscDiagsPorts      = $configFile.Settings.MiscDiagnostics.Ports

        RSS1Enabled         = $configFile.Settings.RSSFeeds.F1.Enabled
        RSS1Name            = $configFile.Settings.RSSFeeds.F1.Name
        RSS1Feed            = $configFile.Settings.RSSFeeds.F1.Feed
        RSS1URL             = $configFile.Settings.RSSFeeds.F1.URL
        RSS1Items           = $configFile.Settings.RSSFeeds.F1.Items

        RSS2Enabled         = $configFile.Settings.RSSFeeds.F2.Enabled
        RSS2Name            = $configFile.Settings.RSSFeeds.F2.Name
        RSS2Feed            = $configFile.Settings.RSSFeeds.F2.Feed
        RSS2URL             = $configFile.Settings.RSSFeeds.F2.URL
        RSS2Items           = $configFile.Settings.RSSFeeds.F2.Items

        IPURLsPath          = $configFile.Settings.IPURLs.Path
        IPURLsAlertsTo      = $configFile.Settings.IPURLs.AlertsTo
        IPURLsNotesFilename = $configFile.Settings.IPURLs.NotesFilename
        CustomNotesFilename = $configFile.Settings.IPURLs.CustomNotesFilename
        IPURLHistory        = $configFile.Settings.IPURLs.History
    
        CnameEnabled        = $configFile.Settings.CNAME.Enabled
        CnameNotes          = ($configFile.Settings.CNAME.Notes).InnerXML
        CnameFilename       = $configFile.Settings.CNAME.Filename
        CnameAlertsTo       = $configFile.Settings.CNAME.AlertsTo
        CnameURLs           = $configFile.Settings.CNAME.URLs
        CnameResolvers      = [string[]]$configFile.Settings.CNAME.Resolvers
        CnameResolverDesc   = [string[]]$configFile.Settings.CNAME.ResolverDesc

        PACEnabled          = $configFile.Settings.PACFile.Enabled
        PACProxy            = $configFile.Settings.PACFile.Proxy
        PACType1Filename    = $configFile.Settings.PACFile.Type1Filename
        PACType2Filename    = $configFile.Settings.PACFile.Type2Filename

        ProxyEnabled        = $configFile.Settings.Proxy.ProxyEnabled
        ProxyHost           = $configFile.Settings.Proxy.ProxyHost
        ProxyIgnoreSSL      = $configFile.Settings.Proxy.IgnoreSSL

        Blogs               = ($configFile.Settings.Blogs).InnerXML
        TeamsEnabled        = $configFile.Settings.Teams.TeamsEnabled
        TeamsURI            = $configFile.Settings.Teams.TeamsURI

        OohEnabled          = $ConfigFile.Settings.OutOfHours.OohEnabled
        OohAlertsTo         = $ConfigFile.Settings.OutOfHours.OohAlertsTo
        OohMorningStart     = $ConfigFile.Settings.OutOfHours.OohMorningStart
        OohMorningEnd       = $ConfigFile.Settings.OutOfHours.OohMorningEnd
        OohEveningStart     = $ConfigFile.Settings.OutOfHours.OohEveningStart
        OohEveningEnd       = $ConfigFile.Settings.OutOfHours.OohEveningEnd

    }
    return $appSettings
}

function Get-StatusDisplay {
    Param(
        [Parameter(Mandatory = $true)] [string]$statusName,
        [Parameter(Mandatory = $true)] [string]$type
    )
    #Icon set
    #
    $icon1 = "<img src='images/1.jpg' alt='Error' style='width:20px;height:20px;border:0;'>"
    $icon2 = "<img src='images/2.jpg' alt='Warning' style='width:20px;height:20px;border:0;'>"
    $icon3 = "<img src='images/3.jpg' alt='OK' style='width:20px;height:20px;border:0;'>"
    #Each service status that is available is mapped to one of the levels - OK (3), warning (2) and error (1)
    #Service status from: https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.servicestatus.tenantcommunications.data.servicestatus?view=o365-service-communications
    switch ($type) {
        "icon" {
            switch ($statusName) {
                "ServiceInterruption" { $StatusDisplay = $icon1 }
                "ServiceDegradation" { $StatusDisplay = $icon1 }
                "RestoringService" { $StatusDisplay = $icon2 }
                "ExtendedRecovery" { $StatusDisplay = $icon1 }
                "Investigating" { $StatusDisplay = $icon2 }
                "ServiceRestored" { $StatusDisplay = $icon3 }
                "FalsePositive" { $StatusDisplay = $icon3 }
                "PIRPublished" { $StatusDisplay = $icon3 }
                "InformationUnavailable " { $StatusDisplay = $icon1 }
                "ServiceOperational" { $StatusDisplay = $icon3 }
                "PostIncidentReviewPublished" { $StatusDisplay = $icon3 }
                #Set default error icon if the status is not listed
                default { $StatusDisplay = $icon1 }
            }
        }
        "class" {
            switch ($statusName) {
                "ServiceInterruption" { $StatusDisplay = "err" }
                "ServiceDegradation" { $StatusDisplay = "err" }
                "RestoringService" { $StatusDisplay = "warn" }
                "ExtendedRecovery" { $StatusDisplay = "err" }
                "Investigating" { $StatusDisplay = "warn" }
                "ServiceRestored" { $StatusDisplay = "ok" }
                "FalsePositive" { $StatusDisplay = "ok" }
                "PIRPublished" { $StatusDisplay = "ok" }
                "InformationUnavailable" { $StatusDisplay = "err" }
                "ServiceOperational" { $StatusDisplay = "ok" }
                "PostIncidentReviewPublished" { $StatusDisplay = "ok" }
                #Set default error colour if the status is not listed
                default { $StatusDisplay = "defcon" }
            }
        }
        "SkuCapabilityStatus" {
            switch ($statusName) {
                "Enabled" { $StatusDisplay = "ok" }
                "Suspended" { $StatusDisplay = "warn" }
                "LockedOut" { $StatusDisplay = "err" }
                default { $StatusDisplay = "warn" }
            }
        }
        "ServicePlanStatus" {
            switch ($statusName) {
                "Success" { $StatusDisplay = "ok" }
                "PendingActivation" { $StatusDisplay = "warn" }
                "PendingInput" { $StatusDisplay = "warn" }
                "PendingProvisioning" { $StatusDisplay = "warn" }
                "Disabled" { $StatusDisplay = "err" }
                default { $StatusDisplay = "warn" }
            }
        }
    }
    return $StatusDisplay
}

function Get-Severity {
    param (
        [parameter(mandatory = $true)] [string]$type,
        [parameter(mandatory = $true)] [string]$severity
    )
    [System.Net.Mail.MailPriority]$returnValue = "Normal"
    switch ($type) {
        "email" {
            #email can have the following priorities : High, Normal, Low
            switch ($severity) {
                "Sev0" { $returnValue = "High" }
                "Sev1" { $returnValue = "High" }
                "Sev2" { $returnValue = "Normal" }
                #Set default error icon if the status is not listed
                default { $returnValue = "Normal" }
            }
        }
    }
    return $returnValue
}

function featureBuilder {
    Param (
        [Parameter(Mandatory = $true)] $strName, 
        [Parameter(Mandatory = $true)] $strFeatures,
        [Parameter(Mandatory = $true)] $strPriority,
        [Parameter(Mandatory = $true)] $intFtCnt
    )
    [array]$rptCard = @()
    [decimal]$decSize = 0
    $decSize = (($intFtCnt * 0.5) + ([math]::ceiling(($strName.length) / 14) * .75) + 0.1) * 2
    [int]$intSize = $decSize
    $tableClass = "class='workload-card-$($strPriority)' style='grid-row: span $($intSize)'"
    $rptCard = @"
    <div $tableClass>
    `t<div class='wkld-name'>$($strName)</div>
    `t$($strFeatures)
    </div>
"@
    return $rptCard
}

function Get-htmlMessage ($msgText) {
    $htmlMessage = $null
    $htmlMessage = $msgText -replace ("`n", '<br>') -replace ([char]8217, "'") -replace ([char]8220, '"') -replace ([char]8221, '"') -replace ('\[', '<b><i>') -replace ('\]', '</i></b>')
    $htmlMessage = $htmlMessage -replace "Title:", "<b>Title</b>:"
    $htmlMessage = $htmlMessage -replace "User Impact:", "<b>User Impact</b>:"
    $htmlMessage = $htmlMessage -replace "More Info:", "<b>More Info</b>:"
    $htmlMessage = $htmlMessage -replace "Next Update By:", "<b>Next Update By</b>:"
    $htmlMessage = $htmlMessage -replace "Current Status:", "<b>Current Status</b>:"
    $htmlMessage = $htmlMessage -replace "Incident Start time:", "<b>Incident Start Time</b>:"
    $htmlMessage = $htmlMessage -replace "Start time:", "<b>Start Time</b>:"
    $htmlMessage = $htmlMessage -replace "Incident End time:", "<b>Incident End Time</b>:"
    $htmlMessage = $htmlMessage -replace "End time:", "<b>End Time</b>:"
    $htmlMessage = $htmlMessage -replace "Scope:", "<b>Scope</b>:"
    $htmlMessage = $htmlMessage -replace "Scope of impact:", "<b>Scope of Impact</b>:"
    $htmlMessage = $htmlMessage -replace "Estimated time to resolve:", "<b>Estimated time to resolve</b>:"
    $htmlMessage = $htmlMessage -replace "Final Status:", "<b>Final Status</b>:"
    $htmlMessage = $htmlMessage -replace "Preliminary Root Cause: ", "<b>Preliminary Root Cause</b>:"
    $htmlMessage = $htmlMessage -replace "Root Cause:", "<b>Root Cause</b>:"
    $htmlMessage = $htmlMessage -replace "Next Steps:", "<b>Next Steps</b>:"
    $htmlMessage = $htmlMessage -replace "Next Update:", "<b>Next Update</b>:"
    $htmlMessage = $htmlMessage -replace "This is the final update for the event.", "<b><u>This is the final update for the event.</u></b>"
    $htmlMessage = $htmlMessage -replace "This is the final update on this incident.", "<b><u>This is the final update on this incident.</u></b>"
    $htmlMessage = $htmlMessage -replace "`n", "<br/>"

    return $htmlMessage
}

function BuildIPURLChanges {
    Param(
        [Parameter(Mandatory = $true)] [array]$changesList,
        [Parameter(Mandatory = $true)] [array]$changesIDX,
        [Parameter(Mandatory = $true)] [array]$changesEPSIDX,
        [Parameter(Mandatory = $true)] [array]$changesFlat,
        [Parameter(Mandatory = $true)] [int]$expandCount,
        [Parameter(Mandatory = $true)] [array]$watchCats,
        [Parameter(Mandatory = $false)] [switch]$email
    )
    if ($email) { $html = $true } else { $html = $false }
    #Use DIVs to build HTML page
    $ipHistoryHTML = ""
    $ipHistoryCnt = 0
    foreach ($change in $changesList) {
        $ipHistoryCnt++
        $changesLast = @($changesIDX | Where-Object { $_.version -In $change.version })
        $inputID = "collapsible$($ipHistoryCnt)"
        if ($html) {
            $ipHistoryHTML += "<table>`r`n"
            $ipHistoryHTML += "<caption>Version: $(Get-Date $change.VersionDate -Format 'dd-MMM-yyyy') : $($changeslast.count) item(s)</caption>`r`n"
            $ipHistoryHTML += "<thead><tr>`r`n"
            $ipHistoryHTML += "<th>Service Area</th>`r`n"
            $ipHistoryHTML += "<th>Disposition</th>`r`n<th>Impact</th>`r`n"
            $ipHistoryHTML += "<th style='max-width:250px'>Add</th>`r`n<th style='max-width:250px'>Remove</th>`r`n"
            $ipHistoryHTML += "<th style='max-width:150px'>Current</th>`r`n`<th style='max-width:150px'>Previous</th>`r`n"
            $ipHistoryHTML += "</tr></thead>`r`n"
            $ipHistoryHTML += "<tbody>`r`n"		
        }
        else {
            $ipHistoryHTML += "<div class='wrap-collabsible'>`r`n"
            $ipHistoryHTML += "<input id='$($InputID)' class='toggle' type='checkbox'"
            if ($ipHistoryCnt -le $expandCount) { $ipHistoryHTML += " checked>" } else { $ipHistoryHTML += ">" }
            $ipHistoryHTML += "<label for='$($inputID)' class='lbl-toggle'>Version: $(Get-Date $change.VersionDate -Format 'dd-MMM-yyyy') : $($changeslast.count) item(s)</label>`r`n"
            $ipHistoryHTML += "<div class='collapsible-content'><div class='content-inner'>`r`n"
            $ipHistoryHTML += "<div class='tableInc'>`r`n"
            $ipHistoryHTML += "<div class='tableInc-header'>`r`n"
            $ipHistoryHTML += "<div class='tableInc-header-l'>Service Area</div>`r`n"
            $ipHistoryHTML += "<div class='tableInc-header-l'>Disposition</div>`r`n<div class='tableInc-header-l'>Impact</div>`r`n"
            $ipHistoryHTML += "<div class='tableInc-header-l' style='max-width:250px'>Add</div>`r`n<div class='tableInc-header-l' style='max-width:250px'>Remove</div>`r`n"
            $ipHistoryHTML += "<div class='tableInc-header-l' style='max-width:150px'>Current</div>`r`n`<div class='tableInc-header-l' style='max-width:150px'>Previous</div>`r`n"
            $ipHistoryHTML += "</div>`r`n"
        }
        $ipHistory = ""
        foreach ($item in $changesLast) {
            $cat = ""
            if ($html) { $ipHistory += "<tr>" } else { $ipHistory += "<div class='tableInc-row'>" }
            $serviceArea = ($changesEPSIDX | Where-Object { $item.endpointsetid -eq $_.id }).ServiceArea
            if ($item.endpointsetid -in ($watchCats.id)) {
                $cat = ($watchCats | where-Object { $_.id -eq $item.endpointsetid }).category
                if ($cat -like 'Optimize') { $suffix = "<font color='red'> ($($cat))</font>" }
                elseif ($cat -like 'Allow') { $suffix = "<font color='blue'> ($($cat))</font>" }
            }
            else { $suffix = "" }
            $serviceArea = "$($serviceArea)$($suffix)"
            if ($html) {
                $ipHistory += "<td>[$($item.endpointsetid)] $($serviceArea)</td>`n`t"
                $ipHistory += "<td>$($item.disposition)</td>`n`t"
            }
            else {
                $ipHistory += "<div class='tableInc-cell-l'>[$($item.endpointsetid)] $($serviceArea)</div>`n`t"
                $ipHistory += "<div class='tableInc-cell-l'>$($item.disposition)</div>`n`t"
            }
            switch ($item.impact) {
                "RemovedIpOrUrl" { $desc = "Removed IP or URL" }
                "AddedIP" { $desc = "Added IP" }
                "AddedUrl" { $desc = "Added URL" }
                "RemovedDuplicateIpOrUrl" { $desc = "Removed Duplicate IP or URL" }
            }
            if ($html) { $ipHistory += "<td>$($desc)</td>`n`t" }
            else { $ipHistory += "<div class='tableInc-cell-l'>$($desc)</div>`n`t" }
            #Get IP and URL changes
            $entry = @()
            $addED, $addIP, $addURL = $null
            $remIP, $remURL = $null
            $entry = @($changesFlat | Where-Object { $_.id -eq $item.id -and $_.action -like 'Add' })
            if ($null -ne $entry.effectivedate) {
                if ((Get-Date $entry.effectivedate) -gt (Get-Date)) { $colour = "<font color='red'>" } else { $colour = "<font color='green'>" }
                $addED = "<b>Effective Date:</b> $($colour)<b>$($entry.effectivedate)</b></font><br/>"
            }
            if (!([string]::IsNullOrEmpty($entry.IPs))) { $addIP = "<b>Add IPs:</b> $($entry.ips)<br/>" }
            if (!([string]::IsNullOrEmpty($entry.urls))) { $addURL = "<b>Add URLs:</b> $($entry.urls)" }
            if ($html) { $ipHistory += "<td  style='max-width:250px'>$($addED)$($addIP)$($addURL)</td>`n`t" }
            else { $ipHistory += "<div class='tableInc-cell-l' style='max-width:250px'>$($addED)$($addIP)$($addURL)</div>`n`t" }
            $entry = @()
            $addED, $addIP, $addURL = $null
            $remIP, $remURL = $null
            $entry = @($changesFlat | Where-Object { $_.id -eq $item.id -and $_.action -like 'Remove' })
            if (!([string]::IsNullOrEmpty($entry.ips))) { $remIP = "<b>Remove IPs:</b> $($entry.ips)<br/>" }
            if (!([string]::IsNullOrEmpty($entry.urls))) { $remURL = "<b>Remove URLs:</b> $($entry.urls)" }
            if ($html) { $ipHistory += "<td  style='max-width:250px'>$($remIP)$($remURL)</td>`n`t" }
            else { $ipHistory += "<div class='tableInc-cell-l' style='max-width:250px'>$($remIP)$($remURL)</div>`n`t" }
            $itemEP, $itemSA, $itemCat, $itemRqd, $itemTCP, $itemUDP, $itemNotes = ""
            $itemCur = ($item.current -replace '@{' -replace '}').Split(";") | ConvertFrom-StringData
            if ($item.Current) {
                if ($null -ne $itemCur.expressroute) { $itemEP = "<b>Express Route:</b> $($itemCur.expressroute)<br/>" }
                if ($null -ne $itemCur.serviceArea) { $itemSA = "<b>Service Area:</b> $($itemCur.serviceArea)<br/>" }
                if ($null -ne $itemCur.category) { $itemCat = "<b>Category:</b> $($itemCur.category)<br/>" }
                if ($null -ne $itemCur.required) { $itemRqd = "<b>Required:</b> $($itemCur.required)<br/>" }
                if ($null -ne $itemCur.tcpPorts) { $itemTCP = "<b>TCP Ports:</b> $($itemCur.tcpPorts)<br/>" }
                if ($null -ne $itemCur.udpPorts) { $itemUDP = "<b>UDP Ports:</b> $($itemCur.udpPorts)<br/>" }
                if ($null -ne $itemCur.notes) { $itemNotes = "<b>Notes:</b> $($itemCur.Notes)<br/>" }
            }
            if ($html) { $ipHistory += "<td  style='max-width:150px'>$($itemEP)$($itemSA)$($itemCat)$($itemRqd)$($itemTCP)$($itemUDP)$($itemNotes)</td>`n`t" }
            else { $ipHistory += "<div class='tableInc-cell-l' style='max-width:150px'>$($itemEP)$($itemSA)$($itemCat)$($itemRqd)$($itemTCP)$($itemUDP)$($itemNotes)</div>`n`t" }
            $itemEP, $itemSA, $itemCat, $itemRqd, $itemTCP, $itemUDP, $itemNotes = ""
            $itemPre = ($item.Previous -replace '@{' -replace '}').Split(";") | ConvertFrom-StringData
            if ($item.Previous) {
                if ($null -ne $itemPre.expressroute) { $itemEP = "<b>Express Route:</b> $($itemPre.expressroute)<br/>" }
                if ($null -ne $itemPre.serviceArea) { $itemSA = "<b>Service Area:</b> $($itemPre.serviceArea)<br/>" }
                if ($null -ne $itemPre.category) { $itemCat = "<b>Category:</b> $($itemPre.category)<br/>" }
                if ($null -ne $itemPre.required) { $itemRqd = "<b>Required:</b> $($itemPre.required)<br/>" }
                if ($null -ne $itemPre.tcpPorts) { $itemTCP = "<b>TCP Ports:</b> $($itemPre.tcpPorts)<br/>" }
                if ($null -ne $itemPre.udpPorts) { $itemUDP = "<b>UDP Ports:</b> $($itemPre.udpPorts)<br/>" }
                if ($null -ne $itemPre.notes) { $itemNotes = "<b>Notes:</b> $($itemPre.Notes)<br/>" }
            }
            if ($html) {
                $ipHistory += "<td  style='max-width:150px'>$($itemEP)$($itemSA)$($itemCat)$($itemRqd)$($itemTCP)$($itemUDP)$($itemNotes)</td>`n`t"
                $ipHistory += "</tr>`n"
            }
            else {
                $ipHistory += "<div class='tableInc-cell-l' style='max-width:150px'>$($itemEP)$($itemSA)$($itemCat)$($itemRqd)$($itemTCP)$($itemUDP)$($itemNotes)</div>`n`t"
                $ipHistory += "</div>`n"
            }
        }
        $ipHistoryHTML += $ipHistory
        if ($html) { $ipHistoryHTML += "</tbody></table><br/>`r`n" }
        else { $ipHistoryHTML += "</div></div></div></div><br/>`r`n" }
    }
    return $ipHistoryHTML
}

function IgnoreSSLWarnings {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type) {
        $certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
        Add-Type $certCallback
    }
    [ServerCertificateValidationCallback]::Ignore()	
}


function SendEmail {
    param (
        [Parameter(Mandatory = $true)] [string]$strMessage,
        [Parameter(Mandatory = $false)][AllowNull()] [System.Management.Automation.PSCredential]$credEmail,
        [Parameter(Mandatory = $true)] $config,
        [Parameter(Mandatory = $false)] [string]$strPriority = "Normal",
        [Parameter(Mandatory = $false)] $subject,
        [Parameter(Mandatory = $false)] [string[]]$emailTo,
        [Parameter(Mandatory = $false)] $attachment = ""
    ) 

    [string]$strSubject = $null
    [string]$strHeader = $null
    [string]$strFooter = $null
    [string]$strSig = $null
    [string]$strBody = $null

    #Build and send email (with attachment)

    $strSubject = "M365 [$($config.tenantshortname)]"
    if ($subject) { $strSubject += ": $($subject)" }
    else { $strSubject += ": Alert [$(Get-Date -f 'dd-MMM-yyy HH:mm:ss')]" }
    $css = Get-Content ..\common\O365email.css

    $strHeader = "<!DOCTYPE html PUBLIC ""-//W3C/DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
    $strHeader += "<head>`n<style type=""text/css"">`n" + $css + "</style></head>`n"
    $strBody = "<body><b>Alert [$(Get-Date -f 'dd-MMM-yyy HH:mm:ss')]</b><br/>`r`n"
    $strBody += $strMessage
    $strSig = "<br/><br/>Kind regards<br/>Powershell scheduled task<br/><br/><b><i>(Automated report/email, do not respond)</i></b><br/>`n"
    $strSig += "<font: xx-small>Generated on: $env:computername by: $($env:userdomain)\$($env:username)</font></br>`n"

    $strFooter += "</body>`n"
    $strFooter += "</html>`n"
    $strHTMLBody = $strHeader + $strBody + $strSig + $strFooter
    [array]$strTo = $emailTo.Split(",")
    $strTo = $strTo.replace('"', '')
    if ($null -eq $config.EmailFrom) { break; }

    # Add OUt of hours contact if enabled
    if ($config.OohEnabled =eq $true)
    {
        $day = (get-date).DayOfWeek
        $time = get-date

        #If ($day -eq 'Saturday' -or $day -eq 'Sunday' -or $day -eq 'Monday') #Bank Holiday
        If ($day -eq 'Saturday' -or $day -eq 'Sunday')
        {
            $strTo += $Config.OohEmailTo
        }
        elseif ($time.timeofday -ge $Config.OohMorningStart -and $time.timeofday -le $Config.OohMorningEnd)
        {
            $strTo += $Config.OohEmailTo
        }
        elseif ($time.timeofday -ge $Config.OohEveningStart -and $time.timeofday -le $Config.OohEveningEnd)
        {
            $strTo += $Config.OohEmailTo
        }
    }

    #Splat the parameters
    $params = @{ }
    $params += @{to = $strTo; subject = $strSubject; body = $strHTMLBody; BodyAsHTML = $true; priority = $strPriority; from = $config.emailfrom; smtpServer = $config.EmailHost; port = $config.emailport }
    If ($credemail -notlike '') {
        $params += @{Credential = $credEmail } 
    }
    If ($config.EmailUseSSL -eq 'True') {
        $params += @{UseSSL = $true } 
    }
    if ($attachment -ne "") {
        $params += @{attachments = $attachment } 
    }
    Send-MailMessage @params
}

function cardBuilder {
    Param (
        [Parameter(Mandatory = $true)] $strName,
        [Parameter(Mandatory = $true)] $strDays,
        [Parameter(Mandatory = $true)] $strMessages,
        [Parameter(Mandatory = $true)] $strAdvisories,
        [Parameter(Mandatory = $true)] $strPriority
    ) 
    [array]$rptCard = @()
    $tableClass = "class='card-type-$($strPriority)'"
    $rptCard = @"
    <div $tableClass>
    `t<div class='card-body'>
    `t`t<div class='card-row'>
    `t`t`t<div class='card-name'>$($strName)</div></div></div>
    `t`t`t<div class='card-body'>
    `t`t`t`t<div class='card-row'>
    `t`t`t`t`t<div class='card-number'>$($strDays)</div>
    `t`t`t`t`t<div class='card-text'>Days Since<br/>Last Incident</div>
    `t`t`t`t</div>
    `t`t`t`t<div class='card-row'>
    `t`t`t`t`t<div class='card-number'>$($strMessages)</div>
    `t`t`t`t`t<div class='card-text'>Recent<br/>Incidents</div>
    `t`t`t`t</div>
    `t`t`t`t<div class='card-row'>
    `t`t`t`t`t<div class='card-number'>$($strAdvisories)</div>
    `t`t`t`t`t<div class='card-text'>Listed Advisories</div>
    `t`t`t`t</div>
    `t`t`t</div>`n`t`t</div>
"@
    return $rptCard
}


function featureBuilder {
    Param (
        [Parameter(Mandatory = $true)] $strName,
        [Parameter(Mandatory = $true)] $strFeatures,
        [Parameter(Mandatory = $true)] $strPriority,
        [Parameter(Mandatory = $true)] $intFtCnt
    )
    [array]$rptCard = @()
    [decimal]$decSize = 0
    $decSize = (($intFtCnt * 0.5) + ([math]::ceiling(($strName.length) / 14) * .75) + 0.1) * 2
    [int]$intSize = $decSize
    $tableClass = "class='workload-card-$($strPriority)' style='grid-row: span $($intSize)'"
    $rptCard = @"
    <div $tableClass>
    `t<div class='wkld-name'>$($strName)</div>
    `t$($strFeatures)
    </div>
"@
    return $rptCard
}

function SkuCardBuilder {
    Param (
        [Parameter(Mandatory = $true)] $strName,
        [Parameter(Mandatory = $true)] $strFeatures,
        [Parameter(Mandatory = $true)] $strPriority,
        [Parameter(Mandatory = $true)] $intFtCnt
    ) 
    [array]$rptCard = @()
    [decimal]$decSize = 0
    $decSize = (($intFtCnt * 1) + ([math]::ceiling(($strName.length) / 20) * .75) + 0.1) * 2
    [int]$intSize = $decSize
    $tableClass = "class='sku-card-$($strPriority)' style='grid-row: span $($intSize)'"
    $rptCard = @"
    <div $tableClass>
    `t<div class='sku-name'>$($strName)</div>
    `t$($strFeatures)
    </div>
"@
    return $rptCard
}

function Get-IncidentInHTML {
    Param (
        [Parameter(Mandatory = $true)] $item,
        [Parameter(Mandatory = $true)] $RebuildDocs,
        [Parameter(Mandatory = $true)] $pathHTMLDocs
    )
    # Get the incident message text and change to something nice in HTML.
    # Message text in advisories has different formatting.
    #Get the latest published date
    #If published in the last 2 hours mins then re-build the html - really need to check the published date when it appears.
    [array]$subMessages = @()
    [string]$htmlBuild = ""
    [int]$pubWindow = 0

    #Main item data
    $url = "docs/$($item.ID).html"
    $htmlHead = "<title>$($item.ID) - $($item.WorkloadDisplayName)</title>"
    $css = Get-Content ..\common\article.css
    $htmlHead += $css
    $htmlBody += "<table class='msg'>"
    $htmlBody += "<tr><th colspan=7 style='font-size:x-large;'>$($item.ImpactDescription)</th></tr>"
    $htmlBody += "<tr><th>ID</th><th>Workload<br/>Feature</th><th>Title</th><th>Classification</th><th>Severity</th><th>Start Time</th><th>Last Updated</th></tr>"
    $htmlBody += "<tr class='msgO'><td>$($item.ID)</td><td>$($item.WorkloadDisplayName)<br/>$($item.FeatureDisplayName)</td><td>$($item.Title)</td><td>$($item.Classification)</td><td>$($item.Severity)</td><td>$(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm')</td><td>$(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')</td></tr>"
    $subMessages = $item | Select-Object -ExpandProperty Messages
    $subMessages = $subMessages | Sort-Object publishedtime -Descending
    $pubWindow = (New-TimeSpan -Start (Get-Date $submessages[0].publishedtime) -End $(Get-Date)).TotalHours
    $updWindow = (New-TimeSpan -Start (Get-Date $item.LastUpdatedTime) -End $(Get-Date)).TotalHours
    if ($pubWindow -le 18 -or $RebuildDocs -or $updWindow -le 72) {
        #Article was updated in the last 2 hours. Lets update it Or force rebuild of docs
        foreach ($message in $subMessages) {
            $htmlBuild = Get-htmlMessage $message.messagetext
            $htmlBuild = "<br/><b>Update:</b> $(Get-Date $message.PublishedTime -f 'dd-MMM-yyyy HH:mm')<br/>" + $htmlBuild
            $htmlSub += $htmlBuild + "<hr><br/>"
        }
        $htmlBody += "<tr><td colspan=7>$($htmlSub)</td></tr>"
        $htmlBody += "</table>"
        ConvertTo-Html -Head $htmlHead -Body $htmlBody | Out-File "$($pathHTMLDocs)\$($item.ID).html"
    }
    #Return a link to the file
    return $url
}
function Get-AdvisoryInHTML {
    Param (
        [Parameter(Mandatory = $true)] $item,
        [Parameter(Mandatory = $true)] $RebuildDocs,
        [Parameter(Mandatory = $true)] $pathHTMLDocs
    )
    #Get the latest published date
    #If published in the last 60 mins then re-build the html - really need to check the published date when it appears.
    [array]$subMessages = @()
    [string]$htmlBuild = ""
    [int]$pubWindow = 0
    #Main item data
    $url = "docs/$($item.ID).html"
    $htmlHead = "<title>$($item.ID) - $($item.Title)</title>"
    $css = Get-Content ..\common\article.css
    $htmlHead += $css
    $htmlBody += "<table class='msg'>"
    $htmlBody += "<tr><th colspan=7 style='font-size:x-large;'>$($item.Title)</th></tr>"
    $htmlBody += "<tr><th>ID</th><th>Workload</th><th>Action</th><th>Classification</th><th>Severity</th><th>Start Time</th><th>Last Updated</th></tr>"
    $htmlBody += "<tr class='msgO'><td>$($item.ID)</td><td>$($item.AffectedWorkloadDisplayNames)</td><td>$($item.ActionType)</td><td>$($item.Classification)</td><td>$($item.Severity)</td><td>$(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm')</td><td>$(Get-Date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')</td></tr>"
    $subMessages = $item | Select-Object -ExpandProperty Messages
    $subMessages = $subMessages | Sort-Object publishedtime -Descending
    $pubWindow = (New-TimeSpan -Start (Get-Date $submessages[0].publishedtime) -End $(Get-Date)).TotalHours
    $updWindow = (New-TimeSpan -Start (Get-Date $item.LastUpdatedTime) -End $(Get-Date)).TotalHours
    if ($pubWindow -le 18 -or $RebuildDocs -or $updWindow -le 2) {
        #Article has been updated in the last 2 hours, or force a rebuild of documents
        foreach ($message in $subMessages) {
            $htmlBuild = $message.messagetext
            $htmlBuild = "<br/><b>Update:</b> $(Get-Date $message.PublishedTime -f 'dd-MMM-yyyy HH:mm')<br/>" + $htmlBuild
            $htmlBuild = $htmlBuild -replace "Title:", "<b>Title</b>:"
            $htmlBuild = $htmlBuild -replace "`n", "<br/>"
            $htmlBuild = $htmlBuild -replace ("`n", '<br>') -replace ([char]8217, "'") -replace ([char]8220, '"') -replace ([char]8221, '"') -replace ('\[', '<b><i>') -replace ('\]', '</i></b>')
            $htmlSub += $htmlBuild + "<br/>"
        }
        if ($item.ExternalLink) { $htmlsub += "<a href='$($item.ExternalLink)' target=_blank>Additional Information</a><br/>" }
        $htmlBody += "<tr><td colspan=7>$($htmlSub)</td></tr>"
        $htmlBody += "</table>"
        ConvertTo-Html -Head $htmlHead -Body $htmlBody | Out-File "$($pathHTMLDocs)\$($item.ID).html"
    }
    #Return a link to the file
    return $url
}

function GetSchedEmailTo {
    Param (
        [Parameter(Mandatory = $true)] $nameScript
    )
    [string]$dow = (get-date).DayOfWeek.ToString().Substring(0, 3)
    [datetime]$dtmNow = Get-Date

    Write-Log "Checking scheduled recipients for $nameScript script."
    $filenameSched = ".\schedule.csv"
    $pathSched = resolve-path $filenameSched

    #Check if schedule file exists
    if (Test-Path $($pathSched)) {
        $schedule = import-csv $filenameSched
        write-log "Importing schedule from $pathSched"
    }
    else {
        write-log "No schedule file found to import at $pathSched"
    }
    $schedule = import-csv $filenameSched
    write-log "Importing schedule from $pathSched"

    [boolean]$chkScript = $false
    [boolean]$chkDay = $false
    [boolean]$chkTime = $false

    $strEmail = @()
    foreach ($entry in $schedule) {
        if (($entry.script -like '*') -or ($entry.script -like $nameScript)) { $chkScript = $true } else { $chkScript = $false }
        if (($entry.day -like '`*') -or ($entry.day -match $dow)) { $chkDay = $true } else { $chkDay = $false }
        if ($entry.starttime -like '`*') { $timeStart = "00:00:00" } else { $timeStart = Get-Date $entry.startTime -Format 'HH:mm:ss' }
        if ($entry.endtime -like '`*') { $timeEnd = "23:59:59" } else { $timeEnd = Get-Date $entry.endTime -Format 'HH:mm:ss' }
        if ($dtmnow -ge $timeStart -and $dtmNow -le $timeEnd) { $chkTime = $true } else { $chkTime = $false }
        If ($chkScript -and $chkDay -and $chkTime) { $strEmail += $entry.email }
    }

    Write-Log "It is $($dow) at $(Get-Date $dtmnow -Format HH:mm:ss)"
    [string]$strEmail2 = '"{0}"' -f ($strEmail -join '","')
    Write-Log "The following will be emailed: $($strEmail2)"

    return $strEmail
}

function TeamsPost {
    param (
		[Parameter(Mandatory = $true)] $config,
        [Parameter(Mandatory = $true)] [array]$item
      )

# Target URI for the Teams channel (Messaging Operations\O365 Alerts)
$uri = $config.TeamsURI
$DashboardURL = $config.hosturl + "/" + $config.DashboardHTML
# create the JSON file with your message
$body = ConvertTo-Json -Depth 4 @{
    title = "Office 365 [$($config.tenantshortname)]: New $($item.Severity) $($item.Classification): $($item.WorkloadDisplayName) - $($item.Status) [$($item.ID)]"
    text = "Alert [$(Get-Date -f 'dd-MMM-yyy HH:mm:ss')]"
    themecolor = "FF0000" 
    sections = @(
         @{
            title = 'Incident Details'
            facts = @(
                @{
                name = 'ID:'
                value = $item.id
                },
                @{
                name = 'Tenant:'
                value = $config.tenantshortname
                },
                @{
                name = 'Feature:'
                value = $item.WorkloadDisplayName
                },
                @{
                name = 'Status:'
                value = $item.status
                },
                @{
                name = 'Severity:'
                value = $item.severity
                },
                @{
                name = 'Classification:'
                value = $item.classification
                },
                @{
                name = 'Start Time:'
                value = $(Get-Date $item.StartTime -f 'dd-MMM-yyyy HH:mm')
                },
                @{
                name = 'Last Updated:'
                value = $item.LastUpdatedTime
                },
                @{
                name = 'End Time:'
                value = $item.EndTime
                },
                @{
                name = 'Incident Title:'
                value = $item.title
                }
     
            )
        }
    )
    potentialAction = @(@{
        '@context' = 'http://schema.org'
        '@type' = 'ViewAction'
        name = 'Office 365 Incident Dashboard'
        target = @("$dashboardurl")
    })

}
# send message to Teams channel
Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json' #-Proxy 'http://appproxy.rbsgrp.net:8080' #-ProxyCredential $creds
}