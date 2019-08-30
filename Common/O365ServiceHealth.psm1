# Shared functions for Office 365 Info powershell
function saveCredentials {
    param (
        [Parameter(Mandatory = $true)] [string]$Password,
        [Parameter(Mandatory = $true)] [boolean]$CreateKey,
        [Parameter(Mandatory = $true)] [string]$KeyPath,
        [Parameter(Mandatory = $true)] [string]$CredsPath
    ) 

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
        TenantName           = $configFile.Settings.Tenant.Name
        TenantShortName      = $configFile.Settings.Tenant.ShortName
        TenantDescription    = $configFile.Settings.Tenant.Description
    
        TenantID             = $configFile.Settings.Azure.TenantID
        AppID                = $configFile.Settings.Azure.AppID
        AppSecret            = $configFile.Settings.Azure.AppSecret
    
        LogPath              = $configFile.Settings.Output.LogPath
        HTMLPath             = $configFile.Settings.Output.HTMLPath
        UseEventLog          = $configFile.Settings.Output.UseEventLog
        EventLog             = $configFile.Settings.Output.EventLog
		HostURL              = $configFile.Settings.Output.HostURL

        EmailHost            = $configFile.Settings.Email.SMTPServer
        EmailPort            = $configFile.Settings.Email.Port
        EmailUseSSL          = $configFile.Settings.Email.UseSSL
        EmailFrom            = $configFile.Settings.Email.From
        EmailUser            = $configFile.Settings.Email.Username
        EmailPassword        = $configFile.Settings.Email.PasswordFile
        EmailKey             = $configFile.Settings.Email.AESKeyFile

        MonitorAlertsTo      = [string[]]$configFile.Settings.Monitor.alertsTo
        MonitorEvtSource     = $configFile.Settings.Monitor.EventSource
  
        WallReportName       = $configFile.Settings.WallDashboard.Name
        WallHTML             = $configFile.Settings.WallDashboard.HTMLFilename
        WallDashCards        = $configFile.Settings.WallDashboard.DashCards
        WallPageRefresh      = $configFile.Settings.WallDashboard.Refresh
        WallEventSource      = $configFile.Settings.WallDashboard.EventSource

        DashboardName        = $configFile.Settings.Dashboard.Name
        DashboardHTML        = $configFile.Settings.Dashboard.HTMLFilename
        DashboardCards       = $configFile.Settings.Dashboard.DashCards
        DashboardRefresh     = $configFile.Settings.Dashboard.Refresh
		DashboardAlertsTo    = $configFile.Settings.Dashboard.AlertsTo
        DashboardEvtSource   = $configFile.Settings.Dashboard.EventSource
        DashboardLogo        = $configFile.Settings.Dashboard.Logo
        DashboardAddLink     = $configFile.Settings.Dashboard.AddLink
		DashboardHistory     = $configFile.Settings.Dashboard.History

        UsageReportsPath     = $configFile.Settings.UsageReports.Path
        UsageEventSource     = $configFile.Settings.UsageReports.EventSource

		DiagnosticsNotes     = ($configfile.Settings.Diagnostics.Notes).InnerXML
		DiagnosticsWeb       = $configfile.Settings.Diagnostics.Web
		DiagnosticsPorts     = $configfile.Settings.Diagnostics.Ports
		DiagnosticsURLs      = $configfile.Settings.Diagnostics.URLs
		DiagnosticsVerbose   = $configfile.Settings.Diagnostics.Verbose


		MaxFeedItems         = $configFile.Settings.IPURLs.MaxFeedItems
		IPURLPath            = $configFile.Settings.IPURLs.Path
		IPURLAlertsTo        = $configFile.Settings.IPURLs.AlertsTo
    
        UseProxy             = $configFile.Settings.Proxy.UseProxy
        ProxyHost            = $configFile.Settings.Proxy.ProxyHost
        ProxyIgnoreSSL       = $configFile.Settings.Proxy.IgnoreSSL

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
    $icon1="<img src='images/1.jpg' alt='Error' style='width:20px;height:20px;border:0;'>"
    $icon2="<img src='images/2.jpg' alt='Warning' style='width:20px;height:20px;border:0;'>"
    $icon3="<img src='images/3.jpg' alt='OK' style='width:20px;height:20px;border:0;'>"
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
    [System.Net.Mail.MailPriority]$returnValue="Normal"
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
    $htmlMessage = $htmlMessage -replace "`n", "<br/>"

    return $htmlMessage
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


function SendReport {
    param (
        [Parameter(Mandatory = $true)] [string]$strMessage,
        [Parameter(Mandatory = $false)][AllowNull()] [System.Management.Automation.PSCredential]$credEmail,
        [Parameter(Mandatory = $true)] $config,
        [Parameter(Mandatory = $false)] [string]$strPriority = "Normal",
        [Parameter(Mandatory = $false)] $subject,
        [Parameter(Mandatory = $false)] [string[]]$emailTo
    ) 

    [string]$strSubject = $null
    [string]$strHeader = $null
    [string]$strFooter = $null
    [string]$strSig = $null
    [string]$strBody = $null

    #Build and send email (with attachment)

    $strSubject = "Office 365 [$($config.tenantshortname)]"
	if ($subject) { $strSubject += ": $($subject)"}
	else {$strSubject += ": Alert [$(get-date -f 'dd-MMM-yyy HH:mm:ss')]"}
    $strHeader = "<!DOCTYPE html PUBLIC ""-//W3C/DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
    $strHeader += "<head>`n<style type=""text/css"">`nbody {font-family: ""Segoe UI"", Tahoma, Geneva, Verdana, sans-serif;font-size: x-small;}`n"
    $strHeader += "table.border {border-collapse:collapse;border:1px solid silver;} table td.border {border:1px solid silver;} table th{color:white;background-color: #003399;}</style></head>`n"
    $strBody = "<body><b>Alert [$(get-date -f 'dd-MMM-yyy HH:mm:ss')]</b><br/>`r`n"
    $strBody += $strMessage
    $strSig = "<br/><br/>Kind regards<br/>Powershell scheduled task<br/><br/><b><i>(Automated report/email, do not respond)</i></b><br/>`n"
    $strSig += "<font: xx-small>Generated on: $env:computername by: $($env:userdomain)\$($env:username)</font></br>`n"

    $strFooter += "</body>`n"
    $strFooter += "</html>`n"
    $strHTMLBody = $strHeader + $strBody + $strSig + $strFooter
    [array]$strTo = $emailTo.Split(",")
    $strTo = $strTo.replace('"', '')
    if ($credEmail -notlike '') {
        #Credentials supplied
        if ($config.EmailUseSSL -eq 'True') {
            Send-MailMessage -To $strTo -Subject $strSubject -Body $strHTMLBody -BodyAsHtml -Priority $strPriority -From $config.EmailFrom -SmtpServer $config.EmailHost -Port $config.EmailPort -UseSSL -Credential $credEmail
        }
        else {
            Send-MailMessage -To $strTo -Subject $strSubject -Body $strHTMLBody -BodyAsHtml -Priority $strPriority -From $config.EmailFrom -SmtpServer $config.EmailHost -Port $config.EmailPort -Credential $credEmail
        }
    }
    else {
        #No credentials
        if ($config.EmailUseSSL -eq 'True') {
            Send-MailMessage -To $strTo -Subject $strSubject -Body $strHTMLBody -BodyAsHtml -Priority $strPriority -From $config.EmailFrom -SmtpServer $config.EmailHost -Port $config.EmailPort -UseSSL
        }
        else {
            Send-MailMessage -To $strTo -Subject $strSubject -Body $strHTMLBody -BodyAsHtml -Priority $strPriority -From $config.EmailFrom -SmtpServer $config.EmailHost -Port $config.EmailPort
        }
    }
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
    $decSize = (($intFtCnt * 0.5) + ([math]::ceiling(($strName.length) / 14) * .75) + 0.1) * 2
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
    $css = get-content article.css
    $htmlHead += $css
    $htmlBody += "<table class='msg'>"
    $htmlBody += "<tr><th colspan=7 style='font-size:x-large;'>$($item.ImpactDescription)</th></tr>"
    $htmlBody += "<tr><th>ID</th><th>Workload<br/>Feature</th><th>Title</th><th>Classification</th><th>Severity</th><th>Start Time</th><th>Last Updated</th></tr>"
    $htmlBody += "<tr class='msgO'><td>$($item.ID)</td><td>$($item.WorkloadDisplayName)<br/>$($item.FeatureDisplayName)</td><td>$($item.Title)</td><td>$($item.Classification)</td><td>$($item.Severity)</td><td>$(get-date $item.StartTime -f 'dd-MMM-yyyy HH:mm')</td><td>$(get-date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')</td></tr>"
    $subMessages = $item | Select-Object -ExpandProperty Messages
    $subMessages = $subMessages | Sort-Object publishedtime -Descending
    $pubWindow = (New-TimeSpan -Start (Get-Date $submessages[0].publishedtime) -End $(Get-Date)).TotalHours
    if ($pubWindow -le 18 -or $RebuildDocs) {
        #Article was updated in the last 2 hours. Lets update it Or force rebuild of docs
        foreach ($message in $subMessages) {
			$htmlBuild = Get-htmlMessage $message.messagetext
            $htmlBuild = "<br/><b>Update:</b> $(get-date $message.PublishedTime -f 'dd-MMM-yyyy HH:mm')<br/>" + $htmlBuild
            #Data is pulled down differently - do Matts replacements still hold?
            #$htmlBuild=$htmlBuild -replace("`n",'<br>') -replace([char]8217,"'") -replace([char]8220,'"') -replace([char]8221,'"') -replace('\[','<b><i>') -replace('\]','</i></b>')
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
    $css = get-content article.css
    $htmlHead += $css
    $htmlBody += "<table class='msg'>"
    $htmlBody += "<tr><th colspan=7 style='font-size:x-large;'>$($item.Title)</th></tr>"
    $htmlBody += "<tr><th>ID</th><th>Workload</th><th>Action</th><th>Classification</th><th>Severity</th><th>Start Time</th><th>Last Updated</th></tr>"
    $htmlBody += "<tr class='msgO'><td>$($item.ID)</td><td>$($item.AffectedWorkloadDisplayNames)</td><td>$($item.ActionType)</td><td>$($item.Classification)</td><td>$($item.Severity)</td><td>$(get-date $item.StartTime -f 'dd-MMM-yyyy HH:mm')</td><td>$(get-date $item.LastUpdatedTime -f 'dd-MMM-yyyy HH:mm')</td></tr>"
    $subMessages = $item | Select-Object -ExpandProperty Messages
    $subMessages = $subMessages | Sort-Object publishedtime -Descending
    $pubWindow = (New-TimeSpan -Start (Get-Date $submessages[0].publishedtime) -End $(Get-Date)).TotalHours
    if ($pubWindow -le 2 -or $RebuildDocs) {
        #Article has been updated in the last 2 hours, or force a rebuild of documents
        foreach ($message in $subMessages) {
            $htmlBuild = $message.messagetext
            $htmlBuild = "<br/><b>Update:</b> $(get-date $message.PublishedTime -f 'dd-MMM-yyyy HH:mm')<br/>" + $htmlBuild
            $htmlBuild = $htmlBuild -replace "Title:", "<b>Title</b>:"
            $htmlBuild = $htmlBuild -replace "`n", "<br/>"
            #Data is pulled down differently - do Matts replacements still hold?
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