
function saveCredentials {
    param (
        [Parameter(Mandatory = $true)] [string]$Password,
        [Parameter(Mandatory = $true)] [bool]$CreateKey,
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
        TenantName        = $configFile.Settings.Tenant.Name
        TenantShortName   = $configFile.Settings.Tenant.ShortName
        TenantDescription = $configFile.Settings.Tenant.Description
    
        LogPath           = $configFile.Settings.Output.LogPath
        HTMLPath          = $configFile.Settings.Output.HTMLPath
        HTMLFileName      = $configFile.Settings.Output.HTMLFilename
        UseEventLog       = $configFile.Settings.Output.UseEventLog
        EventLog          = $configFile.Settings.Output.EventLog
    
        TenantID          = $configFile.Settings.Azure.TenantID
        AppID             = $configFile.Settings.Azure.AppID
        AppSecret         = $configFile.Settings.Azure.AppSecret
    
        WallReportName    = $configFile.Settings.WallDashboard.Name
        WallPageRefresh   = $configFile.Settings.WallDashboard.Refresh
        WallEventSource   = $configFile.Settings.WallDashboard.EventSource

        DashboardName     = $configFile.Settings.Dashboard.Name
        DashboardLogo     = $configFile.Settings.Dashboard.Logo
        DashboardRefresh  = $configFile.Settings.Dashboard.Refresh
        DashboardEvtSource= $configFile.Settings.Dashboard.EventSource

        EmailHost         = $configFile.Settings.Email.SMTPServer
        EmailPort         = $configFile.Settings.Email.Port
        EmailUseSSL       = $configFile.Settings.Email.UseSSL
        EmailFrom         = $configFile.Settings.Email.From
        EmailUser         = $configFile.Settings.Email.Username
        EmailPassword     = $configFile.Settings.Email.PasswordFile
        EmailKey          = $configFile.Settings.Email.AESKeyFile

        MonitorAlertsTo   = [string[]]$configFile.Settings.Monitor.alertsTo
        MonitorEvtSource  = $configFile.Settings.Monitor.EventSource

        UsageReportsPath  = $configFile.Settings.UsageReports.Path
        UsageEventSource  = $configFile.Settings.UsageReports.EventSource
    
        UseProxy          = $configFile.Settings.Proxy.UseProxy
        ProxyHost         = $configFile.Settings.Proxy.ProxyHost
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
    #Each service status that is available is mapped to one of the levels - OK (3), warning (2) and error (1)
    switch ($type) {
        "icon" {
            switch ($statusName) {
                "ExtendedRecovery" { $StatusDisplay = $icon1 }
                "FalsePositive" { $StatusDisplay = $icon3 }
                "Investigating" { $StatusDisplay = $icon2 }
                "RestoringService" { $StatusDisplay = $icon2 }
                "ServiceDegradation" { $StatusDisplay = $icon1 }
                "ServiceInterruption" { $StatusDisplay = $icon1 }
                "ServiceOperational" { $StatusDisplay = $icon3 }
                "ServiceRestored" { $StatusDisplay = $icon3 }
                #Set default error icon if the status is not listed
                default { $StatusDisplay = $icon1 }
            }
        }
        "class" {
            switch ($statusName) {
                "ExtendedRecovery" { $StatusDisplay = "err" }
                "FalsePositive" { $StatusDisplay = "ok" }
                "Investigating" { $StatusDisplay = "warn" }
                "RestoringService" { $StatusDisplay = "warn" }
                "ServiceDegradation" { $StatusDisplay = "err" }
                "ServiceInterruption" { $StatusDisplay = "err" }
                "ServiceOperational" { $StatusDisplay = "ok" }
                "ServiceRestored" { $StatusDisplay = "ok" }
                #Set default error colour if the status is not listed
                default { $StatusDisplay = "defcon" }
            }
        }
    }
    return $StatusDisplay
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

function SendReport {
    param (
        [Parameter(Mandatory = $true)] [string]$strMessage,
        [Parameter(Mandatory = $true)] [security.Securestring]$credEmail,
        [Parameter(Mandatory = $true)] $config,
        [Parameter(Mandatory = $false)] [string]$strPriority = "Normal"
    ) 

    [string]$strSubject = $null
    [string]$strHeader = $null
    [string]$strFooter = $null
    [string]$strSig = $null
    [string]$strBody = $null

    #Build and send email (with attachment)

    $strSubject = "Office 365 Checker: Alert [$(get-date -f 'dd-MMM-yyy HH:mm:ss')]"
    $strHeader = "<!DOCTYPE html PUBLIC ""-//W3C/DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
    $strHeader += "<head>`n<style type=""text/css"">`nbody {font-family: ""Segoe UI"", Tahoma, Geneva, Verdana, sans-serif;font-size: x-small;}`n"
    $strHeader += "table.border {border-collapse:collapse;border:1px solid silver;} table td.border {border:1px solid silver;} table th{color:white;background-color: #003399;}</style></head>`n"
    $strBody = "<body>`n"
    $strBody += $strMessage
    $strSig = "<br/><br/>Kind regards<br/>Powershell scheduled task<br/><br/><b><i>(Automated report/email, do not respond)</i></b><br/>`n"
    $strSig += "<font: xx-small>Generated on: $env:computername by: $($env:userdomain)\$($env:username)</font></br>`n"

    $strFooter += "</body>`n"
    $strFooter += "</html>`n"
    $strHTMLBody = $strHeader + $strBody + $strSig + $strFooter
    [array]$strTo = $config.MonitorAlertsTo.Split(",")
    $strTo = $strTo.replace('"', '')
    if ($config.EmailUseSSL -eq 'True') {
        Send-MailMessage -To $strTo -Subject $strSubject -Body $strHTMLBody -BodyAsHtml -Priority $strPriority -From $config.EmailFrom -SmtpServer $config.EmailHost -Port $config.EmailPort -UseSSL -Credential $credEmail
    }
    else {
        Send-MailMessage -To $strTo -Subject $strSubject -Body $strHTMLBody -BodyAsHtml -Priority $strPriority -From $config.EmailFrom -SmtpServer $config.EmailHost -Port $config.EmailPort -Credential $credEmail
    }
}
