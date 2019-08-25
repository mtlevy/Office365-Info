<#
.SYNOPSIS
	Utility to update an existing profile configuration file.

.DESCRIPTION
	Updates an existing configuration file

.INPUTS
    Existing config.xml file

.OUTPUTS
    Updated config.xml file

.EXAMPLE
    PS C:\> .\ProfileBuilder.ps1 -configXML .\configfile.xml


.NOTES
    Author:  Jonathan Christie
    Email:   jonathan.christie (at) boilerhouseit.com
    Date:    02 Feb 2019
    PSVer:   2.0/3.0/4.0/5.0
    Version: 1.0.1
    Updated:
    UpdNote:

    Wishlist:

    Completed:

    Outstanding:

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [String]$configXML = ""
)

Write-Verbose "Changing Directory to $PSScriptRoot"
Set-Location $PSScriptRoot
$configXML = Resolve-Path $configXML
if (Test-Path $configXML) {
    #Resolve the full path the configuration file
    $configXMLFile = Split-Path $configXML -Leaf
    #Get existing configuration file contents
    $xmlExisting = [xml](Get-Content -Path "$($configXML)")
    #Assign the config to variables. We dont need to but this will allow for checking in future for null values and building a profile via question/input
    $appSettings = [PSCustomObject]@{
        TenantName         = $xmlExisting.Settings.Tenant.Name
        TenantShortName    = $xmlExisting.Settings.Tenant.ShortName
        TenantDescription  = $xmlExisting.Settings.Tenant.Description

        LogPath            = $xmlExisting.Settings.Output.LogPath
        HTMLPath           = $xmlExisting.Settings.Output.HTMLPath
        UseEventLog        = $xmlExisting.Settings.Output.UseEventLog
        EventLog           = $xmlExisting.Settings.Output.EventLog

        TenantID           = $xmlExisting.Settings.Azure.TenantID
        AppID              = $xmlExisting.Settings.Azure.AppID
        AppSecret          = $xmlExisting.Settings.Azure.AppSecret

        WallReportName     = $xmlExisting.Settings.WallDashboard.Name
        WallHTML           = $xmlExisting.Settings.WallDashboard.HTMLFilename
        WallDashCards      = $xmlExisting.Settings.WallDashboard.DashCards
        WallPageRefresh    = $xmlExisting.Settings.WallDashboard.Refresh
        WallEventSource    = $xmlExisting.Settings.WallDashboard.EventSource

        DashboardName      = $xmlExisting.Settings.Dashboard.Name
        DashboardLogo      = $xmlExisting.Settings.Dashboard.Logo
        DashboardRefresh   = $xmlExisting.Settings.Dashboard.Refresh
        DashboardEvtSource = $xmlExisting.Settings.Dashboard.EventSource
        DashboardHTML      = $xmlExisting.Settings.Dashboard.HTMLFilename
        DashboardCards     = $xmlExisting.Settings.Dashboard.DashCards
        DashboardAddLink   = $xmlExisting.Settings.Dashboard.AddLink
        DashboardHistory   = $xmlExisting.Settings.Dashboard.History

        EmailHost          = $xmlExisting.Settings.Email.SMTPServer
        EmailPort          = $xmlExisting.Settings.Email.Port
        EmailUseSSL        = $xmlExisting.Settings.Email.UseSSL
        EmailFrom          = $xmlExisting.Settings.Email.From
        EmailUser          = $xmlExisting.Settings.Email.Username
        EmailPassword      = $xmlExisting.Settings.Email.PasswordFile
        EmailKey           = $xmlExisting.Settings.Email.AESKeyFile

        MonitorAlertsTo    = [string[]]$xmlExisting.Settings.Monitor.alertsTo
        MonitorEvtSource   = $xmlExisting.Settings.Monitor.EventSource

        UsageReportsPath   = $xmlExisting.Settings.UsageReports.Path
        UsageEventSource   = $xmlExisting.Settings.UsageReports.EventSource

        MaxFeedItems       = $xmlExisting.Settings.IPURLs.MaxFeedItems
        IPURLPath          = $xmlExisting.Settings.IPURLs.Path
        IPURLAlertsTo      = $xmlExisting.Settings.IPURLs.AlertsTo

        UseProxy           = $xmlExisting.Settings.Proxy.UseProxy
        ProxyHost          = $xmlExisting.Settings.Proxy.ProxyHost
        ProxyIgnoreSSL     = $xmlExisting.Settings.Proxy.IgnoreSSL
    }

    #set output file
    $xmlNewConfig = @"
<?xml version="1.0"?>
<Settings>
  <Tenant>
    <!-- Basic tenant information. Shortname is used in filenames to help identify tenants-->
    <Name>$($appSettings.TenantName)</Name>
    <!-- Short name is used in filenames to help identify files per tenant-->
    <ShortName>$($appSettings.TenantShortName)</ShortName>
    <Description>$($appSettings.TenantDescription)</Description>
  </Tenant>
  <Azure>
    <!-- Azure AD App information for connectivity to tenant-->
    <TenantID>$($appSettings.TenantID)</TenantID>
    <AppID>$($appSettings.AppID)</AppID>
    <AppSecret>$($appSettings.AppSecret)</AppSecret>
  </Azure>
  <Output>
    <!-- All paths, if not absolute, are relative to the location of the running script-->
    <LogPath>$($appSettings.LogPath)</LogPath>
    <!-- Where any HTML documents should be saved-->
    <HTMLPath>$($appSettings.HTMLPath)</HTMLPath>
    <!-- If using the local event log on the machine that runs the scripts, define which custom event log to use-->
    <UseEventLog>$($appSettings.UseEventLog)</UseEventLog>
    <EventLog>$($appSettings.EventLog)</EventLog>
  </Output>
  <Email>
    <!-- Email server connectivity settings. Can be office365 or other mail system-->
    <!-- SendReport function in the common module can be trimmed if there is no need for username/password (ie internal systems)-->
    <!-- Current settings are required to use exchange online. 'From' should be the authenticated user -->
    <SMTPServer>$($appSettings.EmailHost)</SMTPServer>
    <Port>$($appSettings.EmailPort)</Port>
    <UseSSL>$($appSettings.EmailUseSSL)</UseSSL>
    <From>$($appSettings.EmailFrom)</From>
    <!-- Blank for no authentication (ie internal mail system)-->
    <Username>$($appSettings.EmailUser)</Username>
    <PasswordFile>$($appSettings.EmailPassword)</PasswordFile>
    <AESKeyFile>$($appSettings.EmailKey)</AESKeyFile>
  </Email>
  <Monitor>
    <!-- Where to send monitoring alerts to. Comma separated quoted list "john@home.com","bob@vader.net"-->
    <alertsTo>$($appSettings.MonitorAlertsTo)</alertsTo>
    <!-- Events source to use when logging to the event log-->
    <EventSource>$($appSettings.MonitorEvtSource)</EventSource>
  </Monitor>
  <WallDashboard>
    <Name>$($appSettings.WallReportName)</Name>
    <HTMLFileName>$($appSettings.WallHTML)</HTMLFileName>
    <!-- Always show these cards first on the Wall. This helps with the layout-->
    <DashCards>$($appSettings.WallDashCards)</DashCards>
    <!-- Refresh interval in minutes-->
    <Refresh>$($appSettings.WallPageRefresh)</Refresh>
    <!-- Events source to use when logging to the event log-->
    <EventSource>$($appSettings.WallEventSource)</EventSource>
  </WallDashboard>
  <Dashboard>
    <Name>$($appSettings.DashboardName)</Name>
    <HTMLFileName>$($appSettings.DashboardHTML)</HTMLFileName>
    <!-- Show these dashboard cards in current status -->
    <DashCards>$($appSettings.DashboardCards)</DashCards>
    <!-- Refresh interval in minutes-->
    <Refresh>$($appSettings.DashboardRefresh)</Refresh>
    <!-- Events source to use when logging to the event log-->
    <EventSource>$($appSettings.DashboardEvtSource)</EventSource>
    <Logo>$($appSettings.DashboardLogo)</Logo>
    <AddLink>$($appSettings.DashboardAddLink)</AddLink>
    <!-- Duration to show incidents recently closed (in days)-->
    <History>$($appSettings.DashboardHistory)</History>
  </Dashboard>
  <UsageReports>
    <!-- Where to store the Office 365 Usage Reports (CSV)-->
    <Path>$($appSettings.UsageReportsPath)</Path>
    <!-- Events source to use when logging to the event log-->
    <EventSource>$($appSettings.UsageEventSource)</EventSource>
  </UsageReports>
  <IPURLs>
    <!-- Maximum number of items to return if feed provides more -->
    <MaxFeedItems>$($appSettings.MaxFeedItems)</MaxFeedItems>
    <Path>$($appSettings.IPURLPath)</Path>
    <!-- Where to send updates to IP and URLs to. Comma separated quoted list "john@home.com","bob@vader.net"-->
    <AlertsTo>$($appSettings.IPURLsAlertsTo)</AlertsTo>
  </IPURLs>
  <Proxy>
    <!-- Proxy settings if required-->
	<!-- Use proxy values: true/false-->
    <UseProxy>$($appSettings.UseProxy)</UseProxy>
	<!-- Proxy server FQDN value http://proxyfqdn.domain.com -->
    <ProxyHost>$($appSettings.ProxyHost)</ProxyHost>
	<!-- Ignore SSL: true/false-->
    <IgnoreSSL>$($appSettings.ProxyIgnoreSSL)</IgnoreSSL>
  </Proxy>
</Settings>
"@

    $xmlNewConfig | Set-Content -Path "NEW-$($configXMLFile)"
    #$xmlNewConfig | ConvertTo-Json | Set-Content -Path "New-$($configXMLFile).json"
}