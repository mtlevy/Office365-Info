<#
.SYNOPSIS
	Utility to update an existing profile configuration file. Current version is renamed with date and .bak extension

.DESCRIPTION
	Updates an existing configuration file and backs up the original configuration file with date/time and .bak extension in the same directory as the original.

.INPUTS
    Existing config.xml file

.OUTPUTS
    Updated config.xml file and backup copy of origin

.EXAMPLE
    PS C:\> .\UpdateProfile.ps1 -configXML .\configfile.xml


.NOTES
    Author:  Jonathan Christie
    Email:   jonathan.christie (at) boilerhouseit.com
    Date:    02 Feb 2019
    PSVer:   2.0/3.0/4.0/5.0
    Version: 1.0.8
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
  $configPath = Split-Path $configXML
  #Get existing configuration file contents
  $xmlExisting = [xml](Get-Content -Path "$($configXML)")
  #Assign the config to variables. We dont need to but this will allow for checking in future for null values and building a profile via question/input
  $appSettings = [PSCustomObject]@{
    TenantName             = $xmlExisting.Settings.Tenant.Name
    TenantShortName        = $xmlExisting.Settings.Tenant.ShortName
    TenantMSName           = $xmlExisting.Settings.Tenant.MSName
    TenantDescription      = $xmlExisting.Settings.Tenant.Description

    TenantID               = $xmlExisting.Settings.Azure.TenantID
    AppID                  = $xmlExisting.Settings.Azure.AppID
    AppSecret              = $xmlExisting.Settings.Azure.AppSecret

    LogPath                = $xmlExisting.Settings.Output.LogPath
    HTMLPath               = $xmlExisting.Settings.Output.HTMLPath
    WorkingPath            = $xmlExisting.Settings.Output.WorkingPath
    UseEventLog            = $xmlExisting.Settings.Output.UseEventLog
    EventLog               = $xmlExisting.Settings.Output.EventLog
    HostURL                = $xmlExisting.Settings.Output.HostURL

    EmailEnabled           = $xmlExisting.Settings.Email.Enabled
    EmailHost              = $xmlExisting.Settings.Email.SMTPServer
    EmailPort              = $xmlExisting.Settings.Email.Port
    EmailUseSSL            = $xmlExisting.Settings.Email.UseSSL
    EmailFrom              = $xmlExisting.Settings.Email.From
    EmailUser              = $xmlExisting.Settings.Email.Username
    EmailPassword          = $xmlExisting.Settings.Email.PasswordFile
    EmailKey               = $xmlExisting.Settings.Email.AESKeyFile

    MonitorAlertsTo        = [string[]]$xmlExisting.Settings.Monitor.alertsTo
    MonitorEvtSource       = $xmlExisting.Settings.Monitor.EventSource
    MonitorIgnoreSvc       = [string[]]$xmlExisting.Settings.Monitor.IgnoreSvc
    MonitorIgnoreInc       = [string[]]$xmlExisting.Settings.Monitor.IgnoreInc

    WallReportName         = $xmlExisting.Settings.WallDashboard.Name
    WallHTML               = $xmlExisting.Settings.WallDashboard.HTMLFilename
    WallDashCards          = $xmlExisting.Settings.WallDashboard.DashCards
    WallPageRefresh        = $xmlExisting.Settings.WallDashboard.Refresh
    WallEventSource        = $xmlExisting.Settings.WallDashboard.EventSource

    DashboardName          = $xmlExisting.Settings.Dashboard.Name
    DashboardHTML          = $xmlExisting.Settings.Dashboard.HTMLFilename
    DashboardCards         = $xmlExisting.Settings.Dashboard.DashCards
    DashboardRefresh       = $xmlExisting.Settings.Dashboard.Refresh
    DashboardAlertsTo      = $xmlExisting.Settings.Dashboard.AlertsTo
    DashboardEvtSource     = $xmlExisting.Settings.Dashboard.EventSource
    DashboardLogo          = $xmlExisting.Settings.Dashboard.Logo
    DashboardAddLink       = $xmlExisting.Settings.Dashboard.AddLink
    DashboardHistory       = $xmlExisting.Settings.Dashboard.History
    DashboardRecMsgs       = $xmlExisting.Settings.Dashboard.RecentMsgs

    UsageReportsPath       = $xmlExisting.Settings.UsageReports.Path
    UsageEventSource       = $xmlExisting.Settings.UsageReports.EventSource

    ToolboxName            = $xmlExisting.Settings.Toolbox.Name
    ToolboxHTML            = $xmlExisting.Settings.Toolbox.HTMLFilename
    ToolboxNotes           = ($xmlExisting.Settings.Toolbox.Notes).InnerXML
    ToolboxRefresh         = $xmlExisting.Settings.Toolbox.Refresh


    DiagnosticsEnabled     = $xmlExisting.Settings.Diagnostics.Enabled
    DiagnosticsURLs        = $xmlExisting.Settings.Diagnostics.URLs
    DiagnosticsVerbose     = $xmlExisting.Settings.Diagnostics.Verbose

    MiscDiagnosticsEnabled = $xmlExisting.Settings.MiscDiagnostics.Enabled
    MiscDiagsWeb           = $xmlExisting.Settings.MiscDiagnostics.Web
    MiscDiagsPorts         = $xmlExisting.Settings.MiscDiagnostics.Ports

    RSS1Enabled            = $xmlExisting.Settings.RSSFeeds.F1.Enabled
    RSS1Name               = $xmlExisting.Settings.RSSFeeds.F1.Name
    RSS1Feed               = $xmlExisting.Settings.RSSFeeds.F1.Feed
    RSS1URL                = $xmlExisting.Settings.RSSFeeds.F1.URL
    RSS1Items              = $xmlExisting.Settings.RSSFeeds.F1.Items

    RSS2Enabled            = $xmlExisting.Settings.RSSFeeds.F2.Enabled
    RSS2Name               = $xmlExisting.Settings.RSSFeeds.F2.Name
    RSS2Feed               = $xmlExisting.Settings.RSSFeeds.F2.Feed
    RSS2URL                = $xmlExisting.Settings.RSSFeeds.F2.URL
    RSS2Items              = $xmlExisting.Settings.RSSFeeds.F2.Items

    IPURLsPath             = $xmlExisting.Settings.IPURLs.Path
    IPURLsAlertsTo         = $xmlExisting.Settings.IPURLs.AlertsTo
    IPURLsNotesFilename    = $xmlExisting.Settings.IPURLs.NotesFilename
    CustomNotesFilename    = $xmlExisting.Settings.IPURLs.CustomNotesFilename
    IPURLHistory           = $xmlExisting.Settings.IPURLs.History

    CnameEnabled           = $xmlExisting.Settings.CNAME.Enabled
    CnameNotes             = ($xmlExisting.Settings.CNAME.Notes).InnerXML
    CnameFilename          = $xmlExisting.Settings.CNAME.Filename
    CnameAlertsTo          = $xmlExisting.Settings.CNAME.AlertsTo
    CnameURLs              = $xmlExisting.Settings.CNAME.URLs
    CnameResolvers         = [string[]]$xmlExisting.Settings.CNAME.Resolvers
    CnameResolverDesc      = [string[]]$xmlExisting.Settings.CNAME.ResolverDesc

    PACEnabled             = $xmlExisting.Settings.PACFile.Enabled
    PACProxy               = $xmlExisting.Settings.PACFile.Proxy
    PACType1Filename       = $xmlExisting.Settings.PACFile.Type1Filename
    PACType2Filename       = $xmlExisting.Settings.PACFile.Type2Filename
    
    ProxyEnabled           = $xmlExisting.Settings.Proxy.ProxyEnabled
    ProxyHost              = $xmlExisting.Settings.Proxy.ProxyHost
    ProxyIgnoreSSL         = $xmlExisting.Settings.Proxy.IgnoreSSL

    Blogs                  = ($xmlExisting.Settings.Blogs).InnerXML
  }

  #set output file

  #Set Defaults
  if ([string]::IsNullOrEmpty($appSettings.DashboardRecMsgs)) { $appSettings.DashboardRecMsgs = "7" }
  if ([string]::IsNullOrEmpty($appSettings.UseEventLog)) { $appSettings.UseEventLog = "false" }
  if ([string]::IsNullOrEmpty($appSettings.EmailEnabled)) { $appSettings.EmailEnabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.ToolboxHTML)) { $appSettings.ToolboxHTML = "O365Toolbox.HTML" }
  if ([string]::IsNullOrEmpty($appSettings.ToolboxRefresh)) { $appSettings.ToolboxRefresh = "5" }
  if ([string]::IsNullOrEmpty($appSettings.DiagnosticsEnabled)) { $appSettings.DiagnosticsEnabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.MiscDiagnosticsEnabled)) { $appSettings.MiscDiagnosticsEnabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.MiscDiagsWeb)) { $appSettings.MiscDiagsWeb = "false" }
  if ([string]::IsNullOrEmpty($appSettings.MiscDiagsPorts)) { $appSettings.MiscDiagsPorts = "false" }
  if ([string]::IsNullOrEmpty($appSettings.RSS1Enabled)) { $appSettings.RSS1Enabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.RSS2Enabled)) { $appSettings.RSS2Enabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.CnameEnabled)) { $appSettings.CnameEnabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.PACEnabled)) { $appSettings.PACEnabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.ProxyEnabled)) { $appSettings.ProxyEnabled = "false" }
  if ([string]::IsNullOrEmpty($appSettings.Blogs)) { $appSettings.Blogs = "	<a href=""https://techcommunity.microsoft.com/t5/Office-365-Blog/bg-p/Office365Blog"">Office 365 Blog</a><br /><a href=""https://status.azure.com/en-gb/status"">Azure Status</a><br />" }

  $xmlNewConfig = @"
<?xml version="1.0"?>
<!-- Build date $(Get-Date -f 'dd-MMM-yyyy HH:mm:ss') -->
<Settings>
  <Tenant>
    <!-- Basic tenant information. Shortname is used in filenames to help identify tenants-->
    <Name>$($appSettings.TenantName)</Name>
    <!-- Short name is used in filenames to help identify files per tenant-->
    <ShortName>$($appSettings.TenantShortName)</ShortName>
    <!-- MS name is the name used to create the tenant (which may be used as shortname, above)-->
    <MSName>$($appSettings.TenantMSName)</MSName>
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
    <!-- Where any working files should be saved-->
    <WorkingPath>$($appSettings.WorkingPath)</WorkingPath>
    <!-- If using the local event log on the machine that runs the scripts, define which custom event log to use-->
    <!-- Uses True or False-->
    <UseEventLog>$($appSettings.UseEventLog)</UseEventLog>
    <EventLog>$($appSettings.EventLog)</EventLog>
    <HostURL>$($appSettings.HostURL)</HostURL>
  </Output>
  <Email>
    <!-- Email server connectivity settings. Can be office365 or other mail system-->
    <!-- SendReport function in the common module can be trimmed if there is no need for username/password (ie internal systems)-->
    <!-- Current settings are required to use exchange online. 'From' should be the authenticated user -->
    <!-- Uses True or False-->
    <Enabled>$($appSettings.EmailEnabled)</Enabled>
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
    <!-- Ignore the following Services (from Workload list, feature list) Comma separated quoted list -->
    <IgnoreSvc>$($appSettings.MonitorIgnoreSvc)</IgnoreSvc>
    <!-- Ignore the following Workload Incidents (from Incidents list, monitor alerts) Comma separated quoted list  -->
    <IgnoreInc>$($appSettings.MonitorIgnoreInc)</IgnoreInc>
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
    <!-- Send alert emails to -->
    <AlertsTo>$($appSettings.DashboardAlertsTo)</AlertsTo>
    <!-- Events source to use when logging to the event log-->
    <EventSource>$($appSettings.DashboardEvtSource)</EventSource>
    <Logo>$($appSettings.DashboardLogo)</Logo>
    <AddLink>$($appSettings.DashboardAddLink)</AddLink>
    <!-- Duration to show incidents recently closed (in days)-->
    <History>$($appSettings.DashboardHistory)</History>
    <RecentMsgs>$($appSettings.DashboardRecMsgs)</RecentMsgs>
  </Dashboard>
  <UsageReports>
    <!-- Where to store the Office 365 Usage Reports (CSV)-->
    <Path>$($appSettings.UsageReportsPath)</Path>
    <!-- Events source to use when logging to the event log-->
    <EventSource>$($appSettings.UsageEventSource)</EventSource>
  </UsageReports>
  <Toolbox>
    <Name>$($appSettings.ToolboxName)</Name>
    <HTMLFileName>$($appSettings.ToolboxHTML)</HTMLFileName>
    <!-- Text to add to Diagnostics tab. Will be converted to HTML so can include HTML tags-->
    <Notes>$($appSettings.ToolboxNotes)</Notes>
    <!-- Refresh interval in minutes-->
    <Refresh>$($appSettings.ToolboxRefresh)</Refresh>
  </Toolbox>
  <Diagnostics>
    <!-- Uses True or False-->
    <Enabled>$($appSettings.DiagnosticsEnabled)</Enabled>
    <!-- Run http/https connectivity tests to URLs: true/false-->
    <URLs>$($appSettings.DiagnosticsURLs)</URLs>
    <!-- Show detailed errors for pages: true/false-->
    <Verbose>$($appSettings.DiagnosticsVerbose)</Verbose>
  </Diagnostics>
  <MiscDiagnostics>
    <!-- Additional diagnostic checks that run after the dynamics Microsoft URLs-->
    <!-- Uses True or False-->
    <Enabled>$($appSettings.MiscDiagnosticsEnabled)</Enabled>
    <!-- Run http/https tests for IP connections: true/false-->
    <Web>$($appSettings.MiscDiagsWeb)</Web>
    <!-- Run port connectivity tests: true/false-->
    <Ports>$($appSettings.MiscDiagsPorts)</Ports>
  </MiscDiagnostics>
  <RSSFeeds>
    <!-- Microsoft 365 RSS Feed settings-->
    <F1>
      <!-- Uses True or False-->
      <Enabled>$($appSettings.RSS1Enabled)</Enabled>
      <Name>$($appSettings.RSS1Name)</Name>
      <Feed>$($appSettings.RSS1Feed)</Feed>
      <URL>$($appSettings.RSS1URL)</URL>
      <!-- Maximum number of items to return if feed provides more -->
      <Items>$($appSettings.RSS1Items)</Items>
    </F1>
    <!-- Azure Updates RSS Feed settings-->
    <F2>
      <!-- Uses True or False-->
      <Enabled>$($appSettings.RSS2Enabled)</Enabled>
      <Name>$($appSettings.RSS2Name)</Name>
      <Feed>$($appSettings.RSS2Feed)</Feed>
      <URL>$($appSettings.RSS2URL)</URL>
      <!-- Maximum number of items to return if feed provides more -->
      <Items>$($appSettings.RSS2Items)</Items>
    </F2>
  </RSSFeeds>
  <IPURLs>
    <Path>$($appSettings.IPURLsPath)</Path>
    <!-- Where to send updates to IP and URLs to. Comma separated quoted list "john@home.com","bob@vader.net"-->
    <AlertsTo>$($appSettings.IPURLsAlertsTo)</AlertsTo>
    <!-- Custom CSV file to hold additional information relation to URLs. Matches URL list on ID and URL. System will append short tenant name when loading-->
    <NotesFilename>$($appSettings.IPURLsNotesFilename)</NotesFilename>
    <!-- Custom CSV file to hold additional URLs. System will append short tenant name when loading -->
    <CustomNotesFilename>$($appSettings.CustomNotesFilename)</CustomNotesFilename>
    <!-- Maximum number of items to return if feed provides more -->
    <History>$($appSettings.IPURLHistory)</History>
  </IPURLs>
  <CNAME>
    <!-- CNAME checking enabled -->
    <!-- Uses True or False-->
    <Enabled>$($appSettings.CnameEnabled)</Enabled>
    <!-- Text to add to Information section. Will be converted to HTML so can include HTML tags-->
    <Notes>$($appSettings.CnameNotes)</Notes>
    <!-- Filename to pre-pend to IP lookups ie 'CNAMEs' -->
    <Filename>$($appSettings.CnameFilename)</Filename>
    <!-- Where to send change/error detection to. Comma separated quoted list "john@home.com","bob@vader.net" -->
    <AlertsTo>$($appSettings.CnameAlertsTo)</AlertsTo>
    <!-- URLs to check CNAMEs against. Comma separated quoted list "outlook.office.com","outlook.office365.com" -->
    <URLs>$($appSettings.CnameURLs)</URLs>
    <!-- List of resolvers to test CNAMES ie "dns1.mydomain.com","8.8.8.8" -->
    <Resolvers>$($appSettings.CnameResolvers)</Resolvers>
    <!-- List of descriptions matching the above resolvers ie "Internal DNS","Google DNS" -->
    <ResolverDesc>$($appSettings.CnameResolverDesc)</ResolverDesc>
  </CNAME>
  <PACFile>
    <!-- Proxy .pac file generation required?-->
    <!-- Uses True or False-->
    <Enabled>$($appSettings.PACEnabled)</Enabled>
    <!-- Client proxy server to specificy in .pac file-->
    <Proxy>$($appSettings.PACProxy)</Proxy>
    <!-- If using .pac extension remember to allow on web server as valid extension-->
    <!-- If in doubt use .txt and rename on download-->
    <Type1Filename>$($appSettings.PACType1Filename)</Type1Filename>
    <Type2Filename>$($appSettings.PACType2Filename)</Type2Filename>
  </PACFile>
  <Proxy>
    <!-- Proxy settings if required-->
    <!-- Use proxy values: true/false-->
    <ProxyEnabled>$($appSettings.ProxyEnabled)</ProxyEnabled>
    <!-- Proxy server FQDN value http://proxyfqdn.domain.com:8080 -->
    <ProxyHost>$($appSettings.ProxyHost)</ProxyHost>
    <!-- Ignore SSL: true/false-->
    <IgnoreSSL>$($appSettings.ProxyIgnoreSSL)</IgnoreSSL>
  </Proxy>
  <!-- Blogs is a simple HTML list of useful links -->
  <Blogs>
	$($appSettings.Blogs)
  </Blogs>
</Settings>
"@
  $datetime = Get-Date -Format "yyyyMMddHHmm"
  #Copy existing config file to back
  Copy-Item "$($configPath)\$($configXMLFile)" "$($configPath)\$($configXMLFile)-$($datetime).bak"
  #write new settings to config file
  $xmlNewConfig | Set-Content -Path "$($configPath)\$($configXMLFile)"
}