
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [String]$configXML = "..\config\profile-bhitprod.xml"
)


$swScript = [system.diagnostics.stopwatch]::StartNew()
Write-Verbose "Changing Directory to $PSScriptRoot"
Set-Location $PSScriptRoot
Import-Module "..\common\O365ServiceHealth.psm1"


if ([system.IO.path]::IsPathRooted($configXML) -eq $false) {
    #its not an absolute path. Find the absolute path
    $configXML = Resolve-Path $configXML
}
$config = LoadConfig $configXML


[string]$pathIPURLs = $config.IPURLsPath
[string]$pathHTML = $config.HTMLPath
[string]$rptProfile = $config.TenantShortName


$pathIPURLs = CheckDirectory $pathIPURLs
$pathIPurl = $pathIPURLs + "\O365_endpoints_urls-$($rptProfile).csv"
$flatUrls = Import-Csv $pathIPurl
$diagConnect = $null

$urls = @($flatUrls | Where-Object { $_.category -like 'optimize' -and $_.tcpports -like '*80*' -and $_.url -notmatch '\*' })
$diagConnect = @"

    `$output += "``r``nOptimize category : Direct Connections (HTTP): ``r``n"`r`n
"@
foreach ($url in $urls) {
    $diagConnect += @"
    `$testURL = "http://$($url.url)"
    `$Output += TestConnection `$testURL`r`n
"@
}

$urls = @($flatUrls | Where-Object { $_.category -like 'optimize' -and $_.tcpports -like '*443*' -and $_.url -notmatch '\*' })
$diagConnect += @"

    `$output += "``r``nOptimize Category : Direct Connections (HTTPs): ``r``n"`r`n
"@
foreach ($url in $urls) {
    $diagConnect += @"
    `$testURL = "https://$($url.url)"
    `$Output += TestConnection `$testURL`r`n
"@
}

$diagConnect += @"

#Re-instate proxy
`$output += "Using default proxy for connection:``r``n"`r`n
[System.Net.GlobalProxySelection]::select = `$proxyDefault`r`n

"@

$urls = @($flatUrls | Where-Object { $_.category -like 'optimize' -and $_.tcpports -like '*80*' -and $_.url -notmatch '\*' })
$diagConnect += @"

    `$output += "``r``nOptimize category : Using Proxy Connection (HTTP): ``r``n"`r`n
"@
foreach ($url in $urls) {
    $diagConnect += @"
    `$testURL = "http://$($url.url)"
    `$Output += TestConnection `$testURL`r`n
"@
}

$urls = @($flatUrls | Where-Object { $_.category -like 'optimize' -and $_.tcpports -like '*443*' -and $_.url -notmatch '\*' })
$diagConnect += @"
    
    `$output += "``r``nOptimize Category : Using Proxy Connection (HTTPS): ``r``n"`r`n
"@
foreach ($url in $urls) {
    $diagConnect += @"
    `$testURL = "https://$($url.url)"
    `$Output += TestConnection `$testURL`r`n
"@
}

$urls = @($flatUrls | Where-Object { $_.category -like 'allow' -and $_.tcpports -like '*80*' -and $_.url -notmatch '\*' })
$diagConnect += @"
    
    `$output += "``r``nAllow category : Using Proxy Connection (HTTP): ``r``n"`r`n
"@
foreach ($url in $urls) {
    $diagConnect += @"
    `$testURL = "http://$($url.url)"
    `$Output += TestConnection `$testURL`r`n
"@
}

$urls = @($flatUrls | Where-Object { $_.category -like 'allow' -and $_.tcpports -like '*443*' -and $_.url -notmatch '\*' })
$diagConnect += @"

    `$output += "``r``nAllow Category : Using Proxy Connection (HTTPS): ``r``n"`r`n
"@
foreach ($url in $urls) {
    $diagConnect += @"
    `$testURL = "https://$($url.url)"
    `$Output += TestConnection `$testURL`r`n
"@
}


$diagsStart = @"
`$Output = "Welcome to the diagnostics tool``r``n"
`$tmpOutput = "This tool was run on `$(`$env:computername) at `$(get-date -f 'HH:mm dd-MMM-yyyy')"
`$output += `$tmpOutput + "``r``n``r``n"
`$proxyDefault = [system.net.webproxy]::GetDefaultProxy()
[system.net.globalproxyselection]::Select = [System.Net.GlobalProxySelection]::GetEmptyWebProxy()


function TestConnection {
    param (
        [Parameter(Mandatory = `$true)] [string]`$strWebURL
    )
    `$measure = `$null
    `$testWeb = `$null
    `$testResp = `$null
    try {
        `$measure = Measure-Command {
            [System.Net.ServicePointManager]::DefaultConnectionLimit = 1024
            `$testWeb = [System.Net.WebRequest]::Create(`$strWebURL)
            `$testWeb.AllowAutoRedirect = `$false
            `$testResp = `$testWeb.GetResponse()
        }
        if (`$null -ne `$testResp) {
            if (`$testResp.statusCode -like 'OK') {
                `$tmpOutput = "Good results for `$(`$strWebURL)"
            }
            else {
                `$tmpOutput = "Status Code `$(`$testResp.StatusCode) - `$(`$testResp.StatusDescription) for `$(`$strWebURL). "
            }
            if (`$testResp.ResponseUri.OriginalString -ne `$strWebURL) {
                `$tmpOutput += "Response contains a redirect to alternate web page. "
            }
        }
        else {
            `$tmpOutput = "No response. Can destination be reached for `$(`$strWebURL)"
        }
        `$tmpOutput += ": `$(`$measure.TotalSeconds)``r``n"
    }
    catch {
        `$tmpOutput = "Exception calling `$strWebURL``r``n`$(`$error[0].exception.message)``r``n"
    }
    return `$tmpOutput
}
"@


$diagsEnd = @"
`$outFile = "`$env:temp\clientDiags-`$(get-date -f 'yyyyMMddTHHmmss').txt"
`$output | Out-File `$outFile
Start-Process notepad `$outFile
"@

$clientDiags = $diagsStart + $diagConnect + $diagsEnd
$clientDiags | Out-File "$($pathHTML)\ClientDiags.txt" -Encoding ascii