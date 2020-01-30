$Output = "Welcome to the diagnostics tool"
$tmpOutput = "This tool was run on $($env:computername) at $(get-date -f 'HH:mm dd-MMM-yyyy')"
$output += $tmpOutput + "`r`n`r`n"
$proxyDefault = [system.net.webproxy]::GetDefaultProxy()
[system.net.globalproxyselection]::Select = [System.Net.GlobalProxySelection]::GetEmptyWebProxy()


function TestConnection {
    param (
        [Parameter(Mandatory = $true)] [string]$strWebURL
    )
    $measure = $null
    $testWeb = $null
    $testResp = $null
    try {
        $measure = Measure-Command {
            [System.Net.ServicePointManager]::DefaultConnectionLimit = 1024
            $testWeb = [System.Net.WebRequest]::Create($strWebURL)
            $testWeb.AllowAutoRedirect = $false
            $testResp = $testWeb.GetResponse()
        }
        if ($null -ne $testResp) {
            if ($testResp.statusCode -like 'OK') {
                $tmpOutput = "Good results for $($strWebURL)"
            }
            else {
                $tmpOutput = "Status Code $($testResp.StatusCode) - $($testResp.StatusDescription) for $($strWebURL). "
            }
            if ($testResp.ResponseUri.OriginalString -ne $strWebURL) {
                $tmpOutput += "Response contains a redirect to alternate web page. "
            }
        }
        else {
            $tmpOutput = "No response. Can destination be reached for $($strWebURL)"
        }
        $tmpOutput += ": $($measure.TotalSeconds)`r`n"
    }
    catch {
        $tmpOutput = "Exception calling $strWebURL`r`n$($error[0].exception.message)`r`n"
    }
    return $tmpOutput
}

#Test optimized connections
#(Only direct should pass)
$output += "Direct Connections: `r`n"
$testURL = "http://outlook.office.com"
$Output += TestConnection $testURL
$testURL = "http://outlook.office365.com"
$Output += TestConnection $testURL
$testURL = "https://outlook.office.com"
$Output += TestConnection $testURL
$testURL = "https://outlook.office365.com"
$Output += TestConnection $testURL

#These should fail if non-optimized URLs are forced through proxy
$testURL = "https://portal.office.com"
$Output += TestConnection $testURL
$testURL = "https://admin.microsoft.com"
$Output += TestConnection $testURL


#Re-instate proxy
$output += "Using default proxy for connection:`r`n"
[System.Net.GlobalProxySelection]::select = $proxyDefault
#Run same tests (should all pass)
$testURL = "http://outlook.office.com"
$Output += TestConnection $testURL
$testURL = "http://outlook.office365.com"
$Output += TestConnection $testURL
$testURL = "https://outlook.office.com"
$Output += TestConnection $testURL
$testURL = "https://outlook.office365.com"
$Output += TestConnection $testURL

$outFile = "$env:temp\clientDiags-$(get-date -f 'yyyyMMddTHHmmss').txt"
$output | Out-File $outFile
Start-Process notepad $outFile

