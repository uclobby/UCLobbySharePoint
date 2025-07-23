function Get-UcUPNFromString {
    param (
        [string]$InputStr
    )
    $regexUPN = "([^|]+@[a-zA-Z0-9_\-\.]+\.[a-zA-Z]*)"
    try {
        $RegexTemp = [regex]::Match($InputStr, $regexUPN).captures.groups
        if ($RegexTemp.Count -ge 2) {
            $outUPN = $RegexTemp[1].value
        }
        return $outUPN
    }
    catch {
        return ""
    }
}