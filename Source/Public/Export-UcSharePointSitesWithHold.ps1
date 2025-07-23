function Export-UcSharePointSitesWithHold {
    <#
        .SYNOPSIS
        Report with SharePoint Sites/OneDrive with a hold in place.

        .DESCRIPTION
        This script will generate a csv file with Sites/OneDrives with a hold in place.

        Author: David Paulino

        Requirements:   SharePoint Online PowerShell 
                            Install-Module -Name Microsoft.Online.SharePoint.PowerShell
                            Connect-SPOService -Url https://contoso-admin.sharepoint.com
                        Security & Compliance PowerShell 
                            Install-Module -Name ExchangeOnlineManagement
                            Connect-IPPSSession -UserPrincipalName user.adm@contoso.onmicrosoft.com
        
        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results. By default, it will save on current user Download.

        .PARAMETER IncludeOneDrive
        If OneDrives should also be included in the report.

        .EXAMPLE 
        PS>  Export-UcSharePointSitesWithHold

        .EXAMPLE 
        PS>  Export-UcSharePointSitesWithHold -IncludeOneDrive
    #>

    param(
        [string]$OutputPath,
        [switch]$IncludeOneDrive
    )    
    $startTime = Get-Date
    
    #TODO: SharePoint and Security and Compliance Connectivity check.

    #2025-07-23: All logic to check if we run this and getting the module name moved to the Test-UcPowerShellModule.
    Test-UcPowerShellModule | Out-Null
    #endregion

    $outFile = "SharePointSitesWithHolds_" + (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"
    #Verify if the Output Path exists
    if ($OutputPath) {
        if (!(Test-Path $OutputPath -PathType Container)) {
            Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
            return
        }
        $OutputFilePath = [System.IO.Path]::Combine($OutputPath, $outFile)
    }
    else {                
        $OutputFilePath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFile)
    }

    if ($IncludeOneDrive) {
        Write-Warning "SharePoint Sites and OneDrives"
        $activityMsg = "Checking SharePoint Sites and OneDrives with Holds"
        $SPSites = Get-SPOSite -Limit all -IncludePersonalSite:$true
    }
    else {
        Write-Warning "Getting the SharePoint Sites"
        $activityMsg = "Checking SharePoint Sites with Holds"
        $SPSites = Get-SPOSite -Limit all 
    }
    $SitesProcessed = 1
    $SitesTotal = $SPSites.count
    $row = "SiteURL,HoldID,HoldCreatedDate,HoldCreatedBy,HoldEnabled,Type,IsAdaptivePolicy" + [Environment]::NewLine
    foreach ($SPSite in $SPSites) {
        try {
            $SiteHolds = Invoke-HoldRemovalAction -Action GetHolds -SharePointLocation $SPSite.Url -ErrorAction SilentlyContinue
            Write-Progress -Activity $activityMsg -Status ("Processing site " + $SPSite.Url + " - " + $SitesProcessed + " of " + $SitesTotal)
            foreach ($SiteHold in $SiteHolds) {
                $SiteCompliancePolicy = Get-RetentionCompliancePolicy -Identity $SiteHold
                $row += $SPSite.Url + "," + $SiteHold + "," + $SiteCompliancePolicy.WhenCreatedUTC + "," + $SiteCompliancePolicy.CreatedBy + "," + $SiteCompliancePolicy.Enabled + "," + $SiteCompliancePolicy.Type + "," + $SiteCompliancePolicy.IsAdaptivePolicy
                Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
                $row = ""
            }
            $SitesProcessed++
        }
        catch {
            write-warning ("Failed to get Holds for site: " + $SPSite.Url)
        }
    }

    $endTime = Get-Date
    $totalSeconds = [math]::round(($endTime - $startTime).TotalSeconds, 2)
    $totalTime = New-TimeSpan -Seconds $totalSeconds
    Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
    Write-Host "Execution time:" $totalTime.Hours "Hours" $totalTime.Minutes "Minutes" $totalTime.Seconds "Seconds" -ForegroundColor Green
}