function Export-UcSharePointSitesWithEmptyFiles {
    <#
        .SYNOPSIS
        Generate a report with OneDrive's that have more than a user with access permissions.

        .DESCRIPTION
        This script will check all SharePoint Sites and OneDrives looking for empty files (size = 0), by default will return PDF, but queries can be used.

        Author: David Paulino

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                            "Sites.Read.All"
                        Note: Currently the SharePoint Sites requires to authenticate to Graph API with AppOnly https://learn.microsoft.com/graph/auth/auth-concepts
        
        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results. By default, it will save on current user Download.

        .PARAMETER MultiGeo
        Required if Tenant is MultiGeo

        .PARAMETER IncludeOneDrive
        If OneDrives should also be included in the report.

        .EXAMPLE 
        PS> Export-UcSharePointSitesWithEmptyFiles

        .EXAMPLE 
        PS> Export-UcSharePointSitesWithEmptyFiles -MultiGeo

        .EXAMPLE 
        PS> Export-UcSharePointSitesWithEmptyFiles -IncludeOneDrive
    #>
    param(
        [string]$Query = "PDF",
        [string]$OutputPath,
        [switch]$MultiGeo,
        [switch]$IncludeOneDrive
    )    

    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "Sites.Read.All" -AltScopes ("Sites.ReadWrite.All") -AuthType "Application")) {
        return
    }
    Test-UcPowerShellModule | Out-Null
    #endregion

    $startTime = Get-Date
    #Graph API request is different when the tenant has multigeo
    if ($MultiGeo) {
        $outFile = "SharePointEmptyFiles_MultiGeo_" + (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"
        $GraphRequestSites = "/sites/getAllSites?`$select=id,displayName,isPersonalSite,WebUrl&`$top=999"
    }
    else {
        $outFile = "SharePointEmptyFiles_" + (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"
        $GraphRequestSites = "/sites?`$select=id,displayName,isPersonalSite,WebUrl&`$top=999"
    }
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
    $SharePointSitesProcessed = 0
    $EmptyFilesFound = 0 
    $BatchNumber = 1
    $row = "SiteDisplayName,SiteUrl,FileName,Size,createDate,createdByDisplayName,createdByemail,lastModifiedDate,lastModifiedByDisplayName,lastModifiedByEmail" + [Environment]::NewLine
    do {
        try {
            $ResponseSites = Invoke-UcGraphRequest -Path $GraphRequestSites
            $GraphRequestSites = $ResponseSites.'@odata.nextLink'
            if ($IncludeOneDrive) {
                $tempSites = $ResponseSites.value
            }
            else {
                $tempSites = $ResponseSites.value | Where-Object { $_.isPersonalSite -eq $false }
            }
            #Adding a progress messsage to show status
            foreach ($Site in $tempSites) {
                if ($SharePointSitesProcessed % 10 -eq 0) {
                    Write-Progress -Activity "For empty files" -Status "Batch #$BatchNumber - Number of Sites Processed $SharePointSitesProcessed"
                }
                $GRSharePointDrive = "/sites/" + $Site.id + "/drive/root/search(q='$Query')"
                try {
                    $SPFiles = (Invoke-UcGraphRequest -Path $GRSharePointDrive)
                    if ($SPFiles.value.count -ge 1) {
                        foreach ($SPFile in $SPFiles.value) {
                            if ($SPFile.size -eq 0) {
                                $row += $Site.displayName + "," + $Site.WebUrl + "," + $SPFile.name + "," + $SPFile.size + "," + $SPFile.createdDateTime + "," + $SPFile.createdBy.user.displayName + "," + $SPFile.createdBy.user.email + "," + $SPFile.lastModifiedDateTime + "," + $SPFile.lastModifiedBy.user.displayName + "," + $SPFile.lastModifiedBy.user.email
                                Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
                                $row = ""
                                $EmptyFilesFound++
                            }
                        }
                    }
                    $SharePointSitesProcessed++
                }
                catch { 
                }
            }
            $BatchNumber++
        }
        catch { break }
    } while (![string]::IsNullOrEmpty($GraphRequestSites))
    $endTime = Get-Date
    $totalSeconds = [math]::round(($endTime - $startTime).TotalSeconds, 2)
    $totalTime = New-TimeSpan -Seconds $totalSeconds
    Write-Host "Total of Sites processed: $SharePointSitesProcessed, total empty files: $EmptyFilesFound" -ForegroundColor Cyan
    if ($EmptyFilesFound -gt 0) {
        Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
    }
    Write-Host "Execution time:" $totalTime.Hours "Hours" $totalTime.Minutes "Minutes" $totalTime.Seconds "Seconds" -ForegroundColor Green
}