function Export-UcOneDriveWithMultiplePermissions {
    <#
        .SYNOPSIS
        Generate a report with OneDrive's that have more than a user with access permissions.

        .DESCRIPTION
        This script will check all OneDrives and return the OneDrive that have additional users with permissions.

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

        .EXAMPLE 
        PS> Export-UcOneDriveWithMultiplePermissions

        .EXAMPLE 
        PS> Export-UcOneDriveWithMultiplePermissions -MultiGeo
    #>
    param(
        [string]$OutputPath,
        [switch]$MultiGeo
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
        $outFile = "OneDrivePermissions_MultiGeo_" + (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"
        $GraphPathSites = "/sites/getAllSites?`$select=id,displayName,isPersonalSite,WebUrl&`$top=999"
    }
    else {
        $outFile = "OneDrivePermissions_" + (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"
        $GraphPathSites = "/sites?`$select=id,displayName,isPersonalSite,WebUrl&`$top=999"
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
    $OneDriveProcessed = 0
    $OneDriveFound = 0 
    $BatchNumber = 1
    $row = "OneDriveDisplayName,OneDriveUrl,Role,UserWithAccessDisplayName,UserWithAccessUPN,UserWithAccessSharePointLogin,OneDriveID,PermissionID" + [Environment]::NewLine
    do {
        try {
            $ResponseSites = Invoke-UcGraphRequest  -Path $GraphPathSites 
            $GraphRequestSites = $ResponseSites.'@odata.nextLink'
            #Currently the SharePoint API doenst support filter for isPersonalSite, so we need to filter it 
            $tempOneDrives = $ResponseSites.value | Where-Object { $_.isPersonalSite -eq $true }
            #Adding a progress messsage to show status
            foreach ($OneDrive in $tempOneDrives) {
                if ($OneDriveProcessed % 10 -eq 0) {
                    Write-Progress -Activity "Looking for addtional users in OneDrive permissions" -Status "Batch #$BatchNumber - Number of OneDrives Processed $OneDriveProcessed"
                }
                $GPOneDrivePermission = "/sites/" + $OneDrive.id + "/drive/root/permissions"
                try {
                    $OneDrivePermissions = Invoke-UcGraphRequest -Path $GPOneDrivePermission
                    if ($OneDrivePermissions.count -gt 1) {
                        foreach ($OneDrivePermission in $OneDrivePermissions) {
                            if ($OneDrivePermission.grantedToV2.siteuser.displayName -ne $OneDrive.displayName) {
                                $tempUPN = Get-UcUPNFromString $OneDrivePermission.grantedToV2.siteuser.loginName
                                $row += $OneDrive.displayName + "," + $OneDrive.WebUrl + "," + $OneDrivePermission.roles + ",`"" + $OneDrivePermission.grantedToV2.siteuser.displayName + "`"," + $tempUPN + "," + $OneDrivePermission.grantedToV2.siteuser.loginName + ",`"" + $OneDrive.id + "`"," + $OneDrivePermission.id
                                Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
                                $row = ""
                                $OneDriveFound++
                            }
                        }
                    }
                    $OneDriveProcessed++
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
    Write-Host "Total of OneDrives processed:  $OneDriveProcessed, total OneDrives with additional users with permissions: $OneDriveFound" -ForegroundColor Cyan
    if ($OneDriveFound -gt 0) {
        Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
    }
    Write-Host "Execution time:" $totalTime.Hours "Hours" $totalTime.Minutes "Minutes" $totalTime.Seconds "Seconds" -ForegroundColor Green
}