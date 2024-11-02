## DISCLAIMER:
## Copyright (c) Microsoft Corporation. All rights reserved. This
## script is made available to you without any express, implied or
## statutory warranty, not even the implied warranty of
## merchantability or fitness for a particular purpose, or the
## warranty of title or non-infringement. The entire risk of the
## use or the results from the use of this script remains with you.

<# Sample Config file below. starting with { and ending with }
{
  "AppId": "xxxx-yyyy-zzzz",
  "Thumbprint": "xxxx-yyyy-zzzz",
  "TenantName": "xxxx.onmicrosoft.com",
  "AdminURL": "https://xxxx-admin.sharepoint.com"
}
#>
# Read the configuration from the config.txt file
$configFilePath = "config.txt"
$InputFilePath = "Sites.txt"
$OutPutFolderpath= "C:\Users\smithm\Documents\"
$logfile = 'PermissionReportSelectedSites_Log'+ $((get-date).ToString("dd-MM-yyyy"))+'.txt'
$fullLogFilepath = $OutPutFolderpath+$logfile

Start-Transcript -Path $fullLogFilepath -Append

# Check if the config file exists
if (-not (Test-Path $configFilePath)) {
    Write-Host "Config file not found: $configFilePath" -ForegroundColor Red
    return
}
# Check if the Input file exists
if (-not (Test-Path $InputFilePath)) {
    Write-Host "Input file not found: $InputFilePath" -ForegroundColor Red
    return
}

# Import the CSV file into a variable
$siteUrls = Get-Content -Path $InputFilePath

# Read the configuration from the config file
$config = Get-Content -Raw -Path $configFilePath | ConvertFrom-Json
# Validate the required values
if (-not $config.AppId) {
    Write-Host "Invalid or missing AppId in the config file" -ForegroundColor Red
    return
}

if (-not $config.Thumbprint) {
    Write-Host "Invalid or missing Thumbprint in the config file" -ForegroundColor Red
    return
}

if (-not $config.TenantName) {
    Write-Host "Invalid or missing TenantName in the config file" -ForegroundColor Red
    return
}

if (-not $config.AdminURL) {
    Write-Host "Invalid or missing AdminURL in the config file" -ForegroundColor Red
    return
}

$clientID = $config.AppId
$thumbprint = $config.Thumbprint
$tenantName = $config.TenantName
$AdminURL = $config.AdminURL
#Local variable to create and store output file
$filename = 'PermissionReportSelectedSites'+ $((get-date).ToString("dd-MM-yyyy (HH.mm.ss)"))+'.csv'
$fullpathFileReport = $OutPutFolderpath+$filename
$i = 1
Write-host "Total sites found:" $siteUrls.Count -ForegroundColor Cyan
$results = foreach ($siteUrl in $siteUrls) {
    try {
        Write-Host "Connecting to $i of" $siteUrls.Count "site:" $siteUrl
        Connect-PnPOnline -Url $siteUrl -ClientID $clientID -Thumbprint $thumbprint -Tenant $tenantName -ErrorAction Stop
        $permission = Get-PnPUser -WithRightsAssigned -ErrorAction Stop
        #$permission = Get-PnPUser -ErrorAction Stop
        $result = New-Object System.Collections.ArrayList
        foreach ($perm in $permission)
        {
            $obj = New-Object PSObject -Property @{
                SiteUrl = $siteUrl
                Title = $perm.Title
                LoginName = $perm.LoginName
                Email = $perm.Email
                PrincipalType = $perm.PrincipalType
                Error = $null
            }
            $result.Add($obj) | Out-Null
        }
    }
    catch {
        Write-Host "Error occurred on" $siteUrl
        $obj = New-Object PSObject -Property @{
            SiteUrl = $siteUrl
            Error = $_.Exception.Message
        }
        $result.Add($obj) | Out-Null
    }
    finally {
        Disconnect-PnPOnline
    }

    $result
    $i++
}

# Define the order of columns
$orderedColumns = @("SiteUrl", "Title", "LoginName", "Email", "PrincipalType", "Error")

# Export the sorted results to CSV with specified column order
$results | Select-Object $orderedColumns | Export-Csv -Path $fullpathFileReport -NoTypeInformation

Write-Host "Verify the result in the output file @" $fullpathFileReport -BackgroundColor Cyan
Stop-Transcript
