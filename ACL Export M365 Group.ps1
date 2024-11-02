## DISCLAIMER:
## Copyright (c) Microsoft Corporation. All rights reserved. This
## script is made available to you without any express, implied or
## statutory warranty, not even the implied warranty of
## merchantability or fitness for a particular purpose, or the
## warranty of title or non-infringement. The entire risk of the
## use or the results from the use of this script remains with you.

<# Sample Config file below. starting with { and ending with }
{
  "AppId": "5a78992e-4cd0-4ad4-839f-b4a9c5e4f8c1",
  "Thumbprint": "6DF3D32DA943DDDC0C905A56C2A9C688A1C45C5C",
  "TenantName": "m365x27828082.onmicrosoft.com",
  "AdminURL": "https://m365x27828082-admin.sharepoint.com"
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
$site = "https://mngenvmcap277944.sharepoint.com/"
#Local variable to create and store output file
$filename = 'PermissionReportSelectedSites'+ $((get-date).ToString("dd-MM-yyyy (HH.mm.ss)"))+'.csv'
$fullpathFileReport = $OutPutFolderpath+$filename
$i = 1
Write-host "Total sites found:" $siteUrls.Count -ForegroundColor Cyan
#$results = foreach ($siteUrl in $siteUrls) {
  
        #Write-Host "Connecting to $i of" $siteUrls.Count "site:" $siteUrl
        Connect-PnPOnline -Url $site -ClientID $clientID -Thumbprint $thumbprint -Tenant $tenantName -ErrorAction Stop
        #Connect-PnPOnline  -ClientID $clientID -Thumbprint $thumbprint -Tenant $tenantName -ErrorAction Stop

        # Get the Microsoft 365 group of the site
        $groups = Get-PnPMicrosoft365Group -IncludeSiteUrl 
        
        #-Filter "startswith(displayName, 'PMO')" -ErrorAction Stop
        
        foreach ($group in $groups) {
            # Display group information
            Write-Host "Processing group:" $group.DisplayName $group.SiteUrl -ForegroundColor Yellow
            
            # Get and display owners
            $groupOwners = Get-PnPMicrosoft365GroupOwner -Identity $group.Id
            foreach ($groupOwner in $groupOwners) {
                Write-Host "Owner:" $groupOwner.DisplayName -ForegroundColor Magenta
            }
            
            # Get and display members
            $groupMembers = Get-PnPMicrosoft365GroupMembers -Identity $group.Id
            foreach ($groupMember in $groupMembers) {
                Write-Host "Member:" $groupMember.DisplayName -ForegroundColor Green
            }
        }

  ### Next tasks is to merge this to ACL Export copy ps1 to have M365 Group Owner and Members in the output file
   

#}




Write-Host "Verify the result in the output file @" $fullpathFileReport -BackgroundColor Cyan
Stop-Transcript