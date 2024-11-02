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

# Check if the config file exists
if (-not (Test-Path $configFilePath)) {
    Write-Host "Config file not found: $configFilePath" -ForegroundColor Red
    return
}

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
try {
        Write-host "Connecting to Admin URL" -ForegroundColor Green
        Connect-PnPOnline -Url $AdminURL -ClientID $clientID -Thumbprint $thumbprint -Tenant $tenantName -ErrorAction SilentlyContinue
        Get-PnPTenant | Select StorageQuota
        write-host "If you see the storage quota value, the connection is successfully estabilished. If you see blank value, then the connection is established successfully, however the app doesn't have enough permission" -ForegroundColor Yellow
        }
        Catch
        {
         Write-Host "Error occured while connecting to Admin URL. Connection Test Failed. Refer below error message" -ForegroundColor Red
         Write-host "Error message:" $_.Exception.Message
         return
        }
