### Pre-requisites:
### 1. Download and install PnP PowerShell from https://pnp.github.io/powershell/articles/installation.html
### 2. You need to have tenant admin permission to run this script.
### Owner: Lambert Qin
### Use this script to export all site collections in a tenant.


$tenant = "XXXXX"
$tenantUrl = "$tenant.onmicrosoft.com"
$tenantAdminUrl = "https://XXXXX-admin.sharepoint.com"
$appClientId = "XXXXX"
$appThumbprint = "XXXXX"

$SiteReport = "C:\Users\smithm\Documents\SiteReport-$tenant.csv"

Function Main {
    ProcessTenant -TenantAdminUrl $tenantAdminUrl
}
Function ProcessTenant {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [String]
        $TenantAdminUrl
    )
    Try {
        Write-Host -f Green "Process tenant: $tenant"
        $results = @()
        Connect-PnPOnline -Url $TenantAdminUrl -Tenant $tenantUrl -ClientId $appClientId -Thumbprint $appThumbprint
        Write-Host -ForegroundColor Green "Retrieve site collections... "
        # customize the query to get the site collections you want, refer to https://pnp.github.io/powershell/cmdlets/Get-PnPTenantSite.html
        # $sites = Get-PnPTenantSite
        $sites = Get-PnPListItem -List DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS -Query "<View><Query></Query></View>" -PageSize 2000
        
        Write-Host "--Found $($sites.Count) site collections."

        foreach ($site in $sites) {
            Write-Host "--$($site.FieldValues.SiteUrl)"
            $siteInfo = [ordered] @{  
                Status         = "Not Started"          
                SiteUrl        = $site.FieldValues.SiteUrl
                TemplateName   = $site.FieldValues.TemplateName
                CreatedByEmail = $site.FieldValues.CreatedByEmail
                Created        = $site.FieldValues.Created
                Modified       = $site.FieldValues.Modified
                LastActivityOn = $site.FieldValues.LastActivityOn
                StorageUsed    = $site.FieldValues.StorageUsed
                NumOfFiles     = $site.FieldValues.NumOfFiles
                PageViews      = $site.FieldValues.PageViews
                PagesVisited   = $site.FieldValues.PagesVisited
                Error          = ""
            } 
            $results += New-Object PSObject -Property $siteInfo
        }

        $results | Export-Csv $SiteReport -NoTypeInformation
        Write-Host -ForegroundColor Green "Site Collection Data Exported to CSV!"
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
}

Main