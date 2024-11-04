### Pre-requisites:
### 1. Download and install PnP PowerShell from https://pnp.github.io/powershell/articles/installation.html
### 2. You need to have site collection admin permission for all of the sites to run this script.
### Owner: Lambert Qin
### Use this script to export site collection owners and members information.



$tenant = "XXXXX"
$tenantId = "XXXX"
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
        $sites = Import-Csv $SiteReport
        Write-Host -ForegroundColor Green "Retrieve site collections... "
        Write-Host "--Found $($sites.Count) site collections."
        foreach ($site in $sites) {
            If ("Exported" -ne $site.Status) {
                $siteUrl = $site.SiteUrl
                Write-Host -ForegroundColor Green "Process site collecion: $siteUrl"

                $updatedSite = ProcessSite -SiteObject $site

                $results += New-Object PSObject -Property $updatedSite 
            }
            else {
                $results += New-Object PSObject -Property $site
            }
        }

        $results | Export-Csv $SiteReport -NoTypeInformation
        Write-Host -ForegroundColor Green "Site Collection Data Exported to CSV!"
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
}

Function ProcessSite {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        $SiteObject
    )

    $siteInfo = [ordered] @{  
        Status                      = "Not Started"          
        SiteUrl                     = $SiteObject.SiteUrl
        TemplateName                = $SiteObject.TemplateName
        CreatedByEmail              = $SiteObject.CreatedByEmail
        Created                     = $SiteObject.Created
        Modified                    = $SiteObject.Modified
        LastActivityOn              = $SiteObject.LastActivityOn
        StorageUsed                 = $SiteObject.StorageUsed
        NumOfFiles                  = $SiteObject.NumOfFiles
        PageViews                   = $SiteObject.PageViews
        PagesVisited                = $SiteObject.PagesVisited
        Everyone                    = "False"
        EveryoneExceptExternalUsers = "False"
        Error                       = ""
    } 
    
    Try {
        Connect-PnPOnline -Url $SiteObject.SiteUrl -Tenant $tenantUrl -ClientId $appClientId -Thumbprint $appThumbprint
        #"Everyone" Group's Login ID
        $EveryoneLoginName = "c:0(.s|true"
        $Everyone = Get-PnPUser -Identity $EveryoneLoginName
        if ($Everyone -ne $null) {
            $siteInfo.Everyone = "True"
        }
        #Get "Everyone Except External Users" Login ID
        $EEEULoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenantId"
        $EEEU = Get-PnPUser -Identity $EEEULoginName
        if ($EEEU -ne $null) {
            $siteInfo.EveryoneExceptExternalUsers = "True"
        }
        $siteInfo.Status = "Exported"

        return $siteInfo
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
        $siteInfo.Status = "Error"
        $siteinfo.Error = $_.Exception.Message
        return $siteInfo
    }
}

Main