##########################################################################################################
# Disclaimer
# This sample code, scripts, and other resources are not supported under any Microsoft standard support 
# program or service and are meant for illustrative purposes only.
#
# The sample code, scripts, and resources are provided AS IS without warranty of any kind. Microsoft 
# further disclaims all implied warranties including, without limitation, any implied warranties of 
# merchantability or of fitness for a particular purpose. The entire risk arising out of the use or 
# performance of this material and documentation remains with you. In no event shall Microsoft, its 
# authors, or anyone else involved in the creation, production, or delivery of the sample be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of the 
# use of or inability to use the samples or documentation, even if Microsoft has been advised of 
# the possibility of such damages.
##########################################################################################################

# import the EWS Managed API
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

# read environment settings from local file
$appSettings = Get-Content -Raw -Path .\m365x612691.config.user | ConvertFrom-Json

# authenticate using client credentials OAuth flow
$auth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($appSettings.TenantId)/oauth2/v2.0/token" `
    -Headers @{
        "Content-Type" = "application/x-www-form-urlencoded";
    } `
    -Body @{
        "grant_type" = "client_credentials";
        "client_id" = $appSettings.AppId;
        "client_secret" = $appSettings.AppSecret;
        "scope" = "https://outlook.office365.com/.default";
    }

# create EWS client
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$service.HttpHeaders.Add("Authorization","Bearer $($auth.access_token)")
$service.Url = "https://outlook.office365.com/ews/exchange.asmx"
$service.ImpersonatedUserId = new-object Microsoft.exchange.webservices.data.impersonateduserid([Microsoft.Exchange.WebServices.data.connectingidtype]::SmtpAddress, $appSettings.impersonatedUser)

# get all folder properties
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)
$folders | select *ID, Displayname |fl

# get folder id only
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$folderView.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::IdOnly
$folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)
$folders | select *ID, Displayname |fl

# deep folder inspection
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
$folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)
$folders | select *ID, Displayname |fl

# get folder id, displayname and parentfolderid properties only
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
$propertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
$propertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)
$folderView.PropertySet = $propertySet
$folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)
$folders | select *ID, Displayname |fl

# search folder by name
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,“Projects”) 
$folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $searchFilter, $folderView)
$folders | select *ID, Displayname |fl
