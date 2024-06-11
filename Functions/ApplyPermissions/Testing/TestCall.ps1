$ClientAppId = '0a8dfbc9-0423-495b-a1e6-1055f0ca69c2'
$TenantId = "7e77d071-ed08-468a-bc75-e8254ba77a21"
$Scopes = 'api://mmospfxsecsamplefunction.azurewebsites.net/0a8dfbc9-0423-495b-a1e6-1055f0ca69c2/user_impersonation'
$RedirectUri = "http://localhost"
Import-Module 'MSAL.PS' -ErrorAction 'Stop'
$PublicClient = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientAppId).WithRedirectUri($RedirectUri).Build()

$token = Get-MsalToken -PublicClientApplication $PublicClient -TenantId $TenantId -Scopes $Scopes
$token.AccessToken

$body = @{URL = "https://mmoellermvp.sharepoint.com/sites/SharingDemo"}
Invoke-RestMethod -Uri "http://localhost:7086/api/Function1" -Headers @{Authorization = "Bearer $($token.AccessToken)" } -body $body

$body = @{
					URL = "https://mmoellermvp.sharepoint.com"},
					Descreption = "New Site Desrestion!"
					}
Invoke-RestMethod -Uri "http://localhost:7086/api/Function1" -Headers @{Authorization = "Bearer $($token.AccessToken)" } -body $body -Method Post