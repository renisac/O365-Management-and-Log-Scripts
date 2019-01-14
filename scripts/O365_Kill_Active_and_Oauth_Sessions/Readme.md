## kill_sessions.ps1

### Use case

Kill active O365 session and Oauth session refresh tokens

### Brief description 

Pointer to Microsoft documentation about killing sessions

### Data sensitivity

None.

### Local environment requirements

Powershell.

### License and/or permissions required

None.

### How to run the script

This will kill the refresh tokens: Revoke-AzureADUserAllRefreshToken

https://docs.microsoft.com/en-us/powershell/module/azuread/revoke-azureaduserallrefreshtoken?view=azureadps-2.0

This will kill active O365 sessions: Revoke-SPOUserSession

https://technet.microsoft.com/en-us/library/mt637161.aspx

### Known issues

None.

### Acknowledgements and Credits

Microsoft
