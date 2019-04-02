## References / Links

### Invalidate O365 Sessions

```powershell
Revoke-AzureADUserAllRefreshToken -ObjectId <UserGuid>
```

* Invalidates the refresh tokens issued to applications for a user

```powershell
Revoke-SPOUserSession -User <UserSpn>
```

* Invalidates a user's O365 sessions across all their devices

#### Use case

Invalidates active O365 session and OAuth session refresh tokens

#### Requirements

[Azure AD Module](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2)

[SharePoint Online Module](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)

#### Sources

<https://docs.microsoft.com/en-us/powershell/module/azuread/revoke-azureaduserallrefreshtoken>

<https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/Revoke-SPOUserSession>

---