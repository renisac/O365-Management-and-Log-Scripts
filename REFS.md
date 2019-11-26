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

### HAWK

Designed to ease the burden on O365 administrators who are performing forensic analysis in their M365 tenant. It accelerates the gathering of data from multiple sources in the service.

Installation: 

```powershell
Install-Module -Name Hawk
Import-Module -Name Hawk
```

Getting Started:

```powershell
Start-HawkTenantInvestigation
Start-HawkUserInvestigation -UserPrincipalName <UPN>
```

#### Sources

<https://www.powershellgallery.com/packages/HAWK>

https://github.com/Canthv0/hawk

---

### US-CERT CISA O365 Security Observations:

<https://www.us-cert.gov/ncas/analysis-reports/AR19-133A>

---

### Licensing Comparison Charts

<https://github.com/AaronDinnage/Licensing>
