# Security-Focused O365 Management and Log Resources

Contributions are welcome, whether new code, documentation, or improvements to existing work. Refer to [CONTRIBUTING.md](CONTRIBUTING.md) for details.

### Table of Contents

1. [Objectives](#objectives)
2. [Common EDU License Types](#common-edu-license-types)
3. [Logs Within Office 365](#logs-within-office-365)
4. [Exporting Logs](#exporting-logs)
5. [Resources for Log Processing](#resources-for-log-processing)
6. [Log Examples](#log-examples)
7. [References](#references)

## Objectives

- Describe what O365 logs are available per common EDU license types
- Describe how and with what tools the logs can be accessed
- Describe how the logs may be exported to external tools (e.g. SIEM)
- Provide resources (code/scripts, documentation) that retrieve log information for specific investigative actions
- Nurture community-of-practice contributions to sustain and increase the value of this open resource

## Common EDU License Types

Please note that MS currently (Fall 2018) identifies five (5) license tiers or SKUs:

1. Office 365 Education A1
2. Office 365 Education A3
3. Office 365 Education A5
4. Microsoft 365 Education A3 (O365 A3 + EMS A3 + Windows 10)
5. Microsoft 365 Education A5 (O365 A5 + EMS A5 + Windows 10)

In the fall of 2018, Microsoft updated their names and definitions of paid licensees for the M365 Academic plans. Knowledge Workers are now called Education Qualified Users (EQUs). Student worker employees no longer require a paid license. More details can be seen in this [EES-2018-Changes.pdf](media/EES-2018-Changes.pdf).

### A1: Office 365 Education

A1 is the bottom tier, free licenses for Microsoft O365. No additional security tools are available in this tier.

**Note:** Alumni and other "loosely affiliated" populations are typically not assigned A3/A5 licenses (depending on the institution), so it's safe to assume you'll have users at this tier with fewer security tools.

### A3: Microsoft 365 Education

A3 is the middle tier license for Microsoft O365. A3 includes EMS A3 which adds the following security tools:

- **Azure AD Premium P1/Azure AD Plan 1**: Secure single sign-on to cloud and on-premises app MFA, conditional access, and advanced security reporting

- **Azure Information Protection P1**: Encryption for all files and storage locations; Cloud-based file tracking

- **Microsoft Advanced Threat Analytics**: Protection from advanced targeted attacks leveraging user and entity behavioral analytics

- **Microsoft Intune**: Mobile device and app management to protect corporate apps and data on any device

### A5: Microsoft 365 Education

A5 is the top tier license for Microsoft O365. A5 includes EMS A3 (Intune and Advanced Threat Analytics) + EMS A5 which adds the following security tools:

- **Access Reviews**:
Recertify, attest, and audit users' access to resources to ensure access is appropriate and reviewed on a continual basis

- **Azure AD Premium P2/Azure AD Plan 2 (includes P1 features)**: Identity and access management with advanced protection for users and privileged identities

- **Azure Information Protection P2 (includes P1 features)**: Intelligent classification and encryption for files shared inside and outside your organization

- **Azure Identity Protection**:
Automated responses to detected suspicious actions related to user identities

- **Microsoft Cloud App Security**: Enterprise-grade visibility, control, and protection for your cloud applications

- **Advanced Threat Protection**: Safe Links, Safe Attachments, Phish (spoof) policies

- **Privileged Identity Management**:
Manage, control, and monitor access to roles across AAD, O365, Intune, etc.

- **Threat Intelligence**: Dashboard of attacks detected against your tenant with easy remediation steps.

### Ad-hoc

Organizations can buy licenses ad-hoc through either their reseller or the Microsoft O365 Admin Portal. Please note that while some features can be enabled for the whole tenant by purchasing only 1 license, this can lead to your tenant not being properly licensed or service degradation. You should only enable the features for users for whom you purchased licenses.

### A3/5 comparison chart

Aaron Dinnage, a Microsoft employee helping government sectors, routinely provides updated graphics for plan comparisons. His GitHub repo is linked below.
![Microsoft EMS A3/A5 Comparison Graphic](media/Microsoft_365_Education_on_a_Page(cropped_resized).png)

![Microsoft Office 365 A3/A5 Comparison Graphic](media/Office_365_Education_Licensing.png)

For the latest version of this, and other, charts released by Aaron, please see his [Microsoft Licensing GitHub repo](https://github.com/AaronDinnage/Licensing).

## Logs Within Office 365

### Types of logs

- Email Inbound/Outbound

- Safe Links clicks

- DLP Logs

- Azure AD Logs

- Exchange Online Actions

- OneDrive/SPO

**Note**: The Activities API (AKA: Magic Unicorn Tool) is no longer available.<sup>[3](#footnote3)</sup>

### Log sources

- OCAS - O365 Cloud App Security (formerly known as Advanced Security Management-ASM)

- MCAS - MS Cloud App Security

- [Management API](https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference)

- Powershell
  - [Search-UnifiedAuditLog](https://docs.microsoft.com/en-us/powershell/module/exchange/policy-and-compliance-audit/search-unifiedauditlog?view=exchange-ps)
  - [Get-MessageTrace](https://docs.microsoft.com/en-us/powershell/module/exchange/mail-flow/get-messagetrace?view=exchange-ps)

- Admin Portal
  - [Audit log search](https://docs.microsoft.com/en-us/office365/securitycompliance/search-the-audit-log-in-security-and-compliance)
  - [Message Trace search](https://docs.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/run-a-message-trace-and-view-results)
  - [Safe Links Reports](https://support.office.com/en-us/article/view-reports-for-office-365-advanced-threat-protection-e47e838c-d99e-4c0b-b9aa-e66c4fae902f)
  - Azure AD
    - [Sign-in Logs](https://docs.microsoft.com/en-us/azure/active-directory/reports-monitoring/concept-sign-ins)
    - [Audit Logs](https://docs.microsoft.com/en-us/azure/active-directory/reports-monitoring/concept-audit-logs)

### Log Collection Methods available by SKU

| License | OCAS | MCAS| Management API | PowerShell | Admin Portal |
|:---:|:---:|:---:|:---:|:---:|:---:|
|**A1**   | no| no| **Yes** | **Yes** | **Yes** |
|**A3**   | **Yes**<sup>[1](#footnote1)</sup> | no| **Yes** | **Yes** | **Yes** |
|**A5**   | **Yes**<sup>[1](#footnote1)</sup> | **Yes**<sup>[1](#footnote1)</sup> | **Yes** | **Yes** | **Yes** |
|**Ad-hoc**| **Yes**<sup>[1](#footnote1)</sup> | **Yes**<sup>[1](#footnote1)</sup> | **Yes** | **Yes** | **Yes** |

### Logs Types available by Collection Method

| Log Type | OCAS | MCAS| Management API | PowerShell | Admin Portal |
|:---:|:---:|:---:|:---:|:---:|:---:|
| Inbound/Outbound mail | no | no | no| **Yes**| **Yes**|
| SafeLinks clicks      | ?  | ? |  ?| **Yes**| **Yes**|
| DLP Logs              | no | **Yes**  | **Yes** | **Yes**| **Yes** |
| Azure AD Logs         | **Yes** | **Yes** | **Yes** | **Yes**| **Yes** |
| Exchange Online Actions| **Yes**  | **Yes**  | **Yes**| **Yes**| no|
| OneDrive/SPO Logs     | **Yes**  | **Yes**  | **Yes** | **Yes**| **Yes** |
| Export predefined alerts to SIEM | **Yes** | ?  | ? | ?| ? |
| Export raw data to SIEM        | no | **Yes**  | **Possible**<sup>[2](#footnote2)</sup> | **Possible**<sup>[2](#footnote2)</sup>| no |

## Exporting Logs

### OCAS

#### SIEM Connector

OCAS supports exporting alerts to your SIEM by using the [SIEM Connector](https://docs.microsoft.com/en-us/cloud-app-security/siem) provided by Microsoft. The connector supports exporting as Generic CEF and Micro Focus ArcSight. Both Splunk and ELK appear to be able to import this format; ref: Splunk [Deploy and Use Splunk App for CEF](https://docs.splunk.com/Documentation/CEFapp/2.0.1/DeployCEFapp/Howtheappworks) and ELK: [Cef codec plugin](https://www.elastic.co/guide/en/logstash/current/plugins-codecs-cef.html)

### MCAS

#### SIEM Connector

MCAS logs and alerts can be exported using the [SIEM Connector](https://docs.microsoft.com/en-us/cloud-app-security/siem) provided by Microsoft. The connector supports exporting as Generic CEF and Micro Focus ArcSight. Both Splunk and ELK appear to be able to import this format; ref: Splunk [Deploy and Use Splunk App for CEF](https://docs.splunk.com/Documentation/CEFapp/2.0.1/DeployCEFapp/Howtheappworks) and ELK: [Cef codec plugin](https://www.elastic.co/guide/en/logstash/current/plugins-codecs-cef.html)

#### PowerShell
You can write custom scripts to query MCAS via PowerShell. Microsoft has provided a [PowerShell module](https://github.com/Microsoft/MCAS) to help with querying this data.

### Management API

#### PowerShell

**Note**: that this method can be slow for large tenants. There is a POC that speeds this data collection up and will be shared

Using the Management API, you can write custom scripts to export the data. Microsoft employee [Brandon Koeller](https://github.com/bkoeller) has written a [PowerShell script](https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/O365InvestigationDataAcquisition.ps1) that allows you to export this data to MySQL, Azure Blob Storage, Azure SQL, JSON or CSV. 

#### Splunk Addon

Splunk provides an add on that will ingest the data for your tenant (from the Management API). The add on is available from the the [Splunk base](https://splunkbase.splunk.com/app/4055/) site. Depending on usage, user counts, and audited events, you can expect ~1 GB/day per 10,000 users.

### Message Trace logs

Microsoft 365 provides admins with access to Message trace logs either programatically (over PowerShell) or via a GUI (Security and Compliance Center). 

For help on using the GUI, see [Message trace in the Office 365 Security & Compliance Center
](https://support.office.com/en-us/article/message-trace-in-the-office-365-security-compliance-center-3e64f99d-ac33-4aba-91c5-9cb4ca476803). 

For pull these logs automatically over PowerShell see [Gather Message Trace Logs Script sample](source/MessageTraceLogGatherer)

## Resources for Log Processing

A number of community-developed resources are provided, supporting log query and retrieval, and management actions for specific investigative purposes. View the [source](source) directory. Please consider [contributing your helpful code](CONTRIBUTING.md).

## Log Examples

Various [log examples](example_log_data) are provided. Please consider contributing.

## References

Various [references](REFS.md) are provided to helpful commands, PowerShell modules, and documentation related to O365 Security & Log management. Please consider contributing.

---

### Footnotes

<a name="footnote1">1</a>: Requires a license for each user that is being logged. OCAS/MCAS supports scoped deployments now based on Azure AD group.

<a name="footnote2">2</a>: Local scripting & infrastructure development, operations, and maintenance required.

<a name="footnote3">3</a>: Official statement (July 2018) from Microsoft about the Activities API (AKA: Magic Unicorn Tool):
> Microsoft has always discouraged the use of undocumented APIs, as it can lead to incorrect assumptions, poor user experience or broken functionality. In this case, the particular Outlook Mail REST API referenced and underlying telemetry was built to support service to service communication and does not guarantee the type of complete and accurate activities data that would be necessary to support security investigation scenarios. Following feedback, we are disabling this API.
