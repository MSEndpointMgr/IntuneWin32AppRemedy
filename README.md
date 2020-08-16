# IntuneWin32AppRemedy v2 (28th June 2020)

An Azure automation runbook to automaticly patch Win32 Available apps in MSIntune and integration with Proactive Remediations for End User notifications and local detection. 

## Parameters required at Runtime 

- Old App Intune AppiD
- Old App Intune Name (Version)
- Updated App Intune APPiD
- Updated App Name (Version)

## Parameters required as Automation Account Variables
|Variable Name|Type|Value|Comment|
|---|---|---|---|
|ClientID|String (Encrypted)|AzureAD App ID|The ID of the Azure AD App Registration|
|ClientSecret|String (Encrypted)|App Secret|The Client Secret of the Azure AD App|
|TenantID|String (Encrypted)|Tenant ID|The Azure AD Tenant ID|
|TeamsReporting|Boolean|True/False|Turns Teams Reporting on or off|
|TeamsReportChannelWebHook|String|Company Name|Used as text in notification|
|NotificationAttributeText|String|URL|Used as text in notification|
|NotificationHeroImageUri|String|URL|URL to Image	Hero Image in the notification|
|NotificationLogoImageUri|String|URL|URL to Image	Logo image in the notification|
|Notifications|Boolean|True/False|Turns notification feature on or off |
|CleanOldAppVersion|Boolean|True/False|Remove old app from Intune|

## Implementation
For full instruction and explanations, go to https://msendpointmgr.com/2020/05/26/automated-3rdparty-patch-remediation-in-intune-with-azure-automation/
