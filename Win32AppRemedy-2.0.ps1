<#
    .SYNOPSIS
        Win32AppRemedy - An Azure Automation Runbook - requires the following Automation Variables defined: 
        ClientID - The AppID of your service principal Azure AD App
        ClientSecret - The ClientSecret of your service principal Azure AD App
        TenantID - The AzureAD TenantID of your tenant
        TeamsWebHook - The Webhook url for your Teams Reporting Channel 
        TeamsReporting - (Boolean) True or False to turn Teams Reporting on off
        NotificationAttributeText - for information in the Toast Notification
        NotificationHeroImageURI - url to the Hero image in the Toast Notification
        NotificationLogoImageURI - url to the Logo image in the Toast Notification
        Notifications - (Boolean) True or False to turn Notifications feature on or off
        CleanOldAppVersion - (Boolean) True or False to allow runbook to delete old appversion or not. 
    .DESCRIPTION
        Win32AppRemedy is used to remediate updates for Win32Apps that are assigned as available by using Device Based patching groups. 
        These groups are dynamicly created based on devices that have the previous version installed, app is assigned as required to these groups with a deadline for install. 
        Teams reporting will publish a message card to a chosen teams channel.
    .PARAMETER OrgAppID
        Input parameter Intune AppID for the current application version
    .PARAMETER UpdatedAppID
        Input parameter Intune AppID for the updated application version
    .PARAMETER OrgAppVersion            
        Input parameter - Name of the current application in Intune, Example "Adobe Acrobat Reader DC 20.009.20042"
    .PARAMETER OrgAppVersion            
        Input parameter - Name of the updated application in Intune , Example "Adobe Acrobat Reader DC 20.009.20063"
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-05-20) Initial Version with app remediation groups
        1.0.1 - (2020-06-11) Buxfix for no devices needing updates for Teams Reporting and Restartbehaviortype for App assignment. 
        2.0.0 - (2020-06-28) Added support for deleting old appversion after assignment of new app, added support for Toast notifications with integration with Proactive remediations. 
    #>    
    param (
        [parameter(Mandatory = $true)]
        $OrgAppID,
        [parameter(Mandatory = $true)]
        $UpdatedAppID,
        [parameter(Mandatory = $true)]
        $OrgAppVersion,
        [parameter(Mandatory = $true)]
        $UpdatedAppVersion
)
#region Initialize
#Fetching all variables from the Automation Account (do not change this)
$ClientID = Get-AutomationVariable -Name 'ClientID'
$ClientSecret = Get-AutomationVariable -Name 'ClientSecret'
$TenantID = Get-AutomationVariable -Name 'TenantID'
$TeamsReporting = Get-AutomationVariable -Name 'TeamsReporting'
$AttributionText = Get-AutomationVariable -Name 'NotificationAttributeText'
$HeroImageUri = Get-AutomationVariable -Name 'NotificationHeroImageUri'
$LogoImageUri = Get-AutomationVariable -Name 'NotificationLogoImageUri'
#Custom Settings with defaults set (can be modified)
$EAPublisher = "MSEndpointMgr" #The Publisher name of the proactive remediation package
$Win32AppStartTime = 2 #The number of days before enforced app remediation starts 
$Win32AppDeadLine = 3 #The number of days before deadline for the enforced app remediation (must be higher than or equal to Win32AppStartTime)
$Scheduledtime = "9:0:0" #The time of day for the proactive remediation schedule
$Scheduleinterval = "1"  #The interval for the proactive remeadiation schedule (1 = each day, 7 = once/week)
Write-Output "Declarations done"
#endregion Initialize
#region Functions
function Get-MSGraphAppToken{
    <#
    .SYNOPSIS
        Get app based authentication token for MS Graph and return Header for authentication.
    .DESCRIPTION
        Get app based authentication token for MS Graph and return Header for authentication
    .PARAMETER TenantID
        Azure AD Tenant ID for authentication
    .PARAMETER ClientID
        Azure AD App Client ID 
    .PARAMETER ClientSecret            
        Client Secret for Azure AD Authentication 
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-03-29) Function created
    #>    
[CmdletBinding()]
	param (
		[parameter(Mandatory = $true, HelpMessage = "Your Azure AD Directory ID should be provided")]
		[ValidateNotNullOrEmpty()]
		[string]$TenantID,
		[parameter(Mandatory = $true, HelpMessage = "Application ID for an Azure AD application")]
		[ValidateNotNullOrEmpty()]
		[string]$ClientID,
		[parameter(Mandatory = $true, HelpMessage = "Azure AD Application Client Secret.")]
		[ValidateNotNullOrEmpty()]
		[string]$ClientSecret
	    )
Process {
    $ErrorActionPreference = "Stop"
       
    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
        }
    
    try {
        $MyTokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
        $MyToken =($MyTokenRequest.Content | ConvertFrom-Json).access_token
            If(!$MyToken){
                Write-Warning "Failed to get Graph API access token!"
                Exit 1
            }
        $MyHeader = @{"Authorization" = "Bearer $MyToken" }

       }
    catch [System.Exception] {
        Write-Warning "Failed to get Access Token, Error message: $($_.Exception.Message)"; break
    }
    return $MyHeader
    }
}#end function 
function Invoke-TeamsMessage{
    <#
    .SYNOPSIS
        Send a message card to defined teams channel webhook 
    .DESCRIPTION
        Send a message card to defined teams channel webhook 
    .PARAMETER teamswebhook
        Teams Channel Webhook
    .PARAMETER Text
        Text for message
    .PARAMETER UpdateStatus            
        Text for UpdateStatus 
    .PARAMETER StatusColor            
        Colorcode for message 
    .PARAMETER DeviceName         
        String with devices needing a update
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-03-29) Function created
    #>   
    param (
    [Parameter(Mandatory=$true)]
    $teamswebhook,
    [Parameter(Mandatory=$true)]
    $text,
    [Parameter(Mandatory=$true)]
    $UpdateStatus,
    [Parameter(Mandatory=$true)]
    $StatusColor,
    [Parameter(Mandatory=$true)]
    $DeviceName
    )
    $Now = Get-Date
        $payload = @"
{
    "@type": "MessageCard",
    "@context": "https://schema.org/extensions",
    "summary": "PMPC App Update Status",
    "themeColor": "$StatusColor",
    "title": "$UpdateStatus",
    "sections": [
     {
            "activityTitle": "Patch My PC Status",
            "activitySubtitle": "$Now",
            "facts": [
                {
                    "name": "Devices:",
                    "value": "$DeviceName"
                }
                    ],
                "text": "$text"
            }
    ]
}
"@

Invoke-RestMethod -uri $teamswebhook -Method Post -body $payload -ContentType 'application/json' | Out-Null      
}#endfunction
function Get-IntuneWin32AppAssignment {
    <#
    .SYNOPSIS
        Retrieve all assignments for a Win32 app.
    .DESCRIPTION
        Retrieve all assignments for a Win32 app.
    .PARAMETER ID
        Specify the ID for a Win32 application.
    .PARAMETER ApplicationID
        Specify the Application ID of the app registration in Azure AD. By default, the script will attempt to use well known Microsoft Intune PowerShell app registration.
    .NOTES
        ModifiedBy: Jan Ketil Skanke @JankeSkanke
        Author:      Nickolaj Andersen
        Contact:     @NickolajA
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-04-29) Function created
    #>
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ID,

        [Parameter(Mandatory=$true)]
        $Header
    )    
        try {
        # Attempt to call Graph and retrieve all assignments for Win32 app
        $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($ID)/assignments"
        $Win32AppAssignmentResponse = Invoke-RestMethod -Method "GET" -Uri $Uri -ContentType "application/json" -Headers $Header -ErrorAction Stop
        if ($Win32AppAssignmentResponse.value -ne $null) {
            foreach ($Win32AppAssignment in $Win32AppAssignmentResponse.value) {
                Write-Verbose -Message "Successfully retrieved Win32 app assignment with ID: $($Win32AppAssignment.id)"
                Write-Output -InputObject $Win32AppAssignment
            }
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while retrieving Win32 app assignments for app with ID: $($Win32AppID). Error message: $($_.Exception.Message)"
    }
}#endfunction
function Add-IntuneWin32AppAssignment {
    <#
    .SYNOPSIS
        Assign a Win32App to a Device Group
    .DESCRIPTION
        Assign a Win32App to a Device Group
    .PARAMETER AppID
        Specify the AppID for a Win32 application.
    .PARAMETER GroupID
        Specify the GroupID for a Win32 application assignment
    .PARAMETER Nofification 
        Specify the Notification parameter for the assignment
    .PARAMETER DeliveryOptimization 
        Specify the DeliveryOptimization parameter for the assignment
    .PARAMETER GracePeriod 
        Specify the GracePeriod parameter for the assignment
    .PARAMETER RestartCountDown 
        Specify the RestartCountDown parameter for the assignment
    .PARAMETER RestartSnooze 
        Specify the RestartSnooze parameter for the assignment
    .PARAMETER DeadlineInDays 
        Specify the DeadLine in Days parameter for the assignment
    .NOTES
        Author:      Jan Ketil Skanke
        Contact:     @JankeSkanke
        Created:     2020-05-10
        Updated:     2020-05-10
        Version history:
        1.0.0 - (2020-04-29) Function created
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Header,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$AppID,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupID,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("showAll", "showReboot", "hideAll")]
        [string]$Notification = "showReboot",

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("notConfigured", "foreground")]
        [string]$DeliveryOptimization = "notConfigured",

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$GracePeriod = "1440",

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$RestartCountDown = "30",

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$RestartSnooze = "30",

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int32]$DeadlineInDays = "3",
        
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int32]$StartinDays = "2"

    )
    #Define install deadline 
    [string]$Deadline = Get-Date((Get-Date).AddDays($DeadlineInDays)) -Format "yyyy-MM-ddTHH:mm:ssZ"
    [string]$StartTime = Get-Date((Get-Date).AddDays($StartinDays)) -Format "yyyy-MM-ddTHH:mm:ssZ"
    
    $Win32AppUri =  "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppID)"
    $Win32App = Invoke-RestMethod -Method Get -Uri $Win32AppUri -ContentType "application/json" -Headers $Header
    Write-Output "Restart behavior is set to $($Win32App.installExperience.deviceRestartBehavior)"
    if (($Win32App.installExperience.deviceRestartBehavior -eq "basedOnReturnCode") -or ($Win32App.installExperience.deviceRestartBehavior -eq "force"))
        {
        Write-Output "Assigning app with Restart Grace"
        $Win32AppAssignmentBody = [ordered]@{
            "@odata.type" = "#microsoft.graph.mobileAppAssignment"
            "intent" = "required"
            "source" = "direct"
            "target" = @{
                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                "groupId" = "$GroupID"
            }
            "settings" = @{
                "@odata.type" = "#microsoft.graph.win32LobAppAssignmentSettings"
                "notifications" = "$Notification"
                "deliveryOptimizationPriority" = "$DeliveryOptimization"
        
                "restartSettings" = @{
                    "gracePeriodInMinutes" = "$GracePeriod"
                    "countdownDisplayBeforeRestartInMinutes" = "$RestartCountDown"
                    "restartNotificationSnoozeDurationInMinutes" = "$RestartSnooze"
                }
                "installTimeSettings" = @{
                    "useLocalTime" = $true
                    "startDateTime" = $StartTime
                    "deadlineDateTime" = $Deadline
                }
            }
        }
    }
    else{
        Write-Output "Assigning app without Restart Grace, grace not applicable"
        # Construct table for Win32 app assignment body
        $Win32AppAssignmentBody = [ordered]@{
            "@odata.type" = "#microsoft.graph.mobileAppAssignment"
            "intent" = "required"
            "source" = "direct"
            "target" = @{
                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                "groupId" = "$GroupID"
            }
            "settings" = @{
                "@odata.type" = "#microsoft.graph.win32LobAppAssignmentSettings"
                "notifications" = "$Notification"
                "deliveryOptimizationPriority" = "$DeliveryOptimization"
                "installTimeSettings" = @{
                    "useLocalTime" = $true
                    "startDateTime" = $null
                    "deadlineDateTime" = $Deadline
                }
            }
        }    
    }
    #Write-Output ($Win32AppAssignmentBody | ConvertTo-Json)
    $uri = -join ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/", $AppID ,"/assignments")
    #Write-Output $uri
    try {
        # Attempt to call Graph and create new assignment for Win32 app
        $Win32AppAssignmentResponse = Invoke-RestMethod -Method "POST" -Uri $uri -Body ($Win32AppAssignmentBody | ConvertTo-Json) -ContentType "application/json" -Headers $Header -ErrorAction Stop
        if ($Win32AppAssignmentResponse.id) {
            Write-Output "Successfully created Win32 app assignment with ID: $($Win32AppAssignmentResponse.id)"
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while creating app assignment. Error message: $($_.Exception.Message)"
    }        
}#endfunction
function Remove-IntuneAppAssignment{
    <#
    .SYNOPSIS
        Remove an assignment of a Win32 app to specified AAD groupID
    .DESCRIPTION
        Remove an assignment of a Win32 app to specified AAD groupID
    .PARAMETER AppID
        Specify the ID for the Win32 application.
    .PARAMETER DeviceGroupID
        Specify the Group ID for Device Patching Group
        
        .NOTES
        Author:      Jan Ketil Skanke
        Contact:     @JankeSkanke
        Original Author: @Nickolaja
        Created:     2020-13-05
        Updated:     2020-13-05

        Version history:
        1.0.0 Function Created
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Header,

        [parameter(Mandatory = $true)]
        [string]$AppID,

        [Parameter(Mandatory=$true)]
        [string]$DeviceGroupID
    )
    $uri = -join ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/", $AppID ,"/assignments")
    #Write-Output $uri
    try {
        # Attempt to call Graph and create new assignment for Win32 app
        $Win32AppAssignments = Invoke-RestMethod -Method "GET" -Uri $uri -ContentType "application/json" -Headers $Header -ErrorAction Stop
        foreach ($Win32AppAssignment in $Win32AppAssignments.value){
                $CurrentWin32AppAssignment = Invoke-RestMethod -Method "GET" -uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppID)/assignments/$($Win32AppAssignment.id)" -Headers $Header -ErrorAction Stop
                If (($CurrentWin32AppAssignment.target.'@odata.type') -like "*.groupAssignmentTarget"){
                    if (($CurrentWin32AppAssignment.target.GroupID) -eq $DeviceGroupID){
                        #remove assignment 
                        Invoke-RestMethod -uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppID)/assignments/$($CurrentWin32AppAssignment.id)" -Method "DELETE" -Headers $Header -ErrorAction Stop
                    }
                }
            }    
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while removing app assignment. Error message: $($_.Exception.Message)"
    }        
}#endfunction
function Get-DeviceAppInstallState{
    <#
    .SYNOPSIS
        Retrieve all devices with Win32 app installed
    .DESCRIPTION
        Retrieve all assignments for a Win32 app.
    .PARAMETER ID
        Specify the ID for a Win32 application.
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-05-05) Function created
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Header,    
        
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ID
    )
    try {
        # Attempt to call Graph and retrieve all devices with app installed equals true
        $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($ID)/deviceStatuses?filter=installState eq 'installed'&Select=deviceID"
        $Win32AppDevices = Invoke-RestMethod -Method GET -Uri $Uri -ContentType "application/json" -Headers $Header -ErrorAction Stop
        if ($Win32AppDevices.value -ne $null) {
            foreach ($Win32AppDevice in $Win32AppDevices.value) {
                Write-Output $Win32AppDevice.deviceID
            }
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while getting installstate for app with ID: $($ID). Error message: $($_.Exception.Message)"
    }
}#endfunction 
Function GetNameStringForIntuneSearch{
    <#
    .SYNOPSIS
        Get Application simplified string for search via Graph Api
    .DESCRIPTION
        Get Application simplified string for search via Graph Api
    .PARAMETER name
        Specify the full name for the application
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-05-05) Function created
    #>
    param (
        [string]$name)
        if ($name -match "\d+\.\d+(\.\d+)?(\.\d+)?")
        {
            $result = $name.Split($Matches[0])
            $result = $result[0].Trim()
            return $result
        }
    return $name
}#endfunction
function Invoke-CreateAADGroup{
    <#
    .SYNOPSIS
        Create AzureAD Device Group for Patch Remediation
    .DESCRIPTION
        Create AzureAD Device Group for Patch Remediation
    .PARAMETER DisplayName
        The groups Displayname in Azure AD
    .PARAMETER Description
        The group description on the group in Azure AD
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-05-05) Function created
    #>
    param (
        [Parameter(Mandatory=$true)]
        $Header,

        [Parameter(Mandatory=$true)]
        $Description,

        [Parameter(Mandatory=$true)]
        $DisplayName
    )
    $newGroupJSONObject = @{
        "description" = $Description
        "displayName"= $DisplayName
        "mailEnabled" = $false
        "mailNickname" = "none"
        "securityEnabled" = $true
    } | ConvertTo-Json
    $Group = Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/beta/groups/' -ContentType "application/json" -Headers $Header -Body $newGroupJSONObject 
    Return $Group    
}#endfunction
function Get-AADPatchingGroup{
    <#
    .SYNOPSIS
        Get AzureAD Device Group used for Patch Remediation
    .DESCRIPTION
        Get AzureAD Device Group used for Patch Remediation
    .PARAMETER GroupName
        Name of the Group in Azure AD
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-05-05) Function created
    #>
    param (
        [Parameter(Mandatory=$true)]
        $Header,

        [Parameter(Mandatory=$true)]
        $GroupName
    )
    try {
        # Attempt to call Graph and retrieve AAD Group
        $uri = "https://graph.microsoft.com/beta/groups?`$filter=displayname eq '$GroupName'"
        $Group = Invoke-RestMethod -Method "GET" -Uri $uri -ContentType "application/json" -Headers $Header -ErrorAction Stop
        Return $Group
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while retrieving Group with GroupName: $($GroupName). Error message: $($_.Exception.Message)"
    }
}#endfunction
function Remove-AADPatchingGroup{
    <#
    .SYNOPSIS
        Remove AzureAD Device Group used for Patch Remediation
    .DESCRIPTION
        Remove AzureAD Device Group used for Patch Remediation
    .PARAMETER GroupID
        The GroupID of the Group in Azure AD to remove
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-05-05) Function created
    #>
    param (
        [Parameter(Mandatory=$true)]
        $Header,

        [Parameter(Mandatory=$true)]
        $GroupID
    )
    try {
        # Attempt to call Graph and Delete AAD Group
        $uri = "https://graph.microsoft.com/beta/groups/$($GroupID)"
        $DeleteAADGroupResponse = Invoke-RestMethod -Method "DELETE" -Uri $uri -ContentType "application/json" -Headers $Header -ErrorAction Stop
        Return $DeleteAADGroupResponse
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while retrieving Group with GroupName: $($GroupName). Error message: $($_.Exception.Message)"
    }
}#endfunction
function Add-DeviceToAADGroup{
    <#
    .SYNOPSIS
        Add Devices to AAD Group
    .DESCRIPTION
        Add Devices to AAD Group
    .PARAMETER DeviceID
        Specify the DeviceID to add to group
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-04-29
        Updated:     2020-04-29
        Version history:
        1.0.0 - (2020-05-05) Function created
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Header,

        [parameter(Mandatory = $true)]
        [string]$DeviceID
    )
    try {
        #Get AzureAD Device id from ManagedDeviceID
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($DeviceID)?&select=azureADDeviceId"
        $AADDeviceIDQuery = Invoke-RestMethod -Method Get -Uri $uri -ContentType "application/json" -Headers $Header
        $AADDeviceID =$AADDeviceIDQuery.azureADDeviceId
        #Get AzureAD Directory Object from AzureAD DeviceID
        $DirectoryObjectUrl = -join ("https://graph.microsoft.com/beta/devices?filter=deviceId eq '",$AADDeviceID, "'&select=id")
        $DirectoryObject = Invoke-RestMethod -Method Get -Uri $DirectoryObjectUrl -ContentType "application/json" -Headers $Header
        $DirectoryObjectID =  $DirectoryObject.value.id
        $AddToGroupUri = "https://graph.microsoft.com/beta/groups/$RemediationGroupID/members/`$ref"
        $Body = [ordered]@{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$DirectoryObjectID"
        }
        Write-Output "Adding $AADDeviceID to group $GroupName"
        Invoke-RestMethod -Method Post -Uri $AddToGroupUri -Body ($Body | ConvertTo-Json) -ContentType "application/json" -Headers $Header -ErrorAction Stop
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while adding device with ID: $($DeviceID) to group. Error message: $($_.Exception.Message)"
    }
}#endfunction 
function Invoke-CreateHealthScripts {
    <#
    .SYNOPSIS
        Dynamicly Create Endpoint Analytics Health Scripts
    .DESCRIPTION
        Dynamicly Create Endpoint Analytics Detection and Remediations Scripts
    .PARAMETER OrgAppIntuneID
        Intune AppID for the app that needs an update
    .PARAMETER UpdatedAppIntuneID
        Intune AppID for the updated app
    .PARAMETER OrgAppVersion
        The Application name for the app that needs an update
    .PARAMETER UpdatedAppVersion
        The Application name for the updated app
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-03-29) Function created
    #>   
    param (
        [parameter(Mandatory = $true)]
        $OrgAppIntuneID,
        [parameter(Mandatory = $true)]
        $UpdatedAppIntuneID,
        [parameter(Mandatory = $true)]
        $OrgAppVersion,
        [parameter(Mandatory = $true)]
        $UpdatedAppVersion, 
        [parameter(Mandatory = $true)]
        $LogoImageUri,
        [parameter(Mandatory = $true)]
        $HeroImageUri, 
        [parameter(Mandatory = $true)]
        $AttributionText
    )
#Creating dynamic detectionscript
$DetectionScript ="Function Get-AllInstalledSoftware{
    `$regpath = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*')
    `$propertyNames = 'DisplayName','DisplayVersion','PSChildName','Publisher','InstallDate'
    
    if (-not ([IntPtr]::Size -eq 4)) 
    {
        `$regpath += 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    }
    Get-ItemProperty `$regpath -Name `$propertyNames -ErrorAction SilentlyContinue | .{process{if(`$_.DisplayName) { `$_ } }} | Select-Object DisplayName, DisplayVersion, PSChildName, Publisher, InstallDate | Sort-Object DisplayName
    }#endfunction
    Function DetectInstalledIntuneAppID{
        [CmdletBinding()]
            param (
                [parameter(Mandatory = `$True)]
                [ValidateNotNullOrEmpty()]
                [string]`$AppID
                )
        Process {
            `$Installed = `$false            
            Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\SideCarPolicies\StatusServiceReports\*' -Recurse -ErrorAction SilentlyContinue | ForEach-Object { 
               if((get-itemproperty -Path `$_.PsPath) -match `$AppID){ 
                    if((Get-ItemProperty -Path `$_.PsPath -Name 'Status','AppID') -match 'Installed'){
                        `$Installed = `$true
                    }
                }
            } 
            Return `$Installed
            }
        }#end function

    Function GetNameStringForLocalDetection{
    param (
        [string]`$name)
        if (`$name -match `"\d+\.\d+(\.\d+)?(\.\d+)?`")
        {
            `$result = `$name.Split(`$Matches[0])
            `$result = `$result[0].Trim()
            return `$result
        }
    return `$name
    }#endfunction
        
    #Set variables
    `$OrgAppID = '$OrgAppIntuneID'
    `$UpdatedAppID = '$UpdatedAppIntuneID'
    `$UpdatedAppVersion = '$UpdatedAppVersion'
    `$AppToDetect = GetNameStringForLocalDetection -name `$UpdatedAppVersion

    #Detection
    if (DetectInstalledIntuneAppID -AppID `$UpdatedAppID){
        Write-Output `"App with Appid `$UpdatedAppID is installed`"
        exit 0
    }
    elseif (DetectInstalledIntuneAppID -AppID `$OrgAppID) {
        Write-Output `"`$OrgAppID needs and update. Starting remediation notification`"
        Exit 1
    }
    else {
        `$output = -join ('Intune app ', `$AppToDetect, ' is not installed, checking local inventory for unmanaged app')
        Write-Output `$output
        `$InstalledSoftware = Get-AllInstalledSoftware
        Write-Output `"Checking local installed apps..`"
    if (`$InstalledSoftware.DisplayName -match [regex]::escape(`$AppToDetect)){
        `$output = -join ('Unmanaged application ', `$AppToDetect, ' detected, starting remediation..')
        Write-Output `$output
        `$DetectedApp = `$InstalledSoftware| Where-Object {(`$_.DisplayName -match [regex]::escape(`$AppToDetect))} | Get-Unique
        `$output = -join ('Notifying users to move to managed app, replacing  ', `$DetectedApp.DisplayName, `$DetectedApp.DisplayVersion)
        Write-Output `$output           
        exit 1
    }
    Write-Output `"Application not installed `$AppToDetect`"
    exit 0
}
" 
# Creating dynamic remediation script
$RemediationScript = "function Display-ToastNotification() {
    `$Load = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    `$Load = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    # Load the notification into the required format
    `$ToastXML = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument
    `$ToastXML.LoadXml(`$Toast.OuterXml)
        
    # Display the toast notification
    try {
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(`$App).Show(`$ToastXml)
    }
    catch { 
        Write-Output -Message 'Something went wrong when displaying the toast notification' -Level Warn
        Write-Output -Message 'Make sure the script is running as the logged on user' -Level Warn     
    }
}
# Setting image variables
`$LogoImageUri = `"$LogoImageUri`"
`$HeroImageUri = `"$HeroImageUri`"
`$LogoImage = `"`$env:TEMP\ToastLogoImage.png`"
`$HeroImage = `"`$env:TEMP\ToastHeroImage.png`"
`$UpdatedAppID = '$UpdatedAppIntuneID'
`$OrgAppVersion = '$OrgAppVersion'
`$NewAppVersion = '$UpdatedAppVersion'

#Fetching images from uri
Invoke-WebRequest -Uri `$LogoImageUri -OutFile `$LogoImage
Invoke-WebRequest -Uri `$HeroImageUri -OutFile `$HeroImage

#Defining the Toast notification settings
#ToastNotification Settings
`$Scenario = 'reminder' # <!-- Possible values are: reminder | short | long -->
`$Action = 'companyportal:Applicationid=$UpdatedAppIntuneID' # Generated from the Runbook - Starts Company Portal on the app
        
# Load Toast Notification buttons
`$ActionButtonEnabled = 'True'
`$ActionButtonContent = 'Update now'
`$DismissButtonEnabled = 'True'
`$DismissButtonContent = 'Dismiss'
`$SnoozeButtonEnabled = 'True'
`$SnoozeButtonContent = 'Snooze'

# Load Toast Notification text
`$AttributionText = `"$AttributionText`"
`$HeaderText = 'Application Update needed'
`$TitleText = '$OrgAppVersion needs an update!'
`$BodyText1 = 'For security and stability reasons, we kindly ask you to update to $UpdatedAppVersion as soon as possible. If you do not update the app in $Win32AppStartTime days, the app update will be enforced'
`$BodyText2 = 'Updating your 3rd party apps on a regular basis ensures a secure Windows. Thank you in advance.'
    
# New text options
`$HourText = ' Hour'
`$HoursText = ' Hours'
`$MinutesText = ' Minutes'


# Check for required entries in registry for when using Powershell as application for the toast
# Register the AppID in the registry for use with the Action Center, if required
`$RegPath = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings'
`$App =  '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'

# Creating registry entries if they don't exists
if (-NOT(Test-Path -Path `"`$RegPath\`$App`")) {
    New-Item -Path `"`$RegPath\`$App`" -Force
    New-ItemProperty -Path `"`$RegPath\`$App`" -Name 'ShowInActionCenter' -Value 1 -PropertyType 'DWORD'
}

# Make sure the app used with the action center is enabled
if ((Get-ItemProperty -Path `"`$RegPath\`$App`" -Name 'ShowInActionCenter' -ErrorAction SilentlyContinue).ShowInActionCenter -ne '1') {
    New-ItemProperty -Path `"`$RegPath\`$App`" -Name 'ShowInActionCenter' -Value 1 -PropertyType 'DWORD' -Force
}


# Formatting the toast notification XML
# Create the default toast notification XML with action button and dismiss button
if ((`$ActionButtonEnabled-eq 'True') -AND (`$DismissButtonEnabled -eq 'True')) {
    Write-Output -Message 'Creating the xml for displaying both action button and dismiss button'
[xml]`$Toast = @`"
<toast scenario='`$Scenario'>
    <visual>
    <binding template='ToastGeneric'>
        <image placement='hero' src='`$HeroImage'/>
        <image id='1' placement='appLogoOverride' hint-crop='circle' src='`$LogoImage'/>
        <text placement='attribution'>`$AttributionText</text>
        <text>`$HeaderText</text>
        <group>
            <subgroup>
                <text hint-style='title' hint-wrap='true' >`$TitleText</text>
            </subgroup>
        </group>
        <group>
            <subgroup>     
                <text hint-style='body' hint-wrap='true' >`$BodyText1</text>
            </subgroup>
        </group>
        <group>
            <subgroup>     
                <text hint-style='body' hint-wrap='true' >`$BodyText2</text>
            </subgroup>
        </group>
    </binding>
    </visual>
    <actions>
        <action activationType='protocol' arguments='`$Action' content='`$ActionButtonContent' />
        <action activationType='system' arguments='dismiss' content='`$DismissButtonContent'/>
    </actions>
</toast>
`"@
}
# Snooze button - this option will always enable both action button and dismiss button regardless of config settings
if (`$SnoozeButtonEnabled -eq 'True') {
    Write-Output -Message 'Creating the xml for snooze button'
[xml]`$Toast = @`"
<toast scenario='`$Scenario'>
    <visual>
    <binding template='ToastGeneric'>
        <image placement='hero' src='`$HeroImage'/>
        <image id='1' placement='appLogoOverride' hint-crop='circle' src='`$LogoImage'/>
        <text placement='attribution'>`$AttributionText</text>
        <text>`$HeaderText</text>
        <group>
            <subgroup>
                <text hint-style='title' hint-wrap='true' >`$TitleText</text>
            </subgroup>
        </group>
        <group>
            <subgroup>     
                <text hint-style='body' hint-wrap='true' >`$BodyText1</text>
            </subgroup>
        </group>
        <group>
            <subgroup>     
                <text hint-style='body' hint-wrap='true' >`$BodyText2</text>
            </subgroup>
        </group>
    </binding>
    </visual>
    <actions>
        <input id='snoozeTime' type='selection' title='`$SnoozeButtonText' defaultInput='15'>
            <selection id='15' content='15`$MinutesText'/>
            <selection id='30' content='30`$MinutesText'/>
            <selection id='60' content='1`$HourText'/>
            <selection id='240' content='4`$HoursText'/>
            <selection id='480' content='8`$HoursText'/>
        </input>
        <action activationType='protocol' arguments='`$Action' content='`$ActionButtonContent' />
        <action activationType='system' arguments='snooze' hint-inputId='snoozeTime' content='`$SnoozeButtonContent'/>
        <action activationType='system' arguments='dismiss' content='`$DismissButtonContent'/>
    </actions>
</toast>
`"@

}

#Send the notification
Display-ToastNotification
Exit 0"
$HealthScripts = @($DetectionScript, $RemediationScript)
Return $HealthScripts
}#endfunction
function New-IntuneProactiveRemedation{
    <#
    .SYNOPSIS
        Uploads a new proactive remediation to Intune. 
    .DESCRIPTION
        Uploads a new proactive remediation to Intune. Takes 2 base64 encoded scripts as input. 
    .PARAMETER Publisher
        The name of the publisher of the ProActive Remediation
    .PARAMETER Displayname
        The name of the DisplayName of the ProActive Remediation
    .PARAMETER Description
        Description of the proactive remediation
    .PARAMETER EncodedDetectionScript
        The Base64 encoded script string
    .PARAMETER EncodedRemediationScript
        The Base64 encoded script string
    .PARAMETER runAsAccount
        RunaAs user or system - default is user
    .PARAMETER Header
        The Graph Token Header for authentication
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-03-29) Function created
    #>       
    param (
        [parameter(Mandatory = $true)]
        $Publisher,
        [parameter(Mandatory = $true)]
        $DisplayName,
        [parameter(Mandatory = $false)]
        $Description,
        [parameter(Mandatory = $false)]
        $EncodedDetectionScript,
        [parameter(Mandatory = $true)]
        $EncodedRemediationScript,
        [parameter(Mandatory = $false)]
        $runAsAccount ="user",
        [parameter(Mandatory = $true)]
        $Header
    )
    $Body = ConvertTo-JSON -Depth 3 @{ 
        "displayName" = "$DisplayName"
        "description" = "$Description"
        "publisher"= "$Publisher"
        "runAs32Bit"= "false"
        "runAsAccount"= "$runAsAccount"
        "enforceSignatureCheck"= "false"
        "detectionScriptContent"= "$EncodedDetectionScript"
        "remediationScriptContent" = "$EncodedRemediationScript"
    }
    
    $response = Invoke-RestMethod -Method Post "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/" -Headers $Header -Body $body -ContentType "application/json" 
    return $response.id
}#endfunction
function New-IntuneProctiveRemediationAssignment{
    <#
    .SYNOPSIS
        Assigns a proactive remediation in Intune to all users. 
    .DESCRIPTION
        Assigns a proactive remediation in Intune to all users. 
    .PARAMETER MyProactiveRemediationID
        The ID of the ProactiveRemediation
    .PARAMETER SceduledTime 
        Time of day for when the Proactive Remediation should run. Format ("HH:MM:SS")
    .PARAMETER ScheduleInterval
        How often should the Proactive Remediation run. (Set to 1 for each day, set to 7 for once a week.)
    .PARAMETER Header
        The Graph Token Header for authentication
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-06-27) Function created
    #>       
    param (
        [parameter(Mandatory = $true)]
        $MyProactiveRemediationID,
        [parameter(Mandatory = $false)]
        $scheduledtime ="9:0.0",
        [parameter(Mandatory = $false)]
        $scheduleinterval ="1",
        [parameter(Mandatory = $true)]
        $Header
    )
    $uri = -join("https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/",$MyProactiveRemediationID, "/assign")
    $body = "{`"deviceHealthScriptAssignments`":[
        {`"target`":
            {`"@odata.type`":`"#microsoft.graph.allLicensedUsersAssignmentTarget`"
        },
        `"runRemediationScript`":true,
        `"runSchedule`":
        {`"@odata.type`":`"#microsoft.graph.deviceHealthScriptDailySchedule`",
        `"interval`":$scheduleinterval,
        `"time`":`"$scheduledtime`",
        `"useUtc`":false
        }
        }
        ]
        }"
    try {
        # Attempt to call Graph and assign the proactive remediation package
        Invoke-RestMethod -Method Post $uri -Headers $Header -Body $Body -ContentType "application/json"       
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while assigning remediation script package with ID: $($MyProactiveRemediationID). Error message: $($_.Exception.Message)"
    }
}#endfunction
function Remove-IntuneProactiveRemediation {
    <#
    .SYNOPSIS
        Removes a proactive remediation in Intune based on name
    .DESCRIPTION
        Removes a proactive remediation in Intune based on name, 
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-06-27) Function created
    #>       
    try {
        $ScriptNameToRemove = "Patching $OrgAppVersion"
        $GetScriptIDUri =  -join("https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?filter=startswith(displayName,'", $ScriptNameToRemove ,"')&select=id")
        $ScriptID = Invoke-RestMethod -Method Get $GetScriptIDUri -Headers $Header -ContentType "application/json"
        if ($null -ne $ScriptID.value.id){
            $DeleteScriptUri = -join("https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/", $ScriptID.value.id)
            Invoke-RestMethod -Method Delete $DeleteScriptUri -Headers $Header -ContentType "application/json"
            Write-Output "Old remedationscript deleted"
        }
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $output = -join("Error deleting",$ScriptNameToRemove, " ErrorMessage: ",$ErrorMessage)
        Write-Error -Message $output
    }
}#endfunction
function ConvertTo-Base64{
    <#
    .SYNOPSIS
        Encodes a string of data to Base64
    .DESCRIPTION
        Encodes a string of data to Base64, used for uploading scripts to endpoint analytics and returns the base64 encoded version. 
    .PARAMETER Script
        A powershell script as a single text string is taken as input. 
    .NOTES
        Author:      Jan Ketil Skanke 
        Contact:     @JankeSkanke
        Created:     2020-03-29
        Updated:     2020-03-29
        Version history:
        1.0.0 - (2020-03-29) Function created
    #>   
    [CmdletBinding()]
        param (
            [parameter(Mandatory = $true)]
            [string]$Script
        )
        $commandBytes = [System.Text.Encoding]::UTF8.GetBytes($Script)
        $encodedScript = [Convert]::ToBase64String($commandBytes)
        return $encodedScript
}#endfunction
#endregion Functions

#region Script
#Get Authentication Token for MSGraph
Try
    {
        $Header = Get-MSGraphAppToken -TenantID $tenantId -ClientID $ClientID -ClientSecret $ClientSecret -ErrorAction Stop        
    }
Catch
    {
        $ErrorMessage = $_.Exception.Message
        Write-Error -Message "Connection to Graph failed with $ErrorMessage"
        Break
    }

if ($Header) {
    Write-Output "Connected to Microsoft Graph"
    Write-Output "Patching app $OrgAppVersion ID: $OrgAppID with $UpdatedAppVersion ID:  $UpdatedAppID"
        
    $CurrentAssignment = Get-IntuneWin32AppAssignment -ID $UPDATEDAPPID -Header $Header 
    #If new app is assigned as available - detect all devices with installstate = installed and add to patchinggroup
    #This is based on that other mechanism handles the app update itself and assigns updated app to available and not required
    if ($CurrentAssignment.intent -contains "available"){
        #If notifications is enabled via runbook variable
        if (Get-AutomationVariable -Name 'Notifications'){
            #Dynamicly create Proactive Remediaton Scripts for Intune
            $DynamicHealthScripts = Invoke-CreateHealthScripts -OrgAppIntuneID $OrgAppID -UpdatedAppIntuneID $UpdatedAppID -OrgAppVersion $OrgAppVersion -UpdatedAppVersion $UpdatedAppVersion -LogoImageUri $LogoImageUri -HeroImageUri $HeroImageUri -AttributionText $AttributionText
            $DynamicDetectionScript = $DynamicHealthScripts[0]
            $DynamicRemediationScript = $DynamicHealthScripts[1]

            #Convert to base 64 for upload
            $EncodedDetectionScript = ConvertTo-Base64 -Script $DynamicDetectionScript
            $EncodedRemediationScript = ConvertTo-Base64 -Script $DynamicRemediationScript
            
            #Upload Scripts to Intune 
            $MyProactiveRemediationID = New-IntuneProactiveRemedation -Header $Header -Publisher $EAPublisher -DisplayName "Patching $UpdatedAppVersion" -EncodedDetectionScript $EncodedDetectionScript -EncodedRemediationScript $EncodedRemediationScript

            #AssignToAllUsers
            New-IntuneProctiveRemediationAssignment -MyProactiveRemediationID $MyProactiveRemediationID -scheduledtime $Scheduledtime -scheduleinterval $Scheduleinterval -Header $Header
            
            #Delete Old Remedation script version from Inune 
            Remove-IntuneProactiveRemediation 
        }
        #Detect and remove old patching group
        $OldPatchingGroup = Get-AADPatchingGroup -GroupName "AppRemedy-$($ORGAPPVERSION)" -Header $Header
        if ($OldPatchingGroup.value -ne $null){
            Write-Output "Old Patching group found. Deleting old Patching Group and assignments"
            #removing eventually copied assignement from PMPC
            Write-Output "Removing App Assignement for $($OldPatchingGroup.value.id)"
            Remove-IntuneAppAssignment -AppID $UPDATEDAPPID -DeviceGroupID $OldPatchingGroup.value.id -Header $Header
            #removing group 
            Remove-AADPatchingGroup -GroupID $OldPatchingGroup.value.id -Header $Header
        }
        else{
            Write-Output "No old Patching Group exists, continues..."
        }
        #Creating new patching group 
        Write-Output "App is assigned as available, creating patching group for remediation"
        $GroupName = -join("AppRemedy-", $UPDATEDAPPVERSION)
        $GroupDescription = -join("App Remediation Group for ", $UPDATEDAPPVERSION)
        $RemediationGroup = Invoke-CreateAADGroup -Description $GroupDescription -DisplayName $GroupName -Header $Header
        $RemediationGroupID = $RemediationGroup.id
        #Detect devices with previous versions of the app and add them to the patching group
        $AppSearchName = GetNameStringForIntuneSearch -name $UpdatedAppVersion
        $AppSearchUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$search=`"$AppSearchName`"&Select=id, displayName"
        $AllAppVersionsDetected = Invoke-RestMethod -Method Get -Uri $AppSearchUri -ContentType "application/json" -Headers $Header 
        foreach ($App in  $AllAppVersionsDetected.value | Where-Object {$_.id -notmatch $UpdatedAppID}){
            $CurrentAppID = $app.id
            #Write-Output "CurrentAPPID $CurrentAppID"
            $Devices =  Get-DeviceAppInstallState -ID $CurrentAppID -Header $Header| Get-Unique
            foreach ($Device in $Devices){
                Add-DeviceToAADGroup -DeviceID $Device -Header $Header
            }
        }
    }
    elseif ($CurrentAssignment.intent -notcontains "available" -And $CurrentAssignment.intent -contains "required"){
        Write-Output "App is already assigned as required only, no action needed"
    }
    #Asssigning Updated App to Patching Group as Required with a 2 day deadline
    Add-IntuneWin32AppAssignment -AppID $UPDATEDAPPID -GroupID $RemediationGroupID -Notification showAll -DeliveryOptimization foreground -GracePeriod 1440 -RestartCountDown 20 -RestartSnooze 15 -DeadlineInDays $Win32AppDeadLine -StartinDays $Win32AppStartTime -Header $Header 

    #Delete old app version from Intune if enabled
    if (Get-AutomationVariable -Name 'CleanOldAppVersion'){
        try {
        Write-Output "Deleting old app version $($OrgAppVersion)"
        Invoke-RestMethod -Method Delete -Headers $Header -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($OrgAppID)"
        }
        catch [System.Exception] {
        Write-Warning -Message "An error occurred while removing app $($OrgAppVersion). Error message: $($_.Exception.Message)"
        }   
    }
    #If Teams Reporting is set to True in Runbook Variable
    if ($TeamsReporting){
        $TeamsWebHook = Get-AutomationVariable -Name 'TeamsReportChannelWebHook'
        #Finding app name - stripping out version for search 
        $AppSearchName = GetNameStringForIntuneSearch -name $UpdatedAppVersion
        $AppSearchUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$search=`"$AppSearchName`"&Select=id, displayName"
        $AllAppVersionsDetected = Invoke-RestMethod -Method Get -Uri $AppSearchUri -ContentType "application/json" -Headers $Header 
        $DeviceNames =[System.Collections.ArrayList]@()
        foreach ($App in  $AllAppVersionsDetected.value | Where-Object {$_.id -notmatch $UpdatedAppID}){
            $CurrentAppID = $app.id
            $CurrentAppVersion = $app.displayName
            #Write-Output $CurrentAppID
            $DeviceStateOldUri =  -join('https://graph.microsoft.com/beta/deviceappmanagement/mobileApps/', $CurrentAppID,'/deviceStatuses?filter=installState eq ',"'","installed","'",'&Select=deviceName, deviceID')
            #Write-Output "DeviceStateOLDURI is $DeviceStateOldUri"
            $QueryOldAppDeviceStatuses = Invoke-RestMethod -Method Get -Uri $DeviceStateOldUri -ContentType "application/json" -Headers $Header
            $OldAppDeviceStatuses = $QueryOldAppDeviceStatuses.value.deviceName | Get-Unique
            #Write-Output "Status $OldAppDeviceStatuses"
            if ($OldAppDeviceStatuses.count -ne 0){
                foreach ($device in $OldAppDeviceStatuses) {
                    $DeviceNames.Add($Device) | Out-Null
                }
            }
        }
        if (-not ([string]::IsNullOrEmpty($DeviceNames))){
            $DeviceNames = $DeviceNames | Sort-Object -Unique
            $StatusText =  -join($DeviceNames.count, " devices needs an update to ", $UpdatedAppVersion)
            $StatusColor = "FFFF00"
            $Text = -join($DeviceNames.count, " devices needs an update to ", $UpdatedAppVersion)
            $DeviceNameString = $null
            foreach ($DeviceName in $DeviceNames) {
                Write-Output "DeviceName: $DeviceName"
                $DeviceNameString += " \n\n "
                $DeviceNameString += $DeviceName
            }
            Invoke-TeamsMessage -teamswebhook $TeamsWebHook -UpdateStatus $StatusText -DeviceName $DeviceNameString -text $text -StatusColor $statuscolor| Out-Null
        }
        else {
            $DeviceNames = "No Devices"
            $StatusText = "No devices is currently using a previous version of $UpdatedAppVersion"
            $StatusColor = "008000"
            $Text = "No devices is currently using a previous version of $UpdatedAppVersion"
            Invoke-TeamsMessage -teamswebhook $TeamsWebHook -UpdateStatus $StatusText -DeviceName $DeviceNames -text $text -StatusColor $statuscolor| Out-Null
        }
    }
    else{
        Write-Output "Teams Reporting is not enabled"
    }
}
elseif  (!$Header){
    Write-Warning -Message "Graph Connection failed, check app secret"
    Exit 1
}
#endregion Script
