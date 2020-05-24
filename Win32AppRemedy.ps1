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
#Fetching all variables from the Automation Account
$ClientID = Get-AutomationVariable -Name 'ClientID2'
$ClientSecret = Get-AutomationVariable -Name 'ClientSecret2'
$TenantID = Get-AutomationVariable -Name 'TenantID'
$TeamsWebHook = Get-AutomationVariable -Name 'TeamsReportChannelWebHook'
$TeamsReporting = Get-AutomationVariable -Name 'TeamsReporting'
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
        [int32]$DeadlineInDays = "2"
    )
    #Define install deadline 
    [string]$Deadline = Get-Date((Get-Date).AddDays($DeadlineInDays)) -Format "yyyy-MM-ddTHH:mm:ssZ"
    
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
    
            "restartSettings" = @{
                "gracePeriodInMinutes" = "$GracePeriod"
                "countdownDisplayBeforeRestartInMinutes" = "$RestartCountDown"
                "restartNotificationSnoozeDurationInMinutes" = "$RestartSnooze"
            }
            "installTimeSettings" = @{
                "useLocalTime" = $true
                "startDateTime" = $null
                "deadlineDateTime" = $Deadline
            }
        }
    }
    Write-Output ($Win32AppAssignmentBody | ConvertTo-Json)
    $uri = -join ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/", $AppID ,"/assignments")
    Write-Output $uri
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
    Write-Output $uri
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
}
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
    Add-IntuneWin32AppAssignment -AppID $UPDATEDAPPID -GroupID $RemediationGroupID -Notification showAll -DeliveryOptimization foreground -GracePeriod 1440 -RestartCountDown 20 -RestartSnooze 15 -DeadlineInDays 3 -Header $Header

    #If Teams Reporting is set to True in Runbook Variable
    if ($TeamsReporting){
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
        if ($DeviceNames -ne $null){
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
            $DeviceNames = "Not Applicable"
            $StatusText = "No devices is currently using a previous version of $UpdatedAppVersion"
            $StatusColor = "008000"
            $Text = "No devices is currently using a previous version of $UpdatedAppVersion"
            Invoke-TeamsMessage -teamswebhook $TeamsWebHook -UpdateStatus $StatusText -DeviceName $DeviceNameString -text $text -StatusColor $statuscolor| Out-Null
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