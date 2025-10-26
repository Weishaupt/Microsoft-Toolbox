#For local development connect with Connect-MgGraph with dev account, or app if App Permissions are required
#Connect-MgGraph -TenantId xxx

#ExtensionAttribute 10 is used to store all Admin Units the user/device is member of

$startDate = Get-Date -AsUTC
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

try {
    if ($null -eq (Get-MgContext)) {
        Connect-MgGraph -Identity -NoWelcome
        Write-Output "Connected to Microsoft Graph with Azure Run As Account"
    }
}
catch {
    Write-Error "No connection to Microsoft Graph with Azure Run As Account $($_.Exception.Message)"
    return
}

$requiredScopes = "User.ReadWrite.All", "Device.ReadWrite.All", "AdministrativeUnit.ReadWrite.All"

$acquiredScopes = (Get-MgContext).Scopes
Write-Verbose "Acquired Microsoft Graph scopes:"
foreach ($scope in $acquiredScopes) {
    Write-Verbose "*`t$scope"
}

foreach ($scope in $requiredScopes) {
    if ($scope -notin $acquiredScopes) {
        Write-Error "Not all required scopes available"
        return
    }
}

function Get-DeviceMauAssignments {
    [OutputType([String[]])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DeviceId
    )
    
    $extensionAttributes = (Get-MgDevice -DeviceId $DeviceId -Property "extensionAttributes").AdditionalProperties["extensionAttributes"]
    if ($null -ne $extensionAttributes -and $extensionAttributes.ContainsKey("extensionAttribute10")) {
        [string[]]$extensionAttribute10 = $extensionAttributes["extensionAttribute10"] -split ", "
    }
    else {
        [string[]]$extensionAttribute10 = @()
    }
    
    return $extensionAttribute10
}

function Get-UserMauAssignment {
    [OutputType([String[]])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    $extensionAttributes = (Get-MgUser -UserId $UserId -Property "onPremisesExtensionAttributes").OnPremisesExtensionAttributes

    if ($null -ne $extensionAttributes.ExtensionAttribute10) {
        return $extensionAttributes.ExtensionAttribute10 -split ", "
    }
    else {
        return @()
    }
}

function Set-DeviceMauAssignments {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DeviceId,
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [string[]]$AdministrativeUnits
    )

    #Get current device extension attributes to avoid deleting any of the other 14 attributes 
    $extensionAttributes = (Get-MgDevice -DeviceId $DeviceId -Property "extensionAttributes").AdditionalProperties["extensionAttributes"]

    if ($null -eq $AdministrativeUnits) {
        $extensionAttributes.Remove("extensionAttribute10") | Out-Null
        if ($extensionAttributes.Count -eq 0) {
            #Work around an issue with MS Graph where you cannot just remove the last extensionAttribute but need to send null instead
            $extensionAttributes["extensionAttribute10"] = "null"
        }
    }
    else {
        $extensionAttributes["extensionAttribute10"] = $AdministrativeUnits -join ", "
    }

    $params = @{
        extensionAttributes = $extensionAttributes
    }
    
    Update-MgDevice -DeviceId $DeviceId -BodyParameter $params    
}

function Set-UserMauAssignment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [string[]]$AdministrativeUnits
    )

    $extensionAttributes = (Get-MgUser -UserId $UserId -Property "onPremisesExtensionAttributes").OnPremisesExtensionAttributes

    if ($null -eq $AdministrativeUnits -or $AdministrativeUnits.Count -eq 0) {
        $extensionAttributes.ExtensionAttribute10 = $null
    }
    else {
        $extensionAttributes.ExtensionAttribute10 = $AdministrativeUnits -join ", "
    }

    $params = @{
        OnPremisesExtensionAttributes = $extensionAttributes
    }

    Update-MgUser -UserId $UserId -BodyParameter $params
}

$adminUnits = Get-MgDirectoryAdministrativeUnit -All
$extensionAttribute10Users = Get-MgUser -Filter "onPremisesExtensionAttributes/extensionAttribute10 ne null" -ConsistencyLevel eventual -CountVariable userCount -Property DisplayName, Id, UserPrincipalName, onPremisesExtensionAttributes -All

foreach ($adminUnit in $adminUnits) {
    Write-Output "Processing Administrative Unit $($adminUnit.DisplayName)"
    [object[]]$unitMembers = Get-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $adminUnit.Id
    [object[]]$unitDevices = $unitMembers | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.device" }
    [object[]]$unitUsers = $unitMembers | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user" }
    [object[]]$markedUsers = $extensionAttribute10Users | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute10 -like "*$($adminUnit.Id)*" }

    $unitUserDeviceMapping = [System.Collections.ArrayList]::new()
    foreach ($user in $unitUsers) {
        [object[]]$regDevices = Get-MgUserRegisteredDevice -UserId $user.Id
        if ($regDevices.Count -gt 0) {
            foreach ($device in $regDevices) {
                $userDeviceMapping = [PSCustomObject]@{
                    UserUpn = $user.AdditionalProperties["userPrincipalName"];
                    Device  = $device
                }
                $unitUserDeviceMapping.Add($userDeviceMapping) | Out-Null
            }
        }   
    }

    # Remove Devices from MAU where user is not associated with MAU / Update extensionAttribute
    foreach ($device in $unitDevices) {
        if ($device.Id -notIn ($unitUserDeviceMapping.Device).Id) {
            Write-Output "Removing device '$($device.AdditionalProperties['displayName'])', with id $($device.Id) from administrative unit $($adminUnit.DisplayName)). Device not associated with any user of administrative unit."
            Remove-MgDirectoryAdministrativeUnitMemberByRef -AdministrativeUnitId $adminUnit.Id -DirectoryObjectId $device.Id
            
            [string[]]$assignedMAUs = Get-DeviceMauAssignments -DeviceId $device.Id
            if ($adminUnit.Id -in $assignedMAUs) {
                $assignedMAUs = $assignedMAUs | Where-Object { $_ -ne $adminUnit.Id }
                Write-Output "Updating extensionAttribute10 of device '$($device.AdditionalProperties['displayName'])', with id $($device.Id). Removing '$($adminUnit.DisplayName)' New value '$($assignedMAUs -join ', ')'"
                Set-DeviceMauAssignments -DeviceId $device.Id -AdministrativeUnits $assignedMAUs
            }
        }
    }

    # Add device to MAU where user is associated with MAU / Update extensionAttribute
    foreach ($deviceMapping in $unitUserDeviceMapping) {
        if ($deviceMapping.Device.Id -notIn ($unitDevices).Id) {
            Write-Output "Adding device '$($deviceMapping.Device.AdditionalProperties['displayName'])', with id $($deviceMapping.Device.Id) to administrative unit $($adminUnit.DisplayName). Device is associated with $($deviceMapping.UserUpn)"
            New-MgDirectoryAdministrativeUnitMemberByRef -AdministrativeUnitId $adminUnit.Id -OdataId "https://graph.microsoft.com/v1.0/devices/$($deviceMapping.Device.Id)"
        }
        [string[]]$assignedMAUs = Get-DeviceMauAssignments -DeviceId $deviceMapping.Device.Id
        if ($adminUnit.Id -notin $assignedMAUs) {
            $assignedMAUs += $adminUnit.Id
            Write-Output "Updating extensionAttribute10 of device '$($deviceMapping.Device.AdditionalProperties['displayName'])', with id $($deviceMapping.Device.Id) to administrative unit $($adminUnit.DisplayName) to '$($assignedMAUs -join ", ")'. Device is associated with $($deviceMapping.UserUpn)"
            Set-DeviceMauAssignments -DeviceId $deviceMapping.Device.Id -AdministrativeUnits $assignedMAUs
        }
    }

    # Remove extensionAttribute10 from users not associated with MAU
    foreach ($user in $markedUsers) {
        if ($user.Id -notin ($unitUsers).Id) {
            [string[]]$assignedMAUs = Get-UserMauAssignment -UserId $user.Id
            $assignedMAUs = $assignedMAUs | Where-Object { $_ -ne $adminUnit.Id }
            Set-UserMauAssignment -UserId $user.Id -AdministrativeUnits $assignedMAUs
        }
    }

    # Add extensionAttribute10 to users associated with MAU where it is missing
    foreach ($user in $unitUsers) {
        [string[]]$assignedMAUs = Get-UserMauAssignment -UserId $user.Id
        if ($adminUnit.Id -notin $assignedMAUs) {
            $assignedMAUs += $adminUnit.Id
            Set-UserMauAssignment -UserId $user.Id -AdministrativeUnits $assignedMAUs
        }
    }
}

$endDate = Get-Date -AsUTC
$processingTime = $endDate - $startDate

Write-Output "Done. Processing took $($processingTime.ToString("hh\:mm\:ss")) (hh:mm:ss)."