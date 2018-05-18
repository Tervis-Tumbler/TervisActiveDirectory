function Get-TervisADUser {
    param (
        $Identity,
        $Path,
        $Filter,
        $SearchBase,
        [ValidateSet("Base", "OneLevel", "Subtree")]$SearchScope,
        [String[]]$Properties,
        [Switch]$IncludeMailboxProperties,
        [Switch]$IncludePaylocityEmployee
    )
    $PropertiesIncludingThoseUsedByCustomProperites = $Properties + "msDS-UserPasswordExpiryTimeComputed","LastLogonTimestamp","EmployeeID","Title","Manager","PasswordLastSet","Created","MemberOf","ProtectedFromAccidentalDeletion"

    $ADUserParameters = $PSBoundParameters | ConvertFrom-PSBoundParameters -ExcludeProperty Properties,IncludeMailboxProperties,IncludePaylocityEmployee
    $ADUserParameters |
    Add-Member -MemberType NoteProperty -Name Properties -Value $PropertiesIncludingThoseUsedByCustomProperites
    $ADUserParametersHashTable = $ADUserParameters | ConvertTo-HashTable

    $AddADUserCustomPropertiesParameters = $PSBoundParameters | 
    ConvertFrom-PSBoundParameters -Property IncludeMailboxProperties, IncludePaylocityEmployee -AsHashTable
    
    Get-ADUser @ADUserParametersHashTable | Add-ADUserCustomProperties -PassThru @AddADUserCustomPropertiesParameters
}

function Add-ADUserCustomProperties {
    param (
        [Parameter(ValueFromPipeline)]$ADUser,
        [Switch]$PassThru,
        [Switch]$IncludeMailboxProperties,
        [Switch]$IncludePaylocityEmployee
    )
    process {
        $ADUser | Add-Member -MemberType ScriptProperty -Name PasswordExpirationDate -PassThru -Force -Value {
            [datetime]::FromFileTime($This."msDS-UserPasswordExpiryTimeComputed")
        } |
        Add-Member -MemberType ScriptProperty -Name LastLogon -PassThru -Force -Value {
            [datetime]::FromFileTime($This."lastLogonTimestamp")
        } |
        Add-Member -MemberType ScriptProperty -Name ParentOrganizationalUnitDistinguishedName -Force -Value {
            ($This.DistinguishedName -split "," | Select-Object -Skip 1) -join ","
        }
        
        $ADUser |
        Where-Object { $IncludeMailboxProperties } |
        Add-Member -MemberType ScriptProperty -Name O365Mailbox -PassThru -Force -Value {
            if (Connect-EXOPSSessionWithinExchangeOnlineShell) {
                Get-Mailbox -Identity $This.UserPrincipalName
            } else {
                Import-TervisOffice365ExchangePSSession
                Get-O365Mailbox -Identity $This.UserPrincipalName
            }
        } |
        Add-Member -MemberType ScriptProperty -Name ExchangeRemoteMailbox -PassThru -Force -Value {
            Import-TervisExchangePSSession
            Get-ExchangeRemoteMailbox -Identity $This.UserPrincipalName
        } |
        Add-Member -MemberType ScriptProperty -Name ExchangeMailbox -Force -Value {
            Import-TervisExchangePSSession
            Get-ExchangeMailbox -Identity $This.UserPrincipalName
        } 
        
        $ADUser |
        Where-Object { $IncludePaylocityEmployee } |
        Add-Member -MemberType ScriptProperty -Name PaylocityEmployee -PassThru -Force -Value {
            Get-PaylocityEmployee -EmployeeID $This.EmployeeID
        } |
        Add-Member -Name PaylocityDepartmentCode -MemberType ScriptProperty -PassThru -Force -Value {
            $This.PaylocityEmployee.DepartmentCode
        } |
        Add-Member -Name PaylocityDepartmentName -MemberType ScriptProperty -PassThru -Force -Value {
            $This.PaylocityEmployee.DepartmentName
        } |
        Add-Member -Name PaylocityDepartmentNiceName -MemberType ScriptProperty -Force -Value {
            Get-DepartmentNiceName -PaylocityDepartmentName $this.PaylocityDepartmentName 
        } 

        if ($PassThru) { $ADUser }
    }
}

Function Find-TervisADUsersComputer {
    param (
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$SAMAccountName,
        [String[]]$Properties
    )
    process {
        $ComputerNameFilterString = "*" + $SAMAccountName + "*"
        if ($Properties) { 
            Get-ADComputer -Filter {Name -like $ComputerNameFilterString} -Properties $Properties
        } else {
            Get-ADComputer -Filter {Name -like $ComputerNameFilterString}
        }
    }
}

Function Test-DirectoriesSameSize {
    param (
        [parameter(Mandatory)]$ReferenceDirectory,
        [parameter(Mandatory)]$DifferenceDirectory
    )
    $TotalReferenceDirectory = Get-DirecotrySize -Directory $ReferenceDirectory
    $TotalDifferenceDirectory = Get-DirecotrySize -Directory $DifferenceDirectory

    $TotalReferenceDirectory -eq $TotalDifferenceDirectory
}

Function Get-DirecotrySize {
    param (
        [parameter(Mandatory)][System.IO.DirectoryInfo]$Directory
    )
    Get-ChildItem $Directory -Recurse -Force | 
    Measure-Object -property length -sum | 
    select -ExpandProperty Sum
}

function Remove-TervisADUsersComputer {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$SAMAccountName     
    )

    Write-Verbose "Removing user's Active Directory computer object(s)"
    Find-TervisADUsersComputer $SAMAccountName | Remove-ADComputer -Confirm
}

Function Get-LoggedOnUserName {
    Get-aduser $Env:USERNAME | select -ExpandProperty Name
}

Function Get-ADUserEmailAddressByName {
    param (
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$Name
    )
    Get-ADUser -Filter {Name -eq $Name} -Properties EmailAddress |
    Select -ExpandProperty EmailAddress
}

function Get-ADUserByEmployeeID {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$EmployeeID
    )
    Get-ADUser -Filter {Employeeid -eq $EmployeeID} -Properties EmployeeID
}

Function Test-TervisADComputerIsMac {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADComputer
    )

    $ADComputer.Name -match "-mac"
}

function Get-ADUserLogonFailEventInformation {
    $DomainControllers = (Get-ADGroupMember 'Domain Controllers').name | select -First 2
    Foreach ($DomainController in $DomainControllers) {
        Get-ADUser -Filter {enabled -eq $true} -SearchBase "OU=Departments,DC=tervis,DC=prv" -Properties BadLogonCount, LockedOut, LastLogonDate -Server $DomainController|
        Where BadLogonCount -gt 5 | 
        Select Name, SamAccountName, LockedOut, BadLogonCount, LastLogonDate |
        Add-Member -MemberType NoteProperty -Name DomainController -Value $DomainController -PassThru
    }
}

function Invoke-SwitchComputersCurrentDomainController {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]$ComputerName,
        [Parameter(Mandatory=$true)]$DomainControllerToSwitchTo,
        [switch]$RestartNIC
    )

    $CurrentDomain = $env:USERDOMAIN

    Write-Verbose "$ComputerName`: switching domain controller to $DomainControllerToSwitchTo"
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        nltest /SC_RESET:$Using:CurrentDomain\$Using:DomainControllerToSwitchTo
    }
    
    if ($RestartNIC) {
        Write-Verbose "Restarted connected NIC on $ComputerName"
        sleep -Seconds 5
        Restart-ConnectedNetworkInterface -ComputerName $ComputerName
    }
}

function Invoke-ADAzureSync {
    $Server = Get-AzureADConnectComputerName
    Invoke-Command -ComputerName $Server -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
}

Function Sync-ADDomainControllers {
    param (
        $SleepSeconds = 30,
        [Switch]$Blocking
    )
    $DC = Get-ADDomainController | Select -ExpandProperty HostName
    if ($Blocking) {
        Invoke-Command -computername $DC -ScriptBlock {repadmin /syncall /Aed}
    } else {
        Invoke-Command -ComputerName $DC -ScriptBlock {repadmin /syncall}        
    }
    Start-Sleep $SleepSeconds
}

function Sync-TervisADObjectToAllDomainControllers {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADObject
    )
    process {
        $DomainControllers = Get-ADDomainController -Filter * | select -ExpandProperty HostName
        foreach ($DomainController in $DomainControllers) {
            $ADObject | Sync-ADObject -Destination $DomainController
        }
    }
}

function Get-TervisADComputer {
    param (
        $Identity,
        $Path,
        $Filter,
        $Properties
    )    
    $AdditionalNeededProperties = "lastLogonTimestamp"
    Get-ADComputer @PSBoundParameters | Add-ADComputerCustomProperties
}

function Add-ADComputerCustomProperties {
    param (
        [Parameter(ValueFromPipeline)]$Input
    )

    $Input | 
    Add-Member -MemberType ScriptProperty -Name LastLogon -PassThru -Force -Value {
        [datetime]::FromFileTime($This.“lastLogonTimestamp”) 
    } | 
    Add-Member -MemberType AliasProperty -Name ComputerName -PassThru -Force -Value Name
}

function Remove-TervisADUser {
    [CMDLetBinding()]
    param(
        [Parameter(Mandatory)]$Identity,
        [Switch]$RemoveGroupos
    )
    $ADUser = Get-TervisADUser $Identity -Properties DistinguishedName,ProtectedFromAccidentalDeletion -IncludeMailboxProperties

    $Password = New-RandomPassword
    $SecurePassword = ConvertTo-SecureString $Password -asplaintext -force
    Set-ADAccountPassword -Identity $identity -NewPassword $SecurePassword

    if ($ADUser.ProtectedFromAccidentalDeletion) {
        Set-ADObject -Identity $ADUser.DistinguishedName -ProtectedFromAccidentalDeletion $false
    }

    If ($ADUser.O365Mailbox) {
        $OrganizationalUnit = Get-ADOrganizationalUnit -filter * | 
        where DistinguishedName -like "OU=Shared Mailbox,OU=Exchange,DC=*" | 
        select -ExpandProperty DistinguishedName
    } else {
        $OrganizationalUnit = Get-ADOrganizationalUnit -filter * | 
        where DistinguishedName -like "OU=Company- Disabled Accounts*" | 
        select -ExpandProperty DistinguishedName
    }

    Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $OrganizationalUnit
    $NewDistinguishedName = (Get-ADUser -Identity $Identity).DistinguishedName

    if ($RemoveGroups) {
        $Groups = Get-ADUser $Identity -Properties MemberOf | select -ExpandProperty MemberOf
        foreach ($Group in $Groups) {
            Remove-ADGroupMember -Identity $Group -Members $Identity -Confirm:$false
        }
    }

    Disable-ADAccount $Identity
    Set-ADAccountExpiration $Identity -DateTime (get-date)
}

function New-RandomPassword {
    #https://msdn.microsoft.com/en-us/library/system.web.security.membership.generatepassword(v=vs.110).aspx
    Add-Type -AssemblyName System.Web
    [System.Web.Security.Membership]::GeneratePassword(120,10)
}

function Remove-TervisADComputerObject {
    param(
        [parameter(Mandatory, ValueFromPipeline)]$ComputerName
    )
    process {
        Get-ADComputer -Identity $ComputerName | 
        Remove-ADObject -Recursive -Confirm
    }
}

function Disable-TervisADComputerInactive {
    $AdComputers = Get-TervisADComputerInactive -ThresholdType Disable
    if ($AdComputers) {
        Send-TervisADObjectActionEmail -ADObjects $AdComputers -Action disable -Property Name,LastLogon,Created,Operatingsystem
        $AdComputers | Disable-ADAccount -Confirm:$false
    }
}

function Get-TervisADComputerInactive {
    param (
        [Parameter(Mandatory)][ValidateSet("Disable","Remove")]$ThresholdType
    )
    $ADcomputersToEvaluate = Get-TervisADComputer -Filter * -Properties LastLogonTimestamp,created,enabled,operatingsystem,ProtectedFromAccidentalDeletion,PasswordLastSet

    $ADComputers = if ($ThresholdType -eq "Disable") {
        $ADcomputersToEvaluate |
        Invoke-FilterADObject -LastLogonOlderThanDays 60 -CreatedOlderThanDays 60 -PasswordLastSetOlderThanDays 60 |
        Where-Object Enabled -eq $true 
    } elseif ($ThresholdType -eq "Remove") {        
        $ADcomputersToEvaluate |
        Invoke-FilterADObject -LastLogonOlderThanDays 190 -CreatedOlderThanDays 60 -PasswordLastSetOlderThanDays 190
    }

    $ADComputers |
    Where-Object Name -notlike "TP9*" |
    Where-Object OperatingSystem -notlike "Windows Server*" |
    Where-Object OperatingSystem -NotIn ("RHEL","Mac OS X",$null)
}

function Remove-TervisADComputerInactive {
    $ADComputers = Get-TervisADComputerInactive -ThresholdType Remove    
    Send-TervisADObjectActionEmail -ADObjects $ADComputers -Action remove -Property Name,LastLogon,Created,Operatingsystem
    $ADComputers | Remove-TervisADObject
    $ADComputers | Remove-ADObject -Confirm:$false -Recursive
    $ADComputers | Remove-TervisDNSRecord    
}

function Remove-TervisADObject {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ADObject
    )
    process {
        Push-Location AD:
        $AdObjectACL = Get-Acl ($ADObject).DistinguishedName
        foreach ($AccessRule in $AdObjectACL.Access) {
            if ($AccessRule.IdentityReference.Value -eq 'Everyone' -and $AccessRule.AccessControlType -eq 'Deny' -and $AccessRule.ActiveDirectoryRights -match 'Delete') {
                $AdObjectACL.RemoveAccessRule($AccessRule) | Out-Null
            }
        }
        Pop-Location
        if ($ADObject.ProtectedFromAccidentalDeletion) {
            Set-ADObject -Identity ($ADObject).DistinguishedName -ProtectedFromAccidentalDeletion $false
        }
        
        $ADObject | Remove-ADObject -Confirm:$false -Recursive
    }
}

function Invoke-FilterADObject {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADObject,
        $LastLogonOlderThanDays,
        $CreatedOlderThanDays,
        $PasswordLastSetOlderThanDays
    )
    process {
        $ADObject |
        Where-Object { -not $LastLogonOlderThanDays -or $_.LastLogon -lt (Get-Date).AddDays(-$LastLogonOlderThanDays) } |
        Where-Object { -not $CreatedOlderThanDays -or $_.Created -lt (Get-Date).AddDays(-$CreatedOlderThanDays) } |
        Where-Object { -not $PasswordLastSetOlderThanDays -or $_.PasswordLastSet -lt (Get-Date).AddDays(-$PasswordLastSetOlderThanDays) }
    }
}


function Get-TervisADOrganizationalUnitThatholdsServiceAccounts {
    Get-ADOrganizationalUnit -Filter {Name -eq "Accounts - Service"}
}

filter Select-ADUserUsedAsEndUser {
    $_ |
    Where-Object DistinguishedName -notmatch "CN=Microsoft Exchange System Objects,DC=" |
    Where-Object DistinguishedName -notmatch "OU=Exchange,DC=" |
    Where-Object DistinguishedName -NotMatch $ServiceAccountsOU
}

function Get-ADObjectPropertiesForValidateSet {

    $Properties = $ADuser | gm | Where membertype -eq property | select -ExpandProperty name | sort
    '"' + ($Properties -join "`",`r`n`"") + '"'
}

function Get-TervisADUserInactive {
    param (
        [Parameter(Mandatory)][ValidateSet("Disable","Remove")]$ThresholdType
    )
    $ServiceAccountsOU = Get-ADOrganizationalUnit -Filter {Name -eq "Accounts - Service"}
    $InactivityExceptionsOU = Get-ADOrganizationalUnit -Filter {Name -eq "Inactivity Exceptions"}
    $ADUsersToEvaluate = Get-TervisADUser -Filter *
    $ADUsersInactive = @()

    if ($ThresholdType -eq "Disable") {
        $ADUsersInactive += $ADUsersToEvaluate |
        Where-Object Enabled -eq $true |
        Invoke-FilterADObject -LastLogonOlderThanDays 30 -CreatedOlderThanDays 60 -PasswordLastSetOlderThanDays 90 |
        Where-Object DistinguishedName -notmatch "CN=Microsoft Exchange System Objects,DC=" |
        Where-Object DistinguishedName -notmatch "OU=Exchange,DC=" |
        Where-Object DistinguishedName -NotMatch $ServiceAccountsOU
        
        $ADUsersInactive += Get-TervisADUser -Filter * -SearchBase $ServiceAccountsOU -SearchScope OneLevel |
        Where-Object Enabled -eq $true |
        Where-Object ParentOrganizationalUnitDistinguishedName -eq $ServiceAccountsOU.DistinguishedName |
        Invoke-FilterADObject -LastLogonOlderThanDays 180 -CreatedOlderThanDays 60 -PasswordLastSetOlderThanDays 90         
            
        $ADUsersInactive += $ADUsersToEvaluate | 
        Where-Object Enabled -eq $true |
        Where-Object ParentOrganizationalUnitDistinguishedName -eq $InactivityExceptionsOU.DistinguishedName |
        Invoke-FilterADObject -LastLogonOlderThanDays 365 -PasswordLastSetOlderThanDays 365

    } elseif ($ThresholdType -eq "Remove") {        
        $ADUsersInactive += $ADUsersToEvaluate |
        Invoke-FilterADObject -LastLogonOlderThanDays 190 -CreatedOlderThanDays 90 -PasswordLastSetOlderThanDays 90 |
        Where-Object DistinguishedName -notmatch "CN=Microsoft Exchange System Objects," |
        Where-Object DistinguishedName -notmatch "OU=Exchange,DC=" |
        Where-Object DistinguishedName -notmatch  $ServiceAccountsOU |
        Where-Object Name -NotIn ("krbtgt","Guest","DefaultAccount")

        $ADUsersInactive += $ADUsersToEvaluate | 
        Invoke-FilterADObject -LastLogonOlderThanDays 365 -CreatedOlderThanDays 90 |
        Where-Object DistinguishedName -match $ServiceAccountsOU |
        Where-Object DistinguishedName -notmatch $InactivityExceptionsOU

        $ADUsersInactive += $ADUsersToEvaluate | 
        Invoke-FilterADObject -LastLogonOlderThanDays 425 -PasswordLastSetOlderThanDays 425 |
        Where-Object Enabled -eq $false |
        Where-Object DistinguishedName -match $InactivityExceptionsOU
    }
    $ADUsersInactive
}

function Send-TervisADObjectActionEmail {
    param (
        $ADObjects,
        [ValidateSet("disable","remove")]$Action,
        $Property
    )
    $Body = @"
<html><body>
<h2>The AD objects below are being $($Action)d.</h2>
$(
    $ADObjects |
    Sort-Object -Property Name | 
    Select-Object -Property $Property |
    ConvertTo-Html -As Table -Fragment
)
</body></html>
"@
    $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
    
    $ExportPath = $env:TEMP + "\ADObjects.xml"
    $ADObjects | Export-Clixml -Path $ExportPath
    $MailAttachment = $ExportPath

    Send-TervisMailMessage -To $To -From "$($Action)InactiveADObjects@tervis.com" -Subject "Inactive AD objects to be $($Action)d" -BodyAsHTML -Body $Body -Attachments $MailAttachment

    Remove-Item -Path $MailAttachment -Force -Confirm:$false
}

function Disable-TervisADUserInactive {
    $ADObjects = Get-TervisADUserInactive -ThresholdType Disable
    Send-TervisADObjectActionEmail -ADObjects $ADObjects -Action disable -Property Name, SAMAccountName, Enabled, LastLogon, Created, PasswordLastSet
    $ADObjects | Disable-ADAccount -Confirm:$false
}

function Remove-TervisADUserInactive {
    $AdUsersToDelete = Get-TervisADUserInactive -ThresholdType Remove
    Send-TervisADObjectActionEmail -ADObjects $AdUsersToDelete -Action remove -Property Name, SAMAccountName, Enabled, LastLogon, Created, PasswordLastSet
    
    foreach ($AdUserToDelete in $AdUsersToDelete) {
        Remove-TervisADObject -ADObject $AdUserToDelete

        if (($AdUserToDelete).DistinguishedName -match "OU=Departments,DC=") {
            Remove-TervisPerson -Identity ($AdUserToDelete).SamAccountName -NoUserReceivesData
        } else {
            Remove-ADObject ($AdUserToDelete).DistinguishedName -Confirm:$false -Recursive
        }
    }
}

function Send-TervisInactivityNotification {
    $AdUsersForNotification = @()
    $AdUsersForNotification += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,created,enabled,PasswordLastSet,Manager | 
    Invoke-FilterADObject -LastLogonOlderThanDays 351 -PasswordLastSetOlderThanDays 351 |
        where {
            $_.DistinguishedName -match "OU=Inactivity Exceptions,OU=Accounts - Service,DC=" -and
            $_.Enabled -eq $true
        }
    foreach ($AdUserForNotification in $AdUsersForNotification) {
        if (($AdUserForNotification).Manager) {
            $ManagerEmail = Get-ADUser ($AdUserForNotification).Manager -Properties EmailAddress -ErrorAction SilentlyContinue | Select -ExpandProperty EmailAddress
        }
        If ($ManagerEmail) {
            $To = $ManagerEmail
        } else {
            $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        }
        if (($AdUserForNotification).PasswordLastSet -gt ($AdUserForNotification).LastLogon) {
            $DaysUntilDisabled = (Get-Date).AddDays(-365) - ($AdUserForNotification).PasswordLastSet
        } else {
            $DaysUntilDisabled = (Get-Date).AddDays(-365) - ($AdUserForNotification).LastLogon | Select -ExpandProperty Days
        }
        $Body = "The following user in the Inactive Exceptions OU will be disabled in $DaysUntilDisabled days if the password is not reset. `n" 
        $Body += "User Name `t Tervis Last Logon `t Date Created 't Password Last Set `n"
        $Body += ($AdUserForNotification).name + "`t" + ($AdUserForNotification).LastLogon + "`t" + ($AdUserForNotification).created + "`t" +  ($AdUserForNotification).PasswordLastSet + "`n"
        Send-TervisMailMessage -To $To -From ADUserInactivityNotification@tervis.com -Subject 'Service Account with Expection to be Disabled' -Body $Body
    }
}

function Get-ADUserPhoto {
    param (
        [Parameter(Mandatory)]$Identity,
        $Path = $Home
    )
    $ADUser = get-aduser -Identity $Identity -Properties thumbnailphoto
    [System.Io.File]::WriteAllBytes("$Path\$Identity.jpg", $ADUser.Thumbnailphoto)
}

function Invoke-SyncGravatarPhotosToADUsersInAD {
    $ADUsers = Get-ADUser -Filter {Enabled -eq $true} -Properties ThumbnailPhoto,EmailAddress |
    where {$_.EmailAddress}
    $ADUsersWithGravatarAvatars = $ADUsers |
    Add-Member -MemberType ScriptProperty -Name GravatarAvatarURL -Force -PassThru -Value {
        Get-GravatarAvatarURL -EmailAddress $This.EmailAddress -Size 175 -DefaultType 404
    } |
    Add-Member -MemberType ScriptProperty -Name GravatarAvatarExists -Force -PassThru -Value {
        $Respose = Invoke-WebRequest -Method Head -UseBasicParsing -Uri $This.GravatarAvatarURL
        if ($Respose) { $true } else { $false }
    } |
    Where { $_.GravatarAvatarExists }
    $ADUsersWithGravatarAvatars | Sync-GravatarToADUserPhoto
}

function Sync-GravatarToADUserPhoto {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADUser
    )
    process {
        $Response = Invoke-WebRequest -UseBasicParsing -Uri $ADUser.GravatarAvatarURL
        if ($ADUser.thumbnailphoto.Length -ne $Response.RawContentLength) {
            Set-ADUser -Identity $ADuser.SAMAccountName -Replace @{thumbnailPhoto=$Response.Content}
        }
    }
}

function Install-InvokeSyncGravatarPhotosToADUsersInAD {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = Get-PasswordstateCredential -PasswordID 259
    }
    process {
        Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask "C:\Scripts\" `
            -Credential $ScheduledTaskCredential `
            -FunctionName "Invoke-SyncGravatarPhotosToADUsersInAD" `
            -RepetitionInterval EverWorkdayOnceAtTheStartOfTheDay `
            -ComputerName $ComputerName
    }
}

function Invoke-SyncADUserThumbnailPhotoToOffice365 {
    $ADUsers = Get-ADUser -Filter {
        ThumbnailPhoto -like "*" -and
        Enabled -eq $true
    } -Properties ThumbnailPhoto,EmailAddress
   
    Import-TervisOffice365ExchangePSSession
    $Mailboxes = Get-O365Mailbox
    
    $ADUsersWithMailboxes = $ADUsers |
    where UserPrincipalName -In $Mailboxes.UserPrincipalName
    
    $ADUsersWithMailboxes |
    Add-Member -MemberType ScriptProperty -Name UserPhotoLength -Force -Value {
        $Response = Get-O365UserPhoto -Identity $This.UserPrincipalName
        $Response.PictureData.Length
    }

    $ADUsersWithSmallEnoughThumbnailPhotos = $ADUsersWithMailboxes |
    Where { $_.ThumbnailPhoto.length -lt 21000}

    $ADUsersWithThumbnailPhotosTooLarge = $ADUsersWithMailboxes |
    Where { $_.ThumbnailPhoto.length -ge 21000}

    foreach ($ADUser in $ADUsersWithSmallEnoughThumbnailPhotos) {
        Write-Verbose "UserPrincipalName:$($ADUser.UserPrincipalName) ThumbnailLength:$($ADUser.ThumbnailPhoto.length)"
        Set-O365UserPhoto -Identity $ADUser.UserPrincipalName -PictureData $ADUser.ThumbnailPhoto -Confirm:$False
    }
}

function Install-DisableInactiveADComputersScheduledTask {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = '-Command Disable-InactiveADComputers -NoProfile'
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Disable-InactiveADComputers -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Disable-InactiveADComputers -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt3am -ComputerName $ComputerName
        }
    }
}

function Install-DisableInactiveADUsersScheduledTask {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = '-Command Disable-InactiveADUsers -NoProfile'
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Disable-InactiveADUsers -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Disable-InactiveADUsers -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt3am -ComputerName $ComputerName
        }
    }
}

function Install-RemoveInactiveADComputersScheduledTask {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = '-Command Remove-InactiveADComputers -NoProfile'
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Remove-InactiveADComputers -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Remove-InactiveADComputers -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt3am -ComputerName $ComputerName
        }
    }
}

function Install-RemoveInactiveADUsersScheduledTask {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = '-Command Remove-InactiveADUsers -NoProfile'
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Remove-InactiveADUsers -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Remove-InactiveADUsers -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt3am -ComputerName $ComputerName
        }
    }
}

function Install-SendTervisInactivityNotification {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = '-Command Send-TervisInactivityNotification -NoProfile'
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Send-TervisInactivityNotification -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Send-TervisInactivityNotification -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt3am -ComputerName $ComputerName
        }
    }
}

functin Invoke-TervisActiveDirectoryCleanup {    
    Disable-TervisADUserInactive
    Disable-TervisADComputerInactive
    Remove-TervisADUserInactive
    Remove-TervisADComputerInactive
    #Send-TervisInactivityNotification
}

function Install-TervisActiveDirectoryCleanup {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    
    $InstallPowerShellApplicationParameters = @{
        ModuleName = "TervisSOAMonitoringApplication"
        DependentTervisModuleNames = "TervisMailMessage","TervisOracleSOASuite"
        ScheduledScriptCommandsString = "Invoke-TervisActiveDirectoryCleanup"
        ScheduledTasksCredential = (Get-PasswordstatePassword -ID 259 -AsCredential)
        ScheduledTaskName = "Invoke-TervisOracleSOAJobMonitoring"
        RepetitionIntervalName = "EveryDayEvery15Minutes"
    }

    Install-PowerShellApplication -ComputerName $ComputerName @InstallPowerShellApplicationParameters
}

function Get-ADObjectParentContainer {
    param(
        [Parameter(Mandatory,ValueFromPipeline)]$ObjectPath
    )
    ($ObjectPath.split(",") | select -skip 1 ) -join ","
}

function Invoke-TervisDomainControllerProvision {
    param (
        $EnvironmentName
    )
    Invoke-ApplicationProvision -ApplicationName DomainController -EnvironmentName $EnvironmentName
    $Nodes = Get-TervisApplicationNode -ApplicationName DomainController -EnvironmentName $EnvironmentName
} 

function Get-AzureADConnectComputerName {
    Get-ADComputer -Filter {description -eq 'Azure AD Connect'} | Select -ExpandProperty Name
}


function Get-AvailableSAMAccountName {
    param(
        [parameter(mandatory)]$GivenName,
        [parameter(mandatory)]$Surname
    )

    [string]$FirstInitialSurname = $GivenName[0] + $Surname
    [string]$GivenNameLastInitial = $GivenName + $Surname[0]

    $UserName= if (-not (Get-ADUser -filter {sAMAccountName -eq $FirstInitialSurname})) {
        $FirstInitialSurname
    } elseif (-not (Get-ADUser -filter {sAMAccountName -eq $GivenNameLastInitial})) {
        $GivenNameLastInitial
    } else {
        Throw "Neither $FirstInitialSurname or $GivenNameLastInitial are avaialble SAMAccountNames in Active Directory"
    }

    $UserName.ToLower()
}

function Copy-ADUserGroupMembership {
    param (
        $Identity,
        $DestinationIdentity
    )
    $Groups = Get-ADUser $Identity -Properties MemberOf | Select -ExpandProperty MemberOf
    Foreach ($Group in $Groups) {
        Add-ADGroupMember -Identity $group -Members $DestinationIdentity
    }
}

function Get-ADUserOU {
    param (
        [Parameter(Mandatory)]$SAMAccountName
    )
    $ADUser = Get-ADUser $SAMAccountName
    ($Aduser.DistinguishedName -split "," | select -Skip 1 ) -join ","
}

function Move-UserToCiscoVPNCertificateADGroup {
    param ($SamAccountName)
    Remove-ADGroupMember -Identity CiscoVPN -Members $SamAccountName -Confirm:$false
    Add-ADGroupMember -Identity CiscoVPN-Certificate -Members $SamAccountName
}

function Move-UserToCiscoVPNADGroup {
    param ($SamAccountName)
    Remove-ADGroupMember -Identity CiscoVPN-Certificate -Members $SamAccountName -Confirm:$false
    Add-ADGroupMember -Identity CiscoVPN -Members $SamAccountName
}

function Invoke-GPUpdateForOU {
    param(
        [Parameter(Mandatory)]$SearchBase
    )
    get-adcomputer -SearchBase $SearchBase -Filter * |
    ForEach-Object {
        if (test-netconnection -computername $_.DNSHostNAme) {
        invoke-gpupdate -computer $_.DNSHostNAme -asjob
        }
    }
}

function Remove-ADUserProxyAddress {
    param (
        [Parameter(Mandatory)]$Identity,
        [Parameter(Mandatory)]$ProxyAddress
    )
    Get-ADUser -Identity $Identity -Properties ProxyAddresses | 
    Set-ADUser -Remove @{proxyaddresses=$ProxyAddress}
}

function Add-ADUserProxyAddress {
    param (
        [Parameter(Mandatory)]$Identity,
        [Parameter(Mandatory)]$ProxyAddress
    )
    Get-ADUser -Identity $Identity -Properties ProxyAddresses | 
    Set-ADUser -Add @{proxyaddresses=$ProxyAddress}
}

function Set-ADGroupManagedByAttribute {
    param(
        [parameter(Mandatory)]$ADGroup,
        [parameter(Mandatory)]$GroupManager
    )
    Get-ADGroup $ADGroup | Set-ADGroup -ManagedBy $GroupManager
}

function New-ADUserFilterScriptBlockUnfinished {
    param (
        [Parameter(Position=0)][ValidateSet(
            "AccountExpirationDate",
            "accountExpires",
            "AccountLockoutTime",
            "AccountNotDelegated",
            "adminCount",
            "AllowReversiblePasswordEncryption",
            "AuthenticationPolicy",
            "AuthenticationPolicySilo",
            "BadLogonCount",
            "badPasswordTime",
            "badPwdCount",
            "CannotChangePassword",
            "CanonicalName",
            "Certificates",
            "City",
            "CN",
            "codePage",
            "Company",
            "CompoundIdentitySupported",
            "Country",
            "countryCode",
            "Created",
            "createTimeStamp",
            "Deleted",
            "Department",
            "Description",
            "directReports",
            "DisplayName",
            "DistinguishedName",
            "Division",
            "DoesNotRequirePreAuth",
            "dSCorePropagationData",
            "EmailAddress",
            "EmployeeID",
            "EmployeeNumber",
            "Enabled",
            "facsimileTelephoneNumber",
            "Fax",
            "garbageCollPeriod",
            "GivenName",
            "HomeDirectory",
            "HomedirRequired",
            "HomeDrive",
            "HomePage",
            "HomePhone",
            "Initials",
            "instanceType",
            "isDeleted",
            "KerberosEncryptionType",
            "LastBadPasswordAttempt",
            "LastKnownParent",
            "lastLogoff",
            "lastLogon",
            "LastLogonDate",
            "lastLogonTimestamp",
            "legacyExchangeDN",
            "LockedOut",
            "lockoutTime",
            "logonCount",
            "LogonWorkstations",
            "mail",
            "mailNickname",
            "managedObjects",
            "Manager",
            "MemberOf",
            "MNSLogonAccount",
            "MobilePhone",
            "Modified",
            "modifyTimeStamp",
            "mS-DS-ConsistencyGuid",
            "msDS-ExternalDirectoryObjectId",
            "msDS-User-Account-Control-Computed",
            "msExchALObjectVersion",
            "msExchArchiveGUID",
            "msExchArchiveName",
            "msExchArchiveStatus",
            "msExchCoManagedObjectsBL",
            "msExchELCMailboxFlags",
            "msExchMailboxGuid",
            "msExchMailboxTemplateLink",
            "msExchMobileAllowedDeviceIDs",
            "msExchMobileMailboxFlags",
            "msExchPoliciesIncluded",
            "msExchRecipientDisplayType",
            "msExchRecipientTypeDetails",
            "msExchRemoteRecipientType",
            "msExchShadowProxyAddresses",
            "msExchTextMessagingState",
            "msExchUMDtmfMap",
            "msExchUserAccountControl",
            "msExchUserHoldPolicies",
            "msExchVersion",
            "msExchWhenMailboxCreated",
            "mSMQDigests",
            "mSMQSignCertificates",
            "msTSExpireDate",
            "msTSLicenseVersion",
            "msTSLicenseVersion2",
            "msTSLicenseVersion3",
            "msTSManagingLS",
            "Name",
            "nTSecurityDescriptor",
            "ObjectCategory",
            "ObjectClass",
            "ObjectGUID",
            "objectSid",
            "Office",
            "OfficePhone",
            "Organization",
            "OtherName",
            "PasswordExpired",
            "PasswordLastSet",
            "PasswordNeverExpires",
            "PasswordNotRequired",
            "POBox",
            "PostalCode",
            "PrimaryGroup",
            "primaryGroupID",
            "PrincipalsAllowedToDelegateToAccount",
            "ProfilePath",
            "ProtectedFromAccidentalDeletion",
            "proxyAddresses",
            "publicDelegatesBL",
            "pwdLastSet",
            "SamAccountName",
            "sAMAccountType",
            "ScriptPath",
            "sDRightsEffective",
            "ServicePrincipalNames",
            "showInAddressBook",
            "SID",
            "SIDHistory",
            "SmartcardLogonRequired",
            "sn",
            "State",
            "StreetAddress",
            "Surname",
            "targetAddress",
            "telephoneNumber",
            "terminalServer",
            "textEncodedORAddress",
            "thumbnailPhoto",
            "Title",
            "TrustedForDelegation",
            "TrustedToAuthForDelegation",
            "UseDESKeyOnly",
            "userAccountControl",
            "userCertificate",
            "UserPrincipalName",
            "uSNChanged",
            "uSNCreated",
            "whenChanged",
            "whenCreated"
        )]$Property,
        [Parameter(Position=1)]$Value,
        [Parameter(Mandatory,ParameterSetName="eq")][Switch]$eq,
        [Parameter(Mandatory,ParameterSetName="le")][Switch]$le,
        [Parameter(Mandatory,ParameterSetName="ge")][Switch]$ge,
        [Parameter(Mandatory,ParameterSetName="ne")][Switch]$ne,
        [Parameter(Mandatory,ParameterSetName="lt")][Switch]$lt,
        [Parameter(Mandatory,ParameterSetName="gt")][Switch]$gt,
        [Parameter(Mandatory,ParameterSetName="approx")][Switch]$approx,
        [Parameter(Mandatory,ParameterSetName="bor")][Switch]$bor,
        [Parameter(Mandatory,ParameterSetName="band")][Switch]$band,
        [Parameter(Mandatory,ParameterSetName="recursivematch")][Switch]$recursivematch,
        [Parameter(Mandatory,ParameterSetName="like")][Switch]$like,
        [Parameter(Mandatory,ParameterSetName="notlike")][Switch]$notlike,
        [Parameter(ValueFromPipeline)]$PipelineWhere
    )
    #$PSBoundParameters | ConvertFrom-PSBoundParameters
    $Operator = $PSBoundParameters.Keys | Where-Object {$_ -notin ("Property","Value")}
    $FilterString = "$Property -$Operator $Value"
    $ScriptBlockString = (
        if($PipelineWhere) {
            $PipelineWhere.ToString() + "-and "
        }
    ) + 

    [scriptblock]::Create($ScriptBlockString)
}

function New-ADUserTypeObjects {
    $ADUserTypes = [PSCustomObject]@{
        Name = "EndUser"
        FilterScriptBlock = {
            $_ |
            Where-Object DistinguishedName -notmatch "CN=Microsoft Exchange System Objects,DC=" |
            Where-Object DistinguishedName -notmatch "OU=Exchange,DC=" |
            Where-Object DistinguishedName -NotMatch $ServiceAccountsOU
        }
    },
    [PSCustomObject]@{
        Name = "ServiceAccount"
        FilterScriptBlock = {
            $_ |
            Where-Object DistinguishedName -eq $ServiceAccountsOU
        }
    },
    [PSCustomObject]@{
        Name = "ServiceAccountInactivityExceptions"
        FilterScriptBlock = {
            $_ |
            Where-Object DistinguishedName - $InactivityExceptionsOU
        }
    }

    $InactiveCriteriaTypes = [PSCustomObject]@{
        Name = "Disable"
        ADUserType = [PSCustomObject]@{
            Name = "EndUser"
            LastLogonOlderThanDays = 30
            CreatedOlderThanDays = 60
            PasswordLastSetOlderThanDays = 90
        },
        [PSCustomObject]@{
            Name = "ServiceAccount"
            LastLogonOlderThanDays = 180
            CreatedOlderThanDays = 60
            PasswordLastSetOlderThanDays = 90
            FilterScriptBlock = {
                $_ |
                Where-Object DistinguishedName -notmatch $InactivityExceptionsOU
            }
        },
        [PSCustomObject]@{
            Name = "ServiceAccountInactivityExceptions"
            LastLogonOlderThanDays = 365
            PasswordLastSetOlderThanDays = 365
        }
    },
    [PSCustomObject]@{
        Name = "Remove"
            ADUserType = [PSCustomObject]@{
            Name = "EndUser"
            LastLogonOlderThanDays = 190
            CreatedOlderThanDays = 90
            PasswordLastSetOlderThanDays = 90
        },
        [PSCustomObject]@{
            Name = "ServiceAccount"
            LastLogonOlderThanDays = 365
            CreatedOlderThanDays = 90
        },
        [PSCustomObject]@{
            Name = "InactivityExceptions"
            LastLogonOlderThanDays = 425
            PasswordLastSetOlderThanDays = 425
        }
    }
}