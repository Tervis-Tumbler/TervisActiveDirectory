﻿#Requires -Modules TervisDNS, ActiveDirectory, TervisMailMessage

function Get-TervisADUser {
    param (
        $Identity,
        $Path,
        $Filter,
        $Properties
    )
    
    $AdditionalNeededProperties = "msDS-UserPasswordExpiryTimeComputed,lastLogonTimestamp"
    Get-ADUser @PSBoundParameters | Add-ADUserCustomProperties -PassThru
}

function Add-ADUserCustomProperties {
    param (
        [Parameter(ValueFromPipeline)]$ADUser,
        [Switch]$PassThru
    )
    process {
        $ADUser | Add-Member -MemberType ScriptProperty -Name PasswordExpirationDate -PassThru -Force -Value {
            [datetime]::FromFileTime($This.“msDS-UserPasswordExpiryTimeComputed”)
        } |
        Add-Member -MemberType ScriptProperty -Name TervisLastLogon -PassThru -Force -Value {
            [datetime]::FromFileTime($This."lastLogonTimestamp")
        } |
        Add-Member -MemberType ScriptProperty -Name O365Mailbox -PassThru -Force -Value {
            Import-TervisOffice365ExchangePSSession
            Get-O365Mailbox -Identity $This.UserPrincipalName
        } |
        Add-Member -MemberType ScriptProperty -Name ExchangeMailbox -PassThru:$PassThru -Force -Value {
            Import-TervisExchangePSSession
            Get-ExchangeMailbox -Identity $This.UserPrincipalName
        }
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

function Invoke-TervisADUserShareHomeDirectoryPathAndClearHomeDirectoryProperty {
    param (
        [parameter(Mandatory)]$Identity,       
        [Parameter(Mandatory)]$IdentityOfUserToAccessHomeDirectoryFiles
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties HomeDirectory
    $Path = $ADUser.HomeDirectory 
    $ACL = Get-Acl -Path $Path
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($IdentityOfUserToAccessHomeDirectoryFiles, "Modify","ContainerInherit,ObjectInherit", "None", "Allow")
    $ACL.SetAccessRule($AccessRule)
    Set-Acl -path $Path -AclObject $Acl

    $ADUserToReceiveFiles = Get-ADUser -Identity $IdentityOfUserToAccessHomeDirectoryFiles -Properties EmailAddress
    if ($ADUserToReceiveFiles.EmailAddress) {
        $To = $ADUserToReceiveFiles.EmailAddress
        $Subject = "$($ADUser.SAMAccountName)'s home directory files have been shared with you"
        $Body = @"
$($ADUser.Name)'s home directory files have been shared with you.
You can access the files by going to $($ADUser.HomeDirectory). 
This was done as a part of the termination process for $($ADUser.Name).

If you believe you received this email incorrectly, please contact the Help Desk at x2248.
"@
        Send-TervisMailMessage -from "HelpDeskTeam@Tervis.com" -To $To -Subject $Subject -Body $Body
    }
    $ADUser | Set-ADUser -Clear HomeDirectory
}

function Remove-TervisADUserHomeDirectory {
    param (
        [parameter(Mandatory)]$Identity       
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties HomeDirectory
    Remove-Item -Path $ADUser.HomeDirectory -Confirm:$false -Recurse -Force
    $ADUser | Set-ADUser -Clear HomeDirectory
}

Function Invoke-TervisADUserHomeDirectoryDecomission {
    param (
        [parameter(Mandatory)]$Identity,       
        [Parameter(Mandatory, ParameterSetName="AnotherUserReceivesFiles")]$IdentityOfUserToReceiveHomeDirectoryFiles,                
        [Parameter(Mandatory, ParameterSetName="DeleteUsersFiles")][Switch]$DeleteFilesWithoutMovingThem
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties HomeDirectory

    if (-not $ADUser.HomeDirectory) {
        Throw "$($ADUser.SamAccountName)'s home directory not defined"
    }

    if ($(Test-Path $ADUser.HomeDirectory) -eq $false) {
        Throw "$($ADUser.SamaccountName)'s home directory $($ADUser.HomeDirectory) doesn't exist"
    }

    if ($ADUser.HomeDirectory -notmatch $ADUser.SamAccountName) {
        Throw "$($ADUser.HomeDirectory) doesn't have $($ADUser.SamAccountName) in it"
    }

    if ($DeleteFilesWithoutMovingThem) {
        Remove-TervisADUserHomeDirectory -Identity $Identity
    } else {
        $ADUserToReceiveFiles = Get-ADUser -Identity $IdentityOfUserToReceiveHomeDirectoryFiles -Properties EmailAddress
        
        if (-not $ADUserToReceiveFiles) { "Running Get-ADUser for the identity $IdentityOfUserToReceiveHomeDirectoryFiles didn't find an Active Directory user" }

        $ADUserToReceiveFilesComputer = $ADUserToReceiveFiles | Find-TervisADUsersComputer
        if (-not $ADUserToReceiveFilesComputer ) { Throw "Couldn't find an ADComputer with $($ADUserToReceiveFiles.SamAccountName) in the computer's name. If you know the name of the computer, run: `nInvoke-CopyADUsersHomeDirectoryToADUserToReceiveFilesComputer -Identity $Identity -IdentityOfUserToReceiveHomeDirectoryFiles $IdentityOfUserToReceiveHomeDirectoryFiles -ADUserToReceiveFilesComputerName <COMPUTERNAME>`nwhere <COMPUTERNAME> is replaced by the name of the destination computer. Once completed, rerun Remove-TervisUser." }
        if ($ADUserToReceiveFilesComputer.count -gt 1) { 
            Throw "We found more than one AD computer for $($ADUserToReceiveFiles.SamAccountName). Run: `nFind-TervisADUsersComputer -SamAccountName $($ADUserToReceiveFiles.SamAccountName) -Properties LastLogonDate `nto see the computers. Once the correct computer has been found, run the following command: `nInvoke-CopyADUsersHomeDirectoryToADUserToReceiveFilesComputer -Identity $Identity -IdentityOfUserToReceiveHomeDirectoryFiles $IdentityOfUserToReceiveHomeDirectoryFiles -ADUserToReceiveFilesComputerName <COMPUTERNAME>`nwhere <COMPUTERNAME> is replaced by the name of the destination computer. Once completed, rerun Remove-TervisUser."
        }

        if ($ADUserToReceiveFilesComputer | Test-TervisADComputerIsMac) {
            Throw "ADUserToReceiveFilesComputer: $($ADUserToReceiveFilesComputer.Name) is a Mac, cannot copy the files automatically"            
        }

        Invoke-CopyADUsersHomeDirectoryToADUserToReceiveFilesComputer -ErrorAction Stop -Identity $Identity -IdentityOfUserToReceiveHomeDirectoryFiles $IdentityOfUserToReceiveHomeDirectoryFiles -ADUserToReceiveFilesComputerName $ADUserToReceiveFilesComputer.Name
    }
}

Function Invoke-CopyADUsersHomeDirectoryToADUserToReceiveFilesComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]$Identity,       
        [Parameter(Mandatory)]$IdentityOfUserToReceiveHomeDirectoryFiles,
        [Parameter(Mandatory)]$ADUserToReceiveFilesComputerName
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties HomeDirectory
    $ADUserToReceiveFiles = Get-ADUser -Identity $IdentityOfUserToReceiveHomeDirectoryFiles -Properties EmailAddress

    $PathToADUserToReceiveFilesDesktop = "\\$ADUserToReceiveFilesComputerName\C$\Users\$($ADUserToReceiveFiles.SAMAccountName)\Desktop"

    if ($(Test-Path $PathToADUserToReceiveFilesDesktop) -eq $false) {
        Throw "$PathToADUserToReceiveFilesDesktop doesn't exist so we cannot copy the user's home directory files over"
    }

    $HomeDirectory = Get-Item $ADUser.HomeDirectory
    $PathToFolderToContainUsersCopiedHomeDirectory = "$PathToADUserToReceiveFilesDesktop\$($ADUser.SAMAccountName)"

    $DestinationPath = if ($HomeDirectory.Name -eq $ADUser.SAMAccountName) {
        $PathToADUserToReceiveFilesDesktop
    } else {
        $PathToFolderToContainUsersCopiedHomeDirectory
    }
        
    Copy-Item -Path $HomeDirectory -Destination $DestinationPath -Recurse -ErrorAction SilentlyContinue
    
    if (Test-DirectoriesSameSize -ReferenceDirectory $HomeDirectory -DifferenceDirectory $PathToFolderToContainUsersCopiedHomeDirectory) {
        Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
        $ADUser | Set-ADUser -Clear HomeDirectory
    } else {        
        Throw "Size of $HomeDirectory does not equal $DestinationPath after copying the files"
    }

    if ($ADUserToReceiveFiles.EmailAddress) {
        $To = $ADUserToReceiveFiles.EmailAddress
        $Subject = "$($ADUser.SAMAccountName)'s home directory files have been moved to your desktop"
        $Body = @"
$($ADUser.Name)'s home directory files have been moved to your desktop in a folder named $($ADUser.SAMAccountName). 
This was done as a part of the termination process for $($ADUser.Name).

If you believe you received these files incorrectly, please contact the Help Desk at x2248.
"@
        Send-TervisMailMessage -from "HelpDeskTeam@Tervis.com" -To $To -Subject $Subject -Body $Body
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
    Add-Member -MemberType ScriptProperty -Name TervisLastLogon -PassThru -Force -Value {
        [datetime]::FromFileTime($This.“lastLogonTimestamp”) 
    } | 
    Add-Member -MemberType AliasProperty -Name ComputerName -PassThru -Force -Value Name
}

#function Add-ADComputerCustomProperties {
#    param (
#        [Parameter(ValueFromPipeline,Mandatory)]$Input
#    )
#    process {
#        $Input | Add-member -Name UserNameInComputerName -MemberType ScriptProperty -Force -Value {
#            @($this.name -split "-")[0]         
#        }
#        
#        $Input | Add-member -Name ComputerNameSuffix -MemberType ScriptProperty -Force -Value {
#            @($this.name -split "-")[1]         
#        }
#
#        #$ADComputer | Add-member -Name UserNamesInComputerName -MemberType ScriptProperty -Force -Value {
#        #    $ADSAMAccountNames | where { $this.Name -match $_ }         
#        #}
#
#        #$ADComputer | Add-member -Name ComputersWithSimilarName -MemberType ScriptProperty -Force -Value {
#        #    $ADComputers | where {$_.name -Like "*$($this.UserNameInComputerName)*" -and $_.name -ne $this.name} | select -ExpandProperty name
#        #}
#    }
#}

function Remove-TervisADUser {
    [CMDLetBinding()]
    param(
        [Parameter(Mandatory)]$Identity,
        [Switch]$RemoveGroupos
    )
    $ADUser = Get-TervisADUser $Identity -Properties DistinguishedName,ProtectedFromAccidentalDeletion

    Write-Verbose "Setting a 120 character strong password on the user account"
    $Password = New-RandomPassword
    $SecurePassword = ConvertTo-SecureString $Password -asplaintext -force
    Set-ADAccountPassword -Identity $identity -NewPassword $SecurePassword

    Write-Verbose "Moving user account to the appropriate OU"
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
        Write-Verbose "Removing all AD group memberships"
        $Groups = Get-ADUser $Identity -Properties MemberOf | select -ExpandProperty MemberOf
        foreach ($Group in $Groups) {
            Remove-ADGroupMember -Identity $Group -Members $Identity -Confirm:$false
        }
    }

    Write-Verbose "Disabling AD account"
    Disable-ADAccount $Identity

    Write-Verbose "Setting AD account expiration"
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

function Disable-InactiveADComputers {
    $AdComputersToDisable = Get-TervisADComputer -Filter 'enabled -eq $true' -Properties LastLogonTimestamp,created,enabled,operatingsystem,PasswordLastSet | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-60) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-60) -and 
            $_.Enabled -eq $true -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.Name -notlike "TP9*" -and 
            $_.OperatingSystem -notlike "Windows Server*" -and 
            $_.OperatingSystem -ne "RHEL" -and 
            $_.OperatingSystem -ne "Mac OS X" -and 
            $_.OperatingSystem -ne $null} | 
        Sort Name
    [string]$AdComputersToDisableCount = ($AdComputersToDisable).count
    if ($AdComputersToDisableCount -ge "1") {
        $Body = "The following $AdComputersToDisableCount computers are being disabled. `n" 
        $Body += "Computer Name `t Tervis Last Logon `t Date Created `t Operating System `n"
        foreach ($ADComputer in $AdComputersToDisable) {
            $Body += ($ADComputer).name + "`t" + ($ADComputer).TervisLastLogon + "`t" + ($ADComputer).created + "`t" + ($ADComputer).operatingsystem + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive Computer Accounts to be Disabled' -Body $Body -SmtpServer $SMTPServer
        $AdComputersToDisable | Disable-ADAccount -Confirm:$false
    }
}

function Remove-InactiveADComputers {
    $AdComputersToDelete = @()
    $AdComputersToDelete = Get-TervisADComputer -Filter * -Properties LastLogonTimestamp,created,enabled,operatingsystem,ProtectedFromAccidentalDeletion,PasswordLastSet | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-190) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-190) -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.Name -notlike "TP9*" -and 
            $_.OperatingSystem -notlike "Windows Server*" -and 
            $_.OperatingSystem -ne "RHEL" -and 
            $_.OperatingSystem -ne "Mac OS X" -and 
            $_.OperatingSystem -ne $null} | 
        Sort Name
    [string]$AdComputersToDeleteCount = ($AdComputersToDelete).count
    if ($AdComputersToDeleteCount -ge "1") {
        $Body = "The following $AdComputersToDeleteCount computers are being deleted. `n" 
        $Body += "Computer Name `t Tervis Last Logon `t Date Created `t Operating System `n"
        foreach ($ADComputer in $AdComputersToDelete) {
            $Body += ($ADComputer).name + "`t" + ($ADComputer).TervisLastLogon + "`t" + ($ADComputer).created + "`t" + ($ADComputer).operatingsystem + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive Computer Accounts to be Deleted' -Body $Body -SmtpServer $SMTPServer
        foreach ($AdComputerToDelete in $AdComputersToDelete) {
            Set-Location AD:
            $AdObjectACL = Get-Acl ($AdComputerToDelete).DistinguishedName
            foreach ($AccessRule in $AdObjectACL.Access) {
                if ($AccessRule.IdentityReference.Value -eq 'Everyone' -and $AccessRule.AccessControlType -eq 'Deny' -and $AccessRule.ActiveDirectoryRights -match 'Delete') {
                    $AdObjectACL.RemoveAccessRule($AccessRule) | Out-Null
                }
            }
            Set-Location ($ENV:SystemRoot + '\System32')
            if ($AdComputerToDelete.ProtectedFromAccidentalDeletion) {
                Set-ADObject -Identity ($AdUserToDelete).DistinguishedName -ProtectedFromAccidentalDeletion $false
            }
        }
        $AdComputersToDelete | Remove-ADObject -Confirm:$false -Recursive
        $AdComputersToDelete | Remove-TervisDNSRecord
    }
}

function Disable-InactiveADUsers {
    $AdUsersToDisable = @()
    $AdUsersToDisable = Get-TervisADUser -Filter 'enabled -eq $true' -Properties LastLogonTimestamp,Created,Enabled,PasswordLastSet | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-30) -and 
            $_.Enabled -eq $true -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-90) -and
            $_.DistinguishedName -notmatch "CN=Microsoft Exchange System Objects,DC=" -and 
            $_.DistinguishedName -notmatch "OU=Exchange,DC=" -and 
            $_.DistinguishedName -notmatch "OU=Accounts - Service,DC="}
    $AdUsersToDisable += Get-TervisADUser -Filter 'enabled -eq $true' -Properties LastLogonTimestamp,Created,Enabled,PasswordLastSet | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-180) -and
            $_.Enabled -eq $true -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-90) -and
            $_.DistinguishedName -match "OU=Accounts - Service,DC=" -and
            $_.DistinguishedName -notmatch "OU=Inactivity Exceptions,OU=Accounts - Service,DC="}
    $AdUsersToDisable += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,created,enabled,PasswordLastSet,Manager | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-365) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-365) -and
            $_.DistinguishedName -match "OU=Inactivity Exceptions,OU=Accounts - Service,DC=" -and
            $_.Enabled -eq $true}
    $AdUsersToDisable = $AdUsersToDisable | sort Name
    [string]$AdUsersToDisableCount = ($AdUsersToDisable).count
    if ($AdUsersToDisableCount -ge "1") {
        $Body = "The following $AdUsersToDisableCount users are being disabled. `n" 
        $Body += "User Name `t Tervis Last Logon `t Date Created `n"
        foreach ($ADUser in $AdUsersToDisable) {
            $Body += ($ADUser).name + "`t" + ($ADUser).TervisLastLogon + "`t" + ($ADUser).created + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive User Accounts to be Disabled' -Body $Body -SmtpServer $SMTPServer
        $AdUsersToDisable | Disable-ADAccount -Confirm:$false
    }
}

function Remove-InactiveADUsers {
    $MESUsers = Get-MESOnlyUsers
    $AdUsersToDelete = @()
    $AdUsersToDelete = Get-TervisADUser -Filter * -Properties LastLogonTimestamp,Created,Enabled,ProtectedFromAccidentalDeletion,MemberOf,PasswordLastSet | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-190) -and 
            $_.Created -lt (Get-Date).AddDays(-90) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-90) -and
            $_.DistinguishedName -notmatch "CN=Microsoft Exchange System Objects," -and 
            $_.DistinguishedName -notmatch "OU=Exchange,DC=" -and  
            $_.DistinguishedName -notmatch "OU=Accounts - Service,DC=" -and
            $_.DistinguishedName -notin $MESUsers.DistinguishedName -and
            $_.Name -ne 'krbtgt' -and
            $_.Name -ne 'Guest' -and
            $_.Name -ne 'DefaultAccount'}
    $AdUsersToDelete += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,Created,Enabled,ProtectedFromAccidentalDeletion,MemberOf,PasswordLastSet | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-365) -and 
            $_.Created -lt (Get-Date).AddDays(-90) -and 
            $_.DistinguishedName -match "OU=Accounts - Service,DC=" -and
            $_.DistinguishedName -notmatch "OU=Inactivity Exceptions,OU=Accounts - Service,DC="}
    <#
    $AdUsersToDelete += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,created,enabled,PasswordLastSet,ProtectedFromAccidentalDeletion,MemberOf,Manager | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-425) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-425) -and
            $_.DistinguishedName -match "OU=Inactivity Exceptions,OU=Accounts - Service,DC="}
            #>
    $AdUsersToDelete += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,Created,Enabled,PasswordLastSet,ProtectedFromAccidentalDeletion,MemberOf,Manager | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-425) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-425) -and
            $_.Enabled -eq $false -and
            $_.DistinguishedName -match "OU=Inactivity Exceptions,OU=Accounts - Service,DC="}
    $AdUsersToDelete = $AdUsersToDelete | sort Name
    [string]$AdUsersToDeleteCount = ($AdUsersToDelete).count
    if ($AdUsersToDeleteCount -ge "1") {
        $AdUsersToDelete | Export-Clixml -Path ($env:TEMP + '\ADUsersToDelete.xml')
        $MailAttachment = ($env:TEMP + '\ADUsersToDelete.xml')
        $Body = "The following $AdUsersToDeleteCount users are being deleted. `n" 
        $Body += "User Name `t Tervis Last Logon `t Date Created `n"
        foreach ($ADUser in $AdUsersToDelete) {
            $Body += ($ADUser).name + "`t" + ($ADUser).TervisLastLogon + "`t" + ($ADUser).created + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive User Accounts to be Deleted' -Body $Body -SmtpServer $SMTPServer -Attachments $MailAttachment
        Remove-Item -Path $MailAttachment -Force -Confirm:$false
        foreach ($AdUserToDelete in $AdUsersToDelete) {
            Set-Location AD:
            $AdObjectACL = Get-Acl ($AdUserToDelete).DistinguishedName
            foreach ($AccessRule in $AdObjectACL.Access) {
                if ($AccessRule.IdentityReference.Value -eq 'Everyone' -and $AccessRule.AccessControlType -eq 'Deny' -and $AccessRule.ActiveDirectoryRights -match 'Delete') {
                    $AdObjectACL.RemoveAccessRule($AccessRule) | Out-Null
                }
            }
            Set-Location ($ENV:SystemRoot + '\System32')
            if ($AdUserToDelete.ProtectedFromAccidentalDeletion) {
                Set-ADObject -Identity ($AdUserToDelete).DistinguishedName -ProtectedFromAccidentalDeletion $false
            }
            if (($AdUserToDelete).DistinguishedName -match "OU=Departments,DC=") {
                Remove-TervisUser -Identity ($AdUserToDelete).SamAccountName -NoUserReceivesData
            } else {
                Remove-ADObject ($AdUserToDelete).DistinguishedName -Confirm:$false -Recursive
            }
        }
    }
}

function Send-TervisInactivityNotification {
    $AdUsersForNotification = @()
    $AdUsersForNotification += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,created,enabled,PasswordLastSet,Manager | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-351) -and 
            $_.PasswordLastSet -lt (Get-Date).AddDays(-351) -and
            $_.DistinguishedName -match "OU=Inactivity Exceptions,OU=Accounts - Service,DC=" -and
            $_.Enabled -eq $true}
    $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
    $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
    foreach ($AdUserForNotification in $AdUsersForNotification) {
        if (($AdUserForNotification).Manager) {
            $ManagerEmail = Get-ADUser ($AdUserForNotification).Manager -Properties EmailAddress -ErrorAction SilentlyContinue | Select -ExpandProperty EmailAddress
        }
        If ($ManagerEmail) {
            $To = $ManagerEmail
        } else {
            $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        }
        if (($AdUserForNotification).PasswordLastSet -gt ($AdUserForNotification).TervisLastLogon) {
            $DaysUntilDisabled = (Get-Date).AddDays(-365) - ($AdUserForNotification).PasswordLastSet
        } else {
            $DaysUntilDisabled = (Get-Date).AddDays(-365) - ($AdUserForNotification).TervisLastLogon | Select -ExpandProperty Days
        }
        $Body = "The following user in the Inactive Exceptions OU will be disabled in $DaysUntilDisabled days if the password is not reset. `n" 
        $Body += "User Name `t Tervis Last Logon `t Date Created 't Password Last Set `n"
        $Body += ($AdUserForNotification).name + "`t" + ($AdUserForNotification).TervisLastLogon + "`t" + ($AdUserForNotification).created + "`t" +  ($AdUserForNotification).PasswordLastSet + "`n"
        Send-MailMessage -To $To -From $From -Subject 'Service Account with Expection to be Disabled' -Body $Body -SmtpServer $SMTPServer
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

function Install-MoveMESUsersToCorrectOUScheduledTask {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ScheduledTaskCredential = New-Object System.Management.Automation.PSCredential (Get-PasswordstateCredential -PasswordID 259)
        $Execute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
        $Argument = '-Command Move-MESUsersToCorrectOU -NoProfile'
    }
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        If (-NOT (Get-ScheduledTask -TaskName Move-MESUsersToCorrectOU -CimSession $CimSession -ErrorAction SilentlyContinue)) {
            Install-TervisScheduledTask -Credential $ScheduledTaskCredential -TaskName Move-MESUsersToCorrectOU -Execute $Execute -Argument $Argument -RepetitionIntervalName EveryDayAt2am -ComputerName $ComputerName
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

function Get-ADObjectParentContainer {
    param(
        [Parameter(Mandatory,ValueFromPipeline)]$ObjectPath
    )
    ($ObjectPath.split(",") | select -skip 1 ) -join ","
}

function Move-MESUsersToCorrectOU {
    $MESUserNames = Get-MESUsersWhoHaveLoggedOnIn3Months -DataSource "MESSQL.production.$env:USERDNSDOMAIN" -DataBase MES
    $TargetOU = Get-ADOrganizationalUnit -Filter * | Where DistinguishedName -Match "OU=Users,OU=Production Floor,OU=Operations," |
        Select -ExpandProperty DistinguishedName
    foreach ($MESUser in $MESUserNames) {
        $ErrorActionPreference = "SilentlyContinue"
        $ADUser = Get-TervisADUser $MESUser -Properties LastLogonTimestamp,enabled,ProtectedFromAccidentalDeletion
        $ErrorActionPreference = "Continue"
        if ($ADUser -and (-NOT ($ADUser.DistinguishedName -Match "OU=Users,OU=Production Floor,OU=Operations,"))) {
            if (-NOT(($ADUser).TervisLastLogon -gt (Get-Date).AddDays(-30))) {
                if (-NOT (Test-TervisUserHasOffice365SharedMailbox -Identity ($ADUser).SamAccountName)) {
                    if ($ADUser.ProtectedFromAccidentalDeletion) {
                        Set-ADObject -Identity $ADUser.DistinguishedName -ProtectedFromAccidentalDeletion $false
                    }
                    $ADUser | Move-ADObject -TargetPath $TargetOU -Confirm:$false
                    $UserPrincipalName = $ADUser | Select -ExpandProperty UserPrincipalName
                    $ADUser = Get-ADObject -Filter {UserPrincipalName -eq $UserPrincipalName} 
                    $ADUser | Set-ADObject -ProtectedFromAccidentalDeletion $true
                }
            }
        }
    }
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

function Set-ADGroupManagedByAttribute{
    param(
        [parameter(Mandatory)]$ADGroup,
        [parameter(Mandatory)]$GroupManager
    )
    Get-ADGroup $ADGroup | Set-ADGroup -ManagedBy $GroupManager
}
