#Requires -Modules TervisDNS, ActiveDirectory, TervisMailMessage

#$ADGroups = get-adgroup -Filter *

function Get-TervisADUser {
    param (
        $Identity,
        $Path,
        $Filter,
        $Properties
    )
    
    $AdditionalNeededProperties = "msDS-UserPasswordExpiryTimeComputed,lastLogonTimestamp"
    Get-ADUser @PSBoundParameters | Add-ADUserCustomProperties
}

function Add-ADUserCustomProperties {
    param (
        [Parameter(ValueFromPipeline)]$Input
    )

    $Input | Add-Member -MemberType ScriptProperty -Name PasswordExpirationDate -PassThru -Force -Value {
        [datetime]::FromFileTime($This.“msDS-UserPasswordExpiryTimeComputed”)
    } |
    Add-Member -MemberType ScriptProperty -Name TervisLastLogon -PassThru -Force -Value {
        [datetime]::FromFileTime($This."lastLogonTimestamp")
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
    Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
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
    param (
        [Parameter(Mandatory)]$Server
    )

    $DC = Get-ADDomainController
    Invoke-Command -computername $DC.HostName -ScriptBlock {repadmin /syncall /ed}
    Invoke-Command -ComputerName $Server -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
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

    $Input | Add-Member -MemberType ScriptProperty -Name TervisLastLogon -PassThru -Force -Value {
        [datetime]::FromFileTime($This.“lastLogonTimestamp”)
    }
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
    $ADUser = Get-ADUser $Identity -Properties DistinguishedName,ProtectedFromAccidentalDeletion

    Write-Verbose "Setting a 120 character strong password on the user account"
    $Password = New-RandomPassword
    $SecurePassword = ConvertTo-SecureString $Password -asplaintext -force
    Set-ADAccountPassword -Identity $identity -NewPassword $SecurePassword

    Write-Verbose "Moving user account to the appropriate OU"
    if ($ADUser.ProtectedFromAccidentalDeletion) {
        Set-ADObject -Identity $ADUser.DistinguishedName -ProtectedFromAccidentalDeletion $false
    }
    If (Test-TervisUserHasOffice365Mailbox -Identity $Identity) {
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
    $AdComputersToDisable = Get-TervisADComputer -Filter 'enabled -eq $true' -Properties LastLogonTimestamp,created,enabled,operatingsystem | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-30) -and 
            $_.Enabled -eq $true -and 
            $_.Created -lt (Get-Date).AddDays(-30) -and 
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
    }
    $AdComputersToDisable | Disable-ADAccount -Confirm:$false
}

function Remove-InactiveADComputers {
    $AdComputersToDelete = @()
    $AdComputersToDelete = Get-TervisADComputer -Filter * -Properties LastLogonTimestamp,created,enabled,operatingsystem | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-190) -and 
            $_.Created -lt (Get-Date).AddDays(-30) -and 
            $_.Name -notlike "TP9*" -and 
            $_.OperatingSystem -notlike "Windows Server*" -and 
            $_.OperatingSystem -ne "RHEL" -and 
            $_.OperatingSystem -ne "Mac OS X" -and 
            $_.OperatingSystem -ne $null} | 
        Sort Name
    [string]$AdComputersToDeleteCount = ($AdComputersToDelete).count
    if ($AdComputersToDeleteCount -ge "1") {
        $Body = "The following $AdComputersToDisableCount computers are being deleted. `n" 
        $Body += "Computer Name `t Tervis Last Logon `t Date Created `t Operating System `n"
        foreach ($ADComputer in $AdComputersToDelete) {
            $Body += ($ADComputer).name + "`t" + ($ADComputer).TervisLastLogon + "`t" + ($ADComputer).created + "`t" + ($ADComputer).operatingsystem + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive Computer Accounts to be Deleted' -Body $Body -SmtpServer $SMTPServer
    }
    $AdComputersToDelete | Remove-ADComputer -Confirm:$false
    $AdComputersToDelete | Remove-TervisDNSRecord
}

function Disable-InactiveADUsers {
    $AdUsersToDisable = @()
    $AdUsersToDisable = Get-TervisADUser -Filter 'enabled -eq $true' -Properties LastLogonTimestamp,Created,Enabled | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-30) -and 
            $_.Enabled -eq $true -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.DistinguishedName -notmatch "CN=Microsoft Exchange System Objects,DC=" -and 
            $_.DistinguishedName -notmatch "OU=Exchange,DC=" -and 
            $_.DistinguishedName -notmatch "OU=Accounts - Service,DC="}
    $AdUsersToDisable += Get-TervisADUser -Filter 'enabled -eq $true' -Properties LastLogonTimestamp,Created,Enabled | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-180) -and
            $_.Enabled -eq $true -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.DistinguishedName -match "OU=Accounts - Service,DC="}
    $AdUsersToDisable = $AdUsersToDisable | sort Name
    [string]$AdUsersToDisableCount = ($AdUsersToDisable).count
    if ($AdUsersToDisableCount -ge "1") {
        $Body = "The following $AdUsersToDisableCount users are being disabled. `n" 
        $Body += "User Name `t Tervis Last Logon `t Date Created `t Operating System `n"
        foreach ($ADUser in $AdUsersToDisable) {
            $Body += ($ADUser).name + "`t" + ($ADUser).TervisLastLogon + "`t" + ($ADUser).created + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive User Accounts to be Disabled' -Body $Body -SmtpServer $SMTPServer
    }
    $AdUsersToDisable | Disable-ADAccount -Confirm:$false
}

function Remove-InactiveADUsers {
    $MESUsers = Get-MESOnlyUsers
    $AdUsersToDelete = @()
    $AdUsersToDelete = Get-TervisADUser -Filter * -Properties LastLogonTimestamp,created,enabled,ProtectedFromAccidentalDeletion | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-190) -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.DistinguishedName -notmatch "CN=Microsoft Exchange System Objects," -and 
            $_.DistinguishedName -notmatch "OU=Exchange,DC=" -and  
            $_.DistinguishedName -notmatch "OU=Accounts - Service,DC=" -and
            $_.DistinguishedName -notin $MESUsers.DistinguishedName}
    $AdUsersToDelete += Get-TervisADUser -Filter * -Properties LastLogonTimestamp,created,enabled | 
        where {$_.TervisLastLogon -lt (Get-Date).AddDays(-365) -and 
            $_.Created -lt (Get-Date).AddDays(-60) -and 
            $_.DistinguishedName -match "OU=Accounts - Service,DC="}
    $AdUsersToDelete = $AdUsersToDelete | sort Name
    [string]$AdUsersToDeleteCount = ($AdUsersToDelete).count
    if ($AdUsersToDeleteCount -ge "1") {
        $Body = "The following $AdUsersToDeleteCount users are being deleted. `n" 
        $Body += "User Name `t Tervis Last Logon `t Date Created `t Operating System `n"
        foreach ($ADUser in $AdUsersToDelete) {
            $Body += ($ADUser).name + "`t" + ($ADUser).TervisLastLogon + "`t" + ($ADUser).created + "`n"
        }
        $To = Get-ADGroup -Filter {name -like "it tech*"} -Properties mail | select -ExpandProperty mail
        $From = Get-ADUser -Filter {name -like "*daemon"} -Properties mail | select -ExpandProperty mail
        $SMTPServer = Get-ADObject -Filter {servicePrincipalName -like "*exchangemdb*"} -Properties dNSHostName | select -ExpandProperty dNSHostName
        Send-MailMessage -To $To -From $From -Subject 'Inactive User Accounts to be Deleted' -Body $Body -SmtpServer $SMTPServer
    }
    foreach ($AdUserToDelete in $AdUsersToDelete) {
        if ($AdUserToDelete.ProtectedFromAccidentalDeletion) {
            Set-ADObject -Identity ($AdUserToDelete).DistinguishedName -ProtectedFromAccidentalDeletion $false
        }
        if (($AdUserToDelete).DistinguishedName -match "OU=Departments,DC=") {
            Remove-TervisUser -Identity ($AdUserToDelete).SamAccountName -NoUserReceivesData
        } else {
            Remove-ADUser ($AdUserToDelete).DistinguishedName -Confirm:$false
        }
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
   
    Import-TervisMSOnlinePSSession
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

function Get-ADObjectParentContainer {
    param(
        [Parameter(Mandatory,ValueFromPipeline)]$ObjectPath
    )
    ($ObjectPath.split(",") | select -skip 1 ) -join ","
}
