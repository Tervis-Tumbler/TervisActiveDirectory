﻿#Requires -Modules TervisDNS, ActiveDirectory, TervisMailMessage

#$ADGroups = get-adgroup -Filter *

function Get-TervisADUser {
    param (
        $Identity,
        $Path,
        $Filter,
        $Properties
    )
    
    $AdditionalNeededProperties = "msDS-UserPasswordExpiryTimeComputed"
    Get-ADUser @PSBoundParameters | Add-ADUserCustomProperties
}

function Add-ADUserCustomProperties {
    param (
        [Parameter(ValueFromPipeline)]$Input
    )

    $Input | Add-Member -MemberType ScriptProperty -Name PasswordExpirationDate -PassThru -Force -Value {
        [datetime]::FromFileTime($This.“msDS-UserPasswordExpiryTimeComputed”)
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

Function Remove-TervisADUserHomeDirectory {
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
        Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
        $ADUser | Set-ADUser -Clear HomeDirectory
    } else {
        $ADUserToReceiveFiles = Get-ADUser -Identity $IdentityOfUserToReceiveHomeDirectoryFiles -Properties EmailAddress
        
        if (-not $ADUserToReceiveFiles) { "Running Get-ADUser for the identity $IdentityOfUserToReceiveHomeDirectoryFiles didn't find an Active Directory user" }

        $ADUserToReceiveFilesComputer = $ADUserToReceiveFiles | Find-TervisADUsersComputer
        if (-not $ADUserToReceiveFilesComputer ) { Throw "Couldn't find an ADComputer with $($ADUserToReceiveFiles.SamAccountName) in the computer's name" }
        if ($ADUserToReceiveFilesComputer.count -gt 1) { 
            Throw "We found more than one AD computer for $($ADUserToReceiveFiles.SamAccountName). Run Get-ADUser $($ADUserToReceiveFiles.SamAccountName) | Find-TervisADUsersComputer -Properties lastlogondate to see the computers"
        }

        if ($ADUserToReceiveFilesComputer | Test-TervisADComputerIsMac) {
            Throw "ADUserToReceiveFilesComputer: $($ADUserToReceiveFilesComputer.Name) is a Mac, cannot copy the files automatically"            
        }

        Invoke-CopyADUsersHomeDirectoryToADUserToRecieveFilesComputer -ErrorAction Stop -Identity $Identity -IdentityOfUserToReceiveHomeDirectoryFiles $IdentityOfUserToReceiveHomeDirectoryFiles -ADUserToReceiveFilesComputer $ADUserToReceiveFilesComputer
    }
}

Function Invoke-CopyADUsersHomeDirectoryToADUserToRecieveFilesComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]$Identity,       
        [Parameter(Mandatory)]$IdentityOfUserToReceiveHomeDirectoryFiles,
        [Parameter(Mandatory)]$ADUserToReceiveFilesComputer
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties HomeDirectory
    $ADUserToReceiveFiles = Get-ADUser -Identity $IdentityOfUserToReceiveHomeDirectoryFiles -Properties EmailAddress

    $PathToADUserToReceiveFilesDesktop = "\\$($ADUserToReceiveFilesComputer.Name)\C$\Users\$($ADUserToReceiveFiles.SAMAccountName)\Desktop"

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

Function Test-TervisADComputerIsMac {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ADComputer
    )

    $ADComputer.Name -match "-mac"
}

#function Get-TervisADComputer {
#
#}
#
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