#Requires -Modules TervisDNS, ActiveDirectory, TervisMailMessage

#$ADGroups = get-adgroup -Filter *

function Get-TervisADUser {
    param (
        $Identity,
        $Path,
        $Filter,
        $Properties
    )
    
    $AdditionalNeededProperties = "msDS-UserPasswordExpiryTimeComputed"
    Get-ADUser @PSBoundParameters
}

function Add-ADUserCustomProperties {
    param (
        [Parameter(ValueFromPipeline)]$Input
    )

    $Input | Add-Member -MemberType ScriptProperty -Name PasswordExpirationDate -Value {
        [datetime]::FromFileTime($This.“msDS-UserPasswordExpiryTimeComputed”)
    }
}

Function Find-TervisADUsersComptuer {
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
        [Parameter(Mandatory, ParameterSetName=’ManagerReceivesFiles’)][Switch]$ManagerReceivesFiles,
        [Parameter(Mandatory, ParameterSetName=’AnotherUserReceivesFiles’)]$IdentityOfUserToReceiveHomeDirectoryFiles,        
        [Parameter(Mandatory, ParameterSetName=’DeleteUsersFiles’)][Switch]$DeleteFilesWithoutMovingThem
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties Manager, HomeDirectory

    if ($(Test-Path $ADUser.HomeDirectory) -eq $false) {
        Throw "$($ADUser.SamaccountName)'s home directory $($ADUser.HomeDirectory) doesn't exist"
    }

    $HomeDirectory = Get-Item $ADUser.HomeDirectory

    if ($ADUser.HomeDirectory -notmatch $ADUser.SamAccountName) {
        Throw "$($ADUser.HomeDirectory) doesn't have $($ADUser.SamAccountName) in it"
    }

    if ($DeleteFilesWithoutMovingThem) {
        Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
        $ADUser | Set-ADUser -Clear HomeDirectory
    } else {
        if ($ManagerReceivesFiles) {
            if( -not $ADUser.Manager) { Throw "ManagerReceivesFiles was specified but the user doesn't have a manager in Active Directory" }
            $IdentityOfUserToReceiveHomeDirectoryFiles = $ADUser.Manager
        }        
        $ADUserToReceiveFiles = Get-ADUser -Identity $IdentityOfUserToReceiveHomeDirectoryFiles -Properties EmailAddress
        
        if (-not $ADUserToReceiveFiles) { "Running Get-ADUser for the identity $IdentityOfUserToReceiveHomeDirectoryFiles didn't find an Active Directory user" }

        $ADUserToReceiveFilesComputer = $ADUserToReceiveFiles | Find-TervisADUsersComptuer
        if (-not $ADUserToReceiveFilesComputer ) { Throw "Couldn't find an ADComputer with $($ADUserToReceiveFiles.SamAccountName) in the computer's name" }
        if ($ADUserToReceiveFilesComputer.count -gt 1) { 
            Throw "We found more than one AD computer for $($ADUserToReceiveFiles.SamAccountName). Run Get-ADUser $($ADUserToReceiveFiles.SamAccountName) | Find-TervisADUsersComptuer -Properties lastlogondate to see the computers"                 
        }

        $PathToADUserToReceiveFilesDesktop = "\\$($ADUserToReceiveFilesComputer.Name)\C$\Users\$($ADUserToReceiveFiles.SAMAccountName)\Desktop"

        if ($(Test-Path $PathToADUserToReceiveFilesDesktop) -eq $false) {
            Throw "$PathToADUserToReceiveFilesDesktop doesn't exist so we cannot copy the user's home directory files over"
        }

        $DestinationPath = if ($HomeDirectory.Name -eq $ADUser.SAMAccountName) {
            $PathToADUserToReceiveFilesDesktop
        } else {
            "$PathToADUserToReceiveFilesDesktop\$($ADUser.SAMAccountName)"
        }

        Copy-Item -Path $HomeDirectory -Destination $DestinationPath -Recurse -ErrorAction SilentlyContinue
        
        $TotalHomeDirectorySize = Get-ChildItem $HomeDirectory -Recurse -Force | 
            Measure-Object -property length -sum | 
            select -ExpandProperty Sum

        $TotalCopiedHomeDirectorySize = Get-ChildItem "$PathToADUserToReceiveFilesDesktop\$($ADUser.SAMAccountName)" -Recurse -Force | 
            Measure-Object -property length -sum | 
            select -ExpandProperty Sum
        
        if ($TotalHomeDirectorySize -eq $TotalCopiedHomeDirectorySize ) {
            Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
            $ADUser | Set-ADUser -Clear HomeDirectory
        } else {        
            Throw "TotalHomeDirectorySize: $TotalHomeDirectorySize didn't equal TotalCopiedHomeDirectorySize: $TotalCopiedHomeDirectorySize"
        }

        if ($ADUserToReceiveFiles.EmailAddress) {
            $To = $ADUserToReceiveFilesComputer.EmailAddress
            $Subject = "$($ADUser.SAMAccountName)'s home directory files have been moved to your desktop"
            $Body = @"
$($ADUser.Name)'s home directory files have been moved to your desktop in a folder named $($ADUser.SAMAccountName). 
This was done as a part of the termination process for $($ADUser.Name).

If you believe you received these files incorrectly, please contact the Help Desk at x2248.
"@
            Send-TervisMailMessage -from "HelpDeskTeam@Tervis.com" -To $To -Subject $Subject -Body $Body
        }
    }
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