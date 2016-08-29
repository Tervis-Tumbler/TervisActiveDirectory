#Requires -Modules TervisDNS, ActiveDirectory

#$ADGroups = get-adgroup -Filter *

function Get-TervisADUser {
    param (
        $Identity,
        $Path,
        $Filter,
        $Properties
    )
    
    $AdditionalNeededProperites = "msDS-UserPasswordExpiryTimeComputed"
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
        [Parameter(Mandatory, ParameterSetName=’ManagerRecievesFiles’)][Switch]$ManagerRecievesFiles,
        [Parameter(Mandatory, ParameterSetName=’AnotherUserReceivesFiles’)]$IdentityOfUserToRecieveHomeDirectoryFiles,
        
        [Parameter(Mandatory, ParameterSetName=’ManagerRecievesFiles’)]
        [Parameter(Mandatory, ParameterSetName=’AnotherUserReceivesFiles’)]
        $NameOfServerToSendMailFrom,

        [Parameter(Mandatory, ParameterSetName=’DeleteUsersFiles’)][Switch]$DeleteFilesWithOutMovingThem
    )
    $ADUser = Get-ADUser -Identity $Identity -Properties Manager, HomeDirectory

    if ($(Test-Path $ADUser.HomeDirectory) -eq $false) {
        Throw "$($ADUser.SamaccountName)'s home directory $($ADUser.HomeDirectory) doesn't exist"
    }

    $HomeDirectory = Get-Item $ADUser.HomeDirectory

    if ($ADUser.HomeDirectory -notmatch $ADUser.SamAccountName) {
        Throw "$($ADUser.HomeDirectory) doesn't have $($ADUser.SamAccountName) in it"
    }

    if ($DeleteFilesWithOutMovingThem) {
        Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
        $ADUser | Set-ADUser -Clear HomeDirectory
    } else {
        if ($ManagerRecievesFiles) {
            if( -not $ADUser.Manager) { Throw "ManagerRecievesFiles was specified but the user doesn't have a manager in active directory" }
            $IdentityOfUserToRecieveHomeDirectoryFiles = $ADUser.Manager
        }        
        $ADUserToRecieveFiles = Get-ADUser -Identity $IdentityOfUserToRecieveHomeDirectoryFiles
        
        if (-not $ADUserToRecieveFiles) { "Running Get-ADUser for the identity $IdentityOfUserToRecieveHomeDirectoryFiles didn't find an active directory user" }

        $ADUserToRecieveFilesComputer = $ADUserToRecieveFiles | Find-TervisADUsersComptuer
        if (-not $ADUserToRecieveFilesComputer ) { Throw "Couldn't find an ADComputer with $($ADUserToRecieveFiles.SamAccountName) in the computer's name" }
        if ($ADUserToRecieveFilesComputer.count -gt 1) { 
            Throw "We found more than one AD computer for $($ADUserToRecieveFiles.SamAccountName). Run Get-ADUser $($ADUserToRecieveFiles.SamAccountName) | Find-TervisADUsersComptuer -Properties lastlogondate to see the computers"                 
        }

        $PathToADUserToRecieveFilesDesktop = "\\$($ADUserToRecieveFilesComputer.Name)\C$\Users\$($ADUserToRecieveFiles.SAMAccountName)\Desktop"

        if ($(Test-Path $PathToADUserToRecieveFilesDesktop) -eq $false) {
            Throw "$PathToADUserToRecieveFilesDesktop doesn't exist so we cannot copy the user's home directory files over"
        }

        $DestinationPath = if ($HomeDirectory.Name -eq $ADUser.SAMAccountName) {
            $PathToADUserToRecieveFilesDesktop
        } else {
            "$PathToADUserToRecieveFilesDesktop\$($ADUser.SAMAccountName)"
        }

        Copy-Item -Path $HomeDirectory -Destination $DestinationPath -Recurse
        
        $TotalHomeDirectorySize = Get-ChildItem $HomeDirectory | Measure-Object -property length -sum | select -ExpandProperty Sum
        $TotalCopiedHomeDirectorySize = Get-ChildItem $DestinationPath | Measure-Object -property length -sum | select -ExpandProperty Sum
        if ($TotalHomeDirectorySize -eq $TotalCopiedHomeDirectorySize ) {
            Remove-Item -Path $ADUser.HomeDirectory -Confirm -Recurse -Force
            $ADUser | Set-ADUser -Clear HomeDirectory
        }

        if ($ADUserToRecieveFilesComputer.EmailAddress) {
            $To = $ADUserToRecieveFilesComputer.EmailAddress
            $Subject = "$($ADUser.SAMAccountName)'s home directory files have been moved to your desktop"
            $Body = @"
$($ADUser.Name)'s home directory files have been moved to your desktop in a folder named $($ADUser.SAMAccountName). 
This was done as a part of the termination process for $($ADUser.Name).

If you believe you recieved these files incorreclty please contact the Help Desk at x2248.
"@
            Invoke-Command -ComputerName $NameOfServerToSendMailFrom -ArgumentList $Subject,$Body,$To -ScriptBlock {
                param ($Subject, $Body, $To)
                Send-MailMessage -from "HelpDesk@Tervis.com" -To $To -Subject $Subject -Body $Body -SmtpServer $(Get-TervisDNSMXMailServer)
            }
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