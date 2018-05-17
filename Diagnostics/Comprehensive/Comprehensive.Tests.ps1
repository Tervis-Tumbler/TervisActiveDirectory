$ADUsers = Get-ADUser -Filter * -SearchBase "OU=Departments,DC=tervis,DC=prv" -Properties Manager, EmployeeID, MemberOf, PasswordNeverExpires, Department, TelephoneNumber

foreach ($ADUser in $ADUsers) {
    Describe "Active Directory User" {
        It "$($ADUser.Name) ($($ADUser.samaccountname)) should have an employee ID" {
            
            $ADUserThatShouldHaveEmployeeID = $ADUSer | 
            where DistinguishedName -NotMatch "OU=Store Accounts,OU=Users,OU=Stores,OU=Departments" |
            where {-not ($_.MemberOf -Match "CN=Contractor,")} |
            where {-not ($_.MemberOf -Match "CN=SharedAccountsThatNeedToBeAddressed,")} |
            where {-not ($_.MemberOf -Match "CN=Test Users,")}
            
            if ($ADUserThatShouldHaveEmployeeID) {
                $ADUserThatShouldHaveEmployeeID.EmployeeID | Should Not BeNullOrEmpty
            }
        }

        It "$($ADUser.Name) ($($ADUser.samaccountname)) should have a manager set" {
            $ADUserThatShouldHaveManager = $ADUSer | 
            where DistinguishedName -NotMatch "OU=Store Accounts,OU=Users,OU=Stores,OU=Departments" |
            where {-not ($_.MemberOf -Match "CN=Contractor,")} |
            where {-not ($_.MemberOf -Match "CN=SharedAccountsThatNeedToBeAddressed,")} |
            where {-not ($_.MemberOf -Match "CN=Test Users,")}
            
            if ($ADUserThatShouldHaveManager) {
                $ADUserThatShouldHaveManager.Manager | Should Not BeNullOrEmpty
            }
        }

        It "$($ADUser.Name) ($($ADUser.samaccountname)) should not have the PasswordNeverExpires enabled" {
            $ADUserThatShouldNotHavePasswordNeverExpires = $ADUSer |
            where DistinguishedName -NotMatch "OU=Store Accounts,OU=Users,OU=Stores,OU=Departments" |
            where {-not ($_.MemberOf -Match "CN=SharedAccountsThatNeedToBeAddressed,")}

            if ($ADUserThatShouldNotHavePasswordNeverExpires) {
                $ADUserThatShouldNotHavePasswordNeverExpires.PasswordNeverExpires | Should Be $false
            }
        }

        It "$($ADUser.Name) ($($ADUser.samaccountname)) should have a department set" {
            $ADUser.Department | Should Not BeNullOrEmpty
        }
    }
}