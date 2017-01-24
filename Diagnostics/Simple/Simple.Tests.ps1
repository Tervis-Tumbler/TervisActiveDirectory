Describe "Active Directory Users In Departments OU" {
    $ADUsers = Search-ADAccount -PasswordNeverExpires -SearchBase "OU=Departments,DC=tervis,DC=prv" |

    It "Shouldn't have PasswordSetToNeverExpire set to true" {
        Search-ADAccount -PasswordNeverExpires -SearchBase "OU=Departments,DC=tervis,DC=prv" |
        Should BeNullOrEmpty
    }
}
