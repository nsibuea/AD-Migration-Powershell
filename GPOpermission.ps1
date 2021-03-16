function Get-GPOPermissions {            

param($GpoName)
import-module GroupPolicy            

$permsobj = Get-GPPermissions -Name $GPOName -All
foreach ($perm in $permsobj) {            

    $obj = New-Object -TypeName PSObject -Property @{
   GPOName  = $GPOName
   AccountName = $($perm.trustee.name)
        AccountType = $($perm.trustee.sidtype.tostring())
        Permissions = $($perm.permission)
 }
$obj | Select GPOName, AccountName, AccountType, Permissions            

}
}

$Users = Import-Csv -Path C:\Script\listGPO\listgpo.csv   

$Result = @()
         
$(foreach ($User in $Users)  
{
   
   Get-GPOPermissions -GpoName $User.DisplayName
   

  # $Result += New-Object psobject -Property $Properties

})| Export-csv .\data.csv -noType