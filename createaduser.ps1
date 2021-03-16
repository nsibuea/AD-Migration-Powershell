# Powershell script to create  bulk user upload with a provided CSV file. 
# Attachment file needs to be renamed to csv and path need to be updated. 
# Author - Arnold 
#   

#Import the Active Directory Module 
Import-module activedirectory  
 
#Clear the cmd Window
Clear-Host

$InputCsv = '.\olduser.csv'
$logFile = '.\logfile.csv'


#Verfiy the path to the Input CSV file 
if (!(Test-Path $InputCSV)) 
    {Write-Host "Your Input CSV does not exist, Exiting..." ; Sleep 5 ; Exit} 
If ($NewLog -eq 'Y') 
    { 
        Write-host "Creating a new log file" 
        #Rename Existing file to old. 
        Get-Item $Logfile | Move-Item -force -Destination { [IO.Path]::ChangeExtension( $_.Name, "old" ) } 
    } 


 
#Add the Current Date to the log file 
$Date = Get-Date 
Add-Content $Logfile $Date 


Add-Content -Path $Logfile -Value ("**********************************************")
Add-Content -Path $Logfile -Value ("Starting Adding Users")
Add-Content -Path $Logfile -Value ("**********************************************")
# Clear All Error
$error.clear()
 

  
#Import the list from the user 
$Users = Import-Csv -Path C:\Scripts\addusers\olduser.csv            
foreach ($User in $Users)     
        
{            
           
#    $Password = "pass.123" 

$error.clear()
  
    #Creation of the account with the requested formatting. 
    New-ADUser -Name $User.Name -DisplayName $User.displayName -SamAccountName $User.sAMAccountName -UserPrincipalName $User.UserPrincipalName -AccountPassword (ConvertTo-SecureString $User.Password -AsPlainText -force) -GivenName $User.GivenName  -Surname $User.Surname -Description $User.description `
			   -Enabled: ([System.Convert]::ToBoolean($User.Enabled)) -PasswordNeverExpires: ([System.Convert]::ToBoolean($User.PasswordNeverExpires)) -PasswordNotRequired: ([System.Convert]::ToBoolean($User.PasswordNotRequired)) -CannotChangePassword: ([System.Convert]::ToBoolean($user.CannotChangePassword)) -SmartcardLogonRequired: ([System.Convert]::ToBoolean($User.SmartcardLogonRequired)) `
			   -Path $User.Path -EmployeeNumber $User.EmployeeNumber -Title $User.Title -Department $User.department -Organization $User.Organization -Division $User.Division -Company $User.Company -EmployeeID $User.EmployeeID -Office $User.Office `
			   -City $User.City -Country $User.Country -POBox $User.POBox -StreetAddress $User.StreetAddress -PostalCode $User.PostalCode -State $User.State `
              		   -EmailAddress $User.EmailAddress -Fax $User.Fax -HomePhone $User.HomePhone -MobilePhone $User.MobilePhone -OfficePhone $User.OfficePhone `
               		   -HomeDirectory $User.HomeDirectory -HomeDrive $User.HomeDrive -HomePage $User.HomePage -ProfilePath $User.ProfilePath -ScriptPath $User.ScriptPath `
               		   -Initials $User.Initials -OtherName $User.OtherName
if ($error -ne ""){

    Add-Content -path $Logfile -value "$($Time);$($user.Displayname);import new users;$($error);"

    }
}

Add-Content -Path $Logfile -Value ("**********************************************")
Add-Content -Path $Logfile -Value ("End of Script")
Add-Content -Path $Logfile -Value ("**********************************************")
