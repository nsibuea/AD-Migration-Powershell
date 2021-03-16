# Powershell script to create bulk Group upload with a provided CSV file. 
# Attachment file needs to be renamed to csv and path need to be updated. 
# Author - Arnold 

#Powershell v2.0 
 
#Import the Active Directory Module 
Import-module activedirectory  
 
#Import the list from the user 
$Users = Import-Csv -Path C:\data\ADGroup.csv            
foreach ($User in $Users)     
        
{            
           
 
    #Creation of the account with the requested formatting. 
    New-ADGroup -Name $user.name -GroupScope $user.GroupScope -Description $user.Description -DisplayName $user.DisplayName -GroupCategory $user.GroupCategory `
                -Path $user.Path -SamAccountName $user.samAccountName


}