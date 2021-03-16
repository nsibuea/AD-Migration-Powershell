 <#
    .SYNOPSIS
    Imports Active Directory users to Active Directory groups using a CSV
    .DESCRIPTION
    Author: Daniel Classon
    Version 1.0
 
    This script will take the information in the CSV and add the users specified in the User column and add them to the Group specied in Group column
    .PARAMETER CSV
    Specify the full source to the CSV file i.e c:\temp\members.csv
    .EXAMPLE
    .\add_users_to_multiple_groups.ps1 -CSV c:\temp\members.csv
    .DISCLAIMER
    All scripts and other powershell references are offered AS IS with no warranty.
    These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
    #>
 
[CmdletBinding()]
 
param(
    [Parameter(Mandatory=$True, Helpmessage="Specify full path to CSV (i.e c:\temp\members.csv")]
    [string]$CSV   
)
BEGIN{
    #Checks if the user is in the administrator group. Warns and stops if the user is not.
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator"))
    {
        Write-Warning "You are not running this as local administrator. Run it again in an elevated prompt."
	    Break
    }
    try {
    Import-Module ActiveDirectory
    }
    catch {
    Write-Warning "The Active Directory module was not found"
    }
    try {
    $Users = Import-CSV $CSV
    }
    catch {
    Write-Warning "The CSV file was not found"
    }
}
PROCESS{
 
    foreach($User in $Users){
        try{
            Add-ADGroupMember $User.Group -Members $User.User -ErrorAction Stop -Verbose
        }
        catch{
        }
 
    }
}
END{
 
}