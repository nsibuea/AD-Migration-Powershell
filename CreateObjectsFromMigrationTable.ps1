param([String]$MigrationTable,[String]$LogFile,[String]$Description);

if ([String]::IsNullOrEmpty($MigrationTable))
{
        write-host -ForegroundColor Red "MigrationTable switch must be specified";
        write-host -ForegroundColor Red "Exiting Script";
        return;
}
else
{
    [xml]$MTData = Get-Content $MigrationTable;
}

if ([String]::IsNullOrEmpty($LogFile))
{
    $LogFile = "C:\MigrationTableLogFile.txt";
    set-content $LogFile $NULL;
}
else
{
    set-content $LogFile $NULL;
}

if ([String]::IsNullOrEmpty($Description))
{
    $Description = "This object was created by the CreateObjectsfromMigrationTable.ps1 script. It may need to be renamed.";
}
else
{
    #Do Nothing;
}

import-module ActiveDirectory;

$myDomain = [System.Net.NetworkInformation.IpGlobalProperties]::GetIPGlobalProperties().DomainName;
$myTempDomainDn = "DC=" + [String]::Join(",DC=", $myDomain.Split("."));
$workingContainer = $myTempDomainDn;

try
{
    if (![System.DirectoryServices.DirectoryEntry]::Exists("LDAP://" + $workingContainer))
    {
        write-host -ForegroundColor Red "Could not connect to LDAP path $workingContainer";
        write-host -ForegroundColor Red "Exiting Script";
        return;
    }
}
catch
{
        write-host -ForegroundColor Red "Could not connect to LDAP path $workingContainer";
        write-host -ForegroundColor Red "Exiting Script";
        return;
}
        
foreach ($entity in $MTData.MigrationTable.Mapping)
{

    $objectExists = $NULL

    if ($entity.Type -eq "Unknown")
    {
        add-content -path $LogFile -value ($entity.Type + " entity type " + $entity.Source + " was detected. Usually this message can be safely ignored. Investigate the entity and address accordingly.");
    }
    
    if ($entity.Type -eq "User")
    {
        $samname = $entity.Source.split("@")
                
        try
        {
            $objectExists = get-aduser $samname[0];
        }
        catch
        {
            #Do Nothing
        }
               
        
        if ($objectExists -ne $NULL)
        {
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " already exists.");
        }
        else
        {
            New-ADUser ($samname[0]) -Description $Description;
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " was created.");
        }
        
        $entity.DestinationSameAsSource = $samname[0] + "@" + $myDomain;
    }

    if ($entity.Type -eq "LocalGroup")
    {
        $samname = $entity.Source.split("@");
        
        try
        {
            $objectExists = get-adgroup $samname[0];
        }
        catch
        {
            #Do Nothing
        }
          
        if ($objectExists -ne $NULL)
        {
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " already exists.");
        }
        else
        {
            New-ADGroup $samname[0] -GroupScope DomainLocal -Description $Description;
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " was created.");
        }
        
        $entity.DestinationSameAsSource = $samname[0] + "@" + $myDomain;
    }
    
    if ($entity.Type -eq "GlobalGroup")
    {
        $samname = $entity.Source.split("@");

        try
        {
            $objectExists = get-adgroup $samname[0];
        }
        catch
        {
            #Do Nothing
        }
        
        if ($objectExists -ne $NULL)
        {
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " already exists.");
        }
        else
        {
            New-ADGroup $samname[0] -GroupScope Global -Description $Description;
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " was created.");
        }
        
        $entity.DestinationSameAsSource = $samname[0] + "@" + $myDomain;
    }
    
    if ($entity.Type -eq "UniversalGroup")
    {
        $samname = $entity.Source.split("@");

        try
        {
            $objectExists = get-adgroup $samname[0];
        }
        catch
        {
            #Do Nothing
        }
        
        if ($objectExists -ne $NULL)
        {
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " already exists.");
        }
        else
        {
            New-ADGroup $samname[0] -GroupScope Universal -Description $Description;
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " was created.");
        }
        
        $entity.DestinationSameAsSource = $samname[0] + "@" + $myDomain;
    }

    if ($entity.Type -eq "UNCPath")
    {
        add-content -path $LogFile -value ("The following UNC path exists in at least on GPO: " + $entity.Source + " - Please manually correct the UNC path using the Migration Table Editor.")
    }
    
    if ($entity.Type -eq "Computer")
    {
        $samname = $entity.Source.split("$");
 
        try
        {
            $objectExists = get-adcomputer $samname[0];
        }
        catch
        {
            #Do Nothing
        }
        
        if ($objectExists -ne $NULL)
        {
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " already exists.");
        }
        else
        {
            New-ADComputer $samname[0] -Description $Description;
            add-content -path $LogFile -value ($entity.Type + " " + $samname[0] + " was created.");
            $entity.DestinationSameAsSource = $samname[0] + "$@" + $myDomain
        }
        
        $entity.DestinationSameAsSource = $samname[0] + "@" + $myDomain;
    }
}

$MTData.Save($MigrationTable)

Write-Host ""
Write-Host "A log file was created at the following location: $LogFile and the Migration Table file $MigrationTable was update with new values. Please review these files before proceeding to import GPOs."
Write-Host ""