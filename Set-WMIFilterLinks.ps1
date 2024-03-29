param([String]$ContainerDN,[String]$BackupLocation,[String]$LogFile);

#$BackupLocation = "C:\Temp\GPO_Backups";

if ([String]::IsNullOrEmpty($LogFile))
{
    $LogFile = "C:\LinkWMIFilters.txt";
    set-content $LogFile $NULL;
}
else
{
    set-content $LogFile $NULL;
}

import-module ActiveDirectory;
import-module GroupPolicy;

$myDomain = [System.Net.NetworkInformation.IpGlobalProperties]::GetIPGlobalProperties().DomainName;
$DomainDn = "DC=" + [String]::Join(",DC=", $myDomain.Split("."));
$SystemContainer = "CN=System," + $DomainDn;
$GPOContainer = "CN=Policies," + $SystemContainer;
$WMIFilterContainer = "CN=SOM,CN=WMIPolicy," + $SystemContainer;

try
{
    if (![System.DirectoryServices.DirectoryEntry]::Exists("LDAP://" + $DomainDN))
    {
        write-host -ForegroundColor Red "Could not connect to LDAP path $DomainDN";
        write-host -ForegroundColor Red "Exiting Script";
        return;
    }
}
catch
{
        write-host -ForegroundColor Red "Could not connect to LDAP path $DomainDN";
        write-host -ForegroundColor Red "Exiting Script";
        return;
}

if ([String]::IsNullOrEmpty($BackupLocation))
{
        write-host -ForegroundColor Red "BackupLocation switch must be specified";
        write-host -ForegroundColor Red "Exiting Script";
        return;
}
else
{
    $Manifest = $BackupLocation + "\manifest.xml";
    [xml]$ManifestData = get-content $Manifest;
}
        
foreach ($item in $ManifestData.Backups.BackupInst)
{
$WMIFilterDisplayName = $NULL;
$GPReportPath = $BackupLocation + "\" + $item.ID."#cdata-section" + "\gpreport.xml";
[xml]$GPReport = get-content $GPReportPath;
$WMIFilterDisplayName = $GPReport.GPO.FilterName;
if ($WMIFilterDisplayName -ne $NULL)
    {
    $GPOName = $GPReport.GPO.Name;
    $GPO = Get-GPO $GPOName;
    $WMIFilter = Get-ADObject -Filter 'msWMI-Name -eq $WMIFilterDisplayName';
    $WMIFilterName = $WMIFilter.Name;
    $GPODN = "CN={" + $GPO.Id + "}," + $GPOContainer;
    $WMIFilterLinkValue = "[$myDomain;" + $WMIFilterName + ";0]";
    Set-ADObject $GPODN -Add @{gPCWQLFilter=$WMIFilterLinkValue};
    write-host "The '$WMIFilterDisplayName' WMI Filter has been linked to the following GPO: $GPOName";
    Add-Content $LogFile "The '$WMIFilterDisplayName' WMI Filter has been linked to the following GPO: $GPOName";
    }
}

write-host "A log file of all WMI filters linked has been save here: $LogFile"