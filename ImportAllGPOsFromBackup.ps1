param([String]$BackupLocation,[String]$MigrationTable,[String]$LogFile);

if ([String]::IsNullOrEmpty($BackupLocation))
{
    write-host -ForegroundColor Red "BackupLocation switch must be specified and point to the folder containing a backup of Group Policy Objects with a corresponding manifest.xml file.";
    write-host -ForegroundColor Red "Exiting Script";
    return;
}
else
{
    $Manifest = $BackupLocation + "\manifest.xml";
}

if ([String]::IsNullOrEmpty($MigrationTable))
{
    write-host -ForegroundColor Red "MigrationTable switch must be specified";
    write-host -ForegroundColor Red "Exiting Script";
    return;
}

if ([String]::IsNullOrEmpty($LogFile))
{
    $LogFile = $BackupLocation + "\ImportAllGPOsLog.txt";
    set-content $LogFile $NULL;
}
else
{
    set-content $LogFile $NULL;
}

import-module ActiveDirectory;
import-module GroupPolicy;

[xml]$ManifestData = get-content $Manifest;

foreach ($GPO in $ManifestData.Backups.BackupInst)

{

    $objectExists = get-gpo $GPO.GPODisplayName."#cdata-section" -ea "SilentlyContinue";
    
    if ($ObjectExists -eq $NULL)
    {
        import-gpo -BackupGPOName $GPO.GPODisplayName."#cdata-section" -TargetName $GPO.GPODisplayName."#cdata-section" -CreateIfNeeded -Path $BackupLocation -MigrationTable $MigrationTable | Out-File $LogFile -append;
        write-host "Import of GPO" $GPO.GPODisplayName."#cdata-section" "was successful."
    }
    else
    {
        $TargetGPOName = "DuplicateGPOonImport - " + $GPO.GPODisplayName."#cdata-section";
        import-gpo -BackupGPOName $GPO.GPODisplayName."#cdata-section" -TargetName $TargetGPOName -CreateIfNeeded -Path $BackupLocation -MigrationTable $MigrationTable | Out-File $LogFile -append;
        write-host "A GPO named" $GPO.GPODisplayName."#cdata-section" "was to be imported but a duplicate name existed."
        write-host "A GPO named DuplicateGPOOnImport -" $GPO.GPODisplayName."#cdata-section" "was created instead. Please investigate the neccessity of this GPO."
    }
    
    $ObjectExists = $NULL 

}
