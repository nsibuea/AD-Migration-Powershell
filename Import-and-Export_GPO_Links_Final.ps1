param([String]$Mode,[String]$InputFile,[String]$OutputFile,[String]$LogFile)

if (($InputFile -and $OutputFile) -or ([String]::IsNullOrEmpty($InputFile) -and [String]::IsNullOrEmpty($OutputFile)))
{
write-host -ForegroundColor Red "Either InputFile or OutputFile must be specified and both cannot be specified together"
Write-Host -ForegroundColor Red "Exiting Script"
return
}

if ([String]::IsNullOrEmpty($LogFile))
{
$LogFile = "C:\GPOLinksLog.txt"
}

Set-Content $LogFile $NULL

Import-Module activedirectory

if ($InputFile)
{
    $FileExists = Test-Path $InputFile
    if ($FileExists -eq $false)
    {
    write-host -ForegroundColor Red "Input file does not exist"
    Write-Host -ForegroundColor Red "Exiting Script"
    return
    }

    $thisDomain = Get-ADDomain
    $thisDomainDN = $thisDomain.DistinguishedName
    $header = "objectDN","domainDN","Links"
    $Links = $null
    $Links = import-csv $InputFile -Delimiter "`t" -Header $header
    if ($Links -eq $null)
    {
    write-host -ForegroundColor Red "No input was detected"
    Write-Host -ForegroundColor Red "Exiting Script"
    return
    }

    foreach ($Link in $Links)
    {
    $currentObject = $NULL
    $Link.objectDN = $Link.objectDN -replace $Link.domainDN,$thisDomainDN
    [Array]$objectLinks = $Link.Links.Split("`v")
    $NewLink = $null
    $currentDN = $Link.objectDN
    add-content $LogFile "$currentDN was configured with the following links:"
    foreach ($objectLink in $objectLinks)
        {
        $currentLink = $NULL
        $linkName = $objectLink.TrimEnd("0","1","2")
        $linkName = $linkName.TrimEnd(";")
            if ($linkName)
            {
            add-content $LogFile $linkName
            $currentLink = Get-ADObject -Filter {objectClass -eq "groupPolicyContainer" -and displayName -eq $linkName}
                if ($currentLink -ne $NULL)
                {
                $NewLink = $NewLink + "[LDAP://" + $currentLink + ";" + $objectLink.Substring($objectLink.Length -1,1) + "]"
                }
                else
                {
                Add-Content $LogFile "Error: $linkname does not appear to exist in the destination domain. Please re-import it or create a new GPO with the same name."
                }
            }
        }
        try
        {
        $currentObject = Get-ADObject $Link.objectDN -Properties gpLink
        $currentObject.gpLink = $NewLink
            if ($NewLink)
            {
            Set-ADObject $Link.objectDN -Replace @{gpLink = $NewLink}
            }
            else
            {
            $currentObjectDN = $Link.objectDN
            Add-Content $LogFile "Error: It appears none of the GPO's previously linked to '$currentObjectDN' exist. Please re-import the GPO's to the destination domain."
            }
        }
        catch
        {
        $currentObjectDN = $Link.objectDN
        add-content $LogFile "Error: $currentObjectDN does not exist. Create the object and try again."
        }
            if ($NewLink)
            {
            add-content $LogFile "gPLink will be set to: $NewLink"
            }
            else
            {
            Add-Content $LogFile "gPLink will not be modified on this object."
            }
        add-content $LogFile "---END---"
    }
    write-host -ForegroundColor Yellow "A log file has been saved at $LogFile"
}
else
{
    set-content $OutputFile $null
    $Links = $NULL
    $thisDomain = Get-ADDomain
    $thisDomainDN = $thisDomain.DistinguishedName
    $thisDomainConfigurationPartition = "CN=Configuration," + $thisDomainDN

    $Links = Get-ADObject -Filter {gpLink -LIKE "[*]"} -Properties gpLink
    $Links += Get-ADObject -Filter {gpLink -LIKE "[*]"} -Searchbase $thisDomainConfigurationPartition -Properties gpLink

    $NewLine = $null
    if ($Links)
    {
        foreach ($Link in $Links)
        {
        $NewLine = $null
        $LinkList = $Link.gpLink.Split('\[|\]')

	        foreach ($LinkItem in $LinkList)
	        {
		        if ($LinkItem)
		        {
		        $LinkSplit = $LinkItem.Split(";")
		        $LinkItem = $LinkItem.TrimStart("LDAP://")
		        $LinkItem = $LinkItem.TrimEnd(';0|;1|;2')
		        $LinkItem = get-adobject $LinkItem -Properties displayName
		        $NewLine = $NewLine + $LinkItem.DisplayName + ";" + $LinkSplit[1] + "`v"
		        }
	        }
	
	        $NewLine = $Link.DistinguishedName + "`t" + $thisDomainDN + "`t" + $NewLine
	        add-content $OutputFile $NewLine
        }
            write-host -ForegroundColor Yellow "The output file has been saved at $OutputFile"
    }
    else
    {
        write-host -ForegroundColor Red "No GPO Links exist in this domain"
        write-host -ForegroundColor Red "Exiting script"
        Set-Content $OutputFile $NULL
        return
    }

}