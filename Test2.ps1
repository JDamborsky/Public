function Get-SearchRoot
{
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $root = $domain.GetDirectoryEntry()
    $root.Path
}



function Get-UserLastLogon
{
    [CmdletBinding()]
    param (        
        [String]$UserName
    )
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $root = $domain.GetDirectoryEntry()
    $CurrentSearchRoot = $root.Path

    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
    $searcher.SearchRoot = $CurrentSearchRoot
    #$searcher.PropertiesToLoad.Add("department") > $null
    $result = $searcher.FindOne()
    if ($result -ne $null)
    {
        $resultprop = $result.Properties
        $result.Properties.lastlogontimestamp[0]
       # write-host $result.Properties.Item("lastlogontimestamp")[0]
       # $($result.Properties.Item("lastlogontimestamp")[0])
    }
        
  
}

#$CurrentSearchRoot = Get-SearchRoot
#$CurrentSearchRoot

function Convert-ADDate
{
    [CmdletBinding()]
    param (        
        [String]$ADDate
    )
    $adDate = [DateTime]::FromFileTime($ADDate)
    $adDate
}




$userLastLogon = Get-UserDepartment -UserName "jarled"
$userLastLogon


$adDate = Convert-ADDate -ADDate $userLastLogon
$adDate

