## Create regex to verify if it is a IP-adress
$regex = [regex]"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"

## Create a string with a IP-adress
$ip = "10.2.2.4"
## Check if the string is a IP-adress
if ($regex.IsMatch($ip)) {
    Write-Host "It is a IP-adress"
} else {
    Write-Host "It is not a IP-adress"
}


## /explain
# Path: Test2.ps1
## Create regex to verify if it is a IP-adress
$regex = [regex]"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"



Function Test-IsValidIpv4
{
    [CmdletBinding()]
    param (        
        [String]$IpString
    )
    $pattern = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"
    
    return $($IpString -match $pattern)
}

## /explain
# Path: Test3.ps1
## Create regex to verify if it is a IP-adress

##

Test-IsValidIpv4 -IpString "10.0.0.255"

// function to calculate days between two dates
function days_between($date1, $date2) {
    $ts1 = New-TimeSpan -Start $date1 -End $date2
    return $ts1.Days
}

// empty function with Begin{}
function Test-Begin
{
    [CmdletBinding()]
    param (        
        [String]$IpString
    )
    Begin
    {
        Write-Host "Begin"
    }
    Process
    {
        Write-Host "Process"
    }
    End
    {
        Write-Host "End"
    }
}



// get searchroot for current domain
function Get-SearchRoot
{
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $root = $domain.GetDirectoryEntry()
    $root.Path
}

$CurrentSearchRoot = Get-SearchRoot
$CurrentSearchRoot

function Get-UserDepartment
{
    [CmdletBinding()]
    param (        
        [String]$UserName
    )
    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
    $searcher.SearchRoot = $CurrentSearchRoot
    $searcher.PropertiesToLoad.Add("department") > $null
    $result = $searcher.FindOne()
    if ($result -ne $null)
    {
        $result.Properties.department
    }
}

$userdep = Get-UserDepartment -UserName "jarled"
$userdep


