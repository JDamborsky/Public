Function AddRemoveGroupMember
{
    param (
        [Parameter(Mandatory = $true)]
        [String]$GroupDN,
        [Parameter(Mandatory = $true)]        
        [String]$ItemDN,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Add","Remove")]
		[String]$AddOrRemove,
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PSCredential]$Credentials
    )

    $AdObjectStatus = "OK"
	
	Try
	{
		$CredUser       = $($Credentials.GetNetworkCredential().Username)
		$CredentialsOK  = $True
	}
	catch
	{
		$CredentialsOK  = $False
	}
	
	
	if ($CredentialsOK)
	{				
		Try
		{
			$GroupObject    = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList "LDAP://$GroupDN", "$($Credentials.GetNetworkCredential().Domain)\$($Credentials.GetNetworkCredential().Username)", "$($Credentials.GetNetworkCredential().Password)"

		}
		catch
		{
			$AdObjectStatus = "Group does not exists, or unavailable (cred);"
		}	
	}
	else
	{
		Try
		{
			$GroupObject    = ([ADSI]"LDAP://$GroupDN")
		}
		catch
		{
			$AdObjectStatus = "Group does not exists, or unavailable;"
		}
	}
	
	# Verify if changes are needed at all
	If ($AdObjectStatus -eq "OK")
	{
        $VerifyItemObject           = ([ADSI]"LDAP://$ItemDN")
        $VerifyGroupObject          = ([ADSI]"LDAP://$GroupDN")
        [bool]$IsItemMemberofGroup  = $VerifyGroupObject.IsMember($VerifyItemObject.ADsPath)

		If (($AddOrRemove -eq "Add")    -and ($IsItemMemberofGroup))    { $AdObjectStatus = "<Allready member>" }
		If (($AddOrRemove -eq "Remove") -and (!$IsItemMemberofGroup))   { $AdObjectStatus = "<Not member>" }
	}
		
		# Perform Change
    If (($AdObjectStatus -eq "OK") -and ($AddOrRemove -eq "Add"))
    {
        try {  
            $tmpRes         = $GroupObject.Add("LDAP://$ItemDn") 
            
        } Catch {
            $AdObjectStatus +=  $($_.Exception.Message).Replace('"','')            
        }    
    }

    If (($AdObjectStatus -eq "OK") -and ($AddOrRemove -eq "Remove"))
    {
        try {  
            $tmpRes         = $GroupObject.Remove("LDAP://$ItemDn") 
            
        } Catch {
            $AdObjectStatus = $($_.Exception.Message).Replace('"','')            
        }    
    }
    
    Return $AdObjectStatus
}


#
#$Group = [adsi]::new("LDAP://$($Domain)$($GroupPath)", $CredsUserName, $CredsPassword)


Function Get-AdNetInfoForIP
	{
		[CmdletBinding()]
		param (
			[Parameter()]
			[IPAddress]$Ip4Adress,
            [String]$DomainShortName=""
		)

        $FoundAdNetId   	= ""

        if ($DomainShortName -ne "")
        {
            $DomainLdap     = Get-LdapFromDomainShortName -DomainShortnameToFind $DomainShortName
            $DomainDN       = $DomainLdap.replace("LDAP://","")        
            $subnetsDN     	= "LDAP://CN=Subnets,CN=Sites,CN=Configuration,$DomainDN"

        }
        else {
            $DomainDN      	= $([adsi] "LDAP://RootDSE").Get("rootDomainNamingContext")
            $subnetsDN     	= "LDAP://CN=Subnets,CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
        }
		
        $DomainShortName 	= $DomainDN.Split(",")[0].Trim('DC=')
		$FoundAdNetId       = ""
		
		foreach ($subnet in $([adsi]$subnetsDN).psbase.children)
		{
           	$CurrNetAdr     = ([IPAddress](($subnet.cn -split "/")[0]))
			$CurrAdSn       = ([IPAddress]"$([system.convert]::ToInt64(("1" * [int](($subnet.cn -split "/")[1])).PadRight(32, "0"), 2))")
			if ((([IPAddress]$Ip4Adress).Address -band ([IPAddress]$CurrAdSn).Address) -eq ([IPAddress]$CurrNetAdr).Address)
			{
                $CurrentFoundAdNetId   = $($subnet.cn)
                if ($FoundAdNetId -eq "") {$FoundAdNetId = $CurrentFoundAdNetId}
	            
                If (($([System.Convert]::ToInt32($(($CurrentFoundAdNetId -split "/")[1])))) -ge ($([System.Convert]::ToInt32($(($FoundAdNetId -split "/")[1])))))
                {
    
                    $FoundAdNetId = $CurrentFoundAdNetId
                    $site = [adsi] "LDAP://$($subnet.siteObject)"
                    if ($site.cn -ne $null)
                    {
                        $siteName   		= ([string]$site.cn).toUpper()
                        $siteDescription	= ([string]$site.description)					
                    }
                    
                    $SubNetDescription  = $subnet.description[0]
                    $SubNetLocation     = $subnet.Location[0]
                    $AdSiteForAdress    = @{
                        ip                  = "$Ip4Adress"
                        sn                  = "$CurrAdSn"
                        AdCidr              = "$FoundAdNetId"
                        AdSiteName          = "$siteName"
                        AdSiteDesciption    = "$siteDescription"
                        SubNetDescription   = "$SubNetDescription"
                        SubNetLocation      = "$SubNetLocation"
                        DomainName          = "$DomainShortName"
                        Isfound             = $True
                    }
                #	Break				
                    
                }
			}
		}
		if ($FoundAdNetId -eq "")
		{
			$AdSiteForAdress = @{
				ip                  = "$Ip4Adress"
				sn                  = ""
				AdCidr              = ""
				AdSiteName          = ""
				AdSiteDesciption    = ""
				SubNetDescription   = ""
				SubNetLocation      = ""
                DomainName          = ""
				Isfound             = $False
			}
		}
		$FoundAdNetIdObject = [pscustomobject]$AdSiteForAdress
		Return $FoundAdNetIdObject		
	}


    Function Get-AdObjectProperties
    {
        # Example:  $Result = Get-AdObjectProperty -ObjValueString "testing" -ObjParameter "samaccountname" -objectcategory User -ReturnProperty distinguishedname -DomainShortName 'SIKT' -AdsiSearchString ''
        param (
            [Parameter()]
            [String]$ObjValueString,
            [String]$ObjParameter="samaccountname",
            [ValidateSet("User","person","Computer","Group","All")]
            [String]$objectcategory,
            [String]$ReturnProperty='path',
            [String]$DomainShortName=[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,
            [String]$AdsiSearchString=''
        )
        $ReturnValue            = "Not Found"
        $ObjParameter           = $ObjParameter.ToLower()
        $ReturnProperty         = $ReturnProperty.ToLower()
        if ($objectcategory -eq 'User') {$objectcategory = 'person'}
    
        If ($AdsiSearchString -eq '')
        {
            $AdsiSearchString          =   Get-LdapFromDomainShortName -DomainShortnameToFind  $DomainShortName
        }
        
        If ($objectcategory -ne "All")
        {
            $searcher               =   [adsisearcher]"((objectClass=$objectcategory))" 
        }
        else {
            $searcher               =   [adsisearcher]"((objectClass=*))" 
        }
    
        $searcher.searchRoot    =   [adsi]$AdsiSearchString
    
        If ($ReturnProperty -ne '')
        {
            $ReturnPropertyArray = $ReturnProperty -split ","
            foreach ($ReturnPropertyItem in $ReturnPropertyArray)
            {
                $Searcher.PropertiesToLoad.Add("$ReturnPropertyItem") 
            }
        }
        
        $searcher.Filter        =   "(&(objectClass=$objectcategory)($ObjParameter=$ObjValueString))" 
        $AdSearcResult          =   $Searcher.FindAll()
        If ($AdSearcResult.count -gt 0)
        {
            If ($ReturnProperty -eq 'path')
            {   $ReturnValue = $AdSearcResult.Properties.adspath  }
            else 
            {   $ReturnValue = $AdSearcResult.Properties.$ReturnProperty[0]   }
        }
        
        Return $ReturnValue
    }
    

Function Get-AdObjectProperty
{
    # Example:  $Result = Get-AdObjectProperty -ObjValueString "testing" -ObjParameter "samaccountname" -objectcategory User -ReturnProperty distinguishedname -DomainShortName 'SIKT' -AdsiSearchString ''
    param (
        [Parameter()]
        [String]$ObjValueString,
        [String]$ObjParameter="samaccountname",
        [ValidateSet("User","person","Computer","Group","All")]
        [String]$objectcategory="All",
        [String]$ReturnProperty='path',
        [String]$DomainShortName=[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,
        [String]$AdsiSearchString=''
    )
    $ReturnValue            = "Not Found"
    $ObjParameter           = $ObjParameter.ToLower()
    $ReturnProperty         = $ReturnProperty.ToLower()
    if ($objectcategory -eq 'User') {$objectcategory = 'person'}

    If ($AdsiSearchString -eq '')
    {
        $AdsiSearchString          =   Get-LdapFromDomainShortName -DomainShortnameToFind  $DomainShortName
    }
    If ($objectcategory -eq "All") 
    {
        $searcher               =   [adsisearcher]"(&($ObjParameter=$ObjValueString))"
    }
    else {
        $searcher               =   [adsisearcher]"(&(objectClass=$objectcategory)($ObjParameter=$ObjValueString))"
    }
    #$searcher               =   [adsisearcher]"((objectcategory=$objectcategory)($ObjParameter=$ObjValueString))"  
    $searcher.searchRoot    =   [adsi]$AdsiSearchString
    $searcher.PropertiesToLoad.AddRange(($ReturnProperty))
    #$searcher.Filter        =   "(&(objectCategory=$objectcategory)($ObjParameter=$ObjValueString))" 
    $AdSearcResult          =   $Searcher.FindAll()
    If ($AdSearcResult.count -gt 0)
    {
        If ($ReturnProperty -eq 'path')
        {   $ReturnValue = $AdSearcResult.Properties.adspath  }
        else 
        {   $ReturnValue = $AdSearcResult.Properties.$ReturnProperty[0]   }
    }
    
    Return $ReturnValue
}

Function Get-AdObjectDn
{
    #Example:   $result = Get-AdObjectDn -ObjValueString "testing" -ObjParameter "samaccountname" -objectcategory User     -DomainShortName 'sikt'
    param (
        [Parameter()]
        [String]$ObjValueString,
        [String]$ObjParameter="samaccountname",
        [ValidateSet("User","person","Computer","Group")]
        [String]$objectcategory,
        [String]$DomainShortName=[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,
        [String]$AdsiSearchString=''
    )
    $ObjectDn               = "Not Found"
    $ObjParameter           = $ObjParameter.ToLower()
    if ($objectcategory -eq 'User') {$objectcategory = 'person'}

    If ($AdsiSearchString -eq '')
    {
        $AdsiSearchString          =   Get-LdapFromDomainShortName -DomainShortnameToFind  $DomainShortName
    }
    
    $searcher               =   [adsisearcher]"((objectClass=$objectcategory))" 
    $searcher.searchRoot    =   [adsi]$AdsiSearchString
    $searcher.Filter        =   "(&(objectClass=$objectcategory)($ObjParameter=$ObjValueString))" 
    $AdSearcResult          =   $Searcher.FindAll()
    If ($AdSearcResult.count -gt 0)
    {
        $ObjectDn = $AdSearcResult.Properties.distinguishedname[0]
    }
    
    Return $ObjectDn
}


function Get-AdsiPath
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [Alias("Domain")]
        [String]$DomainDN=$(([AdsiSearcher]"").SearchRoot.Path)
    )

    if ($PSBoundParameters['DomainDN'])
    {
        if ($DomainDN -notlike "*DC=*")
        {
            $DomainDN = Get-LdapFromDomainShortName $DomainDN
        }
        if ($DomainDN -notlike "LDAP://*") {$DomainDN = "LDAP://$DomainDN"}
    }
    return $DomainDN 
}


function Get-AdsiPathForCurrentDomain
{
  $Root = [ADSI]"LDAP://RootDSE"
  $GetAdsiPathStr = 'LDAP://' + $Root.rootDomainNamingContext
  return $GetAdsiPathStr
}


Function Get-GroupsForDomain
{
    #  Example:     $ReturnTable = Get-GroupsForDomain -DomainShortname 'SIKT' -SearchString 'R ISA*'
    #  Example:     $ReturnTable = Get-GroupsForDomain -DomainShortname 'SIKT' -SearcProperty "description" -SearchString '*Ikke tilgang til å endre regler*'


    param (
            [string]$DomainShortname,
            [ValidateSet("name","cn","description","sAMAccountName")]
            [String]$SearchProperty="name",
            [string]$SearchString='')   


    # Create Table to store result
    $GroupTable                     = new-object Data.datatable 
    $GroupTable.Columns.Add("IDx",                      "System.int32")     | Out-Null
    $GroupTable.Columns.Add("DomainName",               "System.String")    | Out-Null
    $GroupTable.Columns.Add("GroupName",                "System.String")    | Out-Null    
    $GroupTable.Columns.Add("GroupCN",                  "System.String")    | Out-Null 
    $GroupTable.Columns.Add("GroupDescription",         "System.String")    | Out-Null    
    $GroupTable.Columns.Add("GroupDistinguishedName",   "System.String")    | Out-Null
    $GroupTable.Columns.Add("GroupsAMAccountName",      "System.String")    | Out-Null
    $GroupTable.Columns.Add("GroupScope",               "System.String")    | Out-Null
    $GroupTable.Columns.Add("GroupType",                "System.String")    | Out-Null
    #$GroupTable.Columns.Add("Processed",            "System.Boolean")   | Out-Null        
    
    # Search AD with ADSI    
    $AdsiSearchList             =   $Null
    $AdsiPathString             =   Get-AdsiPath     $($DomainShortname)
    
    If ($SearchString -eq '')
    { $AdsiSearcher               =   [adsisearcher]"(&(objectCategory=Group)(objectClass=Group))" }
    else 
    { $AdsiSearcher               =   [adsisearcher]"(&(objectCategory=Group)(objectClass=Group)($SearchProperty=$SearchString))"    }
    

    $AdsiSearcher.PropertiesToLoad.AddRange(('name', 'cn', 'description','distinguishedName', 'sAMAccountName', 'grouptype', 'memberof'))
    $AdsiSearcher.searchRoot    =   [ADSI]$AdsiPathString
    $AdsiSearcher.SizeLimit     =   12
    $AdsiSearcher.PageSize      =   11
    $AdsiSearchList             =   $AdsiSearcher.FindAll()

    $Index = 0
    foreach ($GroupItem in $AdsiSearchList)
    {
     #   $GroupItem.properties | format-table
    
        Switch ($($groupItem.Properties.grouptype) )  
       {
            2
                   {    $GroupScope     = "Global"
                        $GroupType      = "Distribution"    }
            4
                   {    $GroupScope     = "Domain local"
                        $GroupType      = "Distribution"    }
            8
                   {    $GroupScope     = "Universal"
                        $GroupType      = "Distribution"    }
            -2147483646
                   {    $GroupScope     = "Global"
                        $GroupType      = "Security"        }
            -2147483644
                   {    $GroupScope     = "Domain local"
                        $GroupType      = "Security"        }
            -2147483640
                   {    $GroupScope     = "Universal"
                        $GroupType      = "Security"        }
            Default
                    {   $GroupScope     = ""
                        $GroupType      = ""                }
        }
        
        $Index++
        $GDR = $GroupTable.NewRow()
        $GDR.Item(0) = $Index
        $GDR.Item(1) = $DomainShortname
        $GDR.Item(2) = $groupItem.Properties.name[0]
        $GDR.Item(3) = $groupItem.Properties.cn[0]

        if ($groupItem.Properties.description)
        {   $GDR.Item(4) = $groupItem.Properties.description[0] }
        else 
        {   $GDR.Item(4) = ''                                   }

        $GDR.Item(5) = $groupItem.Properties.distinguishedname[0].ToString()
        $GDR.Item(6) = $groupItem.Properties.samaccountname[0]
        $GDR.Item(7) = $GroupScope
        $GDR.Item(8) = $GroupType
        $GroupTable.rows.add($GDR)

    }

    $ReturnObject = [Data.datatable]$GroupTable
   
   Return $ReturnObject

}


function Get-MemberOfDN
{
    #  Example  Get-NestedMemberOf -ObjectType User -ObjectName "jardam" -DomainShort "AD"

    param (
        [Parameter()]
        [ValidateSet("Machine","User","Group")]
        [String[]]$ObjectType,
        [String[]]$ObjectName,
        [String[]]$DomainShort  )
    
    $AdsiPathString             =   Get-LdapFromDomainShortName -DomainShortnameToFind $($DomainShort)

    Switch ($ObjectType)
    {
        "Machine"   {$AdsiSearcher  =   [adsisearcher]"(&(objectCategory=computer)(cn=$ObjectName))"}
        "User"      {$AdsiSearcher  =   [adsisearcher]"(&(objectCategory=Person)(objectClass=user)(sAMAccountName=$ObjectName))"}
        "Group"     {$AdsiSearcher  =   [adsisearcher]"(&(objectCategory=Group)(objectClass=Group)(cn=$ObjectName))"}
    }
    
    $AdsiSearcher.PropertiesToLoad.AddRange(('name','memberof'))
    $AdsiSearcher.searchRoot    =   [ADSI]$AdsiPathString
    $AdsiSearchList             =   $AdsiSearcher.FindAll()
    $ActiveGroupName            =   ''
    If ($AdsiSearchList.count -gt 0)
    {        
        foreach ($GroupItem in $AdsiSearchList.properties.memberof)
        {
            $GroupArray += @([PsCustomObject]@{DN=$GroupItem;MemberType='Direct';Executed=0})
        }

        $GroupArray = $GroupArray | Sort-Object DN -Unique
        
    }
    else {
        $GroupArray += @([PsCustomObject]@{Name=$ObjectName;DN='NA';MemberType='NotFound';Executed=0})
        
    }
    
    return  $GroupArray  
}



function Get-NestedMemberOf 
{
    #  Example  Get-NestedMemberOf -ObjectType User -ObjectName "jardam" -DomainShort "AD"

    param (
        [Parameter()]
        [ValidateSet("Machine","User","Group")]
        [String[]]$ObjectType,
        [String[]]$ObjectName,
        [String[]]$DomainShort  )
    
    
    $AdsiPathString      =   Get-AdsiPath  $($DomainShort)
    Switch ($ObjectType)
    {
        "Machine"   {$AdsiSearcher  =   [adsisearcher]"(&(objectCategory=computer)(cn=$ObjectName))"}
        "User"      {$AdsiSearcher  =   [adsisearcher]"(&(objectCategory=Person)(objectClass=user)(sAMAccountName=$ObjectName))"}
        "Group"     {$AdsiSearcher  =   [adsisearcher]"(&(objectCategory=Group)(objectClass=Group)(cn=$ObjectName))"}
    }
    
    $AdsiSearcher.PropertiesToLoad.AddRange(('name','memberof'))
    $AdsiSearcher.searchRoot    =   [ADSI]$AdsiPathString
    $AdsiSearchList             =   $AdsiSearcher.FindAll()
    $ActiveGroupName            =   ''
    If ($AdsiSearchList.count -gt 0)
    {
        
        foreach ($GroupItem in $AdsiSearchList.properties.memberof)
        {
            $GroupArray += @([PsCustomObject]@{Name=$ActiveGroupName;DN=$GroupItem;MemberType='Direct';Executed=0})
        }

        $NestedRemains      = 1
        while ($NestedRemains -eq 1)
        {
            $NestedRemains = 0
            foreach ($GroupArrayItem in $GroupArray)
            {
                if ($GroupArrayItem.Executed -eq 0)
                {
                    $GroupDN                        =   $GroupArrayItem.DN    
                    $AdsiSubSearcher                =   [adsisearcher]"(&(objectCategory=Group)(objectClass=Group)(distinguishedName=$GroupDN))"
                    $AdsiSubSearcher.PropertiesToLoad.AddRange(('name','memberof'))
                    $AdsiSubSearcher.searchRoot     =   [ADSI]$AdsiPathString
                    $AdsiSubSearchList              =   $AdsiSubSearcher.FindAll()
                    foreach ($GroupSubItem in $AdsiSubSearchList.properties.memberof)
                    {
                        $NestedRemains = 1
                        if ($GroupArray.DN -notcontains $GroupSubItem)
                        {
                            $GroupArray += @([PsCustomObject]@{Name=$ActiveGroupName;DN=$GroupSubItem;MemberType='Nested';Executed=0})
                        }    
                    }
                    $GroupArrayItem.Executed        =   1
                }    
            }
        }
        $GroupArray = $GroupArray | Sort-Object DN -Unique
        
        $GroupArray | foreach-object {
        $_.Name = Get-GroupnameFromDn $_.DN $DomainShort
        }
    }
    else {
        $GroupArray += @([PsCustomObject]@{Name=$ObjectName;DN='NA';MemberType='NotFound';Executed=0})
        
    }
    
    return  $GroupArray  
}


function Get-GroupMembersNested
{
    param (
        [String]$GroupName,
        [String]$DomainShort,
        [ValidateSet("Computer","Person","Group","All")]
        [String[]]$ObjectType="All",
        [Bool]$Nested=$true  )
    
    
    $AdsiPathForDomain          =   Get-AdsiPath $($DomainShort)
    $AdsiSearcher               =   [adsisearcher]"(&(objectCategory=Group)(objectClass=Group)(cn=$GroupName))"    
    $AdsiSearcher.searchRoot    =   [ADSI]$AdsiPathForDomain
    $AdsiSearchList             =   $AdsiSearcher.FindAll()

    $ObjArray += @([PsCustomObject]@{Name=$($AdsiSearchList.properties.name);DN=$AdsiSearchList.properties.distinguishedname[0];MemberType='Root';Category="Group";Executed=1})

    foreach ($ObjectItem in $AdsiSearchList.properties.member)
    {   
        $AdObjectItem   = [ADSI]"LDAP://$ObjectItem"
        $Category       = $($($AdObjectItem.objectCategory).Split(",")[0]).Replace("CN=","")

        $ObjArray += @([PsCustomObject]@{Name=$($AdObjectItem.name);DN=$ObjectItem;MemberType='Direct';Category=$Category;Executed=0})
    }

    If ($Nested) {$NestedRemains = 1} else {$NestedRemains = 0}
    
    while ($NestedRemains -eq 1)
    {
        $NestedRemains      = 0
        foreach ($ObjectItem in $ObjArray | where { ($_.Category -eq "Group") -and ($_.Executed -eq 0) })
        {
            $ObjectItem.Executed    = 1
            $AdObjectItem           = [ADSI]"LDAP://$($ObjectItem.DN)"

            foreach ($MemberObjectItem in $AdObjectItem.properties.member)
            {
                if ($ObjArray.DN -notcontains $MemberObjectItem)
                {
                    $MemberAdObjectItem     = [ADSI]"LDAP://$MemberObjectItem"
                    $Category               = $($($MemberAdObjectItem.objectCategory).Split(",")[0]).Replace("CN=","")
                    If ($Category -eq "Group")  {   $NestedRemains  =   1   } 
                    
                    $ObjArray += @([PsCustomObject]@{Name=$($MemberAdObjectItem.name);DN=$MemberObjectItem;MemberType='Nested';Category=$Category;Executed=0})
                }                
            }
        }        
    }

    If ($ObjectType -ne "All")    {$ObjArray = $ObjArray | where { $_.Category -eq $ObjectType} }
    $ObjArray     = $ObjArray | select-object Name, Dn, Membertype, Category |  Sort-Object DN -Unique
    
    return  $ObjArray  
}


Function Get-GroupnameFromDn
{
    param (
        $GroupDN,
        $DomainShortname  )

        $GroupnameFound     = $false
        $RetryCount         = 0
        $DumpToTable        = $null
        $Groupname          = $null

        while ((!$GroupnameFound) -and ($RetryCount -lt 2))
        {
            $RetryCount++
            if ($RetryCount -eq 1)
            {
                #$GroupDN = $GroupDN -replace '\+', '+'
                $GroupDN = ConvertCharsToAdsiSearcher $GroupDN
            }

            $AdsiPathForDomain      =   Get-AdsiPath $($DomainShortname)
            $searcher               =   [adsisearcher]"(&(objectClass=Group)(distinguishedName=$GroupDN))" 
            $searcher.searchRoot    =   [adsi]$AdsiPathForDomain
            $Searcher.SizeLimit     =   12
            $Searcher.PageSize      =   11
            $searcher.PropertiesToLoad.AddRange(('name','samaccountname'))    
            try {
                $DumpToTable            = $Searcher.FindAll() 
            }
            catch {
                $DumpToTable        = 'fail'
            }
            
             
            if ($DumpToTable -ne 'fail')
            {
                foreach ($groupItem in $DumpToTable)
                {
                    $Groupname      = $groupItem.Properties.name
                    $GroupnameFound = $true            
                }
            }
        }

        if ($RetryCount -gt 1)
        {
            write-host " Group not found ->$GroupDN        "            
        }

        return $Groupname
}

function ConvertCharsToAdsiSearcher
{
    param (
            [string]$AdsiSearchString,
            [Bool]$Reverse=$False   )
#   Adsisearcher does fail with some characters

    If ($Reverse -eq $False)
    {
     #   $AdsiSearchString = $AdsiSearchString -replace '\+', '+'
        $AdsiSearchString = $AdsiSearchString -replace '[#(]', '\28'
        $AdsiSearchString = $AdsiSearchString -replace '[#)]', '\29'
    }
    else 
    {
        $AdsiSearchString = $AdsiSearchString.replace('\28','(')
        $AdsiSearchString = $AdsiSearchString.replace('\29',')')
    }

              #  $GroupName = $GroupName -replace ',', '\,'
              #  $GroupName = $GroupName -replace 'ø', '\f8'

    $ReturnString = $AdsiSearchString
    Return $ReturnString
}


Function Get-ClientsInDomain
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$DomainFQDN
    )
    $CurrentDomainId        =   0   # Variable not used ....

    $AdsiString =   Get-LdapFromDomainName $DomainFQDN

    $Counter                =   0
    $StatusCounter          =   1000             
    $searcher               =   [adsisearcher]"((objectcategory=computer))" 
    $searcher.searchRoot    =   [adsi]$AdsiString
    $Searcher.SizeLimit     =   12
    $Searcher.PageSize      =   11
    #$searcher.Filter        =   "(&(objectCategory=computer)(!operatingSystem=*server*)(operatingSystem=*Windows*)(name=$ClientName))" 
    $searcher.Filter        =   "(&(objectCategory=computer)(!operatingSystem=*server*)(operatingSystem=*Windows*))" 
    $searcher.PropertiesToLoad.AddRange(('name','samaccountname','operatingSystem','operatingSystemVersion','lastLogon','memberOf','useraccountcontrol','pwdLastSet','distinguishedname')) 
    $DumpToTable            =   $Searcher.FindAll()

    $dt = new-object Data.datatable 
    $DR = $DT.NewRow()
    $DT.Columns.Add("ID","System.int32")                          | Out-Null
    $DT.Columns.Add("ClientIDX","System.int32")                   | Out-Null
    $DT.Columns.Add("name","System.String")                       | Out-Null
    $DT.Columns.Add("samaccountname","System.String")             | Out-Null
    $DT.Columns.Add("DomainId","System.int32")                    | Out-Null
    $DT.Columns.Add("dn","System.String")                         | Out-Null
    $DT.Columns.Add("operatingSystem","System.String")            | Out-Null
    $DT.Columns.Add("operatingSystemVersion","System.String")     | Out-Null
    $DT.Columns.Add("memberOf","System.String")                   | Out-Null
    $DT.Columns.Add("foretak","System.String")                    | Out-Null
    $DT.Columns.Add("ClientType","System.String")                 | Out-Null
    $DT.Columns.Add("lastLogon","System.DateTime")                | Out-Null
    $DT.Columns.Add("pwdLastSet","System.DateTime")               | Out-Null
    $DT.Columns.Add("Enabled","System.Boolean")                   | Out-Null   
    $DT.Columns.Add("Executed","System.Boolean")                  | Out-Null   
    
    foreach ($ClientItem in $DumpToTable)
    {    
        $AddTotable     = $True
        $DnStr          = $ClientItem.Properties.distinguishedname[0]
        If ($DnStr.Contains('OU=Linux'))    {$AddTotable     = $False}
        If ($DnStr.Contains('OU=Servere'))  {$AddTotable     = $False}
        If ($DnStr.Contains('OU=Citrix'))   {$AddTotable     = $False}

        If ($AddTotable -eq $true)
        {
            
            $Counter++
            if ($Counter -eq $StatusCounter)
            {
                write-host "Counter = $Counter"    
                $StatusCounter = $StatusCounter + 1000
            }

            $MemberOfArray              = $ClientItem.Properties.Item("memberOf")
            $MemberOfString             = $MemberOfArray -join ";"
            
            $lastLogonL                 = $ClientItem.Properties.Item("lastLogon")[0]            
            If ($lastLogonL)       
            {   
            # $lastLogonL             = 0 
                $lastLogonDate          = [DateTime]$lastLogonL
                $lastLogon              = $lastLogonDate.AddYears(1600).ToLocalTime()
            }
            else 
            {   $lastLogon              = $null        }

            $pwdLastSetL                = $ClientItem.Properties.Item("pwdLastSet")[0]            
            If ($pwdLastSetL)      
            { 
            # $pwdLastSetL            = 0 
                $pwdLastSetDate         = [DateTime]$pwdLastSetL
                $pwdLastSet             = $pwdLastSetDate.AddYears(1600).ToLocalTime()
            }
            else 
            {   $pwdLastSet             = $null        }
            
            if ([string]$ClientItem.properties.useraccountcontrol -band 2)
            {   $IsObjectEnabledbit = 0 }
            else
            {   $IsObjectEnabledbit = 1 }

            $Index++
            $DR = $DT.NewRow()
            $DR.Item(0)     = $Index
            $DR.Item(1)     = 0
            $DR.Item(2)     = $ClientItem.Properties.name[0]
            $DR.Item(3)     = $ClientItem.Properties.samaccountname[0]
            $DR.Item(4)     = $CurrentDomainId
            $DR.Item(5)     = $ClientItem.Properties.distinguishedname[0]
            $DR.Item(6)     = $ClientItem.Properties.operatingsystemversion[0]
            $DR.Item(7)     = $ClientItem.Properties.operatingsystem[0]
            $DR.Item(8)     = $MemberOfString
            $DR.Item(9)     = ''
            $DR.Item(10)    = ''
            if ($lastLogon)     { $DR.Item(11)    = $lastLogon  }
            if ($pwdLastSet)    { $DR.Item(12)    = $pwdLastSet }
            $DR.Item(13)    = $IsObjectEnabledbit
            $DR.Item(14)    = 0

            $dt.rows.add($DR)
        }
    
  
    }



    return $dt
}


function Get-LdapFromDomainName
  {
      param (
          [string]$DomainNameToFind
      )
      $DomainArray = @()
      
      $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
      $DomainArray += $CurrentDomain
      
      $TrustedDomainsList = $CurrentDomain.GetAllTrustRelationships()
      foreach ($TrustedDomain in $TrustedDomainsList)
      {
          $DomainArray += $TrustedDomain.Targetname
      }
      
      $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
      $TrustedForestList = $Forest.GetAllTrustRelationships()
      foreach ($TrustedForest in $TrustedForestList)
      {
          $DomainArray += $TrustedForest.Targetname
      }
      
      foreach ($DomainItems in $DomainArray)
      {
          if ($DomainItems.name)
          {$DomainItems = $DomainItems.name}

          if ($DomainItems -like "*$DomainNameToFind*")
          {
              [String]$DomainName = $DomainItems.tostring()
              $LdapStr = "LDAP://"
              $DomainNameParts = $DomainName -split '\.'
              foreach ($DomainNamePart in $DomainNameParts)
              {
                  $LdapStr = $LdapStr + "DC=$DomainNamePart,"
              }
              $LdapStr = $LdapStr.Substring(0, $LdapStr.length - 1)
          }
      }
      
      Return $LdapStr
  }

  Function Get-LdapSearchRootFromDN
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$DNstr
    )
    $LdapSearchRoot = ''

    $DNstrArray = $($DNstr -split ',')
    ForEach ($DNstrItem in $DNstrArray)
    {
        if ($DNstrItem -like 'DC=*')
        {
            $LdapSearchRoot += $DNstrItem + ','
        }
    }
    if ($LdapSearchRoot.length -gt 0) 
    {
        $LdapSearchRoot = "LDAP://" + $($LdapSearchRoot.Substring(0, $LdapSearchRoot.Length - 1))   

    }
    Return $LdapSearchRoot
}



function Get-LdapFromDomainShortName
  {
      param (
          [string]$DomainShortnameToFind
      )
      $DomainArray = @()
      
      $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
      $DomainArray += $CurrentDomain
      
      $TrustedDomainsList = $CurrentDomain.GetAllTrustRelationships()
      foreach ($TrustedDomain in $TrustedDomainsList)
      {
          $DomainArray += $TrustedDomain.Targetname
      }
      
      $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
      $TrustedForestList = $Forest.GetAllTrustRelationships()
      foreach ($TrustedForest in $TrustedForestList)
      {
          $DomainArray += $TrustedForest.Targetname
      }
      
      foreach ($DomainItems in $DomainArray)
      {
          if ($DomainItems -like "*$DomainShortnameToFind*")
          {
              [String]$DomainName = $DomainItems.tostring()
              $LdapStr = "LDAP://"
              $DomainNameParts = $DomainName -split '\.'
              foreach ($DomainNamePart in $DomainNameParts)
              {
                  $LdapStr = $LdapStr + "DC=$DomainNamePart,"
              }
              $LdapStr = $LdapStr.Substring(0, $LdapStr.length - 1)
          }
      }
      
      Return $LdapStr
  }


 
Function Get-PropFromDn
{
    param (
        [String]$DN,
        [String]$PropertyToReturn="samaccountname"  )

        $ItemFound          = $false
        $RetryCount         = 0
        $DumpToTable        = $null
        $ItemProp           = $null

        while ((!$ItemFound) -and ($RetryCount -lt 2))
        {
           
            if ($RetryCount -eq 1)
            {
                #$GroupDN = $GroupDN -replace '\+', '+'
                $DN = ConvertCharsToAdsiSearcher $DN
            }
            $RetryCount++
            $searcher               = [adsisearcher]"(&(distinguishedName=$DN))" 
            $searcher.PropertiesToLoad.AddRange(($PropertyToReturn))    
            try {
                $DumpToTable            = $Searcher.FindAll() 
            }
            catch {
                $DumpToTable        = 'fail'
            }
             
            if ($DumpToTable -ne 'fail')
            {
                foreach ($Item in $DumpToTable)
                {
                    $ItemProp      = $($Item.Properties.$PropertyToReturn)
                    $ItemFound = $true            
                }
            }
        }  
        return $ItemProp
} 



Function Get-TrustedDomainsSIDandDNS
{
    $DomainSIDList = @{}

    $CurrDomainObj      = [ADSI]$(([AdsiSearcher]"").SearchRoot.Path)
    $CurrDomainSid      = (New-Object System.Security.Principal.SecurityIdentifier ($CurrDomainObj.objectsid.value, 0)).Value
    $CurrDomainFQDN     = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
    $DomainSIDList.Add($CurrDomainSid, $CurrDomainFQDN)

    $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    ForEach($Domain in $Forest.Domains) 
    {
        $adsisearcher             = New-Object system.directoryservices.directorysearcher
        $adsisearcher.SearchRoot  = [ADSI]"LDAP://CN=System,$($Domain.GetDirectoryEntry().distinguishedName)"
        $adsisearcher.Filter      = "(objectclass=trustedDomain)"
        ForEach($ExtDomain in $adsisearcher.FindAll()) 
        {
            $name = $ExtDomain.Properties["name"][0]
         #   "Found $($name)"
            $sid = New-Object System.Security.Principal.SecurityIdentifier ($ExtDomain.Properties["securityidentifier"][0], 0)
            if (-not $DomainSIDList.Contains($sid.Value)) 
            {
        #        "Adding $($sid.Value), $($name)"
                $DomainSIDList.Add($sid.Value, $name)
            }
        }
    }

    return $DomainSIDList
}


Function Get-ComputerDn
{
    param (
        [Parameter()]
        [String]$Computername,
        [String]$DomainShortName
    )
    $ClientDN = ""

    $AdsiString             =   Get-LdapFromDomainShortName -DomainShortnameToFind  $DomainShortName
    $searcher               =   [adsisearcher]"((objectcategory=computer))" 
    $searcher.searchRoot    =   [adsi]$AdsiString
    $searcher.Filter        =   "(&(objectCategory=computer)(!operatingSystem=*server*)(operatingSystem=*Windows*)(Name=$Computername))" 
    $AdSearcResult          =   $Searcher.FindAll()
    If ($AdSearcResult.count -eq 1)
    {
        $ClientDN = $AdSearcResult.Properties.distinguishedname[0]
    }
    
    Return $ClientDN
}


Function Get-UserDn
{
    # Example:  $result = Get-UserDn -UserValue 'jardam' -DomainShortName 'ad' -UserSearchparameter "samaccountname"
    param (
        [Parameter()]
        [String]$UserValue,
        [String]$UserSearchparameter="samaccountname",
        [String]$DomainShortName
    )
    $UserDN = ""

    $AdsiString             =   Get-LdapFromDomainShortName -DomainShortnameToFind  $DomainShortName
    $searcher               =   [adsisearcher]"((objectcategory=User))" 
    $searcher.searchRoot    =   [adsi]$AdsiString
    $searcher.Filter        =   "(&(objectCategory=User)($UserSearchparameter=$UserValue))" 
    $AdSearcResult          =   $Searcher.FindAll()
    If ($AdSearcResult.count -eq 1)
    {
        $UserDN = $AdSearcResult.Properties.distinguishedname[0]
    }
    
    Return $UserDN
}


Function Get-AdObjectExistStatus
{
    # Check if a AD-Object exists
    # Example:  $Result = Get-AdObjectExistStatus -ObjValueString 'jardam' -ObjParameter 'samaccountname' -objectcategory User -DomainShortName 'SIKT'
    param (
        [Parameter()]
        [String]$ObjValueString,
        [String]$ObjParameter="samaccountname",
        [ValidateSet("User","Computer","Group")]
        [String]$objectcategory,
        [String]$DomainShortName,
        [String]$AdsiSearchString=''
    )
    $IsObjectFound = $False

    If ($AdsiSearchString -eq '')
    {
        $AdsiSearchString             =   Get-LdapFromDomainShortName -DomainShortnameToFind  $DomainShortName
    }
    
    $searcher               =   [adsisearcher]"((objectcategory=$objectcategory))" 
    $searcher.searchRoot    =   [adsi]$AdsiSearchString
    $searcher.Filter        =   "(&(objectCategory=$objectcategory)($ObjParameter=$ObjValueString))" 
    $AdSearcResult          =   $Searcher.FindAll()
    If ($AdSearcResult.count -gt 0)
    {
        $IsObjectFound = $True
    }
    
    Return $IsObjectFound
}

Function MoveComputerToOU
{
    [CmdletBinding()]
    param (
        [String]$ComputerDn,
        [String]$NewOuDn,
        [String]$Computername="",
        [String]$DomainShortName=""
    )
    
    # Moving machine to new OU, and verifies

    [bool]$MoveIsConfirmed  = $False

    $ComputerObj = [ADSI]"LDAP://$ComputerDn"
    try
    {
        $ComputerObj.psbase.MoveTo([ADSI]"LDAP://$($NewOuDn)")
    }
    catch
    {
        $ResultStr = $_.Tostring()
        $lines = $($ResultStr).Split("`r`n")
        ForEach ($LogLine in $lines) 
        {   
            if ($LogLine -ne "") {  Write-Log -Message "$LogLine" -LogLevel 3  }
            start-sleep -Milliseconds 500        
        }
    
    }
    start-sleep -Seconds 1

    If ($Computername -ne "")
    {
        $NewComputerDN          = Get-ComputerDn -Computername $Computername -DomainShortName $DomainShortName    
        $NewComputerDNArray     = $NewComputerDN.split(",")
        $newComputerOU          = $NewComputerDN.Replace("$($NewComputerDNArray[0]),", "")

        If ($NewOuDn -eq $newComputerOU)
        {        
            $MoveIsConfirmed = $true
            #Write-Log -Message "Move using ADSI success: $($ResultStr) " -LogLevel 1
        }
        else
        {
            #Write-Log -Message "Move using ADSI fail: $($ResultStr) " -LogLevel 3
        }
    }
        
    
    return $MoveIsConfirmed
}


Function MoveObjToOU
{
    #   Moves AD-Object to OU
    #   Depend on:  Write-Log, Get-AdObjectProperty

    [CmdletBinding()]
    param (
        [String]$ObjDn,
        [String]$NewOuDn,
        [ValidateSet("User","Computer","Group")]
        [String]$objectcategory,
        [String]$DomainShortName=""
    )    
    [bool]$MoveIsConfirmed      = $False
    $ObjsAMAccountName          = Get-AdObjectProperty -ObjValueString $ObjDn   -ObjParameter distinguishedName    -objectcategory $objectcategory  -ReturnProperty sAMAccountName -DomainShortName $DomainShortName -AdsiSearchString ''
    $MoveObject                 = [ADSI]"LDAP://$ObjDn"
    try
    {
        $MoveObject.psbase.MoveTo([ADSI]"LDAP://$($NewOuDn)")
    }
    catch
    {
        $ResultStr = $_.Tostring()
        $lines = $($ResultStr).Split("`r`n")
        ForEach ($LogLine in $lines) 
        {   
            if ($LogLine -ne "") {  Write-Log -Message "$LogLine" -LogLevel 3  }
            start-sleep -Milliseconds 500        
        }
    }
    start-sleep -Seconds 1
    
    $VerifyObjDn                = Get-AdObjectProperty -ObjValueString $ObjsAMAccountName   -ObjParameter samAccountname    -objectcategory $objectcategory  -ReturnProperty distinguishedname -DomainShortName $DomainShortName   -AdsiSearchString ''
    $VerifyObjDnArray           = $VerifyObjDn.split(",")
    $VerifyObjOuDn              = $VerifyObjDn.Replace("$($VerifyObjDnArray[0]),", "")

    If ($NewOuDn -eq $VerifyObjOuDn)
    {        
        $MoveIsConfirmed = $true
    }
    else
    {
        $MoveIsConfirmed = $false
    }
    
    return $MoveIsConfirmed
}


Function Set-Extensionattribute    
{
    param (
                    [Parameter(Mandatory = $true)][String]$ComputerName,
                    [Parameter(Mandatory = $true)][String]$DomainShort,                    
                    [Parameter(Mandatory = $true)][String]$ExtAttributeName,
                    [Parameter(Mandatory = $true)][String]$TextToAdd
             )

    [bool]$AdObjectStatus   = $False    
    $AdObjectStatusMsg      = 'Ready'
    $ItemDnPath             = Get-ComputerDn -Computername $ComputerName -DomainShortName $DomainShort

    If ($ItemDnPath -eq "")    {   $AdObjectStatusMsg = "Ad-Objects does not exists, or unavailable"                          }

    IF ($AdObjectStatusMsg -eq "Ready")
    {
        $ItemDnPath = "LDAP://$ItemDnPath"
        $ComputerAdObj = [adsi]$ItemDnPath
        If ($ComputerAdObj.properties.$ExtAttributeName -eq $TextToAdd)
        {
            $AdObjectStatusMsg = "String:  $TextToAdd  Allready Exists on: $ComputerName"
            $AdObjectStatus = $True
            #Break
        }
        else 
        {
            $ComputerAdObj.Put($ExtAttributeName,-join $TextToAdd)
            $ComputerAdObj.SetInfo()        
            $AdObjectStatusMsg = "String:  $TextToAdd  added to: $ComputerName"
            $AdObjectStatus = $True           
            #Break
        }

    }

    If ($AdObjectStatus -ne $true)
    {
        # Exit with failure to stop Task Sequence if Writing to Extensionattribute fails
        #Write-Log -Message  "FAIL to update $MachineName Extensionattribute with: $MacAdressString " -LogLevel 3
     #   exit 20  
    }

    #Write-Log -Message  $AdObjectStatusMsg -LogLevel 1
    [bool]$ReturnValue =  $AdObjectStatus
    return $ReturnValue 
}

Function Get-Extensionattribute    
{
    param (
                    [Parameter(Mandatory = $true)][String]$ComputerName,
                    [Parameter(Mandatory = $true)][String]$DomainShort,                    
                    [Parameter(Mandatory = $true)][String]$ExtAttributeName                    
             )
  
    $AdObjectStatusMsg      = 'Ready'
    $FoundExtAttrTXT        = ''
    $ItemDnPath             = Get-ComputerDn -Computername $ComputerName -DomainShortName $DomainShort

    If ($ItemDnPath -eq "")    {   $AdObjectStatusMsg = "Ad-Objects does not exists, or unavailable"                          }

    IF ($AdObjectStatusMsg -eq "Ready")
    {
        $ItemDnPath         = "LDAP://$ItemDnPath"
        $ComputerAdObj      = [adsi]$ItemDnPath
        $FoundExtAttrTXT    = $ComputerAdObj.properties.$ExtAttributeName
    
    }

    [String]$ReturnValue    =  $FoundExtAttrTXT
    return $ReturnValue 
}


Function Get-DateFromLong
{
    param (
              [Long]$AdStyleDateTime      )

        If (-Not $AdStyleDate) { $AdStyleDate = 0 }
        $LLDate = [DateTime]$AdStyleDate
        $ReturnDateTime = $LLDate.AddYears(1600).ToLocalTime()

        Return $ReturnDateTime
}


Function Get-DNfromForeignSecurityPrincipals
{
    #  Example:  $Result = Get-DNfromForeignSecurityPrincipals -FcpStr "CN=S-1-5-21-1100344877-3013322779-101495848-113313,CN=ForeignSecurityPrincipals,DC=sikt,DC=sykehuspartner,DC=no"
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$FcpStr
    )
    $FoundDNS = ''
    $ReturnDistingueshedName = ""

    # Strip string to SID
    $FcpStr = $($FcpStr -split ',')[0]
    if ($FcpStr -like 'CN=*') { $FcpStr = $FcpStr.Replace('CN=','')}

    #  Find DNS Domain
    if (!$TrustedDomainList) { $TrustedDomainList = Get-TrustedDomainsSIDandDNS   }
    foreach ($TrustedDomainItem in $TrustedDomainList.keys)
    {
        if ( $FcpStr.Substring(0, $TrustedDomainItem.length) -eq $TrustedDomainItem)
        {
            $FoundDNS = $TrustedDomainList[$TrustedDomainItem]
            Break
        }
    }
    try 
    {
        $ReturnDistingueshedName = $([ADSI]"LDAP://$FoundDNS/<SID=$FcpStr>").distinguishedName
        
    }
    catch {
        $ReturnDistingueshedName = 'Unknown'
    }
    
    Return $ReturnDistingueshedName
}



function Get-FQDNfromDomainShortName
  {
      param (
          [string]$DomainShortnameToFind
      )
      $DomainArray          = @()      
      $CurrentDomain        = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
      $DomainArray          += $CurrentDomain
      $TrustedDomainsList   = $CurrentDomain.GetAllTrustRelationships()

      foreach ($TrustedDomain in $TrustedDomainsList)
      {
          $DomainArray += $TrustedDomain.Targetname
      }
      
      $Forest               = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
      $TrustedForestList    = $Forest.GetAllTrustRelationships()
      foreach ($TrustedForest in $TrustedForestList)
      {
          $DomainArray += $TrustedForest.Targetname
      }
      
      foreach ($DomainItems in $DomainArray)
      {
          if ($DomainItems -like "*$DomainShortnameToFind*")
          {
              [String]$FoundFQDN   = $DomainItems.tostring()          
          }
      }
      
      Return $FoundFQDN
  }


  #   TCP-IP functions
  
Function Convert-IPv4AddressToBinaryString {
    Param(
      [IPAddress]$IPAddress='0.0.0.0'
    )
    $addressBytes=$IPAddress.GetAddressBytes()
  
    $strBuilder=New-Object -TypeName Text.StringBuilder
    foreach($byte in $addressBytes){
      $8bitString=[Convert]::ToString($byte,2).PadRight(8,'0')
      [void]$strBuilder.Append($8bitString)
    }
    Write-Output $strBuilder.ToString()
  }
  
  Function ConvertIPv4ToInt {
    [CmdletBinding()]
    Param(
      [String]$IPv4Address
    )
    Try{
      $ipAddress=[IPAddress]::Parse($IPv4Address)
  
      $bytes=$ipAddress.GetAddressBytes()
      [Array]::Reverse($bytes)
  
      [System.BitConverter]::ToUInt32($bytes,0)
    }Catch{
      Write-Error -Exception $_.Exception `
        -Category $_.CategoryInfo.Category
    }
  }
  
  Function ConvertIntToIPv4 {
    [CmdletBinding()]
    Param(
      [uint32]$Integer
    )
    Try{
      $bytes=[System.BitConverter]::GetBytes($Integer)
      [Array]::Reverse($bytes)
      ([IPAddress]($bytes)).ToString()
    }Catch{
      Write-Error -Exception $_.Exception `
        -Category $_.CategoryInfo.Category
    }
  }
  
  Function Add-IntToIPv4Address {
    Param(
      [String]$IPv4Address,
  
      [int64]$Integer
    )
    Try{
      $ipInt=ConvertIPv4ToInt -IPv4Address $IPv4Address `
        -ErrorAction Stop
      $ipInt+=$Integer
  
      ConvertIntToIPv4 -Integer $ipInt
    }Catch{
      Write-Error -Exception $_.Exception `
        -Category $_.CategoryInfo.Category
    }
  }
  
  Function CIDRToNetMask {
    [CmdletBinding()]
    Param(
      [ValidateRange(0,32)]
      [int16]$PrefixLength=0
    )
    $bitString=('1' * $PrefixLength).PadRight(32,'0')
  
    $strBuilder=New-Object -TypeName Text.StringBuilder
  
    for($i=0;$i -lt 32;$i+=8){
      $8bitString=$bitString.Substring($i,8)
      [void]$strBuilder.Append("$([Convert]::ToInt32($8bitString,2)).")
    }
  
    $strBuilder.ToString().TrimEnd('.')
  }
  
  Function NetMaskToCIDR {
    [CmdletBinding()]
    Param(
      [String]$SubnetMask='255.255.255.0'
    )
    $byteRegex='^(0|128|192|224|240|248|252|254|255)$'
    $invalidMaskMsg="Invalid SubnetMask specified [$SubnetMask]"
    Try{
      $netMaskIP=[IPAddress]$SubnetMask
      $addressBytes=$netMaskIP.GetAddressBytes()
  
      $strBuilder=New-Object -TypeName Text.StringBuilder
  
      $lastByte=255
      foreach($byte in $addressBytes){
  
        # Validate byte matches net mask value
        if($byte -notmatch $byteRegex){
          Write-Error -Message $invalidMaskMsg `
            -Category InvalidArgument `
            -ErrorAction Stop
        }elseif($lastByte -ne 255 -and $byte -gt 0){
          Write-Error -Message $invalidMaskMsg `
            -Category InvalidArgument `
            -ErrorAction Stop
        }
  
        [void]$strBuilder.Append([Convert]::ToString($byte,2))
        $lastByte=$byte
      }
  
      ($strBuilder.ToString().TrimEnd('0')).Length
    }Catch{
      Write-Error -Exception $_.Exception `
        -Category $_.CategoryInfo.Category
    }
  }
  
  Function Get-IPv4Subnet {
    [CmdletBinding(DefaultParameterSetName='PrefixLength')]
    Param(
      [Parameter(Mandatory=$true,Position=0)]
      [IPAddress]$IPAddress,
  
      [Parameter(Position=1,ParameterSetName='PrefixLength')]
      [Int16]$PrefixLength=24,
  
      [Parameter(Position=1,ParameterSetName='SubnetMask')]
      [IPAddress]$SubnetMask
    )
    Begin{}
    Process{
      Try{
        if($PSCmdlet.ParameterSetName -eq 'SubnetMask'){
          $PrefixLength=NetMaskToCidr -SubnetMask $SubnetMask `
            -ErrorAction Stop
        }else{
          $SubnetMask=CIDRToNetMask -PrefixLength $PrefixLength `
            -ErrorAction Stop
        }
        
        $netMaskInt=ConvertIPv4ToInt -IPv4Address $SubnetMask     
        $ipInt=ConvertIPv4ToInt -IPv4Address $IPAddress
        
        $networkID=ConvertIntToIPv4 -Integer ($netMaskInt -band $ipInt)
  
        $maxHosts=[math]::Pow(2,(32-$PrefixLength)) - 2
        $broadcast=Add-IntToIPv4Address -IPv4Address $networkID `
          -Integer ($maxHosts+1)
  
        $firstIP=Add-IntToIPv4Address -IPv4Address $networkID -Integer 1
        $lastIP=Add-IntToIPv4Address -IPv4Address $broadcast -Integer -1
  
        if($PrefixLength -eq 32){
          $broadcast=$networkID
          $firstIP=$null
          $lastIP=$null
          $maxHosts=0
        }
  
        $outputObject=New-Object -TypeName PSObject 
  
        $memberParam=@{
          InputObject=$outputObject;
          MemberType='NoteProperty';
          Force=$true;
        }
        Add-Member @memberParam -Name CidrID -Value "$networkID/$PrefixLength"
        Add-Member @memberParam -Name NetworkID -Value $networkID
        Add-Member @memberParam -Name SubnetMask -Value $SubnetMask
        Add-Member @memberParam -Name PrefixLength -Value $PrefixLength
        Add-Member @memberParam -Name HostCount -Value $maxHosts
        Add-Member @memberParam -Name FirstHostIP -Value $firstIP
        Add-Member @memberParam -Name LastHostIP -Value $lastIP
        Add-Member @memberParam -Name Broadcast -Value $broadcast
  
        Write-Output $outputObject
      }Catch{
        Write-Error -Exception $_.Exception `
          -Category $_.CategoryInfo.Category
      }
    }
    End{}
  }



# ------------------------------------------------------------------------------------------------------------




Export-ModuleMember -Function AddRemoveGroupMember
Export-ModuleMember -Function Get-AdNetInfoForIP
Export-ModuleMember -Function Get-AdObjectProperties
Export-ModuleMember -Function Get-AdObjectProperty
Export-ModuleMember -Function Get-AdObjectDn
Export-ModuleMember -Function Get-AdObjectExistStatus
Export-ModuleMember -Function Get-AdsiPath
Export-ModuleMember -Function Get-AdsiPathForCurrentDomain
Export-ModuleMember -Function Get-ClientsInDomain
Export-ModuleMember -Function Get-ComputerDn
Export-ModuleMember -Function Get-UserDn
Export-ModuleMember -Function Get-DateFromLong
Export-ModuleMember -Function Get-DNfromForeignSecurityPrincipals
Export-ModuleMember -Function Get-FQDNfromDomainShortName
Export-ModuleMember -Function Get-GroupMembersNested
Export-ModuleMember -Function Get-GroupsForDomain
Export-ModuleMember -Function Get-LdapFromDomainName
Export-ModuleMember -Function Get-LdapSearchRootFromDN
Export-ModuleMember -Function Get-LdapFromDomainShortName
Export-ModuleMember -Function Get-MemberOfDN
Export-ModuleMember -Function Get-NestedMemberOf 
Export-ModuleMember -Function MoveComputerToOU
Export-ModuleMember -Function MoveObjToOU
Export-ModuleMember -Function Get-PropFromDn
Export-ModuleMember -Function Get-TrustedDomainsSIDandDNS
Export-ModuleMember -Function Get-Extensionattribute
Export-ModuleMember -Function Set-Extensionattribute
Export-ModuleMember -Function Get-GroupnameFromDn

#   TCP-IP functions
Export-ModuleMember -Function Convert-IPv4AddressToBinaryString
Export-ModuleMember -Function ConvertIPv4ToInt
Export-ModuleMember -Function ConvertIntToIPv4
Export-ModuleMember -Function Add-IntToIPv4Address
Export-ModuleMember -Function CIDRToNetMask
Export-ModuleMember -Function NetMaskToCIDR
Export-ModuleMember -Function Get-IPv4Subnet





