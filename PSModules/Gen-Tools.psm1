
function ConvertTo-DataTable
{
            <#
                .SYNOPSIS
                    Converts objects into a DataTable.
            
                .DESCRIPTION
                    Converts objects into a DataTable, which are used for DataBinding.
            
                .PARAMETER  InputObject
                    The input to convert into a DataTable.
            
                .PARAMETER  Table
                    The DataTable you wish to load the input into.
            
                .PARAMETER RetainColumns
                    This switch tells the function to keep the DataTable's existing columns.
                
                .PARAMETER FilterCIMProperties
                    This switch removes CIM properties that start with an underline.
                
                .PARAMETER MatchColumns
                    This switch force only listed properties to be included ( Use |  as delimiter)
            
                .EXAMPLE
                    $DataTable = ConvertTo-DataTable -InputObject (Get-Process)
            #>
    [OutputType([System.Data.DataTable])]
    param (
        $InputObject,
        [ValidateNotNull()]
        [System.Data.DataTable]$Table,
        [switch]$RetainColumns,
        [switch]$FilterCIMProperties,
        [String]$MatchColumns)
    
    if ($null -eq $Table)
    {
        $Table = New-Object System.Data.DataTable
    }
    
    if ($null -eq $InputObject)
    {
        $Table.Clear()
        return @( ,$Table)
    }
    
    if ($InputObject -is [System.Data.DataTable])
    {
        $Table = $InputObject
    }
    elseif ($InputObject -is [System.Data.DataSet] -and $InputObject.Tables.Count -gt 0)
    {
        $Table = $InputObject.Tables[0]
    }
    else
    {
        if (-not $RetainColumns -or $Table.Columns.Count -eq 0)
        {
            #Clear out the Table Contents
            $Table.Clear()
            
            if ($null -eq $InputObject) { return } #Empty Data
            
            $object = $null
            #find the first non null value
            foreach ($item in $InputObject)
            {
                if ($null -ne $item)
                {
                    $object = $item
                    break
                }
            }
            
            if ($null -eq $object) { return } #All null then empty
            
            #Get all the properties in order to create the columns
            foreach ($prop in $object.PSObject.Get_Properties())
            {
                if ('RowError', 'RowState', 'Table', 'ItemArray', 'HasErrors' -contains $prop.Name)
                {
                    continue
                }
                If ($PSBoundParameters.ContainsKey('MatchColumns'))
                {
                    $MatchColumnsArray = $MatchColumns.split("|")
                    if (!$MatchColumnsArray.contains($prop.Name))
                    {
                        continue
                    }
                }
                if (-not $FilterCIMProperties -or -not $prop.Name.StartsWith('__')) #filter out CIM properties
                {
                    
                    #Get the type from the Definition string
                    $type = $null
                    
                    if ($null -ne $prop.Value)
                    {
                        try { $type = $prop.Value.GetType() }
                        catch { Out-Null }
                    }
                    #write-host "$($prop.Name) - $type"
                    If ($type.FullName -eq 'System.DBNull')
                    {
                        $type = 'string'
                    }
                    #write-host "$($prop.Name) - $type"
                    if ($null -ne $type) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
                    {
                        [void]$table.Columns.Add($prop.Name, $type)
                    }
                    else #Type info not found
                    {
                        [void]$table.Columns.Add($prop.Name)
                    }
                }
            }
            
            if ($object -is [System.Data.DataRow])
            {
                foreach ($item in $InputObject)
                {
                    #$Table.Rows.Add($item)
                    $Table.ImportRow($item)
                }
                return @( ,$Table)
            }
        }
        else
        {
            $Table.Rows.Clear()
        }
        
        foreach ($item in $InputObject)
        {
            $row = $table.NewRow()
            
            if ($item)
            {
                foreach ($prop in $item.PSObject.Get_Properties())
                {
                    if ($table.Columns.Contains($prop.Name))
                    {
                        $row.Item($prop.Name) = $prop.Value
                    }
                }
            }
            [void]$table.Rows.Add($row)
        }
    }
    
    return @( ,$Table)
}


function Write-Log
{
    param (
    [Parameter(Mandatory = $true)]
    [string]$Message,
             
    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [int]$LogLevel = 1,

    [Parameter()]
    [bool]$Output = $True
    )

    if ($Global:ScriptLogFilePath -eq $null)
    {            
        $date               = Get-Date
        $DateStr            = $date.toString("ddMM")
        $ScriptPath         = @(Get-PSCallStack)[0].invocationinfo.PsScriptRoot
        $ScriptName         = Split-Path @(Get-PSCallStack)[0].invocationinfo.PSCommandpath -Leaf
        $ScriptName         = $ScriptName.Replace('.ps1','')
        $LoggFileFolder     = "$ScriptPath\logs"    
        if (!(Test-Path $LoggFileFolder))        {            New-Item -ItemType directory -Path $LoggFileFolder        }
        $Global:ScriptLogFilePath = "$LoggFileFolder\$ScriptName-$DateStr.log"        

    }

    $TimeGenerated  = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line           = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat     = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line           = $Line -f $LineFormat

    Start-Sleep -Milliseconds 200
    Add-Content -Value $Line -Path $Global:ScriptLogFilePath | Out-Null  
    
    if ($Output)
    {
        $CurrentFunction    = [string]$(Get-PSCallStack)[1].FunctionName + ':'
        If ($CurrentFunction -eq "<ScriptBlock>:")
        {       $CurrentFunction = 'Line: ' + $(Get-PSCallStack)[1].ScriptLineNumber }
        else {  $CurrentFunction = $CurrentFunction  + $(Get-PSCallStack)[1].ScriptLineNumber }
                 
        #$Message = $($Global:LogStopWatch.Elapsed.ToString('mm\:ss\.ff')) + ': ' + $Message + ' (' + $CurrentFunction +')'
        $Message =  $Message + ' (' + $CurrentFunction +')'

        Switch ($LogLevel) {
            2           {write-host -ForegroundColor    Yellow  $Message}
            3           {write-host -ForegroundColor    Red     $Message}
            Default     {write-host -ForegroundColor    Green   $Message}
        }
    }
    
}

function Get-FunctionName ([int]$StackNumber = 1) {
    return [string]$(Get-PSCallStack)[$StackNumber].FunctionName
}


function Export-TableToExcel
	{
		param (
			$TableToWrite,
			[string]$SheetName = 'TableDump',
			[bool]$CloseWb = $false)
		
	#	Add-Type -AssemblyName System.Windows.Forms
	#	[System.Windows.Forms.Clipboard]::Clear()
        $dataTable                     = new-object Data.datatable 
		$HeaddingRow = 1
		If (!$Global:wb)
		{
			$Global:Excel = New-Object -ComObject excel.application
			$Global:Excel.visible = $True
			$Global:Excel.DisplayAlerts = $False
			$Global:Excel.EnableEvents = $False
			$Global:Excel.AskToUpdateLinks = $False
			$Global:wb = $Global:Excel.Workbooks.Add()
		}
		else
		{
			$Global:wb.Worksheets.Add()
		}
		
		$ws = $Global:wb.Worksheets.Item(1)
		$ws.Name = $SheetName
		
		#$dataTable = $TableToWrite[0].Table
		#$dataTable = $TableToWrite[0].Table
		if ($TableToWrite.rows.count -lt 1)
		{
			$dataTable = $TableToWrite[0].Table
		}
		else
		{
			#$dataTable = $TableToWrite.table
            $dataTable = $TableToWrite
		}

        #$dataTable = $TableToWrite


		$rowDT = $dataTable.Rows.Count;
		
		$colDT = $dataTable.Columns.Count;
		
		$tableArray = New-Object 'object[,]' $rowDT, $colDT;
		
		
		for ($i = 0; $i -lt $rowDT; $i++)
		{
			#Write-Progress -Activity "Transforming DataTable" -status "Row $i" -percentComplete ($i / $rowDT*100)
			for ($j = 0; $j -lt $colDT; $j++)
			{
				$tableArray[$i, $j] = $dataTable.Rows[$i].Item($j).ToString();
			}
		}
		
		$rowOffset = 1; $colOffset = 1; # 1,1 = "A1"
		
		# Write out the header column names
		for ($j = 0; $j -lt $colDT; $j++)
		{
			$ws.cells.item($rowOffset, $j + 1) = $dataTable.Columns[$j].ColumnName;
		}
		$headerRange = $ws.Range($ws.cells.item($rowOffset, $colOffset), $ws.cells.item($rowOffset, $colDT + $colOffset - 1));
		#$headerRange.Font.Bold = $false
		#$headerRange.Interior.Color = $headingColour
		#$headerRange.Font.Name = $headingFont
		$rowOffset++;
		
		# Extract the data to Excel
		$tableRange = $ws.Range($ws.cells.item($rowOffset, $colOffset), $ws.cells.item($rowDT + $rowOffset - 1, $colDT + $colOffset - 1));
		$tableRange.Cells.Value2 = $tableArray;
		
		
		$usedRange = $ws.range($ws.Cells.Item(1, 1), $ws.Cells.Item($rowDT + $rowOffset - 1, $colDT + $colOffset - 1))
		$usedRange.Select() | Out-Null
		$usedRange.Activate() | Out-Null
		$ListObject = $ws.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $usedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = $SheetName + 'List'
		$ListObject.TableStyle = "TableStyleMedium9"
		#$usedRange.Columns | %{ $_.AutoFit() | Out-Null }
		$usedRange.EntireColumn.AutoFit() | Out-Null
		#$usedRange.Columns.AutoFit() | Out-Null
		
		
		If ($CloseWb)
		{
			$Global:wb.Close
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WS)
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Global:wb)
			Remove-Variable wb -scope 'Global' -ErrorAction SilentlyContinue
		}
		
		
	}
Function Get-ClientProgWmi 
{
    #   Example
    #   $MachineName = @('PC184000','PC189895')
    #   $TempTable = Get-ClientProgWmi -MachineName $MachineName


    [CmdletBinding()]
    Param
    ( 
        [Parameter(ValueFromPipeline)]
        $MachineNames
    )

    Begin
    {
        $AplTable         = new-object Data.datatable 
        $AplTable.Columns.Add("Computer","System.String")               | Out-Null    
        $AplTable.Columns.Add("DisplayName","System.String")            | Out-Null
        $AplTable.Columns.Add("DisplayVersion","System.String")         | Out-Null
        $AplTable.Columns.Add("Publisher","System.String")              | Out-Null
    }

    process 
    {
        Foreach ($MachineName in $MachineNames)
        {
            If (Test-Connection -ComputerName $MachineName -Count 2 -Delay 1 -Quiet)
            {
                $TempTableArp       = Get-CIMInstance -ClassName Win32_InstalledWin32Program -NameSpace root\cimv2 -ComputerName $MachineName #|  Select Vendor, Name, Version
                $TempTableAppv      = Invoke-Command{get-appvclientPackage -all} -computername $MachineName #| Format-Table Vendor, Name, Version
            
                Foreach ($ArpItem in $TempTableArp)
                {        
                    If ($ArpItem.Name)
                    {                
                        $ArpI = $AplTable.NewRow()            
                        $ArpI.Item(0)       = $MachineName
                        $ArpI.Item(1)       = $ArpItem.Name
                        $ArpI.Item(2)       = $ArpItem.Version        
                        $ArpI.Item(3)       = $ArpItem.Vendor

                        $AplTable.rows.add($ArpI)            
                    }        
                }

                Foreach ($ArpItem in $TempTableAppv)
                {        
                    If ($ArpItem.Name)
                    {                
                        $ArpI = $AplTable.NewRow()    
                        $ArpI.Item(0)       = $MachineName
                        $ArpI.Item(1)       = $ArpItem.Name
                        $ArpI.Item(2)       = $ArpItem.Version        
                        $ArpI.Item(3)       = 'App-V'
                        $AplTable.rows.add($ArpI)            
                    }        
                }

            }
        }
    }

    END
    {
        Return $AplTable
    }

  #  
}


Function NewFileVersionName
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$OldFileName
    )
    $FileVersionCount   = 1
    $NewFileNameFound   = $False
    # Get New XLS name to save to
    $NameArray = $OldFileName.split('.')
    
    While ($NewFileNameFound -eq $False)
    {
        $NewFileName         = $($NameArray[0]).Replace("$(Get-Date -Format dd_MM)_",'')
        $NewFileName = -join($NewFileName , '_' , $(Get-Date -Format dd_MM),"_$FileVersionCount" , '.' ,  $($NameArray[1]))
        if ((Test-Path -Path $NewFileName -PathType Leaf) -eq $False) {$NewFileNameFound   = $True}
        $FileVersionCount++
    }
    
    #Logg "NewFileName = $NewFileName"
    Return $NewFileName
}

Function RunCmd
{
    [CmdletBinding()]
    param (
        [String]$Filename ,
        [String]$FilePath ,
        [String]$Arguments     
    )

    #  Example of usage
    #
    #   $RunResult = RunCmd -Filename "ping.exe" -FilePath "c:\windows" -Arguments "localhost"
    #   write-host $RunResult[0]  # Standard output
    #   write-host $RunResult[1]  # Error Output
    #   write-host $RunResult[2]  # ExitCode

    $RunResultArray = @()

    $pinfo                          = New-Object System.Diagnostics.ProcessStartInfo
   
    $pinfo.FileName                 = $Filename
    $pinfo.WorkingDirectory         = $FilePath
    $pinfo.Arguments                = $Arguments
    $pinfo.RedirectStandardError    = $true
    $pinfo.RedirectStandardOutput   = $true
    $pinfo.UseShellExecute          = $false
    
    $p                              = New-Object System.Diagnostics.Process
    $p.StartInfo                    = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()

    $stdout                         = $p.StandardOutput.ReadToEnd()
    $stderr                         = $p.StandardError.ReadToEnd()

    #   Write-Host "stdout: $stdout"
    #   Write-Host "stderr: $stderr"
    #   Write-Host "exit code: " + $p.ExitCode

    $RunResultArray = @($stdout, $stderr, $p.ExitCode)
    
    return $RunResultArray
    
}

function Test-Port($server, $port) {
    $client = New-Object Net.Sockets.TcpClient
    try {
        $client.Connect($server, $port)
        $true
    } catch {
        $false
    } finally {
        $client.Dispose()
    }
}  

Function WriteToRegistry
{
    [CmdletBinding()]
    param (
        [String]$RegPath,
        [String]$RegKey,
        [String]$RegValue,
        [Parameter()]
        [ValidateSet("String","ExpandString","Binary","DWord","MultiString","Qword","Unknown")]
        [String]$RegType
    )

    if (!$(Test-Path "$RegPath")) 
    { 
        New-Item -path "$RegPath" -Force | Out-Null
    }

    $RegObj = Get-ItemProperty -Path $RegPath -Name $RegKey -ErrorAction Ignore
    If (-not($RegObj))
    {
        New-ItemProperty -Path $RegPath -Name $RegKey -Value $RegValue -Type $RegType | Out-Null
    }
    else {
        Set-ItemProperty -Path $RegPath -Name $RegKey -Value $RegValue -Type $RegType | Out-Null
    }
}

Function Test-RegistryValue($regkey, $name) 
{
    try
    {
        $exists = Get-ItemProperty $regkey $name -ErrorAction SilentlyContinue
 #       Write-Host "Test-RegistryValue: $exists"
        if (($exists -eq $null) -or ($exists.Length -eq 0))
        {
            return $false
        }
        else
        {
            return $true
        }
    }
    catch
    {
        return $false
    }
}

function Test-Cred
{
	
	[CmdletBinding()]
	[OutputType([String])]
	Param (
		[Parameter(
				   Mandatory = $false,
				   ValueFromPipeLine = $true,
				   ValueFromPipelineByPropertyName = $true
				   )]
		[Alias(
			   'PSCredential'
			   )]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential][System.Management.Automation.Credential()]
		$Credentials
	)
	
	
	
	#  Example
	#  $UserCredOK = $False
	#  while ($UserCredOK -eq $False)
	#  {
	#      Try
	#      {
	#          $Credentials = Get-Credential "ad\jardam" -ErrorAction Stop
	#      }
	#      Catch
	#      {
	#          $ErrorMsg = $_.Exception.Message
	#          Write-Warning "Failed to validate credentials: $ErrorMsg "
	#          Pause
	#          Break
	#      }
	#      $UserCredOK = $Credentials | Test-Cred
	#  }
	#  write-host "UserCredOK = $UserCredOK"
	
	
	
	
	$Domain = $null
	$Root = $null
	$Username = $null
	$Password = $null
	
	# Checking module
	Try
	{
		# Split username and password
		$Username = $credentials.username
		$Password = $credentials.GetNetworkCredential().password
		$DomainName = $credentials.GetNetworkCredential().Domain
		
		
		
		# Get Domain
		#$Root = "LDAP://" + ([ADSI]'').distinguishedName
		$Root = Get-LdapFromDomainName $DomainName
		$Domain = New-Object System.DirectoryServices.DirectoryEntry($Root, $UserName, $Password)
	}
	Catch
	{
		Write-Warning $_.Exception.Message
		Continue
	}
	
	If (!$domain)
	{
		Write-Warning "Something went wrong"
	}
	Else
	{
		If ($domain.name -ne $null)
		{
			return $True
		}
		Else
		{
			return $False
		}
	}
}


Function Test-IsValidIpv4
{
    [CmdletBinding()]
    param (        
        [String]$IpString
    )
    $pattern = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"
    
    return $($IpString -match $pattern)
}

Function Test-IsValidMacAdress
{
    [CmdletBinding()]
    param (        
        [String]$MacAdress
    )
    $pattern = "^[a-fA-F0-9:]{17}|[a-fA-F0-9]{12}$"
    
    return $($MacAdress -match $pattern)
}

function Get-RandomPassword
{
	param (
		#[Parameter(Mandatory)]
		[ValidateRange(4, [int]::MaxValue)]
		[int]$length,
		[int]$upper = 1,
		[int]$lower = 1,
		[int]$numeric = 1,
		[int]$special = 1
	)
	if ($upper + $lower + $numeric + $special -gt $length)
	{
		throw "number of upper/lower/numeric/special char must be lower or equal to length"
	}
	$uCharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	$lCharSet = "abcdefghijklmnopqrstuvwxyz"
	$nCharSet = "0123456789"
	$sCharSet = "/*-+,!?=()@;:._"
	$charSet = ""
	if ($upper -gt 0) { $charSet += $uCharSet }
	if ($lower -gt 0) { $charSet += $lCharSet }
	if ($numeric -gt 0) { $charSet += $nCharSet }
	if ($special -gt 0) { $charSet += $sCharSet }
	
	$charSet = $charSet.ToCharArray()
	$rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
	$bytes = New-Object byte[]($length)
	$rng.GetBytes($bytes)
	
	$result = New-Object char[]($length)
	for ($i = 0; $i -lt $length; $i++)
	{
		$result[$i] = $charSet[$bytes[$i] % $charSet.Length]
	}
	$password = (-join $result)
	$valid = $true
	if ($upper -gt ($password.ToCharArray() | Where-Object { $_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
	if ($lower -gt ($password.ToCharArray() | Where-Object { $_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
	if ($numeric -gt ($password.ToCharArray() | Where-Object { $_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }
	if ($special -gt ($password.ToCharArray() | Where-Object { $_ -cin $sCharSet.ToCharArray() }).Count) { $valid = $false }
	
	if (!$valid)
	{
		$password = Get-RandomPassword $length $upper $lower $numeric $special
	}
	return $password
}

function Get-WinRMstatusForMachine
{
	[CmdletBinding()]
	param (
		[Parameter()]
		[String]$MachineName
	)
	
	[bool]$WinRmStatus = $False
	$FoundMachineName = ""
	
	try
	{
		$FoundMachineName = $((Get-CimInstance -ComputerName $MachineName -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).PSComputerName)
		if ($FoundMachineName -eq $MachineName)
		{
			$WinRmStatus = $true
		}
		
	}
	catch
	{
		$WinRmStatus = $False
	}
	
	Return $WinRmStatus
}



Function Set-Audio 
{
    [CmdletBinding()]
	param (
        [Parameter()]
        [ValidateSet("Audio","Mic")]
        [String]$Device="Audio",
        [Parameter()]
        [Bool]$Mute,
        [Parameter()]
        [ValidateRange(0,100)]
        [Int]$Vol
    )
	
	# Example
	# Set-Audio -Device Audio -Vol 55
	# Set-Audio -Device Mic -Mute $true

    $loaded = [appdomain]::currentdomain.getassemblies()
    $foundType = $loaded | Where { ($_.GetExportedTypes()).Name -eq $Audio }

    if (!$foundType)
    {
        Add-Type -TypeDefinition @'

        using System.Runtime.InteropServices;
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IAudioEndpointVolume {
        // f(), g(), ... are unused COM method slots. Define these if you care
        int f(); int g(); int h(); int i();
        int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
        int j();
        int GetMasterVolumeLevelScalar(out float pfLevel);
        int k(); int l(); int m(); int n();
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
        int GetMute(out bool pbMute);
        }
        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IMMDevice {
        int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
        }
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IMMDeviceEnumerator {
        int f(); // Unused
        int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
        }
        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
        public class Audio {
        static IAudioEndpointVolume Vol() {
            var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
            IMMDevice dev = null;
            Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
            IAudioEndpointVolume epv = null;
            var epvid = typeof(IAudioEndpointVolume).GUID;
            Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
            return epv;
        }
        public static float Volume {
            get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
            set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
        }
        public static bool Mute {
            get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
            set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
        }
        }
        public class Mic {
        static IAudioEndpointVolume Vol() {
            var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
            IMMDevice dev = null;
            Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 1, /*eMultimedia*/ 1, out dev));
            IAudioEndpointVolume epv = null;
            var epvid = typeof(IAudioEndpointVolume).GUID;
            Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
            return epv;
        }
        public static float Volume {
            get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
            set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
        }
        public static bool Mute {
            get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
            set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
        }
        }
'@
    }
    If ($PSBoundParameters.ContainsKey('Mute'))
    {
        switch ($Device) {
            "Audio"     {   [audio]::Mute = $Mute }
            "Mic"       {   [Mic]::Mute = $Mute }    
        }        
    }

    If ($PSBoundParameters.ContainsKey('Vol'))
    {
        $Level = $Vol / 100
        switch ($Device) {
            "Audio"     {   [audio]::Volume = $Level }
            "Mic"       {   [Mic]::Volume = $Level }    
        }        
    }

}




Export-ModuleMember -Function ConvertTo-DataTable
Export-ModuleMember -Function Export-TableToExcel
Export-ModuleMember -Function Get-ClientProgWmi 
Export-ModuleMember -Function Get-FunctionName
Export-ModuleMember -Function Get-RandomPassword
Export-ModuleMember -Function Get-WinRMstatusForMachine
Export-ModuleMember -Function NewFileVersionName
Export-ModuleMember -Function RunCmd
Export-ModuleMember -Function Set-Audio 
Export-ModuleMember -Function Test-Cred									   
Export-ModuleMember -Function Test-IsValidIpv4
Export-ModuleMember -Function Test-IsValidMacAdress
Export-ModuleMember -Function Test-Port
Export-ModuleMember -Function Test-RegistryValue
Export-ModuleMember -Function Write-Log
Export-ModuleMember -Function WriteToRegistry




