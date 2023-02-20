function RetrievePackages
{
	param($path, $registry, $computer)
	
	$packages = @()
	$key = $registry.OpenSubKey($path) 
	$subKeys = $key.GetSubKeyNames() |% {
		$subKeyPath = $path + "\\" + $_ 
		$packageKey = $registry.OpenSubKey($subKeyPath) 
		$package = New-Object PSObject 
		$package | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($packageKey.GetValue("DisplayName"))
		$package | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($packageKey.GetValue("DisplayVersion"))
		$package | Add-Member -MemberType NoteProperty -Name "UninstallString" -Value $($packageKey.GetValue("UninstallString")) 
		$package | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($packageKey.GetValue("Publisher")) 
		$package | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computer
		$packages += $package	
	}
	return $packages
}

function Get-InstalledSoftwares
{
	[CmdletBinding()]
	param
	(
		[string] $Computer
	)

	$installedSoftwares = @{}
	$path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
	if($Computer -eq $env:COMPUTERNAME)
	{
		$registry32 = [microsoft.win32.registrykey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
	}
	else
	{
		$registry32 = [microsoft.win32.registrykey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer , [Microsoft.Win32.RegistryView]::Registry32)
	}
	$packages = RetrievePackages $path $registry32 $Computer
	if($Computer -eq $env:COMPUTERNAME)
	{
		$registry64 = [microsoft.win32.registrykey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
	}
	else
	{
		$registry64 = [microsoft.win32.registrykey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer , [Microsoft.Win32.RegistryView]::Registry64)
	}
	$packages += RetrievePackages $path $registry64 $Computer

	$packages.Where({$_.DisplayName}) |% { 
		if(-not($installedSoftwares.ContainsKey($_.DisplayName)))
		{
			$installedSoftwares.Add($_.DisplayName, $_) 
		}
	}

	return $installedSoftwares

}

function Get-InstalledSoftwaresOnServers
{
	[CmdletBinding()]
	param
	(
		[string[]] $Servers,
		[string] $Output
	)

	$console = $Host.UI.RawUI
	$windowSizeOld = $console.WindowSize
	$bufferSizeOld = $console.BufferSize
	$bufferSizeNew = $bufferSizeOld

	$bufferSizeNew.Width = 300
	$console.BufferSize = $bufferSizeNew
	if(-not([string]::IsNullOrEmpty($Output)))
	{
		if(-not(Test-Path $Output))
		{
			New-Item -ItemType File -Path $Output -Force | Out-Null
		}
		Start-Transcript $Output
	}

	$Servers |% {
		Write-Host "Computer name : $($_)"
		$installedSoftwares = Get-InstalledSoftwares $_
		$installedSoftwares.GetEnumerator() |% { $_.Value } | select DisplayName, DisplayVersion, Publisher, UninstallString | ft -AutoSize -Wrap -GroupBy Publisher
	}
	if(-not([string]::IsNullOrEmpty($Output)))
	{
		Stop-Transcript
	}
	
	$console.BufferSize = $bufferSizeOld	
}

Export-ModuleMember -Function RetrievePackages
Export-ModuleMember -Function Get-InstalledSoftwares
Export-ModuleMember -Function Get-InstalledSoftwaresOnServers