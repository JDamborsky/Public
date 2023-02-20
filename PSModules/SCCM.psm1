
# Example:  Trigger-AppInstallation -Computername spro8 -AppName "7-zip" -Method Install
Function Trigger-AppInstallation
{
 
Param
(
    [String][Parameter(Mandatory=$True, Position=1)] $Computername,
    [String][Parameter(Mandatory=$True, Position=2)] $AppName,
    [ValidateSet("Install","Uninstall")]
    [String][Parameter(Mandatory=$True, Position=3)] $Method
)
 
Begin {
    $Application = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName $Computername | Where-Object {$_.Name -like $AppName})
    
    $Args = @{EnforcePreference = [UINT32] 0
    Id = "$($Application.id)"
    IsMachineTarget = $Application.IsMachineTarget
    IsRebootIfNeeded = $False
    Priority = 'High'
    Revision = "$($Application.Revision)" }
 
}
 
Process
{
    Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Application -ComputerName $Computername -MethodName $Method -Arguments $Args
}
 
End {}
 
}

#Example: Get-SCCMDeviceCollectionDeployment -DeviceName spro -SiteCode sv8 -computername sccm03 
function Get-SCCMDeviceCollectionDeployment {
    <#
    .SYNOPSIS
        Function to retrieve a Device targeted application(s)

    .DESCRIPTION
        Function to retrieve a Device targeted application(s).
        The function will first retrieve all the collection where the Device is member of and
        find deployment advertised to those.

    .PARAMETER Devicename
        Specifies the SamAccountName of the Device.
        The Device must be present in the SCCM CMDB

    .PARAMETER SiteCode
        Specifies the SCCM SiteCode

    .PARAMETER ComputerName
        Specifies the SCCM Server to query

    .PARAMETER Credential
        Specifies the credential to use to query the SCCM Server.
        Default will take the current user credentials

    .PARAMETER Purpose
        Specifies a specific deployment intent.
        Possible value: Available or Required.
        Default is Null (get all)
    .EXAMPLE
        Get-SCCMDeviceCollectionDeployment -DeviceName MYCOMPUTER01 -Credential $cred -Purpose Required

    .NOTES
        Francois-Xavier cat
        lazywinadmin.com
        @lazywinadmin

        CHANGE HISTORY
            1.0 | 2015/09/03 | Francois-Xavier Cat
                Initial Version
            1.1 | 2017/09/15 | Francois-Xavier Cat
                Update Comment based help
                Update Crendential parameter type
                Update Verbose messages

        SMS_R_SYSTEM: https://msdn.microsoft.com/en-us/library/cc145392.aspx
        SMS_Collection: https://msdn.microsoft.com/en-us/library/hh948939.aspx
        SMS_DeploymentInfo: https://msdn.microsoft.com/en-us/library/hh948268.aspx
    .LINK
        https://github.com/lazywinadmin/PowerShell
#>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory)]
        [System.String]$DeviceName,

        [Parameter(Mandatory)]
        [System.String]$SiteCode,

        [Parameter(Mandatory)]
        [System.String]$ComputerName,

        [Alias('RunAs')]
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [ValidateSet('Required', 'Available')]
        [System.String]$Purpose
    )

    BEGIN {
        $FunctionName = (Get-Variable -Scope 1 -Name MyInvocation -ValueOnly).MyCommand.Name

        Write-Verbose -Message "[$FunctionName] Create splatting"

        # Define default properties
        $Splatting = @{
            ComputerName = $ComputerName
            NameSpace    = "root\SMS\Site_$SiteCode"
        }

        IF ($PSBoundParameters['Credential']) {
            Write-Verbose -Message "[$FunctionName] Append splatting"
            $Splatting.Credential = $Credential
        }

        Switch ($Purpose) {
            "Required" { $DeploymentIntent = 0 }
            "Available" { $DeploymentIntent = 2 }
            default { $DeploymentIntent = "NA" }
        }

        Write-Verbose -Message "[$FunctionName] Define helper functions"
        Write-Verbose -Message "[$FunctionName] helper function: Get-SCCMDeploymentIntentName"
        Function Get-SCCMDeploymentIntentName {
            PARAM(
                [Parameter(Mandatory)]
                $DeploymentIntent
            )
            PROCESS {
                if ($DeploymentIntent -eq 0) { Write-Output "Required" }
                if ($DeploymentIntent -eq 2) { Write-Output "Available" }
                if ($DeploymentIntent -ne 0 -and $DeploymentIntent -ne 2) { Write-Output "NA" }
            }
        } #Function Get-DeploymentIntentName

        Write-Verbose -Message "[$FunctionName] helper function: Get-SCCMDeploymentTypeName"
        function Get-SCCMDeploymentTypeName {
            <#
            https://msdn.microsoft.com/en-us/library/hh948731.aspx
            #>
            PARAM ($TypeID)
            switch ($TypeID) {
                1 { "Application" }
                2 { "Program" }
                3 { "MobileProgram" }
                4 { "Script" }
                5 { "SoftwareUpdate" }
                6 { "Baseline" }
                7 { "TaskSequence" }
                8 { "ContentDistribution" }
                9 { "DistributionPointGroup" }
                10 { "DistributionPointHealth" }
                11 { "ConfigurationPolicy" }
            }
        }

    }
    PROCESS {
        TRY {

            Write-Verbose -Message "[$FunctionName] Retrieving Device '$DeviceName'..."
            $Device = Get-WMIObject @Splatting -Query "Select * From SMS_R_SYSTEM WHERE Name='$DeviceName'"

            Write-Verbose -Message "[$FunctionName] Retrieving collection(s) where the device is member..."
            Get-WmiObject -Class sms_fullcollectionmembership @splatting -Filter "ResourceID = '$($Device.resourceid)'" | ForEach-Object -Process {

                Write-Verbose -Message "[$FunctionName] Retrieving collection '$($_.Collectionid)'..."
                $Collections = Get-WmiObject @splatting -Query "Select * From SMS_Collection WHERE CollectionID='$($_.Collectionid)'"

                Foreach ($Collection in $collections) {
                    IF ($DeploymentIntent -eq 'NA') {
                        Write-Verbose -Message "[$FunctionName] DeploymentIntent is not specified"
                        $Deployments = (Get-WmiObject @splatting -Query "Select * From SMS_DeploymentInfo WHERE CollectionID='$($Collection.CollectionID)'")
                    }
                    ELSE {
                        Write-Verbose -Message "[$FunctionName] DeploymentIntent '$DeploymentIntent'"
                        $Deployments = (Get-WmiObject @splatting -Query "Select * From SMS_DeploymentInfo WHERE CollectionID='$($Collection.CollectionID)' AND DeploymentIntent='$DeploymentIntent'")
                    }

                    Foreach ($Deploy in $Deployments) {
                        Write-Verbose -Message "[$FunctionName] Retrieving DeploymentType..."
                        $TypeName = Get-SCCMDeploymentTypeName -TypeID $Deploy.DeploymentTypeid
                        if (-not $TypeName) { $TypeName = Get-SCCMDeploymentTypeName -TypeID $Deploy.DeploymentType }

                        # Prepare output
                        Write-Verbose -Message "[$FunctionName] Preparing output..."
                        $Properties = @{
                            DeviceName           = $DeviceName
                            ComputerName         = $ComputerName
                            CollectionName       = $Deploy.CollectionName
                            CollectionID         = $Deploy.CollectionID
                            DeploymentID         = $Deploy.DeploymentID
                            DeploymentName       = $Deploy.DeploymentName
                            DeploymentIntent     = $deploy.DeploymentIntent
                            DeploymentIntentName = (Get-SCCMDeploymentIntentName -DeploymentIntent $deploy.DeploymentIntent)
                            DeploymentTypeName   = $TypeName
                            TargetName           = $Deploy.TargetName
                            TargetSubName        = $Deploy.TargetSubname

                        }

                        #Output the current object
                        Write-Verbose -Message "[$FunctionName] Output information"
                        New-Object -TypeName PSObject -prop $Properties

                        # Reset TypeName
                        $TypeName = ""
                    }
                }
            }
        }
        CATCH {
            $PSCmdlet.ThrowTerminatingError()
        }
    }
}


#Example:  Get-SCCMUserCollectionDeployment -UserName jarle -SiteCode sv8 -computername sccm03 -Purpose Available
function Get-SCCMUserCollectionDeployment {
    <#
    .SYNOPSIS
        Function to retrieve a User's collection deployment

    .DESCRIPTION
        Function to retrieve a User's collection deployment
        The function will first retrieve all the collection where the user is member of and
        find deployments advertised on those.

        The final output will include user, collection and deployment information.

    .PARAMETER Username
        Specifies the SamAccountName of the user.
        The user must be present in the SCCM CMDB

    .PARAMETER SiteCode
        Specifies the SCCM SiteCode

    .PARAMETER ComputerName
        Specifies the SCCM Server to query

    .PARAMETER Credential
        Specifies the credential to use to query the SCCM Server.
        Default will take the current user credentials

    .PARAMETER Purpose
        Specifies a specific deployment intent.
        Possible value: Available or Required.
        Default is Null (get all)

    .EXAMPLE
        Get-SCCMUserCollectionDeployment -UserName TestUser -Credential $cred -Purpose Required

    .NOTES
        Francois-Xavier cat
        lazywinadmin.com
        @lazywinadmin

        SMS_R_User: https://msdn.microsoft.com/en-us/library/hh949577.aspx
        SMS_Collection: https://msdn.microsoft.com/en-us/library/hh948939.aspx
        SMS_DeploymentInfo: https://msdn.microsoft.com/en-us/library/hh948268.aspx
    .LINK
        https://github.com/lazywinadmin/PowerShell
#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory)]
        [Alias('SamAccountName')]
        $UserName,

        [Parameter(Mandatory)]
        $SiteCode,

        [Parameter(Mandatory)]
        $ComputerName,

        [Alias('RunAs')]
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [ValidateSet('Required', 'Available')]
        $Purpose
    )

    BEGIN {
        # Verify if the username contains the domain name
        #  If it does... remove the domain name
        # Example: "FX\TestUser" will become "TestUser"
        if ($UserName -like '*\*') { $UserName = ($UserName -split '\\')[1] }

        # Define default properties
        $Splatting = @{
            ComputerName = $ComputerName
            NameSpace    = "root\SMS\Site_$SiteCode"
        }

        IF ($PSBoundParameters['Credential']) {
            $Splatting.Credential = $Credential
        }

        Switch ($Purpose) {
            "Required" { $DeploymentIntent = 0 }
            "Available" { $DeploymentIntent = 2 }
            default { $DeploymentIntent = "NA" }
        }

        Function Get-DeploymentIntentName {
            PARAM(
                [Parameter(Mandatory)]
                $DeploymentIntent
            )
            PROCESS {
                if ($DeploymentIntent -eq 0) { Write-Output "Required" }
                if ($DeploymentIntent -eq 2) { Write-Output "Available" }
                if ($DeploymentIntent -ne 0 -and $DeploymentIntent -ne 2) { Write-Output "NA" }
            }
        }#Function Get-DeploymentIntentName


    }
    PROCESS {
        # Find the User in SCCM CMDB
        $User = Get-WMIObject @Splatting -Query "Select * From SMS_R_User WHERE UserName='$UserName'"

        # Find the collections where the user is member of
        Get-WmiObject -Class sms_fullcollectionmembership @splatting -Filter "ResourceID = '$($user.resourceid)'" |
            ForEach-Object -Process {

                # Retrieve the collection of the user
                $Collections = Get-WmiObject @splatting -Query "Select * From SMS_Collection WHERE CollectionID='$($_.Collectionid)'"


                # Retrieve the deployments (advertisement) of each collections
                Foreach ($Collection in $collections) {
                    IF ($DeploymentIntent -eq 'NA') {
                        # Find the Deployment on one collection
                        $Deployments = (Get-WmiObject @splatting -Query "Select * From SMS_DeploymentInfo WHERE CollectionID='$($Collection.CollectionID)'")
                    }
                    ELSE {
                        $Deployments = (Get-WmiObject @splatting -Query "Select * From SMS_DeploymentInfo WHERE CollectionID='$($Collection.CollectionID)' AND DeploymentIntent='$DeploymentIntent'")
                    }

                    Foreach ($Deploy in $Deployments) {

                        # Prepare Output
                        $Properties = @{
                            UserName             = $UserName
                            ComputerName         = $ComputerName
                            CollectionName       = $Deploy.CollectionName
                            CollectionID         = $Deploy.CollectionID
                            DeploymentID         = $Deploy.DeploymentID
                            DeploymentName       = $Deploy.DeploymentName
                            DeploymentIntent     = $deploy.DeploymentIntent
                            DeploymentIntentName = (Get-DeploymentIntentName -DeploymentIntent $deploy.DeploymentIntent)
                            TargetName           = $Deploy.TargetName
                            TargetSubName        = $Deploy.TargetSubname

                        }

                        # Output the current Object
                        New-Object -TypeName PSObject -prop $Properties
                    }
                }
            }
    }
}



Export-ModuleMember -Function Get-SCCMUserCollectionDeployment 
Export-ModuleMember -Function Get-SCCMDeviceCollectionDeployment
Export-ModuleMember -Function Trigger-AppInstallation


