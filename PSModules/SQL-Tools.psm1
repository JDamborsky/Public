



Function SetConnStr-LAB
{
    $global:SqlSccmConStr       = 'Data Source=SCCM03;Initial Catalog=CM_SV8;Integrated Security=True;User ID=;Password='
    $global:SqlWin10BaseStr     = 'Data Source=SCCM03;Initial Catalog=W10Mapping;Integrated Security=True;User ID=;Password='
    $Global:MultiUserLimit      = 3
    $Global:RunEnv              = 'LAB'
}


Function OpenSqlCon{

    # Get DB-server name
    $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    Switch ($CurrentDomain) {        
        'damborsky.com'                 {SetConnStr-LAB}
        Default                         {SetConnStr-Sikt}
    }

    #connect to database
    $global:connection      = New-Object System.Data.SqlClient.SqlConnection($global:SqlWin10BaseStr)
    $global:connection.Open()
        
    #connect to database
    $global:connectionSccm  = New-Object System.Data.SqlClient.SqlConnection($global:SqlSccmConStr)
    $global:connectionSccm.Open()
}




Function CloseSqlCon{

    $global:connection.Close()
    $global:connectionSccm.Close()
    $global:connection      = $null
    $global:connectionSccm  = $null
}

function ExecuteSqlQueryCommand{
    param (
        [Parameter(Mandatory = $true)]
        [String[]]$ExecuteQueryString,
        $ConnectionObject
        )
 
    #build query object
    $command = $ConnectionObject.CreateCommand()
    $command.CommandText = $ExecuteQueryString
    $command.CommandTimeout = 9000
    
    #run query
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    try {
        $adapter.Fill($dataset) | out-null
    }
    catch
    {
        $SqlErr = $_
         
        Write-host $ExecuteQueryString
        Write-host $SqlErr
        logg $ExecuteQueryString
        logg $SqlErr

    }

    #return the first collection of results or an empty array
    If ($dataset.Tables[0] -ne $null) {$table = $dataset.Tables[0]}
    ElseIf ($table.Rows.Count -eq 0) { $table = New-Object System.Collections.ArrayList }
    
    return $table    
}

Function ExecuteSqlQuery
{
    param (
        [Parameter(Mandatory = $true)]
        [String[]]$ExecuteQueryString    )

    if (!$global:SqlWin10BaseStr)  {OpenSqlCon}  
    $TmpValue = ExecuteSqlQueryCommand $ExecuteQueryString $global:connection
    return $TmpValue
}

Function ExecuteSccmSqlQuery
{
    param (
        [Parameter(Mandatory = $true)]
        [String[]]$ExecuteQueryString       )

    if (!$global:connectionSccm)  {OpenSqlCon}    
    $TmpValue = ExecuteSqlQueryCommand $ExecuteQueryString $global:connectionSccm
    return $TmpValue
}



Export-ModuleMember -Function ExecuteSccmSqlQuery
Export-ModuleMember -Function ExecuteSqlQuery
Export-ModuleMember -Function ExecuteSqlQueryCommand
Export-ModuleMember -Function CloseSqlCon
Export-ModuleMember -Function OpenSqlCon
