Set-PSDebug -strict

function Process-Account
{
    param(
        [int]      $Year = 2016,
        [string]   $connStr = "data source=UCLAWSQL1;initial catalog=Admissions;Integrated Security=true;",
        [string]   $cmdStr = "BulkAddAdmits", 
        [string]   $admittedDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl",
        [string]   $studentDn = "OU=Students,OU=Network Users,DC=lawnet,DC=lcl",
        [string]   $homeRoot = "\\uclawdata1\StuHome\",
        [string]   $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl",
        [string]   $admittedNotComingOu = "OU=Admitted Students - Not Coming,OU=Network Users,DC=lawnet,DC=lcl",
        [DateTime] $moveStart = "8/1/" + [DateTime]::Now.Year.ToString(),
        [DateTime] $moveEnd = "9/1/" + [DateTime]::Now.Year.ToString(),
        [string]   $upnDomain = "lawnet.lcl",
        [string]   $emailDomain = "lawnet.ucla.edu"
    )
    
    #Build group tables
    $admitTable = New-Object "System.Collections.Generic.Dictionary[string,string]"
    $studentTable = New-Object "System.Collections.Generic.Dictionary[string,string]"
    
    $admitTable.Add("JD Class of {0}" -f $Year, "CN=Admitted Students - JD,OU=Network Users,DC=lawnet,DC=lcl")
    $admitTable.Add("Transfer Class of {0}" -f ($Year - 1), "CN=Admitted Students - Transfer 2L,OU=Network Users,DC=lawnet,DC=lcl")
    $admitTable.Add("Transfer Class of {0}" -f ($Year - 2), "CN=Admitted Students - Transfer 3L,OU=Network Users,DC=lawnet,DC=lcl")
    $admitTable.Add("JD Class of {0}/Visiting Student" -f ($Year - 2), "CN=Admitted Students - Transfer 3L,OU=Network Users,DC=lawnet,DC=lcl")
    $admitTable.Add("LLM Class of {0}" -f ($Year - 2), "CN=Admitted Students - LLM,OU=Network Users,DC=lawnet,DC=lcl")
    $admitTable.Add("Exchange Student", "CN=Admitted Students - LLM,OU=Network Users,DC=lawnet,DC=lcl")
    
    $studentTable.Add("JD Class of {0}" -f $Year, "CN=Class of {0},OU=Network Users,DC=lawnet,DC=lcl" -f $Year)
    $studentTable.Add("Transfer Class of {0}" -f ($Year - 1), "CN=Class of {0},OU=Network Users,DC=lawnet,DC=lcl" -f ($Year - 1))
    $studentTable.Add("Transfer Class of {0}" -f ($Year - 2), "CN=Class of {0},OU=Network Users,DC=lawnet,DC=lcl" -f ($Year - 2))
    $studentTable.Add("JD Class of {0}/Visiting Student" -f ($Year - 2), "CN=Class of {0},OU=Network Users,DC=lawnet,DC=lcl" -f ($Year -2))
    $studentTable.Add("LLM Class of {0}" -f ($Year - 2), "CN=LLM Students,OU=Network Users,DC=lawnet,DC=lcl")
    $studentTable.Add("Exchange Student", "CN=LLM Students,OU=Network Users,DC=lawnet,DC=lcl")
    
    
    $cn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $cmd = New-Object System.Data.SqlClient.SqlCommand($cmdStr, $cn)
    $da = New-Object System.Data.SqlClient.SqlDataAdapter
    $ds = New-Object System.Data.DataSet

    $da.SelectCommand = $cmd
    $da.Fill($ds) | Out-Null

    $cn.Close()
    
    $admitGrp = $admitTable.Values | Sort-Object -Unique
    
    
    $ds.Tables[0].Rows | ForEach-Object { try { Add-NewAdmit $_ $admitTable.Item($_.Description) $admittedDn $studentTable.Item($_.Description)$studentDn $homeRoot <#$googleAppDn#> $moveStart $moveEnd $upnDomain $emailDomain } catch {Write-Host "Error creating account: $_.CommonName" $error[0]} }
    
	#$ds.Tables[2].Rows | ForEach-Object { Add-SIR $_ $admitTable.Item($_.Description) $admittedDn $googleAppDn}
    
    $ds.Tables[1].Rows | ForEach-Object { try{ Disable-Account $_.CommonName $admittedDn $studentDn $googleAppDn $admittedNotComingOu $admitGrp } catch {Write-Host "Error running Disable-Account" $error[0]}}
    

    $ds.Dispose()
    
    #in case we still are ready, we will move all new accounts
    if($moveStart -le [DateTime]::Now -and $moveEnd -gt [DateTime]::Now)
    {
        $admittedOu = [ADSI]("LDAP://$admittedDn")
        
        foreach($admitType in $admitTable.Keys)
        {
            #add check to make sure the account is in googleAppsGroup
            $admittedOu.Children | Where-Object { $_.description -eq $admitType } | ForEach-Object { Move-Account $_ $admitTable.Item($_.description) $studentTable.Item($_.description) $studentDn }
        }
    }
}

function Load-Module
{
    param(
        [string] $scriptPath
    )
    
    if((!$scriptPath) -or $scriptPath -eq "")
    {
        $scriptLoc = Split-Path -Parent $MyInvocation.ScriptName
    }
    else
    {
        $scriptLoc = $scriptPath
    }
    
    if(!$scriptLoc.EndsWith("\"))
    {
        $scriptFile = ($scriptLoc + "\LawAccount.psm1")
    }
    else
    {
        $scriptFile = ($scriptLoc + "LawAccount.psm1")
    }
    
    if(Test-Path -LiteralPath $scriptFile -PathType Leaf)
    {  
        if(Get-Module | Where-Object { $_.Name -eq "LawAccount" })
        {
            Remove-Module -Name "LawAccount"
        }
        
        Import-Module -Name $scriptFile
    }
    else
    {
        throw ("Unable to locate script module " + $scriptFile)
    }
    
    if(!(Get-Command | Where-Object {$_.Name -eq "Enable-MailUser"}))
    {
        ImportSystemModules
    }
}

###EXECUTION###

[int]$global:createdCount = 0
[int]$global:failCount = 0
[string]$global:usersCreated = " "
[string]$global:usersFail = " "
$Year = 2016
$now = [DateTime]::Now.Date.ToString("yyyy.MM.dd")
$scriptpath = $MyInvocation.MyCommand.Path
$path = (Split-Path $scriptpath) + "\logs\NewAdmitLog - $now.txt"

Load-Module
startLog $path $now
Process-Account $Year
$err = writeErrors
stop-transcript
sendLog $path $now $err
