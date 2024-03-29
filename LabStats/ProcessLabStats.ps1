#Set-PSDebug -strict
#Process Head Count on All Lab Machines

function getMax 
{           
            $query = "SELECT Max(logontime) FROM $global:db"
            $cmdmax = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
            try{[datetime]$max = $cmdmax.ExecuteScalar()}
            catch
            {
                Write-Host "Getmax query failed, getting entries from last 4 months." $error[0]
                $max = (Get-Date).AddDays(-120)
            }
            return $max
}

function singleMachine{
    param ($cmdstr, $cmdstr2)

    $labMachine =  "stulab132" #"stulab101" stulab135 #"lawnet7"
    $cmd = New-Object System.Data.SqlClient.SqlCommand($cmdStr, $conn)
    $cmd2 = New-Object System.Data.SqlClient.SqlCommand($cmdStr2, $conn)

    #get last logon entry, to knwo where to start. 
    $max= getMax $conn
    Process-Logon $labMachine $cmd $max
    Process-Logoff $labMachine $cmd2 $max
}

function allMachines {
    param ($cmdstr, $cmdstr2)

    ########### Run Query on all lab machines ###########
    [string[]]$labsOU = "OU=Labs - Win 7,OU=Network Resources,DC=lawnet,DC=lcl", "OU=Clinical,OU=Network Resources,DC=lawnet,DC=lcl"
    foreach($ou in $labsOU){
        $labmachineCollection = [ADSI]("LDAP://$ou")
        
        foreach($machine in $labmachineCollection.children) {
            $labMachine = $machine.cn
            $bool = $TRUE;
            [string[]]$exclude = "stulabfrontdesk", "mootcourt7", "BLSA7", "APALJ7"
            foreach ($ex in $exclude){ if ($ex -eq $labmachine) {$bool = $FALSE}}
            
            if($bool){
               #get last logon entry, to knwo where to start. 
                $max= getMax
                
                $cmd = New-Object System.Data.SqlClient.SqlCommand($cmdStr, $conn)
                Process-Logon $labMachine $cmd $max
                
                #repeat get max with logoff
                $cmd2 = New-Object System.Data.SqlClient.SqlCommand($cmdStr2, $conn)
                Process-Logoff $labMachine $cmd2 $max
                break
            } #endif
        }#endforeach Machine
        break
    }#endfoeach OU#>
}

function Process-All{
    param(
        [string] $connStr = "data source=UCLAWSQL1;initial catalog=General;Integrated Security=true;",
        [string] $cmdStr = "IF NOT EXISTS ( select * from $global:db where computername=@computername and logontime=@logontime and logonid=@logonid) begin INSERT INTO $global:db (logontime,username,computername,logonid) VALUES (@logontime, @username, @computername,@logonid) end", 
        [string] $cmdStr2 = "UPDATE $global:db SET logofftime=@logofftime WHERE computername=@computername and username=@username and logonid=@logonid"
    )
    #open DB connection
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.open()
    
    #singleMachine $cmdstr $cmdstr2
    allMachines $cmdstr $cmdstr2
    
    ##### run stored proc on sql to calc session length ##########
    $sproccmd = New-Object System.Data.SqlClient.SqlCommand("exec dbo.calcLogonLength",$conn)
    try{$sproccmd.ExecuteNonQuery()}
    catch{Write-Host "Failed stored procedure: " $error[0]}
    #############################>

    $conn.close()

}#end Process-All


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
        $scriptFile = ($scriptLoc + "\LabStats.psm1")
    }
    else
    {
        $scriptFile = ($scriptLoc + "LabStats.psm1")
    }
    
    if(Test-Path -LiteralPath $scriptFile -PathType Leaf)
    {  
        if(Get-Module | Where-Object { $_.Name -eq "LabStats" })
        {
            Remove-Module -Name "LabStats"
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

#Process-Lab-Machines
[string]$global:failures = $null
[string]$global:db = "[General].[dbo].Lab_Stats"
[string]$global:server = "UCLAWSQL1"
$error.clear() 
$now = [DateTime]::Now.Date.ToString("yyyy.MM.dd")
$scriptpath = $MyInvocation.MyCommand.Path
$path = (Split-Path $scriptpath) + "\logs\LabStatsLog - $now.txt"

Load-Module
#startLog $path $now
Process-All
$err = writeErrors
#stop-transcript
try {sendLog $path $now $err}
catch{Write-Host "error sending log." $error[0]}

#>