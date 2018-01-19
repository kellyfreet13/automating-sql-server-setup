$driveletter = $pwd.Drive.Name

# fool-proofing the sqlps import
$pspath = $driveletter + ":\Program Files (x86)\Microsoft SQL Server\130\Tools\PowerShell\Modules"
$testpath = Test-Path $pspath
if ($testpath){
    # be extra sure it'll work
    if(!$env:PSModulePath.Contains($driveletter + ":\Program Files (x86)\Microsoft SQL Server\130\Tools\PowerShell\Modules")){
        $env:PSModulePath += ";" + $driveletter + ":\Program Files (x86)\Microsoft SQL Server\130\Tools\PowerShell\Modules"
    }
    Import-Module "sqlps"
} else {
    Read-Host -Prompt "Are you sure you have SQL Server 2016 installed?`n Press enter to exit script"
    Exit
}

$smo = 'Microsoft.SqlServer.Management.Smo.'
$wmi = New-Object ($smo + 'Wmi.ManagedComputer').
 
# site audit uses sqlexpress, baby
$sqlinstance = 'SQLEXPRESS'
$uri = "ManagedComputer[@Name='" + (get-item env:\computername).Value + "']/ServerInstance[@Name='$($sqlinstance)']/ServerProtocol[@Name='Tcp']"
$Tcp = $wmi.GetSmoObject($uri)
$Tcp.IsEnabled = $true
$Tcp.Alter()

# set tcp ports without shitty vbscript
$validips = $Tcp.IPAddresses | ? {$_.Name -like 'IP*'}
$ipproperties = $validips.IPAddressProperties
$ipproperties | foreach { if($_.Name -like 'TcpPort') {$_.Value = "1433"}}
$Tcp.Alter()

# restart the windows service for SQL Express
Restart-Service -Name 'MSSQL$SQLEXPRESS'

# note: SQL Server Express 2016 connection string ends in 13

# allow sql server access through the windows firewall
New-NetFirewallRule -DisplayName "MSSQL" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow

