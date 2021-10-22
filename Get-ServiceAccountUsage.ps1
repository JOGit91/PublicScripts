<# 

This script searches all computers in an OU for services that are using a specific account name. 

User inputs OU and Service Account name

$OU should be in in this format: "OU=Servers,OU=Contoso,DC=Contoso,DC=local"

$ServicAccount should be in this format: "contoso.local\\svcaccount"
	Note: When including domain and using backslash (\), make sure to double up on backslash. This is due to how
            Powershell processes the single backslash. Double backslash will process normally. If you get a ton
            of WARNING output, you probably only did one backslash.
	Note: wildcards (%) can be used to search for just portions of a service account name: "%svcaccount%"

Jake Ouellette 10.22.21

#>

#Parameters
$ServiceAccount = 
$OU = 

#End of required parameters, do not edit below this comment
$Computers = Get-ADComputer -Filter * -SearchBase $OU | Select Name | Sort-Object Name
$Computers = $Computers.Name
Foreach ($Computer in $Computers)
{
Try {
    if (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
    {
      Write-host "Checking Services on $Computer"
$Test = Get-WmiObject Win32_Service -EA Stop -ComputerName $Computer -filter "STARTNAME LIKE '$ServiceAccount'" |  select DisplayName,Name,StartName,ProcessID,StartMode,State
if ($Test -ne $null){

Write-host "Found Services running under $service on $computer" -ForegroundColor green
$Test | Format-Table | Out-String|% {Write-Host $_}

}
Else{
Write-host "No Services running under $service on $computer" -ForegroundColor red
}

    } else {
       Write-Host "$Computer is not responding to ping, skipping" -ForegroundColor yellow
    }
}
Catch{
Write-Warning "$Computer is responding to ping, but cannot be processed. Likely RPC service is not running or $Computer is not running Windows. Investigate manually."
}
} 
