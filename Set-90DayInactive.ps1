<# 

This script can be ran once manually or set up as a scheduled task to run on a schedule and regularly scan for inactive accounts. 

Creates a logfile to track changes and accounts processed by the script

Adds "DISABLED_90D_INACTIVE $shortdate - $($userinfo.description)" to description field of accounts that it disables

Accounts set with "Password Never Expires" are skipped

Accounts with .DA or .EA in their name are considered admin accounts and are skipped

System accounts are identified and skipped according to this expression ($userinfo.samaccountname -match "^IUSR*"-or $userinfo.samaccountname -match "^IWAM*" -or $userinfo.samaccountname -like '*`$'

Final 6 lines of script are configuration for SMTP to email the report if desired. 

This script was not made by me, I take no credit for its creation. 

#>

$startscript = Get-Date
$domain = (Get-ADDomain).dnsroot

$path = "c:\scripts\90dayinactive\logs"
$logfilename = "$($startscript.tostring("yyyyMMddmmss"))"
$Logfile = "$path\$domain-90daydisable-$logfilename.txt"


$count_new = 0
$count_new = 0
$count_admin = 0
$count_passwordneverexpires = 0
$count_inactive = 0
$count_system = 0
$hostname = $env:COMPUTERNAME

Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}

logwrite "$domain - 90 day inactive accounts."
logwrite "Start Script - $startscript"
logwrite "---"
logwrite "Type,Samaccountname,Givenname,Surname,Created,LastLogon,PasswordLastSet,PasswordAge,Passwordexpired,Action"

$startscript = Get-Date

$body = "$domain - 90 day inactive accounts.`r`n"

$90Days = (get-date).adddays(-90)
$inactive = Get-ADUser -properties enabled, lastlogondate,lastlogontimestamp, passwordlastset,enabled,passwordneverexpires,whencreated -filter {(lastlogondate -notlike "*" -OR lastlogondate -le $90days) -AND (passwordlastset -le $90days) -AND (enabled -eq $True) -and (whencreated -le $90days)} | select-object name, SAMaccountname, passwordExpired, PasswordNeverExpires, logoncount, whenCreated, lastlogondate, PasswordLastSet, lastlogontimestamp
$todaydate = Get-Date
$shortdate = $todaydate.ToString("yyMMdd")

foreach ($user in $inactive)
{
$testthisuser = $user.SamAccountName
$newaccounts = ($todaydate).AddDays(-90)
$userinfo = get-aduser $testthisuser  -properties whencreated,lastlogondate,description,passwordneverexpires,passwordlastset,passwordexpired
$newdescription = "DISABLED_90D_INACTIVE $shortdate - $($userinfo.description)" 
$passwordage = $userinfo.PasswordLastSet
if ($passwordage -eq $null)
{$passwordage = "NeverSet"}
else
{
 $passwordage = [math]::round( ((get-date)-$passwordage).totaldays)}

if ($userinfo.samaccountname -like "*.DA" -or $userinfo.samaccountname -like "*.EA")
{
$count_admin ++
$desc = "Admin_Account"
$message = "IGNORED_ADMIN_ACCOUNT"
}

elseif ($userinfo.passwordneverexpires -eq $true)
{
$count_passwordneverexpires ++
$desc = "PASSWORD_NEVER_EXPIRES"
$message = "IGNORED_PWD_NEV_EXP"
}

elseif ($userinfo.samaccountname -match "^IUSR*"-or $userinfo.samaccountname -match "^IWAM*" -or $userinfo.samaccountname -like '*`$')
{
$count_system ++
$desc = "SYSTEM_ACCOUNT"
$message= "IGNORED_SYS_ACCOUNT"
}

elseif ($userinfo.whencreated -gt $newaccounts)
{
$dayssincecreate = [math]::round( ((get-date)-($userinfo.whencreated)).totaldays)
"Ignore new account. $testthisuser $($userinfo.whencreated) - Created $dayssincecreate days ago"
$count_new ++
$desc = "NEW_ACCOUNT"
$message = "IGNORED_NEW_ACCOUNT"
} 

elseif ($passwordage -lt 90)
{
$desc = "PASSWORD_AGE"
$message = "PASSWORD_RECENTLY_SET"
} 



else 
{
$desc="INACTIVE"
$message="Account_disabled" 
$count_inactive ++
}


if (($message -eq "Account_disabled") -and ($userinfo.PasswordExpired -eq $true))
{
try {
     write-host -ForegroundColor yellow "$count_inactive Disabling $($userinfo.SamAccountName) - Passwordage: $passwordage"
     Disable-ADAccount -identity $userinfo.SamAccountName
     set-aduser -identity $userinfo.samaccountname -Description "$newdescription"
     }
catch 
    {
     $message= $_.exception.message 
    }

logwrite "$desc, $testthisuser,$($userinfo.givenname),$($userinfo.surname),$($userinfo.whencreated),$($userinfo.lastlogondate),$($userinfo.passwordlastset),$passwordage,$($userinfo.passwordexpired),$message"
}


}

logwrite "---"
logwrite "$count_admin inactive Admin accounts."
logwrite "$count_passwordneverexpires inactive accounts with PasswordNeverExpires." 
logwrite "$count_system inactive system accounts."
logwrite "$count_new inactive newly created accounts."
logwrite "$count_inactive inactive accounts found.  Accounts disabled."
logwrite "---"

$body += "Inactive script complete`r`n"
$body += "$count_admin inactive Admin accounts.`r`n"
$body += "$count_passwordneverexpires inactive accounts with PasswordNeverExpires.`r`n" 
$body += "$count_system inactive system accounts.`r`n"
$body += "$count_new inactive newly created accounts.`r`n"
$body += "$count_inactive inactive accounts found and disabled.`r`n"
$body += "-----`r`n"



$endscript = get-date
$timespan = $endscript - $startscript
logwrite "Script time to Run = $($timespan.minutes) minutes $($timespan.seconds) seconds"
logwrite "$hostname"
$body += "$hostname`r`n"
$body += "Script time to Run = $($timespan.minutes) minutes $($timespan.seconds) seconds`r`n"
#--------------------
Start-Sleep 5
#--------------------
$from    = "AD-Account-Report@domain.com"
$to      = "example@domain.com"
$bcc    =  "example2@domain.com"
$subject = "$Domain - 90 day inactive account report - $count_inactive accounts disabled."
$smtp    = "your-smtp-server-here"
Send-MailMessage -From $from -To $to -Subject $subject -Body $body -Attachments $logfile -SmtpServer $smtp -port 25
