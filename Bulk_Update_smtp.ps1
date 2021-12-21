<# Written by Jake Ouellette 9.24.20
.NAME
    Bulk_Add_smtp
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(416,267)
$Form.text                       = "Bulk Add smtp address"
$Form.TopMost                    = $false

$Lbl_OU                          = New-Object system.Windows.Forms.Label
$Lbl_OU.text                     = "User OU you would like to edit:"
$Lbl_OU.AutoSize                 = $true
$Lbl_OU.width                    = 25
$Lbl_OU.height                   = 10
$Lbl_OU.location                 = New-Object System.Drawing.Point(11,26)
$Lbl_OU.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Domain to add:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(11,107)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OUText                          = New-Object system.Windows.Forms.MaskedTextBox
$OUText.multiline                = $false
$OUText.text                     = "`OU=Users,DC=domain,DC=local"
$OUText.width                    = 384
$OUText.height                   = 21
$OUText.location                 = New-Object System.Drawing.Point(8,50)
$OUText.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$startButton                     = New-Object system.Windows.Forms.Button
$startButton.text                = "OK"
$startButton.width               = 60
$startButton.height              = 30
$startButton.location            = New-Object System.Drawing.Point(163,190)
$startButton.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$startButton.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#03d8e0")

$NewDomainTxt                    = New-Object system.Windows.Forms.MaskedTextBox
$NewDomainTxt.multiline          = $false
$NewDomainTxt.text               = "@seconddomain.com"
$NewDomainTxt.width              = 370
$NewDomainTxt.height             = 25
$NewDomainTxt.location           = New-Object System.Drawing.Point(11,129)
$NewDomainTxt.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$Form.controls.AddRange(@($Lbl_OU,$Label2,$OUText,$startButton,$NewDomainTxt))

$startButton.Add_Click({ btnOK_Click })

function btnOK_Click {


$users = Get-ADUser -Filter * -SearchBase $OUText.text -Properties mail

foreach ($user in $users ){

   $Secondary = "smtp:" + $User.SamAccountName + $NewDomainTxt.text

   Set-ADUser $User.SamAccountName -Add @{'ProxyAddresses'=$Secondary}

}

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Operation Completed!",0,"Done",0x0)

$form.Close()

}


[void]$Form.ShowDialog()