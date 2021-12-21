<#
.NAME
    365 Email Search and Purge
   
.SYNOPSIS
    Search Office 365 tenant for an email and purge from all mailboxes. This tool features a GUI for usage.
    
.PARAMETERS
    Define a name for the search. Email search queries include date received, email subject, and sender.
    
.NOTES
    Features a pop out window to review search results before initiating purge. Also includes search status and purge status indicators.
    Best run using ISE.

.ISSUES
    All input fields must contain a value.
    Clicking on the purge puts the prompt to continue in the background powershell window. Keep an eye out for it.
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Window                          = New-Object system.Windows.Forms.Form
$Window.ClientSize               = New-Object System.Drawing.Point(1268,534)
$Window.text                     = "365 Email Search and Purge"
$Window.TopMost                  = $false

$Get_365_Creds                   = New-Object system.Windows.Forms.Button
$Get_365_Creds.text              = "Enter 365 Admin Credentials"
$Get_365_Creds.width             = 200
$Get_365_Creds.height            = 30
$Get_365_Creds.location          = New-Object System.Drawing.Point(20,24)
$Get_365_Creds.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Get_Search_Name                 = New-Object system.Windows.Forms.TextBox
$Get_Search_Name.multiline       = $false
$Get_Search_Name.width           = 359
$Get_Search_Name.height          = 20
$Get_Search_Name.location        = New-Object System.Drawing.Point(38,136)
$Get_Search_Name.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Provide a name for this search:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(16,109)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ToolTip1                        = New-Object system.Windows.Forms.ToolTip
$ToolTip1.ToolTipTitle           = "Search Name"

$Get_Search_Date                 = New-Object system.Windows.Forms.TextBox
$Get_Search_Date.multiline       = $false
$Get_Search_Date.width           = 357
$Get_Search_Date.height          = 20
$Get_Search_Date.location        = New-Object System.Drawing.Point(38,193)
$Get_Search_Date.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Get_Search_Subject              = New-Object system.Windows.Forms.TextBox
$Get_Search_Subject.multiline    = $false
$Get_Search_Subject.width        = 354
$Get_Search_Subject.height       = 20
$Get_Search_Subject.location     = New-Object System.Drawing.Point(40,249)
$Get_Search_Subject.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Get_Search_Address              = New-Object system.Windows.Forms.TextBox
$Get_Search_Address.multiline    = $false
$Get_Search_Address.width        = 353
$Get_Search_Address.height       = 20
$Get_Search_Address.location     = New-Object System.Drawing.Point(40,303)
$Get_Search_Address.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "When was the email received? (mm/dd/yyyy):"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(16,167)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "What is the email Subject?"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(15,224)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "What email address did the email come from?"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(16,282)
$Label4.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ToolTip2                        = New-Object system.Windows.Forms.ToolTip
$ToolTip2.ToolTipTitle           = "Email Subject"

$IT_Glue                         = New-Object system.Windows.Forms.Button
$IT_Glue.text                    = "Take me to the IT Glue article!"
$IT_Glue.width                   = 240
$IT_Glue.height                  = 30
$IT_Glue.location                = New-Object System.Drawing.Point(233,24)
$IT_Glue.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Button_Start_Search             = New-Object system.Windows.Forms.Button
$Button_Start_Search.text        = "Start Search"
$Button_Start_Search.width       = 167
$Button_Start_Search.height      = 30
$Button_Start_Search.location    = New-Object System.Drawing.Point(40,334)
$Button_Start_Search.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button_Search_Status            = New-Object system.Windows.Forms.Button
$Button_Search_Status.text       = "Refresh Search Status"
$Button_Search_Status.width      = 169
$Button_Search_Status.height     = 30
$Button_Search_Status.location   = New-Object System.Drawing.Point(230,334)
$Button_Search_Status.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ToolTipStartSearch              = New-Object system.Windows.Forms.ToolTip

$Connection_Status               = New-Object system.Windows.Forms.Label
$Connection_Status.text          = "Please enter credentials..."
$Connection_Status.AutoSize      = $true
$Connection_Status.width         = 25
$Connection_Status.height        = 10
$Connection_Status.location      = New-Object System.Drawing.Point(29,73)
$Connection_Status.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Connection_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#de1717")

$Label_Status                    = New-Object system.Windows.Forms.Label
$Label_Status.text               = "Search not started"
$Label_Status.AutoSize           = $true
$Label_Status.width              = 25
$Label_Status.height             = 10
$Label_Status.location           = New-Object System.Drawing.Point(321,373)
$Label_Status.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Label_Status.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#d50303")

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "Search Status:"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(230,373)
$Label5.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Get_Results                     = New-Object system.Windows.Forms.Button
$Get_Results.text                = "Review Search Results"
$Get_Results.width               = 165
$Get_Results.height              = 30
$Get_Results.location            = New-Object System.Drawing.Point(41,377)
$Get_Results.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $true
$TextBox1.width                  = 754
$TextBox1.height                 = 489
$TextBox1.location               = New-Object System.Drawing.Point(492,22)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox1.ScrollBars             = "vertical"

$Button_Purge                    = New-Object system.Windows.Forms.Button
$Button_Purge.text               = "PURGE"
$Button_Purge.width              = 355
$Button_Purge.height             = 29
$Button_Purge.enabled            = $false
$Button_Purge.location           = New-Object System.Drawing.Point(40,451)
$Button_Purge.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Button_Purge.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#e11111")

$Button_Confirm_Purge            = New-Object system.Windows.Forms.Button
$Button_Confirm_Purge.text       = "I have reviewed the results and am ready to purge"
$Button_Confirm_Purge.width      = 408
$Button_Confirm_Purge.height     = 33
$Button_Confirm_Purge.location   = New-Object System.Drawing.Point(19,413)
$Button_Confirm_Purge.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button_Status_Purge             = New-Object system.Windows.Forms.Button
$Button_Status_Purge.text        = "Refresh Purge Status"
$Button_Status_Purge.width       = 148
$Button_Status_Purge.height      = 30
$Button_Status_Purge.location    = New-Object System.Drawing.Point(40,487)
$Button_Status_Purge.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label_Purge                     = New-Object system.Windows.Forms.Label
$Label_Purge.text                = "Purge Status:"
$Label_Purge.AutoSize            = $true
$Label_Purge.width               = 25
$Label_Purge.height              = 10
$Label_Purge.location            = New-Object System.Drawing.Point(196,497)
$Label_Purge.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label_PStat                     = New-Object system.Windows.Forms.Label
$Label_PStat.text                = "Purge not Started"
$Label_PStat.AutoSize            = $true
$Label_PStat.width               = 25
$Label_PStat.height              = 10
$Label_PStat.location            = New-Object System.Drawing.Point(281,497)
$Label_PStat.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Label_PStat.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#d22222")

$ToolTip1.SetToolTip($Label1,'This is personal preference and is used for differentiating one search from another. Cannot be left blank.')
$ToolTip1.SetToolTip($Label3,'If the subject contains an apostrophe, add another apostrophe next to the existing one. Otherwise, the search will fail.')
$ToolTipStartSearch.SetToolTip($Button_Start_Search,'Let the hunt begin!')
$Window.controls.AddRange(@($Get_365_Creds,$Get_Search_Name,$Label1,$Get_Search_Date,$Get_Search_Subject,$Get_Search_Address,$Label2,$Label3,$Label4,$IT_Glue,$Button_Start_Search,$Button_Search_Status,$Connection_Status,$Label_Status,$Label5,$Get_Results,$TextBox1,$Button_Purge,$Button_Confirm_Purge,$Button_Status_Purge,$Label_Purge,$Label_PStat))

$IT_Glue.Add_Click({ Button_Click_Article })
$Get_365_Creds.Add_Click({ Button_Click_Creds })
$Button_Start_Search.Add_Click({ Button_Click_Start })
$Button_Search_Status.Add_Click({ Button_Click_Status })
$Get_Results.Add_Click({ button_Get_Results })
$Button_Confirm_Purge.Add_Click({ Button_Click_Confirm_Purge })
$Button_Purge.Add_Click({ Button_Click_Purge })
$Button_Status_Purge.Add_Click({ Btn_Click_Purge_Status })


function Btn_Click_Purge_Status {

$searchname = $Get_Search_Name.text

$purgename = $Searchname+"_Purge"

$purge = (Get-ComplianceSearchAction -Identity $purgename)

if ($purge.Status -ne "Completed"){

$Label_PStat.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#E07A0E")

$Label_PStat.text               = "Purging..."

}

Else{

$Label_PStat.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#3BDE17")

$Label_PStat.text               = "Done!"

}

}


function Button_Click_Purge {

New-ComplianceSearchAction -SearchName $Get_Search_Name.text -Purge -PurgeType HardDelete

$Label_PStat.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#E07A0E")

$Label_PStat.text               = "Purging..."

}


function Button_Click_Confirm_Purge {

$button_purge.enabled = $true

}

function button_Get_Results {

$textbox1.Text                   = ""

$results = Get-ComplianceSearch -Identity $Get_Search_Name.text | format-list -property successresults | out-string

$textbox1.text = $results

}

function Button_Click_Status {

$search = Get-ComplianceSearch $Get_Search_Name.text

if ($search.Status -ne "Completed"){

$Label_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#E07A0E")

$Label_Status.text               = "Searching..."

}

Else{

$Label_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#3BDE17")

$Label_Status.text               = "Done!"

}


}

function Button_Click_Start {

$Subject = $Get_Search_Subject.text

$From = $Get_Search_Address.text

$Received = $Get_Search_Date.text

New-ComplianceSearch -Name $Get_Search_Name.text -ExchangeLocation all -ContentMatchQuery "received:$received AND subject:$subject AND from:$from"

Start-ComplianceSearch -Identity $Get_Search_Name.text

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Search Started! Use the 'Check Search Status' button to keep an eye on the progress",0,"Started",0x0)

$Label_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#E07A0E")

$Label_Status.text               = "Searching..."

}



function Button_Click_Creds {


    try {

    $Connection_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#E07A0E")
    $Connection_Status.text          = "Connecting, Please wait..."

    Import-Module ExchangeOnlineManagement

    Connect-IPPSSession -Credential $UserCredential

    $Connection_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#3BDE17")
    $Connection_Status.text  = "Connected!"

    $Get_365_Creds.Enabled = $false

       
    }
    catch {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        $MessageBody = $($_.Exception.Message)
        $MessageTitle = "Error"

        $Result = [System.Windows.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
        $Result
        $Connection_Status.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#de1717")
        $Connection_Status.text = "Failed to connect"

    }
    

    
}



function Button_Click_Article { 
    
    Start-Process https://tinyurl.com/y3dvkb3n
    
}


[void]$Window.ShowDialog()
