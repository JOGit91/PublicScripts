# created by Jake Ouellette
# This script automates the process of installing mscomct2.ocx. Checks bitness of Office as well as bitness of Windows OS to determine eligibility as well as the correct install location.
# 64 bit Office is not eligible to run mscomct2.ocx. Those systems are logged and then skipped.

$bitness = get-itemproperty HKLM:\Software\Microsoft\Office\14.0\Outlook -name Bitness
if($bitness -eq $null) {
$bitness = get-itemproperty HKLM:\Software\Microsoft\Office\15.0\Outlook -name Bitness}
if($bitness -eq $null) {
$bitness = get-itemproperty HKLM:\Software\Microsoft\Office\16.0\Outlook -name Bitness}
if($bitness -eq $null) {
$bitness = get-itemproperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Outlook -name Bitness}
if($bitness -eq $null) {
$bitness = get-itemproperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Outlook -name Bitness}
if($bitness -eq $null) {
$bitness = get-itemproperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook -name Bitness}

Test-Path $bitness

If($bitness -eq "x86")

{

    if ((gwmi win32_operatingsystem | select osarchitecture).osarchitecture -eq "64-bit")
        {
          #64 bit logic here
         if ((Test-Path C:\Windows\SysWoW64\mscomct2.ocx -PathType Leaf) -eq "true"){
         C:\Windows\SysWoW64\regsvr32.exe mscomct2.ocx
            }

    else{

    #copy file from location and then run C:\Windows\SysWoW64\regsvr32.exe mscomct2.ocx
    Copy-Item "\\NetworkShare\mscomct2.ocx" -Destination "C:\Windows\SysWoW64"

    C:\Windows\SysWoW64\regsvr32.exe mscomct2.ocx

    }

}
else
{
    #32 bit logic here
    if ((Test-Path C:\Windows\System32\mscomct2.ocx -PathType Leaf) -eq "true"){
    C:\Windows\System32\regsvr32.exe mscomct2.ocx
    }

    else{

    #copy file from location and then run C:\Windows\System32\regsvr32.exe mscomct2.ocx
    Copy-Item "\\NetworkShare\mscomct2.ocx" -Destination "C:\Windows\System32"

    C:\Windows\System32\regsvr32.exe mscomct2.ocx

    }

}

}

Else
{

$env:computername+ " skipped due to 64 bit Office" | out-file \\NetworkShare\Skipped.txt -append

}