'IE Version
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer"
strValueName = "svcVersion"
strValueRollBackName = "Version"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
Wscript.Echo "Installed IE Version: " & strValue
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueRollBackName,strValue
Wscript.Echo "IE Roll Back Version: " & strValue & vbCr & vbLf

'Check Ie version Update information. 
'Crude way to check through each version
'Checking Version 11
Wscript.Echo "Checking version 11 setting" 
Wscript.Echo "Blank values mean not set"
strKeyPathSetup11 = "SOFTWARE\Microsoft\Internet Explorer\Setup\11.0"
strIE11Offer = "DoNotOfferIE11AU"
strIE11UpdatesHidden = "IEUpdatesHidden"
strIE11Allow = "DoNotAllowIE11"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup11,strIE11Offer,dwValue
Wscript.Echo "IE 11 Do Not Offer is: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup11,strIE11Allow,dwValue
Wscript.Echo "IE 11 Do Not Allow: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup11,strIE11UpdatesHidden,dwValue
Wscript.Echo "IE 11 Updates Hidden: " & dwValue
Wscript.Echo ""

Wscript.Echo "Checking version 10 setting" 
Wscript.Echo "Blank values mean not set"
strKeyPathSetup10 = "SOFTWARE\Microsoft\Internet Explorer\Setup\10.0"
strIE10Offer = "DoNotOfferIE10AU"
strIE10UpdatesHidden = "IEUpdatesHidden"
strIE10Allow = "DoNotAllowIE10"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup10,strIE10Offer,dwValue
Wscript.Echo "IE 10 Do Not Offer is: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup10,strIE10Allow,dwValue
Wscript.Echo "IE 10 Do Not Allow: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup10,strIE10UpdatesHidden,dwValue
Wscript.Echo "IE 10 Updates Hidden: " & dwValue
Wscript.echo "" & vbCr & vbLf

Wscript.Echo "Checking version 9 setting" 
Wscript.Echo "Blank values mean not set"
strKeyPathSetup90 = "SOFTWARE\Microsoft\Internet Explorer\Setup\9.0"
strIE90Offer = "DoNotOfferIE90AU"
strIE90UpdatesHidden = "IEUpdatesHidden"
strIE90Allow = "DoNotAllowIE90"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup90,strIE90Offer,dwValue
Wscript.Echo "IE 9 Do Not Offer is: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup90,strIE90Allow,dwValue
Wscript.Echo "IE 9 Do Not Allow: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup90,strIE90UpdatesHidden,dwValue
Wscript.Echo "IE 9 Updates Hidden: " & dwValue
Wscript.echo  "" & vbCr & vbLf

Wscript.Echo "Checking version 8 setting" 
Wscript.Echo "Blank values mean not set"
strKeyPathSetup80 = "SOFTWARE\Microsoft\Internet Explorer\Setup\8.0"
strIE80Offer = "DoNotOfferIE80AU"
strIE80UpdatesHidden = "IEUpdatesHidden"
strIE80Allow = "DoNotAllowIE80"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup80,strIE80Offer,dwValue
Wscript.Echo "IE 8 Do Not Offer is: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup80,strIE80Allow,dwValue
Wscript.Echo "IE 8 Do Not Allow: " & dwValue
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPathSetup80,strIE80UpdatesHidden,dwValue
Wscript.Echo "IE 8 Updates Hidden: " & dwValue
Wscript.echo "" & vbCr & vbLf


Set objFSO = CreateObject("Scripting.FileSystemObject")
file = "C:\Program Files\Internet Explorer\iexplore.exe"
Wscript.Echo "IE file version is " & objFSO.GetFileVersion(file)




wscript.echo "Press enter to exit"
Input = wscript.stdin.Read(1)