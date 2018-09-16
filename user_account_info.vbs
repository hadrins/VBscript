'this should be run from the CLI with cscript
'otherwise it will error on the last statement. 

'On Error Resume Next



Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("PC")
For Each strComputer In arrComputers
   Wscript.Echo " " &vbCrLf &_
    "==========================================" &vbCrLf &_
    "Computer: " & strComputer &vbCrLf &_
    "==========================================" &vbCrLf &_
	Wscript.Echo

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      Wscript.Echo "AccountType: " & objItem.AccountType &vbCrLf &_
      "Caption: " & objItem.Caption &vbCrLf &_
      "Description: " & objItem.Description &vbCrLf &_
      "Disabled: " & objItem.Disabled &vbCrLf &_
      "Domain: " & objItem.Domain &vbCrLf &_
      "FullName: " & objItem.FullName &vbCrLf &_
      "Lockout: " & objItem.Lockout &vbCrLf &_
      "Name: " & objItem.Name &vbCrLf &_
      "PasswordChangeable: " & objItem.PasswordChangeable &vbCrLf &_
      "PasswordExpires: " & objItem.PasswordExpires &vbCrLf &_
      "PasswordRequired: " & objItem.PasswordRequired &vbCrLf &_
      "SID: " & objItem.SID &vbCrLf &_
      "SIDType: " & objItem.SIDType &vbCrLf &_ 
      "Status: " & objItem.Status &vbCrLf &_
	  Wscript.Echo
   Next
Next



Wscript.Echo "Press enter to exit"
Input = wscript.stdin.Read(1)