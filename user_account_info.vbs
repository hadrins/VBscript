'On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("W7")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccountType: " & objItem.AccountType &vbCrLf &_
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
	  WScript.Echo
   Next
Next



wscript.echo "Press enter to exit"
Input = wscript.stdin.Read(1)