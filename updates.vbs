Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Set objSettings = objAutoUpdate.Settings

Select Case objSettings.NotificationLevel
    Case 0
        Wscript.Echo "Notification level: Automatic Updates is not configured by the user " & _
            "or by a Group Policy administrator."
    Case 1
        Wscript.Echo "Notification level: Automatic Updates is disabled."
    Case 2
        Wscript.Echo "Notification level: Automatic Updates prompts users to approve updates " & _
            "before downloading or installing."
    Case 3
        Wscript.Echo "Notification level: Automatic Updates automatically downloads " & _
             "updates, but prompts users to approve them before installation."
    Case 4
        Wscript.Echo "Notification level: Automatic Updates automatically installs " & _
            "updates per the schedule specified by the user."
    Case Else
        Wscript.Echo "Notification level could not be determined."
End Select

wscript.echo "Press enter to exit"
Input = wscript.stdin.Read(1)

