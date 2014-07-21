
Set objWMIservice = GetObject("winmgmts:\\.\root\cimv2")
set colitems = objWMIservice.ExecQuery("Select * from Win32_BIOS",,48)
For each objitem in colitems
Wscript.echo "Dell Service Tag: " & objitem.serialnumber
Next 

