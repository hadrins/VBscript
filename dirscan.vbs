' FILE: dirscan.vbs..
' DESC: Given a root directory, this script will scan for..
' subdirectories and will output the size of each..
' subdirectory...
' AUTH: TY WSH in 21 days..
'
On Error Resume Next
Dim RootDir, FileSystem, RootFolder, SubFolders, Folder
Dim FolderSize, Tmp
' 1. Get the root directory parameter
' if no parameter is specified, quit and shw usage of the script
if Wscript.Arguments.count <> 1 then 
	' Show the usage 
	Wscript.Echo "Usage:dirscan.vbs [root directory]"
	Wscript.Echo ""
	Wscript.Echo "Given a root directory,dirscan will scan"
	Wscript.Echo "all directories and output the size of"
	Wscript.Echo "each subdirecotry."
' And quit
Wscript.quit 0
Else
	RootDir = Wscript.Arguments(0)
	Wscript.Echo "Root Directory is: " & RootDir
	Wscript.echo " "
end If 
Set FileSystem = CreateObject("Scripting.FileSystemObject")
Set RootFolder = FileSystem.GetFolder(RootDir)
If Err.Number <> 0 Then
	Wscript.Echo "(" & Err.Number & ") " & Err.Description
	Wscript.Echo ""
	Wscript.Echo "The path you entered is invalid, please " & _
	"select a different path."
	Wscript.Quit Err.Number
End If
Set SubFolders = RootFolder.SubFolders
For Each Folder In SubFolders
	FolderSize = Folder.Size
	Tmp = FormatNumber (FolderSize, 0, 0, 0, -1)
	Tmp = Right (Space(20) & Tmp, 20)
	Wscript.Echo Tmp & " " & Folder.Path
Next

wscript.echo "Press enter to exit"
Input = wscript.stdin.Read(1)