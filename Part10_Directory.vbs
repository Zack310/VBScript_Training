Dim cd1, cd2, cd3, cd4, obj

cd1 = CreateObject("WScript.Shell").CurrentDirectory
cd2 = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
cd3 = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Wscript.ScriptFullName)
cd4 = Replace(WScript.ScriptFullName,WScript.ScriptName, "")

wscript.echo cd1
wscript.echo cd2
wscript.echo cd3
wscript.echo cd4

Set obj=CreateObject("Wscript.shell")
msgbox obj.CurrentDirectory&"\a.txt"
obj.run"C:\Users\%username%\Desktop\Vbscript\a.txt"
