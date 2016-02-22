'Part25 get file system object

Dim fso, file
Set fso = createobject("Scripting.FileSystemObject")
Set file = fso.getfile("c:\PS\test1.txt")

file = file.path &vbLf
file = file & "create data" & file.Datacreated &vbLf

msgBox file