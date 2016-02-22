'Part27 get file system object

Dim fso, folder, data,
Set fso = createobject("Scripting.Filesystemobject")
Set folder = fso.Getfolder("C:\users\satoei1\desktop")

Data = "Attributes : " & folder.attributes &vbLf
Data = Data & "DataCreated : " & folder.DateCreated &vbLf
Data = Data & "DataLastModified : " & folder.DatelastModified &vbLf
Data = Data & "Drive : " & folder.drive &vbLf
Data = Data & "IsRootFolder : " & folder.IsRootFolder &vbLf
Data = Data & "Name : " & folder.Name &vbLf
Data = Data & "ParentFolder : " & folder.ParentFolder &vbLf
Data = Data & "Path : " & folder.Path &vbLf
Data = Data & "ShortName : " & folder.ShortName &vbLf
Data = Data & "ShortPath : " & folder.ShortPath &vbLf
Data = Data & "Size : " & folder.Size &vbLf
Data = Data & "Type : " & folder.Type &vbLf

msgBox Data
