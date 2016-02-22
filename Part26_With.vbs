'Part26 With

Dim obj
Set obj = createObject("Wscript.shell")
With obj
.run "notepad.exe"
.run "mspaint.exe"
End With

Wscript.echo "Completed all application"









