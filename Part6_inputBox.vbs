Dim name

name=inputBox("What is your age","Information")
if isNumeric(name) And name="15" then
	msgbox "Correct"
Elseif name<>"15" then
	msgbox "Worng Age"
Elseif Not isNumeric(name) then
	msgbox "Please tyie in a number value."
End If 


name=inputBox("What is your name?","Information")

if name="Sato" Or name="Ben" then
	msgbox "Hello"
	Else
	msgbox "Wrong"
End if


If name="Sato" Or name="Ben" then
	msgbox "Please Type your name"
ElseIf name="Sato" then 
	msgbox "Hello"
ElseIf name<>"Sato" then
	msgbox "Intuder"

End if
