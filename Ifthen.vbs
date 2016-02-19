'Option Exlicit
Dim a
a=MsgBox("Guess a Button",vbAbortRetryIgnore)

If a=vbRetry Or a-vbAbort Then
	msgbox"correct"
Elseif a=vbIgnore Then
	msgbox "Wrong"
End If