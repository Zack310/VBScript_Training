'Option Exlicit

Dim birth
birth=Msgbox("Is it your birthday?" ,vbYesNo+vbQuestion, "Tell me")
if birth=vbYes then 
	msgbox "Happy Brithday!!!!" ,vbinformation
Elseif birth=vbNo then 
	msgbox "Oops, My god." ,vbCritical
End if

a=Msgbox("Pandora Radio" _
	,vbAbortRetryignore+vbExclamation+vbDefaultButton3+vbSystemModal, "Gadget:")
if a=vbAbort then msgbox "QUit" ,vbCritical


