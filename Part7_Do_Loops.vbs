Dim a, pass

Do
pass=inputbox("Password")
if pass="aaa" then
Exit Do
Elseif pass="" then
	msgbox "Don't leave filed blank."
Elseif pass<>"aaa" then
	msgbox "Incorrect",vbcritical
end If 
loop
msgbox "Correct!"


a=0
Do while a<6
a=a+1
msgbox a
loop

msgbox "Loop Ended"