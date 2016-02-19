set x=createobject("wscript.shell")
x.run "notepad.exe"
wscript.sleep 2000
x.sendkeys "Hello There"
x.sendkeys "{Enter}"
x.sendkeys "How are you doing"

x.sendkeys "%"
X.sendkeys "%fs"
wscript.sleep 500
x.sendkeys "Test.vbs"
wscript.sleep 300
x.sendkeys "{enter}"
