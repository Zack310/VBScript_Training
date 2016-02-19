'Part17
Dim obj, path, desk
Set obj = createObject("Wscript.shell")

'msgBox obj.specialfolders("Desktop")
'msgBox obj.specialfolders("Desktop")&"\VBscript"
'obj.run obj.specialfolders("Desktop")&"\VBScript"

'path = obj.specialfolders("Desktop")&"\VBScript"
'obj.run path

desk = obj.specialfolders("Desktop")
obj.run desk&"\VBScript"

'List of Special Folders:
'AllUsersDesktop
'AllUsersStartMenu
'AllUsersPrograms
'AllUsersStartup
'Desktop
'Favorites
'Fonts
'MyDocuments
'NetHood
'PrintHood
'Programs
'Recent
'SendTo
'StartMenu
'Startup
'Templates
