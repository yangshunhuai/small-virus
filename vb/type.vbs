dim a
Set ws=CreateObject("wscript.shell")
a=inputbox("content")
ws.run"notepad.exe"
wscript.sleep 1000
do
ws.sendkeys a
ws.sendkeys"{ENTER}"
wscript.sleep 1000
loop