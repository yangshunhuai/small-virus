dim ws,speaker
Set ws=CreateObject("wscript.shell")
Set speaker=CreateObject("sapi.spvoice")
On Error Resume next
wscript.sleep 7500
ws.run"补丁.vbs"
do
speaker.speak"我是补丁！来打我呀！"
msgbox("我是补丁！来打我呀！")
loop