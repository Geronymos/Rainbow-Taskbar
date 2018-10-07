Set WshShell = CreateObject("WScript.Shell")

wshShell.Run "ping-display.hta"


do


  wscript.sleep 5000


  wshShell.AppActivate "pingdisplay"


loop