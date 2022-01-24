On Error Resume Next 
 
Dim wsh,s,xTimes,tle
 
set wsh=createobject("wscript.shell") 
xTimes = inputbox("请输入重发消息的次数")
 
'clipboardData.SetData s
 
for i=1 to Cint(xtimes)
wscript.sleep 30
wsh.AppActivate(cStr(tle)) 
wsh.sendKeys "^v"
wsh.sendKeys "%s" 
next 
 
wscript.quit