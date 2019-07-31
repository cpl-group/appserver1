<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->


<%
loadNewXML "daniell"
setBuilding "mm","10.0.7.110"
setBuilding "33","10.0.7.21"
setKeyValue "login","login"
setKeyValue "name","name"
setKeyValue "roleid",23242
setKeyValue "um","um"
setKeyValue "eri","eri"
setKeyValue "opslog","opslog"
setKeyValue "ts","ts"
setKeyValue "corp","corp"
setKeyValue "it","it"
setKeyValue "admin","admin"
'setView "grgergrerg", "10.0.55.55"
setGroup("genergy")
response.write hasGroup("genergy")
%>

<%=getXMLUserName()%><br>_______________<br>
<%
response.write getIPfor("j")
dim k, i
k = getAllBuildings()
for i = 0 to ubound(k)-1
  response.write k(i,0)&"|"&k(i,1)&"<br>"
next
printxml()
%>