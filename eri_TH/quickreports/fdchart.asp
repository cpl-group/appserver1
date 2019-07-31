<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim pid, rid, building, billingid, meterid, startdate, enddate, utilityid, groupname, interval

pid = request("pid")'13
rid = 0
building = request("building")'"113DORM"
billingid = 0
meterid = 0
startdate = request("startdate")'"9/1/2003"
enddate = request("enddate")'"9/30/2003"
utilityid = 2
groupname = 0
interval = 2
'response.write "0, ""bottom"", "&pid&", "&rid&", "&building&", "&billingid&", "&meterid&", "&groupname&", """&startdate&""", """&enddate&""""&"<br>"
'response.end
dim ip
if trim(building)<>"" and trim(building)<>"0" and instr(building,"|")=0 then
  ip = getBuildingIP(building)
else
  ip = getPidIP(pid)
end if

dim lmp, myHeight, myWidth


 Set lmp = CreateObject("lmpchartFordam.lmpcontrol")
lmp.setLocalIP 0, ip
lmp.utility = utilityid
lmp.interval = interval
lmp.loadaggs = false
lmp.projectionSeries = -1
lmp.SmallSize = false
lmp.setCost 0, 0
lmp.setCost 1, 0
'response.write "0, ""bottom"", "&pid&", "&rid&", "&building&", "&billingid&", "&meterid&", "&groupname&", """&startdate&""", """&enddate&""""&getLocalConnect(building)&"<br>"
'response.end
lmp.setSeries 0, "top", pid, rid, building, billingid, meterid, groupname, startdate, enddate

response.BinaryWrite(lmp.returnimage())
%>