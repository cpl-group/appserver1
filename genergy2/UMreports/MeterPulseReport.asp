<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim meterid, startdate, enddate, i, smonth, sday, syear, emonth, eday, eyear, genergy1
genergy1 = request("genergy1")
meterid = Request("meterid")
startdate = Request("startdate")
enddate = Request("enddate")
smonth = Request("smonth")
sday = Request("sday")
syear = Request("syear")
emonth = Request("emonth")
eday = Request("eday")
eyear = Request("eyear")
bldg = Request("bldg")
startdate = smonth&"/"&sday&"/"&syear
enddate = emonth&"/"&eday&"/"&eyear
if not isdate(startdate) then startdate = Request("startdate")
if not isdate(enddate) then enddate = Request("enddate")
if trim(startdate)="" then startdate = dateadd("d",-1,date())
if trim(enddate)="" then enddate = dateadd("d",1,startdate)
if datediff("d", startdate, enddate)<0 then enddate = dateadd("d",7,startdate)
Dim cnn1, rst1, sql
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

'Get meter info
dim meternum, datasource, utilitydisplay, billingname, buildingname, measure, bldg, sqlstr
sqlstr = "SELECT m.bldgnum, m.meternum, m.datasource, u.utilitydisplay, u.measure, l.billingname, b.strt FROM meters m, tblLeasesutilityprices lup, tblutility u, tblleases l, buildings b WHERE m.leaseutilityid=lup.leaseutilityid and lup.billingid=l.billingid and lup.utility=u.utilityid and m.bldgnum=b.bldgnum and m.meterid="&meterid


rst1.open sqlstr, getConnect(0,bldg,"billing")

if not rst1.eof then
  meternum = rst1("meternum")
  datasource = rst1("datasource")
  utilitydisplay = rst1("utilitydisplay")
  billingname = rst1("billingname")
  buildingname = rst1("strt")
  measure = rst1("measure")
  bldg = rst1("bldgnum")
end if
rst1.close

cnn1.open getLocalConnect(bldg)
'date ranges
dim mindate, maxdate, datamax, dataavg
if trim(datasource)="" then
  response.write "No datasource."
  response.end
end if
sqlstr="SELECT isnull(max(date),0) as [max], isnull(min(date),0) as [min], isnull(max("&measure&"),0) as datamax, isnull(avg("&measure&"),0) as dataavg FROM ["&datasource&"] WHERE meterid="&meterid&" and date between '"&startdate&"' and '"&enddate&"'"

rst1.open sqlstr, getConnect(0,bldg,"IntervalData")
if not rst1.eof then
  mindate = rst1("min")
  maxdate = rst1("max")
  datamax = rst1("datamax")
  dataavg = rst1("dataavg")
end if
rst1.close
%>
<html>
<head>
<title>Meter Pulse Data</title>
</head>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<body>
<form name="form1" action="MeterPulseReport.asp">
<input type="hidden" name="meterid" value="<%=meterid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="genergy1" value="<%=genergy1%>">
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#6699CC" class="standardheader">Meter Pulse Data</td></tr>
<tr bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <table border=0 cellpadding="0" cellspacing="0">
    <tr><td align="right"><b>Meter Number:</b></td>
        <td>&nbsp;<%=meternum%></td></tr>
    <tr><td align="right"><b>Building:</b></td>
        <td>&nbsp;<%=buildingname%></td></tr>
    <tr><td align="right"><b>Tenant:</b></td>
        <td>&nbsp;<%=billingname%></td></tr>
    <tr><td align="right"><b>Utility:</b></td>
        <td>&nbsp;<%=utilitydisplay%></td></tr>
    <tr><td align="right"><b>Data Range:</b></td>
        <td>&nbsp;<%if mindate<0 and maxdate<0 then%><%=mindate%> to <%=maxdate%><%else%>none<%end if%></td></tr>
    <tr><td align="right"><b>Data Range Max:</b></td>
        <td>&nbsp;<%=datamax%></td></tr>
    <tr><td align="right"><b>Data Range Avrage:</b></td>
        <td>&nbsp;<%=dataavg%></td></tr>
<!--     <tr><td align="right"><b>Dates Selected:</b></td>
        <td>&nbsp;<%=startDate%> to <%=enddate%></td></tr> -->
  </table>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr><td>
    <select name="smonth">
      <%for i = 1 to 12%>
        <option value="<%=i%>"<%if i=month(startdate) then response.write " selected"%>><%=monthname(i)%></option>
      <%next%>
    </select>
    <select name="sday">
      <%for i = 1 to 31%>
        <option value="<%=i%>"<%if i=day(startdate) then response.write " selected"%>><%=i%></option>
      <%next%>
    </select>
    <select name="syear">
      <%for i = year(enddate)-7 to year(enddate)%>
        <option value="<%=i%>"<%if i=year(startdate) then response.write " selected"%>><%=i%></option>
      <%next%>
    </select>
to
    <select name="emonth">
      <%for i = 1 to 12%>
        <option value="<%=i%>"<%if i=month(enddate) then response.write " selected"%>><%=monthname(i)%></option>
      <%next%>
    </select>
    <select name="eday">
      <%for i = 1 to 31%>
        <option value="<%=i%>"<%if i=day(enddate) then response.write " selected"%>><%=i%></option>
      <%next%>
    </select>
    <select name="eyear">
      <%for i = year(enddate)-7 to year(enddate)%>
        <option value="<%=i%>"<%if i=year(enddate) then response.write " selected"%>><%=i%></option>
      <%next%>
    </select>
    <input type="submit" name="action" value="Set Dates">
    </td>
  </tr>
  </table>
  </td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="1" bgcolor="#eeeeee">
<%
dim currentdate, fieldcount, intervalcount
currentdate = 0

sqlstr="SELECT * FROM ["&trim(datasource)&"] WHERE meterid="&meterid&" and date>='"&startdate&"' and date<'"&dateadd("d",1,enddate)&"' ORDER BY date"

rst1.open sqlstr, getConnect(0,bldg,"Intervaldata")
response.write "<tr>"
if not rst1.eof then
  fieldcount = rst1.Fields.Count-1
  for i=0 to fieldcount%>
    <td bgcolor="#cccccc"><%=rst1.Fields.Item(i).name%></td>
  <%
  next
end if
response.write "</tr>"
do until rst1.eof
  if datediff("d",rst1("date"),currentdate)<>0 then 
    if intervalcount<>"" then 
    %><tr><td colspan="<%=fieldcount + 1%>" bgcolor="#cccccc">Total intervals for <%=currentdate%>= <%=intervalcount%></td></tr><%
	end if 
  	intervalcount = 1
    currentdate = rst1("date")
    %><tr><td colspan="<%=fieldcount%>" bgcolor="#eeeeee"><%=weekdayname(weekday(currentdate))%>&nbsp;<%=monthname(month(currentdate))%>&nbsp;<%=day(currentdate)%></td></tr><%
  else 
  	 intervalcount = intervalcount + 1
  end if
  response.write "<tr>"
  for i=0 to fieldcount%>
    <td bgcolor="white"><%=rst1(i)%></td>
  <%
  next
  response.write "</tr>"
 rst1.movenext
loop
    if intervalcount<>"" then 
    %><tr><td colspan="<%=fieldcount + 1%>" bgcolor="#cccccc">Total intervals for <%=currentdate%>= <%=intervalcount%></td></tr><%
	end if 

rst1.close
%>
</table>
</form>
</body>
</html>