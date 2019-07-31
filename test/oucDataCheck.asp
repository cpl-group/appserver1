<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim startdate, enddate, groupname, bldg, datasource
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")
bldg = request("bldg")
if trim(bldg)="" then bldg="MM"
dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)
if not(isdate(startdate)) then
  startdate = dateadd("d",-1,date())
end if
if not(isdate(enddate)) then enddate = startdate
datasource = "oucdata1"
dim portfolio
%>
<html>
<head>
<title>Meter Data Source Checking</title>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr>
  <td bgcolor="#3399CC">
    <span class="standardheader">Meter Datasource Data Check</span></td>
  </tr>
<tr>
  <td bgcolor="#eeeeee">
<form action="oucDataCheck.asp" method="get">
<select name="bldg" onchange="document.forms[0].submit()">
<%
  rst1.open "SELECT * FROM "&makeIPUnion("buildings","")&" b, portfolio p WHERE p.id=b.portfolioid ORDER BY name, strt", application("cnnstr_supermod")
  do until rst1.eof
    if portfolio<>rst1("name") then
      portfolio = rst1("name")
      %><optgroup label="<%=portfolio%>"><%
    end if
    %><option value="<%=rst1("bldgnum")%>" <%if bldg=rst1("bldgnum") then response.write "SELECTED"%>><%=rst1("strt")%>, (<%=rst1("bldgnum")%>)</option><%
    rst1.movenext
  loop
  rst1.close
%>
</select>&nbsp;Check&nbsp;Dates&nbsp;<input type="text" name="startdate" value="<%=startdate%>" size="9" maxlength="10">&nbsp;through&nbsp;<input type="text" name="enddate" value="<%=enddate%>" size="9" maxlength="10">&nbsp;<input type="submit" value="reload"></td></tr>
<tr><td>
<%
rst1.open "SELECT distinct datasource FROM meters WHERE bldgnum='"&bldg&"'", cnn1
if not rst1.eof then datasource = rst1("datasource")
rst1.close
rst1.open "SELECT name FROM sysobjects WHERE name='"&datasource&"'", cnn1
if rst1.eof then datasource = ""
rst1.close
if trim(datasource)<>"" then
  do until datediff("d",startdate,enddate)<0
    rst1.open "SELECT count(m.meterid) as c FROM meters m LEFT JOIN (SELECT meterid, count(date) as mcount FROM "&datasource&" WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY meterid) ds ON ds.meterid=m.meterid WHERE m.bldgnum='"&bldg&"' and online=1 and isnull(mcount,0)<96", cnn1
    %><b><%=startdate%> (<%=rst1("c")%>)</b><UL><%
    rst1.close
    rst1.open "SELECT m.meterid, m.meternum, isnull(mcount,0) as entries FROM meters m LEFT JOIN (SELECT meterid, count(date) as mcount FROM "&datasource&" WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY meterid) ds ON ds.meterid=m.meterid WHERE m.bldgnum='"&bldg&"' and online=1 and isnull(mcount,0)<96"
    do until rst1.eof
      %><li><%=rst1("meterid")%>,&nbsp;<%=rst1("meternum")%>&nbsp;&nbsp;&nbsp;<%=rst1("entries")%>/96 </li><%
      rst1.movenext
    loop
    response.write "</ul>"
    rst1.close
    startdate = dateadd("d",1,startdate)
  loop
else
  response.write "No datasource found for this building"
end if
%>
</form>
</td></tr>
</table>
</body>
</html>
