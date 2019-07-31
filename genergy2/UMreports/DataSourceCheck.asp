<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim startdate, enddate, groupname, bldg, datasource,pid, sqlstr
pid = request("pid") 
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")
bldg = request("bldg")


if trim(bldg)="" then bldg=""
dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
if bldg<>"" then cnn1.open getLocalConnect(bldg)
'if pid="" then pid=getpid(bldg)
if not(isdate(startdate)) then
  startdate = dateadd("d",-30,date())
end if
if not(isdate(enddate)) then enddate = date()
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
<form action="DataSourceCheck.asp" method="get">
<select name="bldg" onchange="document.forms[0].submit()">
<%
  if trim(pid) <> "" then 
  	sqlstr = "SELECT * FROM buildings b, portfolio p WHERE p.id=b.portfolioid and p.id = "&trim(pid)& " ORDER BY name, strt"
  else
	sqlstr = "SELECT * FROM buildings b, portfolio p WHERE p.id=b.portfolioid ORDER BY name, strt"
  end if 
  rst1.open sqlstr, getConnect(0,0,"dbCore")
  do until rst1.eof
    if portfolio<>rst1("name") then
      portfolio = rst1("name")
      %><optgroup label="<%=portfolio%>"><%
    end if
    %><option <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=rst1("bldgnum")%>" <%if bldg=rst1("bldgnum") then response.write "SELECTED"%>><%=rst1("strt")%>, (<%=rst1("bldgnum")%>)</option><%
    rst1.movenext
  loop
  rst1.close
%>
</select>&nbsp;Check&nbsp;Dates&nbsp;<input type="text" name="startdate" value="<%=startdate%>" size="9" maxlength="10">&nbsp;through&nbsp;<input type="text" name="enddate" value="<%=enddate%>" size="9" maxlength="10">&nbsp;<input type="submit" value="reload"></td></tr>
<tr><td>
<%
if bldg<>"" then
	sqlstr = "SELECT distinct datasource, isnull(date_offline,'1/1/1900') as dateoff, offline FROM dbBilling.dbo.meters m INNER JOIN dbo.buildings b ON b.bldgnum=m.bldgnum INNER JOIN dbIntervalData.dbo.sysobjects s ON s.name=m.datasource WHERE m.bldgnum='"&bldg&"'"
	rst1.open sqlstr, getLocalConnect(bldg)
	dim offline, dateoff
	if not rst1.eof then
		offline = rst1("offline")
		dateoff = rst1("dateoff")
	end if
	do until rst1.eof
	  datasource = datasource &"|"& "["&rst1("datasource")&"]"
	  rst1.movenext
	loop
	rst1.close
	if trim(datasource)<>"" then
	  datasource = split(mid(datasource,2),"|")
	  do until datediff("d",startdate,enddate)<0
	    sqlstr = "SELECT count(m.meterid) as c FROM dbBilling.dbo.meters m LEFT JOIN (SELECT meterid, count(date) as mcount FROM "&join(datasource," WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY meterid UNION SELECT meterid, count(date) as mcount FROM ")&" WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY meterid) ds ON ds.meterid=m.meterid WHERE m.bldgnum='"&bldg&"' and online=1 and isnull(mcount,0)<96"
		rst1.open sqlstr, getLocalConnectCom(bldg)
	 '  response.write getLocalConnectCom(bldg)
	 	
		'response.write getpid(bldg)
	   'response.end
	    %><b><%=startdate%> (<%=rst1("c")%>)
		<%if datediff("d",dateoff,cdate(startdate))>=0 and lcase(offline)="true" then%><span class="grayout">Building offline as of <%=dateoff%><%'dateoff="1/1/2200"%></span><%end if%>
		 </b><UL><%
	    rst1.close
		sqlstr = "SELECT b.portfolioid, m.meterid, m.meternum, l.billingid, lup.leaseutilityid, isnull(mcount,0) as entries, isnull(manualentry,0) as me FROM dbBilling.dbo.meters m INNER JOIN dbBilling.dbo.tblleasesutilityprices lup on lup.leaseutilityid=m.leaseutilityid INNER JOIN dbBilling.dbo.tblleases l on l.billingid=lup.billingid INNER JOIN dbBilling.dbo.buildings b on b.bldgnum=l.bldgnum LEFT JOIN (SELECT meterid, count(date) as mcount FROM  dbo."&join(datasource," WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY meterid UNION SELECT meterid, count(date) as mcount FROM ")&" WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY meterid) ds ON ds.meterid=m.meterid WHERE m.bldgnum='"&bldg&"' and online=1 and isnull(mcount,0)<96"
	    rst1.open sqlstr, getConnect(pid,bldg,"intervaldata")
		do until rst1.eof
	      %><li><a <%if cint(rst1("me")) then%>style="color:red"<%end if%> href="/genergy2/setup/contentfrm.asp?action=meteredit&pid=<%=rst1("portfolioid")%>&bldg=<%=bldg%>&tid=<%=rst1("billingid")%>&lid=<%=rst1("leaseutilityid")%>&meterid=<%=rst1("meterid")%>" target="new"><%=rst1("meterid")%></a>,&nbsp;<%=rst1("meternum")%>&nbsp;&nbsp;&nbsp;<%=rst1("entries")%>/96</li><%
	    
		  rst1.movenext
	    loop
	    response.write "</ul>"
	    rst1.close
	    startdate = dateadd("d",1,startdate)
	  loop
	else
	  response.write "No datasource found for this building"
	end if
end if
%>
</form>
<span style="color:red">&nbsp;&nbsp;&nbsp;&nbsp;*Any meters that show up in red are manual read meters.</span>
</td></tr>
</table>
</body>
</html>
