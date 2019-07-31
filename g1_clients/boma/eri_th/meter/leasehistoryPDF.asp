<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, leaseid, title, units, utility, usagedivisor
bldg = request.querystring("B")
leaseid = request.querystring("leaseid")
units = "Usage"
utility = request("utility")
if instr(leaseid,"u")>0 then 
  utility = mid(leaseid,2)
  leaseid = 0
end if

dim cnn1, cmd1, rst1, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)
'response.write(getLocalConnect(bldg))
'response.end
dim DBmainIP
DBmainIP = "["&application("superIP")&"].mainmodule.dbo."

if leaseid<>0 then
	title = "Tenant"
  rst1.open "SELECT measure, lup.utility FROM tblleasesutilityprices lup, "&DBmainIP&"tblutility u WHERE lup.utility=u.utilityid and leaseutilityid="&leaseid, cnn1
  if not rst1.eof then 
    units=rst1("measure")
    utility = rst1("utility")
  end if
  rst1.close
  'if instr(request.servervariables("SCRIPT_NAME"),"/genergy2/")<>0 then
	  sql="SELECT top 24 isnull(TotalAmt,0) as TotalAmt, b.postdate, b.BillYear, b.BillPeriod, sum(isnull(m.Used,0)) as KWH, sum(Demand_P) as demand_P, m.ypid, m.leaseutilityid FROM tblmetersbyperiod m INNER JOIN tblbillbyperiod b on b.ypid=m.ypid and b.LeaseUtilityId=m.LeaseUtilityId WHERE m.leaseutilityid="&leaseid&" and b.bldgnum='" & bldg & "' and b.posted = 1 GROUP BY b.BillYear, b.BillPeriod, m.ypid, m.leaseutilityid, TotalAmt, b.postdate order by b.billyear desc ,b.billperiod desc"
  'else
  '	sql="SELECT top 12 isnull(TotalAmt,0) as TotalAmt, b.BillYear, b.BillPeriod, sum(isnull(m.KWHUsed,0)) as KWH, sum(Demand_P) as demand_P, m.ypid, m.leaseutilityid FROM tblmetersbyperiod m INNER JOIN tblbillbyperiod b on b.ypid=m.ypid and b.LeaseUtilityId=m.LeaseUtilityId WHERE m.leaseutilityid="&leaseid&" and b.bldgnum='" & bldg & "' and b.posted = 1 GROUP BY b.BillYear, b.BillPeriod, m.ypid, m.leaseutilityid, TotalAmt order by m.ypid desc"
  'end if
else
	title = "Building"
  rst1.open "SELECT measure FROM tblutility u WHERE utilityid="&utility, application("cnnstr_supermod")
  if not rst1.eof then units=rst1("measure")
  rst1.close
	sql = "select top 24 0 as leaseutilityid, sum(kwh) as kwh,sum(demand_p) as demand_p,sum(totalamt) as totalamt,ypid,billyear,billperiod, max(postdate) as postdate from (select sum(used) as kwh,sum(demand_p) as demand_p ,b.totalamt as totalamt ,m.ypid,billyear,billperiod, postdate from tblmetersbyperiod m, (SELECT DISTINCT ypid, leaseutilityid,SUM(totalamt) AS totalamt,utility, postdate FROM tblbillbyperiod WHERE  bldgnum ='"&bldg&"' and posted=1 and utility="&utility&" GROUP BY ypid,leaseutilityid,utility, postdate) B where m.ypid=b.ypid and b.leaseutilityid=m.leaseutilityid group by b.totalamt,m.ypid,billyear,billperiod, postdate)Z group by ypid,billyear,billperiod order by billyear desc,billperiod desc"

end if
if not isnumeric(utility) then utility=0
dim charge1, charge2, charge1FLD, charge2FLD
select case utility
case 3
  usagedivisor = 100
  charge1 = "Water Charge"
  charge2 = "Sewer Charge"
  charge1FLD = "energydetail"
  charge2FLD = "demanddetail"
case else
  usagedivisor = 1
  charge1 = "Energy Charge"
  charge2 = "Demand Charge"
  charge1FLD = "energy"
  charge2FLD = "demand"
end select
'response.write sql
'response.end
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<html>

<head>
<title>Meter Details - Usage History</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">
</head>

<script> 
function loadentry(luid, ypid){

	var temp = 'meterbillsummaryPDF.asp?utility=<%=utility%>&l=' +luid+'&Y='+ypid+'&building=<%=server.urlencode(request("b"))%>'
	parent.document.frames.panel_2.location = temp
}
function settonull(){
	var temp = '/null.htm'
	parent.document.frames.panel_2.location = temp
}
</script>
<body bgcolor="#FFFFFF"onLoad="settonull()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#6699CC" width="23%" height="2"><span class="standardheader"><%=title%> 
      Details - Usage History</span></td>
    <td bgcolor="#6699CC" width="23%"><div align="right"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><strong><font size="1">(<i>Click 
        Any Bill Period Row To View Meter Details Below</i>)</font></strong></font></div></td>
  </tr>
  <tr> 
    <td width="46%" colspan="2">&nbsp;</td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
	<tr bgcolor="#6699CC" > 
		<%
		dim columnwidth
		if utility = 3 then 
			columnwidth = "25%"
		else 
			columnwidth = "20%"
		end if
		%>
		<td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial">Bill Period</font></b></td>
		<td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial"><%=units%></font></b></td>
    <%if utility<>3 then%><td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial">Demand</font></b></td><%end if%>
		<td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial">Total Amount</font></b></td>
		<td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial">Post Date</font></b></td>
	</tr>
</table>
<table width = "100%" cellpadding = 0 cellspacing = 0>

<%while not rst1.eof

			if isnull(rst1("postdate")) then
				
			%>
	<tr valign="top" border="0" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=rst1("leaseutilityid")%>','<%=rst1("ypid")%>')"> 
		<td width = "<%=response.write(columnwidth)%>" align="right">
			<b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("billyear")%>/<%=rst1("billperiod")%></font></b>
		</td>
		<td width = "<%=response.write(columnwidth)%>" align="right">
			<b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(cdbl(rst1("kwh"))/usagedivisor,2)%></font></b>
		</td>
		<%
		if utility<>3 then
			%>
			<td width = "<%=response.write(columnwidth)%>" align="right">
				<b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("demand_P"),2)%></font></b>
			</td>
			<%
		end if
		%>
		<td width = "<%=response.write(columnwidth)%>" align="right">
			<b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatcurrency(rst1("TotalAmt"))%></font></b>
		</td>
		<td width = "<%=response.write(columnwidth)%>" align="right">
			<b><font size="1" face="Arial, Helvetica, sans-serif">
			<%
			dim postdate
			postdate = rst1("postdate")
			if isnull(postdate) then
				response.write("no post date")
			else
				response.write(postdate)
			end if
			%>
			</font></b>
		</td>
	</tr>



<%
end if
rst1.movenext
wend
set cnn1 = nothing
%>
</table>
</body>

</html>
