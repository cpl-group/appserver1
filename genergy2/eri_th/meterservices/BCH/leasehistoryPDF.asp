<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, leaseid, title, units, utility, usagedivisor,demo, billingid
bldg = request.querystring("B")
leaseid = request.querystring("leaseid")
units = "Usage"
utility = request("utility")
demo = request("demo")


if instr(leaseid,"u")>0 then 
  utility = mid(leaseid,2)
  leaseid = 0
  billingid = 0
elseif leaseid = "A" then 
	utility = 0 
	leaseid = 0
	billingid = 0
elseif instr(leaseid,"_")>0 then 
	billingid = split(leaseid,"_")(1)
	leaseid = 0
	utility = 0 
end if

if trim(utility)="" then utility = 0
dim cnn1, cmd1, rst1, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)

if leaseid<>0 then
	title = "Tenant"
  rst1.open "SELECT measure, lup.utility FROM tblleasesutilityprices lup, tblutility u WHERE lup.utility=u.utilityid and leaseutilityid="&leaseid, cnn1
  if not rst1.eof then 
    units=rst1("measure")
    utility = rst1("utility")
  end if
  rst1.close
  sql="SELECT top 24 isnull(TotalAmt,0) as TotalAmt, b.postdate, b.BillYear, b.BillPeriod, sum(isnull(used,0)+isnull(usedoff,0)+isnull(usedint,0)) as KWH, sum(demand_p+isnull(demand_int,0)+isnull(demand_off,0)) as demand_P, m.ypid, m.leaseutilityid FROM tblmetersbyperiod m INNER JOIN tblbillbyperiod b on b.id=m.bill_id and b.ypid=m.ypid and b.LeaseUtilityId=m.LeaseUtilityId WHERE b.reject=0 and m.leaseutilityid="&leaseid&" and b.bldgnum='" & bldg & "' and b.posted = 1 GROUP BY b.BillYear, b.BillPeriod, m.ypid, m.leaseutilityid, TotalAmt, b.postdate order by b.billyear desc ,b.billperiod desc"
elseif billingid <> 0 then 
   sql="select top 24 lup.billingid, b.billyear, b.billperiod, SUM(totalamt) as TotalAMT from tblbillbyperiod b inner join tblleasesutilityprices lup on lup.leaseutilityid = b.leaseutilityid where b.reject = 0 and lup.billingid = '"&billingid&"' and b.posted = 1 group by billyear, billperiod, billingid order by b.billyear desc ,b.billperiod desc"	
elseif utility<>0 then 
	title = "Building"
	rst1.open "SELECT measure FROM tblutility u WHERE utilityid="&utility, getConnect(0,0,"dbCore")
	if not rst1.eof then units=rst1("measure")
	rst1.close
	sql = "select top 24 0 as leaseutilityid, sum(kwh) as kwh,sum(demand_p) as demand_p,sum(totalamt) as totalamt,ypid,billyear,billperiod, max(postdate) as postdate from (select sum(used+isnull(usedoff,0)+isnull(usedint,0)) as kwh,sum(demand_p+isnull(demand_int,0)+isnull(demand_off,0)) as demand_p ,b.totalamt as totalamt ,m.ypid,billyear,billperiod, postdate from tblmetersbyperiod m, (SELECT DISTINCT ypid, leaseutilityid,SUM(totalamt) AS totalamt,utility, postdate, id FROM tblbillbyperiod WHERE reject=0 and bldgnum ='"&bldg&"' and posted=1 and utility="&utility&" GROUP BY ypid,leaseutilityid,utility, postdate, id) B where b.id=m.bill_id and m.ypid=b.ypid and b.leaseutilityid=m.leaseutilityid group by b.totalamt,m.ypid,billyear,billperiod, postdate)Z group by ypid,billyear,billperiod order by billyear desc,billperiod desc"
else
   sql="select top 24 b.billyear, b.billperiod, SUM(totalamt) as TotalAMT from tblbillbyperiod b where b.reject = 0 and b.bldgnum = '"&bldg&"' and b.posted = 1 group by billyear, billperiod order by b.billyear desc ,b.billperiod desc"
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

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<html>

<head>
<title>Meter Details - Usage History</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head><style type="text/css">
<!--
BODY {
	SCROLLBAR-FACE-COLOR: #336699;
	SCROLLBAR-HIGHLIGHT-COLOR: #336699;
	SCROLLBAR-SHADOW-COLOR: #333333;
	SCROLLBAR-3DLIGHT-COLOR: #333333;
	SCROLLBAR-ARROW-COLOR: #333333;
	SCROLLBAR-TRACK-COLOR: #333333;
	SCROLLBAR-DARKSHADOW-COLOR: #333333;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>

<script> 
function loadentry(luid, ypid,billingid, b, bp, by){

	var temp = 'meterbillsummaryPDF.asp?demo=<%=demo%>&utility=<%=utility%>&l=' +luid+'&Y='+ypid+'&building='+b+'&bp='+bp+'&by='+by+'&billingid='+billingid
	parent.document.frames.panel_2.location = temp
}
function settonull(){
	var temp = '/null.htm'
	parent.document.frames.panel_2.location = temp
}
</script>
<body bgcolor="#FFFFFF"onLoad="settonull()">
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
	<tr bgcolor="#336699" > 
		<%
		dim columnwidth
		if utility = 3 then 
			columnwidth = "25%"
		else 
			columnwidth = "20%"
		end if
		%>
		<td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Bill Period</font></b></td>
	<%if utility<>0 then%><td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial"  color="#FFFFFF"><%=units%></font></b></td><%end if%>
    <%if utility<>3 and utility <> 0 then%><td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Demand</font></b></td><%end if%>
		<td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total Amount</font></b></td>
	<%if utility<>0 then%><td width = "<%=response.write(columnwidth)%>" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Post Date</font></b></td><%end if%>
	</tr>
</table>
<table width = "100%" cellpadding = 0 cellspacing = 0>

<%

Dim ypid, billperiod, billyear

	while not rst1.eof
	if utility <> 0 then
		leaseid = rst1("leaseutilityid")
		ypid    = rst1("ypid") 
		billperiod = rst1("billperiod")
	    billyear = rst1("billyear")
	else
		leaseid = 0
		ypid = 0
		billperiod = rst1("billperiod")
	    billyear = rst1("billyear")
	end if 

%>


	<tr valign="top" border="0" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=leaseid%>','<%=ypid%>','<%=billingid%>','<%=bldg%>',<%=billperiod%>,<%=billyear%>)"> 
		<td width = "<%=response.write(columnwidth)%>" align="right">
			<b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("billyear")%>/<%=rst1("billperiod")%></font></b>
		</td>
		<% if utility <> 0 then %>
		<td width = "<%=response.write(columnwidth)%>" align="right">
			<b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(cdbl(rst1("kwh"))/usagedivisor,2)%></font></b>
		</td>
		<% end if %>
		<%
		if utility<>3 and utility <> 0 then
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
	<% if utility <> 0 then %>
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
	<% end if%>
	</tr>



<%
rst1.movenext
wend
set cnn1 = nothing
%>
</table>
</body>

</html>
