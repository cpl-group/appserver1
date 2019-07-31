<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, meterid, luid, billingid, billingtype, utility, measure, units
bldg = Request("bldg")
meterid = request("meterid")
billingid = Request.QueryString("billingid")
utility = request("utility")

select case cint(utility)
case 1
	units = "lbs"
case 2
	units = "KW"
case 3
	units = "Mlbs"
case 4
	units = "Mlbs"
case 5
	units = "Mlbs"
case 6
	units = "Tons"
case 9
	units = "KVA"
case else
	units = "(Usage) un"
end select


dim cnn1, cmd1, rst1, rst2, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)
rst1.open "SELECT measure FROM tblutility WHERE utilityid="&utility, getConnect(0,bldg,"billing")
if not rst1.eof then measure = rst1("measure")
rst1.close

if trim(billingid)<>"" then
    rst1.open "SELECT BillingName, leaseutilityid FROM tblLeases l , tblleasesutilityprices lup WHERE l.billingid=lup.billingid AND lup.utility="&utility&" AND l.billingId="&billingid, cnn1
    if not(rst1.eof) then
        luid = cint(rst1("leaseutilityid"))
    end if
    rst1.close
    billingtype = "Tenant "
elseif trim(bldg)<>"" then
    billingtype = "Building "
elseif trim(meterid)<>"" then
    billingtype = "Meter "
else
    billingtype = " "
end if
%>
<html>
<head>
<title>Meter Details - Usage History</title>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<script> 
function loadentry(ypid, bldg)
{   <%if luid<>"" then%>
        var temp = '/genergy2/eri_th/meterservices/meterbillsummaryPDF.asp?Y='+ ypid +'&luid=<%=luid%>&meterid=<%=meterid%>&bldg=<%=bldg%>&billingid=<%=billingid%>&utility=<%=utility%>';
    <%else%>
        var temp = 'bldgdetail.asp?Y='+ ypid +'&meterid=<%=meterid%>&bldg=<%=bldg%>&billingid=<%=billingid%>&utility=<%=utility%>';
    <%end if%>
    //alert(temp)
	document.location.href = temp;
}
</script>
<body bgcolor="#FFFFFF" onload="parent.closeLoadBox('loadFrame2')" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%

if luid<>"" then
    sql = "select top 12 t.datestart, t.dateend, sum(isnull(t.energy,0)) as consumption, isnull(t.totalamt,0) as TotalBillAmt, sum(isnull(m.demand_P,0)) as demand, sum(isnull(m.Used,0)) as TotalUsage, t.ypid FROM tblmetersbyperiod m join tblbillbyperiod t on t. leaseutilityid=m. leaseutilityid and t.ypid=m.ypid  where t.reject=0 and t.leaseutilityid="& luid &" and t.dateend >= '01/01/2002' group by  t.datestart, t.dateend, t.ypid, t.totalamt order by dateend desc"
    sql = "select bp.datestart, bp.dateend, bp.ypid, totalamt as TotalBillAmt, m.used as TotalUsage, m.demand FROM tblbillbyperiod b, billyrperiod bp, (select leaseutilityid,sum(used) as used ,sum(demand_p)as demand,billyear ,billperiod from tblmetersbyperiod group by leaseutilityid,billyear,billperiod)m where b.reject=0 and b.ypid=bp.ypid and b.billyear=m.billyear and b.billperiod=m.billperiod and b.leaseutilityid=m.leaseutilityid and b.leaseutilityid="&luid&" and b.billyear>=2002 order by b.billyear desc, b.billperiod desc"
else
  select case utility
    case 6
      sql = "SELECT TOP 12 0 as OnPeak, 0 as OffPeak, datestart, dateend, isnull(totalbillamt,0) as TotalBillAmt, 0 as demand, isnull(totaltonhrs,0) as TotalUsage, ub.ypid FROM utilitybill_chilledwater ub, billyrperiod b WHERE ub.ypid=b.ypid and b.bldgnum='"&bldg&"' and b.utility=3 ORDER BY dateend DESC"
    case 3
      sql = "SELECT TOP 12 0 as OnPeak, 0 as OffPeak, datestart, dateend, isnull(avgcost,0) as consumption, isnull(totalbillamt,0) as TotalBillAmt, 0 as demand, isnull(totalccf,0) as TotalUsage, ub.ypid FROM utilitybill_coldwater ub, billyrperiod b WHERE ub.ypid=b.ypid and b.bldgnum='"&bldg&"' and b.utility=3 ORDER BY dateend DESC"
    case 1
      sql = "SELECT TOP 12 0 as OnPeak, 0 as OffPeak, datestart, dateend, isnull(avgcost,0) as consumption, isnull(totalbillamt,0) as TotalBillAmt, 0 as demand, isnull(MLbUsage,0) as TotalUsage, ub.ypid FROM utilitybill_steam ub, billyrperiod b WHERE ub.ypid=b.ypid and b.bldgnum='"&bldg&"' and b.utility=1 ORDER BY dateend DESC"
    case 2
      sql = "SELECT TOP 12 datestart, dateend, sum(isnull(totalbillamt,0)) as TotalBillAmt, sum(isnull(totalkw,0)) as demand, sum(isnull(totalkwh,0)) as TotalUsage, ub.ypid FROM utilitybill ub, billyrperiod b WHERE ub.ypid=b.ypid and b.bldgnum='"&bldg&"' GROUP BY datestart, dateend, ub.ypid ORDER BY dateend DESC"
  end select
end if
'response.write sql&"<br>"&cnn1&utility
'response.end
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#000000" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><%'billingtype%>Billing History</b></font></td>
    <td bgcolor="#000000" width="46%" height="2">
      <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location='options.asp?bldg=<%=bldg%>&meterid=<%=meterid%>&luid=<%=luid%>&utility=<%=utility%>'" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Options</a></b></font></div>
    </td>
  </tr>
  <tr>
    <td width="46%">&nbsp;</td>
    <td width="46%">&nbsp;</td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0099FF"> 
    <td width="15%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">From</font></b></td>
    <td width="15%" align="center"><b><font face="Arial" size="1" color="#FFFFFF">To</font></b></td>
    <%if trim(utility)="2" then%>
    <td width="20%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total <%=units%></font></b></td>
    <%end if%>
    <td width="20%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total <%=measure%></font></b></td>
    <td width="30%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total Billed Amount</font></b></td>
  </tr>
</table>

<%
while not rst1.eof
%>

<div align="left">

  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
    <tr valign="top" style="cursor:hand" onmouseover="this.style.backgroundColor='lightgreen'" onmouseout="this.style.backgroundColor='CCCCCC'; " onclick="javascript:loadentry('<%=rst1("ypid")%>','<%=bldg%>')" bgcolor="#CCCCCC"> 
      <td width="15%" align="center"><div align="right"><font color="#000000"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("datestart")%></font></b></font></div></td>
      <td width="15%" align="right"><b><font face="Arial, Helvetica, sans-serif" size="1" color="#000000"><%=rst1("dateend")%></font></b></td>
    <%if trim(utility)="2" then%>
      <td width="20%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif" color="#000000"><%=formatnumber(rst1("demand"),0)%></font></b></td>
    <%end if%>
      <td width="20%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif" color="#000000"><%=formatnumber(rst1("TotalUsage"),0)%></font></b></td>
      <td width="30%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">$<%=formatnumber(rst1("TotalBillAmt"))%></font></b></td>
    </tr>
  </table>


</div>


<%
rst1.movenext
wend
rst1.close
set cnn1 = nothing

%>
<p>&nbsp;</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td bgcolor="#000000" align="center">
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click any Bill Period 
    row for Detialed information</i></b></font></p>
</td></tr></table>
</body>

</html>
