<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
bldg = Request("B")
meterid = request("m")
luid = request("luid")

dim billingtype
if trim(luid)<>"" then
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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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
        var temp = '/eri_th/meter/meterbillsummary2.asp?Y='+ ypid +'&l=<%=luid%>&m=<%=meterid%>&b='+bldg;
    <%else%>
        var temp = 'bldgdetail.asp?Y='+ ypid +'&m=<%=meterid%>&b='+bldg;
    <%end if%>
    //alert(temp)
	document.location.href = temp;
}
</script>
<body bgcolor="#FFFFFF" onload="parent.closeLoadBox('loadFrame2')" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%


Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

if luid<>"" then
    sql = "select top 12 sum(m.onpeak) as OnPeakKWH,sum(m.offpeak) as OffPeakKWH, t.datestart, t.dateend, sum(t.energy) as energy ,t.totalamt as TotalBillAmt, sum(m.demand_P) as TotalKW, sum(m.KWHUsed) as TotalKWH, t.ypid from tblmetersbyperiod m join tblbillbyperiod t on t. leaseutilityid=m. leaseutilityid and t.ypid=m.ypid  where t.leaseutilityid="& luid &" and t.dateend >= '01/01/2002' group by  t.datestart, t.dateend, t.ypid, t.totalamt order by dateend desc"
else
    sql = "select top 12 billyrperiod.billperiod, billyrperiod.billyear,sum(u.OnPeakKWH) as OnPeakKWH, sum(u.OffPeakKWH) as OffPeakKWH, sum(u.TotalKW) as TotalKW, sum(u.TotalKWH) as TotalKWH, sum(u.TotalBillAmt) as TotalBillAmt,datestart, dateend,u.ypid from utilitybill u join billyrperiod on billyrperiod.ypid = u.ypid where u.ypid in (SELECT ypid FROM utilitybill where ypid in (select ypid from billyrperiod where bldgnum='"&bldg&"')) group by datestart, dateend, u.ypid, billyear, billperiod order by billyear desc, billperiod desc"
end if
'response.write sql
'response.end
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#000000" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><%'billingtype%>Billing History</b></font></td>
    <td bgcolor="#000000" width="46%" height="2">
      <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location='<%="options2.asp?b=" & request("b") & "&m=" & request("m") & "&luid=" & luid%>'" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return 
        To Options</a></b></font></div>
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
    <td width="20%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total KW</font></b></td>
    <td width="20%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total KWH</font></b></td>
    <td width="30%" align="center"><b><font size="1" face="Arial" color="#FFFFFF">Total Billed Amount</font></b></td>
  </tr>
</table>

<%
while not rst1.eof
%>

<div align="left">

  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
    <tr valign="top" style="cursor:hand" onmouseover="this.style.backgroundColor='lightgreen'" onmouseout="this.style.backgroundColor='CCCCCC'; " onclick="javascript:loadentry('<%=rst1("ypid")%>','<%=request("b")%>')" bgcolor="#CCCCCC"> 
      <td width="15%" align="center"><div align="right"><font color="#000000"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("datestart")%></font></b></font></div></td>
      <td width="15%" align="right"><b><font face="Arial, Helvetica, sans-serif" size="1" color="#000000"><%=rst1("dateend")%></font></b></td>
      <td width="20%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif" color="#000000"><%=formatnumber(rst1("totalKW"),0)%></font></b></td>
      <td width="20%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif" color="#000000"><%=formatnumber(rst1("TotalKWH"),0)%></font></b></td>
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
