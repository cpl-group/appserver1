<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Meter Details - Usage History</title>
</head>
<script> 
function loadentry(luid, ypid){

	var temp = 'pgibill.asp?l=' +luid+'&Y='+ypid

	
	document.location = temp
}
</script>
<body bgcolor="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request("B")
meterid = request("M")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql="SELECT top 12  ypid FROM tblbillbyperiod where bldgnum='" & bldg & "' group by ypid order by ypid desc"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Meter 
      Details - </b>Usage History</font></td>
  </tr>
  <tr> 
    <td width="46%">&nbsp;</td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0099FF"> 
    <td width="14%" height="1%" align="center"><b><font size="1" face="Arial">Bill 
      Period</font></b></td>
    <td width="8%" height="1%" align="center"><b><font face="Arial" size="1">&nbsp;</font></b></td>
    <td width="8%" height="1%" align="center"><b><font size="1" face="Arial">Previous</font></b></td>
    <td width="20%" height="1%" align="center"><b><font size="1" face="Arial">Current</font></b></td>
    <td width="13%" height="1%" align="center"><b><font size="1" face="Arial">On 
      Peak</font></b></td>
    <td width="12%" height="1%" align="center"><b><font size="1" face="Arial">Off 
      Peak</font></b></td>
    <td width="13%" height="1%" align="center"><b><font size="1" face="Arial">Kwhr</font></b></td>
    <td width="12%" height="1%" align="center"><b><font size="1" face="Arial">Demand</font></b></td>
  </tr>
</table>

<%
while not rst1.eof

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * from tblmetersbyperiod where meterid=" & Meterid & " and ypid=" & rst1("ypid")
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
%>

<div align="left">

  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
    <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=rst2("leaseutilityid")%>','<%=rst2("ypid")%>')"> 
 <td width="14%" height="1%" align="center">
        <div align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></b></div>
      </td>
      <td width="8%" height="1%" align="right"><b><font face="Arial, Helvetica, sans-serif" size="1"></font></b></td>
      <td width="8%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("Prevkwh"),0)%></font></b></td>
      <td width="20%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("currentkwh"),0)%></font></b></td>
      <td width="13%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("onpeak"),0)%></font></b></td>
      <td width="12%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("offpeak"),0)%></font></b></td>
      <td width="13%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("kwhused"),0)%></font></b></td>
      <td width="12%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("demand_P"),2)%></font></b></td>
    </tr>
  </table>


</div>


<%
rst1.movenext
else
rst1.movenext
end if
wend
set cnn1 = nothing
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click any Bill Period 
    row for billing information</i></b></font></p>
</div>
</body>

</html>
