<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body bgcolor="#FFFFFF">
<%
bldg = Request("B")
meterid = request("M")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql="SELECT distinct ypid FROM tblbillbyperiod where bldgnum='" & bldg & "' and datediff(year,datestart,getdate())<=12  and dateend < getdate() "

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#0099FF"> 
      <div align="center"><font face="Arial" size="4" color="#000000">Meter 
        Details </font></div>
    </td>
  </tr>
</table>
<p>&nbsp;</p><table border="1" width="100%" height="1">
  <tr bgcolor="#0099FF"> 
    <td width="14%" height="1%" align="center"><font size="2" face="Arial">Bill Period</font></td>
    <td width="8%" height="1%" align="center"><font face="Arial">&nbsp;</font></td>
    <td width="8%" height="1%" align="center"><font size="2" face="Arial">Previous</font></td>
    <td width="20%" height="1%" align="center"><font size="2" face="Arial">Current</font></td>
    <td width="13%" height="1%" align="center"><font size="2" face="Arial">On Peak</font></td>
    <td width="12%" height="1%" align="center"><font size="2" face="Arial">Off Peak</font></td>
    <td width="13%" height="1%" align="center"><font size="2" face="Arial">Kwhr</font></td>
    <td width="12%" height="1%" align="center"><font size="2" face="Arial">Demand</font></td>
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

<table border="0" width="100%" height="1">
  <tr>
    <td width="14%" height="1%" align="center"><a href="pgibilldemo.asp?l=<%=rst2("leaseutilityid")%>&amp;Y=<%=rst2("ypid")%>"><font size="2" face="Arial"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></a></td>
    <td width="8%" height="1%" align="right"></td>
    <td width="8%" height="1%" align="right"><font size="2"><font face="Arial"><%=formatnumber(rst2("Prevkwh"),0)%></font></font></td>
    <td width="20%" height="1%" align="right"><font size="2"><font face="Arial"><%=formatnumber(rst2("currentkwh"),0)%></font></font></td>
    <td width="13%" height="1%" align="right"><font size="2"><font face="Arial"><%=formatnumber(rst2("onpeak"),0)%></font></font></td>
    <td width="12%" height="1%" align="right"><font size="2"><font face="Arial"><%=formatnumber(rst2("offpeak"),0)%></font></font></td>
    <td width="13%" height="1%" align="right"><font size="2"><font face="Arial"><%=formatnumber(rst2("kwhused"),0)%></font></font></td>
    <td width="12%" height="1%" align="right"><font size="2"><font face="Arial"><%=formatnumber(rst2("demand_P"),2)%></font></font></td>
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

</body>

</html>
