<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="./adovbs.inc" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#0099FF"> 
      <div align="center"><font face="Arial" size="4" color="#000000">Meter 
        Details </font></div>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="605" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
      <%
leaseid = Request("l")
ypid = request("y")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getconnect(0,0,"Engineering") 

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
%>
      <div align="center"> 
        <p align="left"> <b><font face="Arial" size="3">&nbsp;Tenant #  / Tenant Name</font></b> </p>
        <table border="1" width="600" height="1">
          <tr bgcolor="#0099FF"> 
            <td width="7%" height="1%" align="center"><b><font face="Arial">Period</font></b></td>
            <td width="14%" height="1%" align="center"><b><font face="Arial">Energy 
              Charge</font></b></td>
            <td width="12%" height="1%" align="center"><b><font face="Arial">Demand 
              Charge</font></b></td>
            <td width="10%" height="1%" align="center"><b><font face="Arial">Admin 
              Fee</font></b></td>
            <td width="10%" height="1%" align="center"><b><font face="Arial">Service 
              Fee</font></b></td>
            <td width="10%" height="1%" align="center"><b><font face="Arial">Sales 
              Tax</font></b></td>
            <td width="10%" height="1%" align="center"><b><font face="Arial">Total 
              Amt</font></b></td>
          </tr>
          <tr> 
            <td width="7%" height="1%" align="center"><font size="2" face="Arial"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></td>
            <td width="14%" height="1%" align="center">
              <p align="right"><font size="2" face="Arial"><%=FormatCurrency(rst2("energy"),2)%></font></td>
            <td width="12%" height="1%" align="center">
              <p align="right"><font size="2" face="Arial"><%=FormatCurrency(rst2("demand"),2)%></font></td>
            <td width="10%" height="1%" align="center">
              <p align="right"><font size="2" face="Arial"><%=Formatpercent(rst2("Adminfee"),2)%></font></td>
            <td width="10%" height="1%" align="center">
              <p align="right"><font size="2" face="Arial"><%=FormatCurrency(rst2("Addonfee"),2)%></font></td>
            <td width="10%" height="1%" align="center">
              <p align="right"><font size="2" face="Arial"><%=FormatCurrency(rst2("tax"),2)%></font></td>
            <td width="10%" height="1%" align="center">
              <p align="right"><font size="2" face="Arial"><%=FormatCurrency(rst2("totalamt"),2)%></font></td>
          </tr>
        </table>
         </div>
      <font face="Arial">
     <br> 
      </font> 
	  <table border="1" width="600" height="1" cellpadding="0" cellspacing="0" align="center">
        <tr bgcolor="#0099FF"> 
          <td width="20%" height="1%" align="center"><font face="Arial">Meter</font></td>
          <td width="20%" height="1%" align="center"><font face="Arial">On Peak KWH</font></td>
          <td width="20%" height="1%" align="center"><font face="Arial">Off Peak KWH</font></td>
          <td width="20%" height="1%" align="center"><font face="Arial">KWH</font></td>
          <td width="20%" height="1%" align="center"><font face="Arial">Demand</font></td>
        </tr>
        <%
cmd1.CommandText = "select * from tblmetersbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst1 = cmd1.Execute

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
%>
        <tr> 
          <td width="20%" height="1%" align="center"> 
            <p align="left"><a href="pgimeterdemo.asp?b=<%=rst1("bldgnum")%>&amp;m=<%=rst1("meterid")%>"><font size="2" face="Arial"><%=rst1("Meternum")%> 
              </font></a>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><%=formatnumber(rst1("onpeak"),0)%> </font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><%=formatnumber(rst1("offpeak"),0)%> </font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><%=formatnumber(rst1("kwhused"),0)%> </font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><%=formatnumber(rst1("demand_P"),0)%> </font>
          </td>
        </tr>
        <%
tot_onpeak = tot_onpeak + rst1("onpeak")
tot_offpeak= tot_offpeak+ rst1("offpeak")
tot_kwhused= tot_kwhused + rst1("kwhused")
tot_demand_p= tot_demand_p + rst1("demand_P")

rst1.movenext
wend

else
end if
%>
        <tr bgcolor="#CCCCCC"> 
          <td width="20%" height="1%" align="center"> 
            <div align="center"></div>
            <p align="center"><font face="Arial">Totals</font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><b><%=formatnumber(tot_onpeak,0)%> </b> </font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><b><%=formatnumber(tot_offpeak,0)%> </b> </font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><b><%=formatnumber(tot_kwhused,0)%> </b> </font>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><font size="2" face="Arial"><b><%=formatnumber(tot_demand_P,0)%> </b> </font>
          </td>
        </tr>
      </table>
      <%
set cnn1 = nothing
%>
    </td>
  </tr>
</table>
</body>
</html>




