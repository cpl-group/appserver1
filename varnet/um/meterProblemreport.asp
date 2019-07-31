<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">


<%@Language="VBScript"%>
<%
bldg= Request.Querystring("bldg")
billyear= Request.Querystring("year")
billperiod= Request.Querystring("period")
qtable="peakdemand"

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql = "SELECT DISTINCT meters.meternum,tblleases.billingName,tblleases.flr, peakdemand.billyear,peakdemand.billperiod, peakdemand.usernote as pnote,consumption.usernote as cnote FROM peakdemand JOIN meters ON peakdemand.meterid = meters.meterid JOIN consumption ON consumption.meterid = meters.meterid join tblleasesutilityprices tp on meters.leaseutilityid=tp.leaseutilityid join tblleases on tp.billingid=tblleases.billingid  WHERE peakdemand.billyear = '" & billyear & "' AND peakdemand.billperiod = '" & billperiod & "' AND   consumption.billyear = '" & billyear & "' AND consumption.billperiod = '" & billperiod & "' AND (peakdemand.usernote IS NOT NULL OR consumption.usernote IS NOT NULL) AND meters.bldgnum = '" & bldg & "'"

rst1.Open sql, cnn1, 0, 1, 1
if rst1.EOF then
%>
<table width="100%" border="0">
  <tr>
    <td>NO PROBLEMS FOUND</td>
  </tr>
</table>
<%
else

%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">Genergy 
      Meter <%=qtable%> Problem Report for Building Number <%=bldg%></font></b></font></td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%" bgcolor="#CCCCCC" height="28"><font face="Arial, Helvetica, sans-serif">Meter 
            # </font></td>
			
          <td width="15%" bgcolor="#CCCCCC" height="28"><font face="Arial, Helvetica, sans-serif">Tenant 
            Name</font></td>
          <td width="7%" bgcolor="#CCCCCC" height="28"><font face="Arial, Helvetica, sans-serif">Floor</font></td>
          <td width="10%" bgcolor="#CCCCCC" height="28"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif">Period</font></div>
          </td>
          <td width="20%" bgcolor="#CCCCCC" height="28"><font face="Arial, Helvetica, sans-serif">User 
            Note (Consumption) </font></td>
          <td width="40%" bgcolor="#CCCCCC" height="28"><font face="Arial, Helvetica, sans-serif">User 
            Note (Peakdemand) </font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <% While not rst1.EOF %>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr align="left" valign="top"> 
          <td width="8%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("meternum")%></font></td>
		  <td width="15%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("billingname")%></font></td>
		  <td width="7%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("flr")%></font></td>
          <td width="10%" height="37"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("billperiod")%>/<%=rst1("billyear")%></font></div>
          </td>
          <td width="20%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("cnote")%></font></td>
          <td width="40%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("pnote")%></font></td>
        </tr>
      </table>
      <hr>
      <% 
		rst1.movenext
		Wend
		%>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<%
end if
%>
</body>
</html>