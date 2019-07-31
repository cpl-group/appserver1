<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Floor Detail</title>
</head>

<body bgcolor="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request("B")
fl = request("F")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql="exec sp_floor_detail '" & bldg & "', '"&fl&"'"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
riser=rst1("riser")
'response.write riser
'response.end
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Floor 
      Detail  as of: <%=rst1("date")%>, 
      <%=rst1("time")%></font></b></font></td>
  </tr>
  
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0099FF"> 
    <td width="9%" height="1%" align="center"><b><font size="1" face="Arial">Floor</font></b></td>
    <td width="16%" height="1%" align="center"><b><font face="Arial" size="1">Riser</font></b></td>
<td width="21%" height="1%" align="center"><b><font size="1" face="Arial">Current 
      Demand KW</font></b></td>
    <td width="20%" height="1%" align="center"><b><font size="1" face="Arial">Delivered 
      WSQFT</font></b></td>
   
  </tr>
</table>



<div align="left">
<table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
<%
while not rst1.eof
%>
<form name="form1" method="post" action=""> 
  <tr>
        <td width="9%" height="19" align="center"> 
          <div align="right"><b><a href="floorsummary.asp?b=<%=bldg%>&f=<%=fl%>"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("floor")%></font></a></b></div>
      </td>
        <td width="16%" height="19" align="right"><b><a href="riser.asp?b=<%=bldg%>&r='<%=riser%>'"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("riser")%></font></a></b></td>
  <td width="21%" height="19" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("current_demand_kw"),2)%></font></b></td>
        <td width="20%" height="19" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("delivered_wsqft"),2)%></font></b></td>
      
    </tr>
 </form>
<% 
		rst1.movenext
		Wend
		
		%>
 </table>
</div>


<%
rst1.close
set cnn1 = nothing
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click any Floor or Riser
    for  detailed information</i></b></font></p>
</div>
</body>

</html>
