<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Tenant Selection</title>
</head>

<body bgcolor="#FFFFFF">
<%
bldg = Request("bldg")

tmpMoveFrame =  "parent.frames.tenant.location = " & Chr(34) & _
                  "tenant.asp?bldg=" & bldg & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf

Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"
strsql = "SELECT * FROM Buildings WHERE (BldgNum = '" & bldg & "')"
rst1.Open strsql, cnn1, 0, 1, 1

if not rst1.eof then
%>
<table border="1" width="100%" bordercolor="#000000" height="100%">
  <tr> 
    <td width="27%" height="23"> <font size="2" face="Arial, Helvetica, sans-serif"> 
      Building Address</font></td>
    <td width="73%" height="23"><font size="2"><%=rst1("strt") & " - " & rst1("city") & " - " & rst1("state") & " - " & rst1("zip")%></font></td>
  </tr>
  <tr> 
    <td width="27%"><font size="2" face="Arial, Helvetica, sans-serif">Billing 
      Address.</font></td>
    <td width="73%"><font size="2"><%=rst1("btbldgname") & " - " & rst1("btcity") & " - " & rst1("btstate") & " - " & rst1("btzip")%></font></td>
  </tr>
  <tr> 
    <td width="27%"><font size="2" face="Arial, Helvetica, sans-serif">Service 
      Class</font></td>
    <td width="73%"><font size="2"><%=rst1("ratebldg")%></font></td>
  </tr>
  <tr> 
    <td width="27%"><font size="2" face="Arial, Helvetica, sans-serif">Con Ed. 
      Account</font></td>
    <td width="73%">
      <p align="left"><font size="2"><%=rst1("elecacctnum")%></font>
    </td>
  </tr>
  <%
rst1.close
set cnn1 = nothing
end if
%>
</table>
</body>

</html>
