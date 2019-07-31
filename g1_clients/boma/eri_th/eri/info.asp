<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>New Page 1</title>

<meta name="Microsoft Theme" content="none, default">
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


<%
Tenant = Request("qcatnr")
Set cnn2 = Server.CreateObject("ADODB.Connection")
openStr= getconnect(0,0,"Engineering") 
cnn2.Open openStr

Set rst1 = Server.CreateObject("ADODB.Recordset")

sql2 = "SELECT * FROM tenant_history WHERE (tenant_no='" & Tenant & "')  order by date_event "  
'response.write sql2
'response.end
rst1.Open sql2, cnn2,adOpenstatic %>

<body bgcolor="#0099FF">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
  <tr bgcolor="#999999"> 
    <td width="10%" colspan="10"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF" size="2"><i>Tenant 
      # Acme<%=left (rst1("Tenant_no"),4)%></i></font></b></td>
  </tr>
  <tr bgcolor="#0099FF" style="font-family:Arial, Helvetica, sans-serif; font-size:13;color:white"> 
    <td width="8%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">Date 
      Event</font></i></b></td>
    <td width="8%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Previous 
      Charge</font></b></td>
    <td width="8%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Surveyed 
      KWH</font></b></td>
    <td width="8%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Surveyed 
      KW</font></b></td>
    <td width="8%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">% 
      Rate</font></i></b></td>
    <td width="8%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">% 
      MAC</font></i></b></td>
    <td width="8%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">One 
      Time Charge</font></i></b></td>
    <td width="8%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">Rate 
      Increase</font></i></b></td>
    <td width="8%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">New 
      Monthly Charge</font></i></b></td>
    <td width="28%" align="center"><b><i><font face="Arial, Helvetica, sans-serif" size="1">Note</font></i></b></td>
  </tr>
</table>
<div style="overflow:auto;height:220">
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
    <%
Do While Not rst1.EOF
 %>
    <tr style="font-family:Arial, Helvetica, sans-serif; font-size:13" bgcolor="#FFFFFF"> 
      <td width="8%" align="right"><i><%=rst1("date_event")%></i></td>
	  <td width="8%" align="right"><i> 
        <%if cDBL(rst1("org_monthly"))<>0 then%>
        <%=Formatcurrency(rst1("org_monthly"),2)%> 
        <%else%>
        <%="-"%> 
        <%end if%>
        </i></td>
	  <td width="8%" align="center"><i><%=rst1("sur_kwh")%></i></td>
	  <td width="8%" align="center"><i><%=rst1("sur_kw")%></i></td>
	  <td width="8%" align="right"><i><%=FormatPercent(rst1("rate"),2)%></i>&nbsp;</td>
	  <td width="8%" align="right"><i><%=FormatPercent(rst1("fuel"),2)%></i></td>
	  <td width="8%" align="right"><i> 
        <%if rst1("org_monthly")=0  and trim(rst1("code"))  <> "9999" then response.write Formatcurrency(rst1("Charge"),2) %>
        </i>&nbsp;</td>
	  <td width="8%" align="right"><i> 
        <%if rst1("org_monthly")<>0  then response.write Formatcurrency(cDBL(rst1("Charge"))-cDBL(rst1("org_monthly")),2)%>
        </i>&nbsp;</td>
	  <td width="8%" align="right"><i> 
        <%if rst1("org_monthly")<>0  or trim(rst1("code"))  = "9999" then response.write Formatcurrency(rst1("charge"),2) %>
        </i>&nbsp;</td>
	  <td width="28%"><i><%=rst1("note")%></i>&nbsp;</td>
</tr>
      <%
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn2.Close
Set cnn2 = Nothing
 %>
    </table>
</div>
</body>
</html>
