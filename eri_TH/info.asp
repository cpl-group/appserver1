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
openStr=  getconnect(0,0,"Engineering") 
cnn2.Open openStr

Set rst1 = Server.CreateObject("ADODB.Recordset")

sql2 = "SELECT * FROM tenant_history WHERE (tenant_no='" & Tenant & "') "  

rst1.Open sql2, cnn2,adOpenstatic %>

<body>

<div align="center">
 <center>
 <table border="1" width="100%">
 <tr>
 <td width="10%" align="center" bgcolor="#0066CC"><b><font face="Arial" color="#FFFF00"><i>Tenant #</i></font></b></td>
 <td width="10%" align="center" bgcolor="#0066CC"><b><font face="Arial" color="#FFFF00"><i>Date Event</i></font></b></td>
 <td width="10%" align="center" bgcolor="#0066CC"><b><font face="Arial" color="#FFFF00"><i>% Rate</i></font></b></td>
 <td width="10%" align="center" bgcolor="#0066CC"><b><font face="Arial" color="#FFFF00"><i>% MAC</i></font></b></td>
 <td width="10%" align="center" bgcolor="#0066CC"><b><font face="Arial" color="#FFFF00"><i>Charge</i></font></b></td>
 <td width="50%" align="center" bgcolor="#0066CC"><b><font face="Arial" color="#FFFF00"><i>Note</i></font></b></td>
 </tr>
 <tr>
     
 <%
Do While Not rst1.EOF
 %> 
 <td width="10%" align="center"><font face="Arial"><i><%=rst1("Tenant_no")%></i></font></td>
 <td width="10%" align="center"><font face="Arial"><i><%=rst1("date_event")%></i></font></td>
      
      <%if IsNull(rst1("rate")) Then %>
 <td width="10%" align="center">&nbsp;</td>
      <%else %> 
 <td width="10%" align="center"><font face="Arial"><i><%=FormatPercent(rst1("rate"),2)%></i></font></td>
 	   <% end if %>
 
  	   <%if IsNull(rst1("fuel")) Then %>
 <td width="10%" align="center">&nbsp;</td>
      <%else %>     
 <td width="10%" align="center"><font face="Arial"><i><%=FormatPercent(rst1("fuel"),2)%></i></font></td>
 		<% end if %>
 		
 	  <%if IsNull(rst1("Charge")) Then %>
 <td width="10%" align="center">&nbsp;</td>
      <%else %> 
 <td width="10%" align="center"><font face="Arial"><i><%=Formatcurrency(rst1("charge"),2)%></i></font></td>
 	   <% end if %>
 	  
 <td width="50%" align="center"><font face="Arial"><i><%=rst1("note")%></i></font></td>
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
 </center>
</div>


</body>

</html>
