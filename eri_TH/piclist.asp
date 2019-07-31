
<!-- #include file="secure.inc" -->
<html>
<head>
<title>Holiday Photo Catalog Main Menu</title>
<!-- #include file="./adovbs.inc" -->
<link rel="stylesheet" type="text/css" 
      href="../holiday/holiday.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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


<body>

<div align="left">
  <table border="1" width="100%" height="20">
    <tr>
      <td width="10%" height="20" align="center" bgcolor="#003399"><p align="center"><b><font color="#00FF00">Tenant #</font></b></td>
      <td width="30%" height="20" align="center" bgcolor="#003399"><p align="center"><b><font color="#00FF00">Tenant Name</font></b></td>
      <td width="10%" height="20" align="center" bgcolor="#003399"><b><font color="#00FF00">Sqft</font></b></td>
      <td width="10%" height="20" align="center" bgcolor="#003399"><b><font color="#00FF00">Monthly Charge</font></b></td>
      <td width="10%" height="20" align="center" bgcolor="#003399"><b><font color="#00FF00">Yearly Charge</font></b></td>
      <td width="10%" height="20" align="center" bgcolor="#003399"><b> <font color="#00FF00"> $ / Sqft</font></b></td>
      <td width="10%" height="20" align="center" bgcolor="#003399"><b><font color="#00FF00">Lease Exp. Date</font></b></td>
      <td width="10%" height="20" align="center" bgcolor="#003399"><b><font color="#00FF00">Move Out Date</font></b></td>
    </tr>
    
<%
bldg = Request("qcatnr")

Set cnn1 = Server.CreateObject("ADODB.Connection")
openStr= "data Source=eri;user Id=web"
cnn1.Open openStr

Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tenant_info WHERE (bldg_no='" & bldg & "') order by tenantname"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.

rst1.movefirst

Response.Write "<script>" & vbCrLf

If rst1.EOF Then
  Response.Write "parent.frames.info.location = " & _
                 Chr(34) & "blank.htm" & Chr(34) & vbCrLf
Else
  Response.Write "parent.frames.info.location = " & _
                  Chr(34) & _
                  "info.asp" & _
                  "?qcatnr=" & rst1("Tenant_no") & _
                  Chr(34) & vbCrLf
End If

Response.Write "</script>" & vbCrLf

Do While Not rst1.EOF
%> 

<tr>
      <td width="10%" height="20" align="center"><p align="center">
      <%
Response.write "<a href=" & chr(34) & "info.asp?" & "qcatnr=" & rst1("Tenant_no") & chr(34) & " target=info>" & rst1("Tenant_no")
%>
        </td>
      <td width="30%" height="20" align="center"><p align="left"><%=rst1("Tenantname")%></td>
      
      <% If IsNull(rst1("sqft")) then %>
        	<td width="10%" height="20" align="center"></td>
			<% else %>        	
          <td width="10%" height="20" align="center"><p align="right"><%=Formatnumber(rst1("sqft"),0)%></td>
      		<%end If%>
      		
        	<% If IsNull(rst1("ccm")) then %>
        	<td width="10%" height="20" align="center"></td>
			<% else %>        	
          <td width="10%" height="20" align="center"><p align="right"><%=FormatCurrency(rst1("ccm"),2)%></td>
      		<%end If%>
      		
      		<% If IsNull(rst1("ccy")) then %>
        	<td width="10%" height="20" align="center"></td>
			<% else %>        	
          <td width="10%" height="20" align="center"><p align="right"><%=FormatCurrency(rst1("ccy"),2)%></td>
      		<%end If%>

   
      		<% If IsNull(rst1("cost_sqft")) then %>
        	<td width="10%" height="20" align="center"></td>
			<% else %>        	
          <td width="10%" height="20" align="center"><p align="right"><%=formatcurrency(rst1("cost_sqft"),2)%></td>
      		<%end If%>     		
      		
      		
      		<td width="10%" height="20" align="center"><%=rst1("lease_exp_date")%></td>
      		<td width="10%" height="20" align="center"><p align="right"><%=rst1("Move_out_date")%></td>
    </tr

</table <%
  rst1.MoveNext  
Loop

'Close and destroy the recordset and connection objects.
rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
%>>
  </table>
</div>
</body>






























