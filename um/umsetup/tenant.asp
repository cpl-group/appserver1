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

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"
strsql = "SELECT * FROM Tenants WHERE (BldgNum = '" & bldg & "') ORDER BY TenantName"
rst1.Open strsql, cnn1, 0, 1, 1

if not rst1.eof then

Response.Write "<script>" & vbCrLf

Response.Write "parent.frames.lease.location = " & Chr(34) & _
                  "leases_info.asp?bldg=" & rst1("bldgnum") & "&" & "ten=" & rst1("tenantnum") & Chr(34) & vbCrLf
                  
 Response.Write "</script>" & vbCrLf                 
                  %>
<table border="1" width="100%" height="41" bordercolor="#000000">
  <tr bgcolor="#66CCFF"> 
    <td width="20%" height="10"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Tenant 
        #</font>
    </td>
    <td width="80%" height="10"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Tenant 
        Name</font>
    </td>
  </tr>
  <%
 while not rst1.eof
 %>
  <tr> 
    <td width="20%" height="19">
      <p align="left"><font size="2"> 
        <%Response.write "<a href=" & chr(34) & "leases_info.asp?bldg=" & rst1("bldgnum") & "&" & "ten=" & rst1("tenantnum") & Chr(34) & "target=lease>"%>
        <%=rst1("tenantnum")%> </font>
    </td>
    <td width="80%" height="19">
      <p align="left"><font size="2"><%=rst1("tenantname")%></font>
    </td>
  </tr>
  <%
rst1.movenext
wend
rst1.close
set cnn1 = nothing
end if
%>
</table>
</body>

</html>
