
<%@Language="VBScript"%>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

if request("infotype")="bld" then 

strsql = "Update buildings Set bldgname='" & Request.Form("bldgname") & "', strt='" & Request.Form("strt") & "', city='" & Request.Form("city") &"', state='" & Request.Form("state") & "', zip='" & Request.Form("zip") & "' where bldgnum=" & Request.Form("bldgnum")

else

strsql = "Update buildings Set btbldgname='" & Request.Form("btbldgname") & "', btstrt='" & Request.Form("btstrt") & "', btcity='" & Request.Form("btcity") &"', btstate='" & Request.Form("btstate") & "', btzip='" & Request.Form("btzip") & "' where bldgnum=" & Request.Form("bldgnum") 

end if

cnn1.execute strsql

set cnn1=nothing

tmp =  "window.close()"
%>
<html>
<head>
<title>Updating Building Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<%
Response.Write "<script>" & vbCrLf

Response.Write tmp
Response.Write "</script>" & vbCrLf
%>

</body>
</html>
