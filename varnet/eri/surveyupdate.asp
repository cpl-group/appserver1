<%@Language="VBScript"%>
<%
	
choice=Request("submit")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

if(choice="Update") then

	strsql = "Update tblTenantSurvey Set surveydate='" & Request.Form("surveydate") & "', location='" & Request.form("location") &"', floor='" & Request.Form("floor") & "', orderno='" & Request.Form("orderno") & "' where tenant_no='" & Request.Form("tenant_no") & "' and id="& Request("survey_id")

else
    strsql = "Insert into tblTenantSurvey (tenant_no, location, floor, orderno, surveydate) "_
	& "values ("_
	& "'" & Request.Form("tenant_no") & "', "_
	& "'" & Request.form("location") & "', "_
	& "'" & Request.Form("floor") & "', "_
	& "'" & Request.Form("orderno") & "', "_
	& "'" & Request.Form("surveydate") & "')"

	
end if	

cnn1.execute strsql
set cnn1=nothing
tmp =  "window.close()"
%>
<html>
<head>
<title>Updating Tenant Survey Information</title>
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
