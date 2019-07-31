<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim choice, cnn1, rst1, tmp, strsql
choice=Request("submit")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")
if(choice="Update") then
	strsql = "Update tblTenantSurvey Set surveydate='" & request("surveydate") & "', location='" & request("location") &"', floor='" & request("floor") & "', orderno='" & request("orderno") & "' where tenant_no='" & request("tenant_no") & "' and id="& Request("survey_id")
	tmp =  "window.close()"
elseif choice="Delete" then
	strsql = "exec sp_Delete_Survey_Location '"&Request("survey_id")&"'"
	tmp = "document.location.href='survey_detail.asp?tenant_no="&request("tenant_no")&"'"
	'survey_detail.asp?tenant_no=10 1800&surveyid=1480
	'survey_detail.asp?tenant_no=10+1800&surveyid=1475
	'survey_detail.asp?tenant_no=10 1800&surveyid=0
else
    strsql = "Insert into tblTenantSurvey (tenant_no, location, floor, orderno, surveydate) "_
	& "values ("_
	& "'" & request("tenant_no") & "', "_
	& "'" & request("location") & "', "_
	& "'" & request("floor") & "', "_
	& "'" & request("orderno") & "', "_
	& "'" & request("surveydate") & "')"
	tmp =  "window.close()"
end if	

cnn1.execute strsql
logger strsql
set cnn1=nothing
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
