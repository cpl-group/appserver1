<%@Language="VBScript"%>
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"

	strsql = "Insert into employees (login, password, name, um, eri, opslog, ts,corp,it,admin) "_
	& "values ("_
	& "'" & Request.Form("login") & "', "_
	& "'" & Request.Form("password") & "', "_
	& "'" & Request.Form("name") & "', "_
	& "'" & Request.Form("um") & "', "_
	& "'" & Request.Form("eri") & "', "_
	& "'" & Request.Form("opslog") & "', "_
	& "'" & Request.Form("ts") & "', "_
	& "'" & Request.Form("corp") & "', "_
	& "'" & Request.Form("it") & "', "_
	& "'" & Request.Form("admin") & "')"
cnn1.execute strsql

set cnn1=nothing

'Response.Write strsql

Response.Redirect "usrlist.asp"
			
%>