<%@Language="VBScript"%>
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"

if Request.Form("login")= "delete" then

	strsql = "Delete from employees where employeeid=" & Request.Form("id")

else

	strsql = "Update employees Set login='" & Request.Form("login") & "', password='" & Request.Form("password") & "',name='" & Request.Form("name")& "', um=" & Request.Form("um") & ", eri=" & Request.Form("eri") & ", opslog=" & Request.Form("opslog")& ", ts=" & Request.Form("ts") & ", corp=" & Request.Form("corp") & ", it=" & Request.Form("it") & ", admin=" & Request.Form("admin") & " where employeeid=" & Request.Form("id")

end if

cnn1.execute strsql

set cnn1=nothing

'Response.Write strsql


Response.Redirect "usrlist.asp"
			
%>
