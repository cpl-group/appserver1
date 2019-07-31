<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")

strsql = "Insert into tenant_history (tenant_no, date_event, rate, fuel, charge, note,sur_kwh,sur_kw,sqft) values ('" & Request("tenant_no") & "', '" & Request.Form("date_event") & "', " & Request.Form("rate")  & ", " & Request.Form("fuel")& ", " & CCur(Formatnumber(Request.Form("charge"))) & ", '" & Request.Form("note") &"','" & Request.Form("sur_kwh") &"','" & Request.Form("sur_kw") &"','" &Request.Form("sqft") &"')"
'response.write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing

'Response.Write strsql


Response.Redirect "info.asp?qcatnr=" & Request.Form("Tenant_no")
		
%>
