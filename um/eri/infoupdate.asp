<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")


strsql = "Update tenant_history Set date_event='" & Request.Form("date_event") & "', rate=" & (Left(Request.Form("rate"),(instr(Request.Form("rate"),"%"))-1))/100  & ", fuel=" & (Left(Request.Form("fuel"),(instr(Request.Form("fuel"),"%"))-1))/100 & ", charge=" & CCur(Formatnumber(Request.Form("charge"))) & ", note='" & Request.Form("note") &"', sur_kwh='" & Request.Form("sur_kwh") &"', sur_kw='" & Request.Form("sur_kw") &"',sqft='" & Request.Form("sqft") &"' where id='" & Request.Form("id") & "'"
Response.Write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing




Response.Redirect "info.asp?qcatnr=" & Request.Form("Tenant_no")
			
%>
