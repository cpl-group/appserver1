<%@Language="VBScript"%>
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

strsql = "Insert into tenant_history (tenant_no, date_event, rate, fuel, charge, note,sur_kwh,sur_kw,sqft) values ('" & Request.Form("tenant_no") & "', '" & Request.Form("date_event") & "', " & Request.Form("rate")  & ", " & Request.Form("fuel")& ", " & CCur(Formatnumber(Request.Form("charge"))) & ", '" & Request.Form("note") &"','" & Request.Form("sur_kwh") &"','" & Request.Form("sur_kw") &"',Request.Form("sqft") &"')"
'response.write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing

'Response.Write strsql


Response.Redirect "info.asp?qcatnr=" & Request.Form("Tenant_no")
		
%>
