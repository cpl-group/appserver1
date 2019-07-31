<%@Language="VBScript"%>
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

strsql = "Update tenant_info Set tenantname='" & Request.Form("tname") & "', sqft=" & Request.Form("sqft") & ", last_sur_kwh=" & Request.Form("last_sur_kwh") & ", last_sur_kw=" & Request.Form("last_sur_kw") & ", base_hours=" & Request.Form("base_hours") & ", notes='" & Request.Form("notes") & "' where tenant_no='" & Request.Form("tenant_id") & "'"

cnn1.execute strsql

set cnn1=nothing

'Response.Write strsql


Response.Redirect "tenant_survey.asp?tenant_no=" & Request.Form("tenant_id") & "&bldg=" & Request.Form("bldg_no")
			
%>
