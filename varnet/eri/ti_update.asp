<%@Language="VBScript"%>
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

strsql = "Update tenant_info Set tenant_no='" & Request.Form("tenant_no") & "', tenantname='" & Request.Form("tenantname") & "', effective_date='" & Request.Form("effective_date") &"', lease_exp_date='" & Request.Form("lease_exp_date") & "', move_out_date='" & Request.Form("move_out_date") & "', sqft=" & Request.Form("sqft") & ", eri_base_month=" & CCur(Formatnumber(Request.Form("eri_base_month"))) & ", eri_base_date='" & Request.Form("eri_base_date") & "',ccy=" & CCur(Formatnumber(Request.Form("ccy"))) & ", ccm=" & CCur(Formatnumber(Request.Form("ccm"))) & ", last_sur_kwh=" & Request.Form("last_sur_kwh") & ", last_sur_kw=" & Request.Form("last_sur_kw") & ", bldg_rate='" & Request.Form("bldg_rate") & "', base_hours=" & Request.Form("base_hours") & ", cost_sqft=" & CCur(Formatnumber(Request.Form("cost_sqft"))) & ", notes='" & Request.Form("notes") & "' where tenant_no='" & Request.Form("tenant_id") & "'"

cnn1.execute strsql

set cnn1=nothing




Response.Redirect "piclist.asp?bldg=" & Request.Form("bldg_no")
			
%>
