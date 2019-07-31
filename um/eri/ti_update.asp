<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")

strsql = "Update tenant_info " & _
		 "Set tenant_no='" & Request.Form("tenant_no") & "', tenantname='" & Request.Form("tenantname") & "', "  & _
			"effective_date='" & Request.Form("effective_date") & "', " & _
			"lease_exp_date='" & Request.Form("lease_exp_date") & "', " & _
			"move_out_date='" & Request.Form("move_out_date") & "', " & _
			"sqft=" & Request.Form("sqft") & ", " & _
			"eri_base_month=" & CCur(Formatnumber(Request.Form("eri_base_month"))) & ", " & _
			" eri_base_date='" & Request.Form("eri_base_date") & "', " & _
			" ccy=" & CCur(Formatnumber(Request.Form("ccy"))) & ", " & _
			" ccm=" & CCur(Formatnumber(Request.Form("ccm"))) & ", " & _
			" last_sur_kwh=" & Request.Form("last_sur_kwh") & ", " & _
			" last_sur_kw=" & Request.Form("last_sur_kw") & ", " & _
			" bldg_rate='" & Request.Form("bldg_rate") & "', " & _
			" base_hours=" & Request.Form("base_hours") & ", " & _
			" cost_sqft=" & CCur(Formatnumber(Request.Form("cost_sqft"))) & ", notes='" & Request.Form("notes") & "', " & _
			" floor ='" & Request.Form("floor") & "', " 
			If  Request.Form("Online")="on" then
				strsql= strsql & " online =1 "  
			else
				strsql= strsql & " online =0 "  
			end if	
			
		strsql = strsql & " where tenant_no='" & Request.Form("tenant_no") & "'"
			

'response.write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing




Response.Redirect "info.asp?qcatnr=" & Request.Form("tenant_no")
			
%>
