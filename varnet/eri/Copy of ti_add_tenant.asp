<%@Language="VBScript"%>
<%

Response.Write Request.Form("tenantname")	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

	strsql = "Insert into Tenant_info (bldg_no,tenant_no,tenantname,effective_date,lease_exp_date,move_out_date,sqft,eri_base_month,ccy,ccm,last_sur_kwh,last_sur_kw,bldg_rate,base_hours,cost_sqft,notes) "_
	& "values ("_
	& "'" & Request.Form("bldg_no") & "', "_
	& "'" & Request.Form("tenant_no")& "', "_
	& "'" & Request.Form("tenantname") & "', "_
	& "'" & Request.Form("effective_date") & "', "_
	& "'" & Request.Form("lease_exp_date") & "', "_
	& "'" & Request.Form("move_out_date") & "', "_
	& Request.Form("sqft") & ", "_
	& Request.Form("eri_base_month") & ", "_
	& Request.Form("ccy") & ", "_
	& Request.Form("ccm")& ", "_
	& Request.Form("last_sur_kwh")& ", "_
	& Request.Form("last_sur_kw")& ", '"_
	& Request.Form("bldg_rate") & "', "_
	& Request.Form("base_hours") & ", "_ 
	& Request.Form("cost_sqft")& ", '"_ 
	& Request.Form("notes")& "')"
	
cnn1.execute strsql

set cnn1=nothing


tmpMoveFrame =  "parent.frames.title.location = " & Chr(34) & _
                  "title.asp?bldg=" & Request.Form("bldg_no") & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf

Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf
			
%>
