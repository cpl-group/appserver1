<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

Response.Write Request.Form("tenantname")	
	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")

	Dim dEri_Base_Month
	Dim intLast_Sur_Kwh
	Dim dLast_sur_KW 
	Dim intBase_Hrs
	
	intLast_Sur_Kwh=Request.Form("last_sur_kwh")
	dEri_Base_Month= Request.Form("eri_base_month")
	dLast_sur_KW = Request.Form("last_sur_kw")
	intBase_Hrs	=Request.Form("base_hours")

	If dEri_Base_Month = "" then
		dEri_Base_Month = "0"
	End If
	If intLast_Sur_Kwh = "" then
		intLast_Sur_Kwh = "0"
	End If
	If dLast_sur_KW= "" Then 
		dLast_sur_KW = "0"
	End If
	If intBase_Hrs = "" Then
		intBase_Hrs = "0"
	End IF
	


	strsql = "Insert into Tenant_info (bldg_no,tenant_no,tenantname,effective_date,lease_exp_date,move_out_date, " & _
										" eri_base_date,sqft,eri_base_month,ccy,ccm,last_sur_kwh,last_sur_kw,bldg_rate," & _ 
										" base_hours,cost_sqft,notes,floor,online) " _
	& "values ("_
	& "'" & Request.Form("bldg_no") & "', " _
	& "'" & Request.Form("tenant_no")& "', "_
	& "'" & Request.Form("tenantname") & "', "_
	& "'" & Request.Form("effective_date") & "', "_
	& "'" & Request.Form("lease_exp_date") & "', "_
	& "'" & Request.Form("move_out_date") & "', "_
	& "'" & Request.Form("eri_base_date") & "', "_
	& Request.Form("sqft") & ", "_
	& dEri_Base_Month & ", "_
	& Request.Form("ccy") & ", "_
	& Request.Form("ccm")& ", "_
	& intLast_Sur_Kwh & ", "_
	& dLast_sur_KW & ", '"_
	& Request.Form("bldg_rate") & "', "_
	& intBase_Hrs & ", "_ 
	& Request.Form("cost_sqft")& ", '"_ 
	& Request.Form("notes") & "', '" _
	& Request.Form("floor") & "', " 
	
	if Request.Form("Online") = True then
		strsql = strsql & "1)"
	else
		strsql = strsql & "0)"
	End If
	

cnn1.execute strsql

set cnn1=nothing


tmpMoveFrame =  "parent.frames.title.location = " & Chr(34) & _
                  "title.asp?bldg=" & Request.Form("bldg_no") & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf

Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf
			
%>
