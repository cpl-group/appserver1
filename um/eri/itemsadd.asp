<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

ie2=Request.Form("ie2")
if(ie2="on") then
	ie2=1
else
	ie2=0
end if
bho=Request.Form("bho")	
if(bho="on") then
	bho=1
else
	bho=0
end if



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")




	strsql = "Insert into tblSurveyItem (type, description,amps, volt, ph, pf, watt, qty,  monthfactor, adjfactor, houron, houroff, intense, base, surveyid) "_
	& "values ("_
	& "'" & Request.Form("type1") & "', "_
	& "'" & Request.Form("description") & "', "_
	&  request.form("amps") & ", "_
	&  Request.Form("volt") & ", "_
	&  Request.Form("ph") & ", "_
	&  Request.Form("pf") & ", "_
	&  Request.Form("watt")& ", "_
	& Request.Form("qty")& ", "_
	& Request.Form("mf") & ", "_
	&  Request.Form("adj") & ", "_ 
	&  Request.Form("hon") & ", "_ 
	&  Request.Form("hoff") & ", "_ 
	 
	&  ie2 & ", "_ 
	&  bho & ", "_ 
	& Request.Form("id")& ")"
	
	
cnn1.execute strsql
'Response.Write strsql
'response.end
set cnn1=nothing



tmpMoveFrame =  "parent.frames.tenant.location = ""survey_detail.asp?tenant_no=" & request("tenant_no") & "&surveyid=" & request("id") & "&xscroll="&request("xscroll")&"&yscroll="&request("yscroll")& "0"""

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
		
%>
