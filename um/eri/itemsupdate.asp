<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<%
key=Request.Form("key")	

ie=Request.Form("ie")
if(ie="on") then
	ie=1
else
	ie=0
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

strsql = "Update tblSurveyItem Set type='" & Request.Form("type1") & "', description='" & Request.Form("description2") &"', amps=" & Request.Form("amps") & ", volt=" & Request.Form("volt") & ", ph=" & Request.Form("ph") & ", pf=" & Request.Form("pf") & ", watt=" & Request.Form("watt") & ", qty=" & Request.Form("qty") & ", monthfactor=" & Request.Form("mf") & ", adjfactor=" & Request.Form("adj") & ", houron=" & Request.Form("hon") & ", houroff=" & Request.Form("hoff") & ", intense=" & ie & ", base=" & bho & " where id=" & key

id=Request.Form("id")
survey_id=id

'Response.Write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing
tmpMoveFrame =  "parent.frames.details.location = ""surveyitems.asp?tenant_no=" & tenant_no & "&orderno="&orderno&"&id=" & id & "&survey_id="&survey_id& "&location=" & location & "&xscroll="&request("xscroll")&"&yscroll="&request("yscroll")&""""

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


			
%>
