<%@Language="VBScript"%>
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
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

strsql = "Update tblSurveyItem Set type='" & Request.Form("type1") & "', description='" & Request.Form("description2") &"', amps=" & Request.Form("amps") & ", volt=" & Request.Form("volt") & ", ph=" & Request.Form("ph") & ", pf=" & Request.Form("pf") & ", watt=" & Request.Form("watt") & ", qty=" & Request.Form("qty") & ", monthfactor=" & Request.Form("mf") & ", adjfactor=" & Request.Form("adj") & ", houron=" & Request.Form("hon") & ", houroff=" & Request.Form("hoff") & ", intense=" & ie & ", base=" & bho & " where id=" & key

id=Request.Form("id")
survey_id=id

'Response.Write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing
tmpMoveFrame =  "parent.frames.details.location = " & Chr(34) & _
				  "surveyitems.asp?tenant_no=" & tenant_no & _  
				  "&orderno="&orderno&"&id=" & id & _  
				  "&survey_id="&survey_id& _                
				  "&location=" & location & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


			
%>
