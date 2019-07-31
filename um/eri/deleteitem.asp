<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"Engineering")
id1=request("surveyid")
key= Request("key")

sqlstr = "delete tblsurveyitem where id=" & key & ""
cnn1.Execute sqlstr 
'tmpMoveFrame =  "document.location = ""surveyitems.asp?surveyid="&id1 & "&xscroll="&xscroll&"&yscroll="&yscroll&""""
id=Request.Form("id")
survey_id=id

tmpMoveFrame =  "document.location = ""surveyitems.asp?tenant_no=" & tenant_no & "&orderno="&orderno&"&id=" & id & "&survey_id="&survey_id& "&location=" & location & "&xscroll="&request("xscroll")&"&yscroll="&request("yscroll")&""""
				  
				  
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 				  
%>