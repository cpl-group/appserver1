<%@Language="VBScript"%>
<%
temp=Request.Form("temp")
item=Request.Form("item")
dir=Request.Form("dir")
'Response.Write temp&" "&item&" "&dir
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")

	strsql = "Insert into tblSurveyLib (type, description, amps, volt, ph, pf, watt,monthfactor, adjfactor) "_
	& "values ("_
	& "'" & Request.Form("type1") & "', "_
	& "'" & Request.Form("description") & "', "_
	& "'" & Request.Form("amps") & "', "_
	& "'" & Request.Form("volt") & "', "_
	& "'" & Request.Form("ph") & "', "_
	& "'" & Request.Form("pf") & "', "_
	& "'" & Request.Form("watt")& "', "_
	& "'" & Request.Form("mf") & "', "_
	& "'" & Request.Form("adj") & "')"
'cnn1.execute strsql
'Response.Write strsql
set cnn1=nothing
tmpMoveFrame =  "parent.frames.main.location = " & Chr(34) & _
				"library.asp?type1=" & temp & _  
				"&item="&item&"&dir=" & dir & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
		
%>
