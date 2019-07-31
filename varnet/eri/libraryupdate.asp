<%@Language="VBScript"%>
<%
temp=Request.Form("temp")
item=Request.Form("item")
dir=Request.Form("dir")

'Response.Write temp&"@ "&item&"@ "&dir
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

strsql = "Update tblSurveyLib Set amps='" & Request.Form("amps") & "', volt='" & Request.Form("volt") & "', ph='" & Request.Form("ph") & "', pf='" & Request.Form("pf") & "', watt='" & Request.Form("watt") & "' where type='"& Request.Form("type1") &"' and description='"& Request.Form("description") &"'"

'Response.Write strsql
cnn1.execute strsql

set cnn1=nothing
tmpMoveFrame =  "parent.frames.main.location = " & Chr(34) & _
				  "library.asp?type1=" & temp & _  
				  "&item="&item&"&dir=" & dir & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
