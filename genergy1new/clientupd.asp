<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
bldg=Request.Form("bldgnum")
addr=Request.Form("address")
city=Request.Form("city")
state=Request.Form("state")
phone=Request.Form("phone")
fax=Request.Form("fax")
zip=Request.Form("zip")
sqft=Request.Form("sqft")
fl=Request.Form("fl")
c1=Request.Form("name1")
cp1=Request.Form("phone1")
bid=request.form("bid")
logo=request.form("logourl")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_lighting")

strsql = "update clients set corp_name='" & bldg & "',address='" & addr & "',city='" & city & "',state='" & state & "',zip='" & zip & "',contact='" & c1 & "',contactphone= '" & cp1 & "', logo='"& logo &"' where id='" &bid&"'"
'response.write strsql
'response.end
cnn1.execute strsql

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "clientview.asp?id="& bid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


set cnn1=nothing

%>