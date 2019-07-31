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
logo=Request.Form("logourl")


Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_lighting")


'strsql = "INSERT nodes (label, clientid) values ('"&Clients&"') "
strsql = "insert clients (corp_name,address,city,state,zip,contact,contactphone,logo)values ('" & bldg& "', '" & addr & "', '" & city & "', '" & state & "', '" & zip & "','" & c1 & "', '" & cp1 & "', '" & logo & "')"
'response.write strsql
'response.end
cnn1.execute strsql


sqlstr = "select max(id) as id from clients"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1
maxid=rst1("id")

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "clientview.asp?id="& maxid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

rst1.close
set cnn1=nothing

%>