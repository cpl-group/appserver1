
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
bldg=Request.Form("bldg")
t=Request.Form("type")
floor=Request.Form("floor")
room=Request.Form("room")
fixqty=Request.Form("fixqty")
lampqty=Request.Form("lampqty")
comments=Request.Form("comments")
fid=Request.Form("id")
balqty=Request.Form("balqty")


Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

strsql = "update fixtures set typeid='" &t & "',bldgnum='" & bldg& "',floor='" & floor & "',room='" &room & "',fixtureqty='" & fixqty& "',lampqty='" &lampqty & "',comments='" &comments & "',ballast_qty='" &balqty & "'  where id='" &fid & "' "
'response.write strsql
'response.end
cnn1.execute strsql




tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "fixture.asp?id="& fid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


set cnn1=nothing

%>