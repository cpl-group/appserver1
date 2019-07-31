<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
cust=Request.Form("cust")
sit=Request.Form("sit")
eb=Request.Form("eb")
sm=Request.Form("manager")


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"Intranet")

strsql = "insert mktlog (contact,situation,enteredby,salesmanager)values ('" & cust& "', '" &sit & "','" & eb & "',"& sm & ")"
'response.write strsql
'response.end
cnn1.execute strsql


strsql = "select max (id) as id from mktlog"
rst.Open strsql, cnn1, 0, 1, 1


if not rst.eof then
	mkid = rst("id")
end if

set cnn1=nothing



tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "mktview.asp?mkid="& mkid&"&cust="&cust& chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
