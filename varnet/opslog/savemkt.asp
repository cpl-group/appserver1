<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
cust=Request.Form("cust")
sit=Request.Form("sit")
eb=Request.Form("eb")


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

strsql = "insert mktlog (contact,situation,enteredby)values ('" & cust& "', '" &sit & "','" & eb & "' )"
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