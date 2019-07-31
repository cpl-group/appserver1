<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
value=Request.Form("value")
value=Formatcurrency(value, 2)
user=Request.Form("user")
if (Request.Form("job")="") then
	Response.Redirect  "timedetail.asp"
end if
job=Request.Form("job")

id=Request.Form("id")
temp=Request.Form("date")
description=Request.Form("description")
customer=Request.Form("customer")
contact=Request.Form("contact")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")

	

if (description="") then
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT description FROM [Job Log] where([Entry id]='"& job &"')"
 	rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
	description=rst2("description")
end if

	strsql = "Insert into invoice_submission (date, jobno, description, hours, overt, expense, value,invoice_date, invoice_comment, matricola) "_
	& "values ("_
	& "'" & Request.Form("date") & "', "_
	& "'" & Request.Form("job") & "', "_
	& "'" & description & "', "_
	& "'" & Request.Form("hrs") & "', "_
	& "'" & Request.Form("ot") & "', "_
	& "'" & Request.Form("exp") & "', "& value &" ,"_
	& "'" & Request.Form("invday") & "',"_
	& "'" & Request.Form("comment") & "',"_
	& "'ghnet\" & user & "')"

cnn1.execute strsql

set cnn1=nothing
tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "corpmain.asp?description="& Request.Form("des") &"&day="& Request.Form("invday") &"&job="& job & "&customer="& customer &"&contact="& contact & chr(34) & vbCrLf 


Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>