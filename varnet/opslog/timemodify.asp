<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
action=Request.Form("modify")
value=Request.Form("value")
value=Formatcurrency(value, 2)
user="ghnet\"&Session("login")
if (Request.Form("job")="") then
	Response.Redirect  "timedetail.asp"
end if
job=Request.Form("job")

id=Request.Form("id")
temp=Request.Form("date")
description=Request.Form("description")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")

	

if (description="") then
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT description FROM [Job Log] where([Entry id]='"& job &"')"
 	rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
'	Response.Write(job)
	description=rst2("description")
'	Response.Write(description)
end if
if(action = "Save") then
    if (Request.Form("hrs")=0) then
		Response.Redirect  "timedetail.asp?day="&temp&"&job="&job
	end if
	
	
		
	strsql = "Insert into Times (date, jobno, description, hours, overt, expense, value, matricola) "_
	& "values ("_
	& "'" & Request.Form("date") & "', "_
	& "'" & Request.Form("job") & "', "_
	& "'" & description & "', "_
	& "'" & Request.Form("hrs") & "', "_
	& "'" & Request.Form("ot") & "', "_
	& "'" & Request.Form("exp") & "', "& value &" ,"_
	& "'" & user & "')"

'Response.Write strsql	
cnn1.execute strsql


else		
id=Request.Form("id")
if (Request.Form("hrs")=0) then
	Response.Redirect  "timedetail.asp?day="&temp&"&id="&id
end if	
strsql = "Update Times Set date='" & Request.Form("date") & "', jobno='" & Request.Form("job") & "', description='" & description & "', hours='" & Request.Form("hrs") & "', overt='" & Request.Form("ot") & "', expense='" & Request.Form("exp") & "', value=" & value & " where (matricola='"& user &"' and id='"& id &"')"

'Response.Write strsql
cnn1.execute strsql

'set cnn1=nothing

end if
tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "time.asp" & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>