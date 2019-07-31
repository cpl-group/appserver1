<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
action=Request.Form("modify")
v=Request.Form("v")
value=Formatnumber(v, 2)
billh=trim(Request.Form("billh"))

if billh="" then
	billh=0
end if
billable=Formatnumber(billable, 2)
id=Request.Form("id")
temp=Request.Form("date")
'response.write(temp)
job=Request.Form("job")
user="ghnet\"&Session("login")
if (Request.Form("description")="") then
	Response.Redirect  "opstimesheet.asp?day="&temp&"&id="&id
end if
if (Request.Form("hours")=0) then
'	Response.Redirect  "opstimesheet.asp?day="&temp&"&id="&id
end if

description=Request.Form("description")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")
	
strsql = "Update invoice_submission Set date='" & Request.Form("date") & "', description='" & description & "', hours='" & Request.Form("hours") & "', hours_bill='" & billh & "', overt='" & Request.Form("overt") & "',  expense='" & Request.Form("expense") & "', value=" & value &" where ( id='"& id &"')"

cnn1.execute strsql
'response.write strsql
set cnn1=nothing

tmpMoveFrame =  "parent.frames.oplog.location = " & Chr(34) & _
				  "timesheetsearch.asp?job="&job & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
'Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
flag=1
tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "timesheetmain.asp?job="&job &"&flag="&flag& chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


%>