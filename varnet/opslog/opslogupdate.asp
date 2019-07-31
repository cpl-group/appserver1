<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
email=Request.Form("email")
job=Request.Form("job")
entry_type=Request.Form("entrytype")
refnum=Request.Form("refnum")
cid=Request.Form("cid")
cname=Trim(Request.Form("cname"))
reqname=Request.Form("reqname")
reqphone=Request.Form("reqphone")
refby=Request.Form("refby")
customerphone=Request.Form("customerphone")
customerfax=Request.Form("customerfax")
floorroom=Request.Form("floorroom")
recdate=Request.Form("recdate")
enddate=Request.Form("enddate")
schdate=Request.Form("stdate")
description=Request.Form("description")
enteredby=Request.Form("EnteredBy")
manager=Request.Form("mid")
m=Request.Form("mid")
status=Request.Form("status")
percentcomp=Request.Form("percentcomp")
billdate=Request.Form("billdate")
comments=Request.Form("comments")
amt=Request.Form("amt")
jobtype=Request.Form("cost")
secamt=Request.Form("secamt")
seccost=Request.Form("seccost")
prob=Request.Form("probability")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

strsql = "Update [job log] Set [entry type]='" & entry_type & "', customer='" & cid & "', [contact name]='" & cname & "', [requested by name]='" &reqname & "', [requested by phone]='" & reqphone & "', [referred by]='" & refby & "', ChgOrderRefNum='" & refnum & "', [phone number]='" & customerphone & "', [fax number]='" &customerfax & "', [floor/room]='" & floorroom & "', [recording date]='" & recdate & "',  [requested target date]='" & enddate & "', [scheduled date]='" & schdate & "', description='" &description & "', [Entered By]='" & enteredby & "',[manager]='" & manager & "', [current status]='" & status & "', [% completed]='" & percentcomp & "', billdate='" & billdate & "', comments='" &comments & "', jobtype='" &jobtype & "', amt='" &amt & "', secamt='" &secamt & "',sectype='" &seccost & "',email='" &email & "',probability='" &prob & "' where ([entry id]='"& job &"')"

'strsql2 = "Update customers Set [companyname]='" & customer &"' where (customerid='"& cid &"')"


'response.write strsql
'response.end
'Response.Write customer
cnn1.execute strsql
'cnn1.execute strsql2
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "opslogview.asp?job="& job & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>