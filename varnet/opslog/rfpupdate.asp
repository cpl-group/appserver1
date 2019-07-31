<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%

rfp=Request.Form("rfp")
entry_type=Request.Form("entrytype")
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
prob=Request.Form("prob")
comments=Request.Form("comments")
amt=Request.Form("amt")
jobtype=Request.Form("cost")
amt2=Request.Form("amt2")
jobtype2=Request.Form("cost2")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")

strsql = "Update rfplog Set [entry type]='" & entry_type & "', customer=convert(int," & cid & "), [contact name]='" & cname & "', [requested by name]='" &reqname & "', [requested by phone]='" & reqphone & "', [referred by]='" & refby & "', [phone number]='" & customerphone & "', [fax number]='" &customerfax & "', [floor/room]='" & floorroom & "', [recording date]='" & recdate & "',estcdate='" & enddate & "', [scheduled date]='" & schdate & "', description='" &description & "', [Entered By]='" & enteredby & "',salesmanager='" & manager & "', [current status]='" & status & "', [% completed]='" & percentcomp & "', comments='" &comments & "', probability='" &prob & "',proposal='" &jobtype & "', amt='" &amt & "',proposal2='" & jobtype2 & "', amt2='" &amt2 & "' where ([entry id]='"& rfp &"')"

'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing
if status="Proposal Accepted" then
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "opslogview.asp?job="& rfp & chr(34) & vbCrLf 
else
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "rfpview.asp?rfp="& rfp & chr(34) & vbCrLf 
end if
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf


 

%>