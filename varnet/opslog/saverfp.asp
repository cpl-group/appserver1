<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%


type1=Request.Form("type1")
refnum=Request.Form("refnum")
customer=Request.Form("customer")
cid=Request.Form("cid")
contactname=Trim(Request.Form("contactname"))
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
manager=Request.Form("manager")
m=Request.Form("mid")
status=Request.Form("status")
billdate=Request.Form("billdate")
comments=Request.Form("comments")
amt=Request.Form("amt")
jobtype=Request.Form("cost")
prob=Request.Form("prob")
mkid=Request.Form("mkid")
amt2=Request.Form("amt2")
jobtype2=Request.Form("cost2")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")


strsql="insert [job log] ([entry type],customer,[contact name],[requested by name],[requested by phone],[referred by],[phone number],[fax number], [floor/room], description,[Entered By],manager,comments,[requested target date],[scheduled date],[% completed])values ('" &type1 &"','" & customer & "', '" & contactname & "', '" &reqname & "', '" & reqphone & "', '" & refby & "', '" & customerphone & "','" &customerfax & "','" & floorroom & "', '" & description & "', '" & enteredby & "','" & manager& "', '" & comments & "','" &enddate & "','" & schdate & "','0')"
cnn1.execute strsql

Set rst1 = Server.CreateObject("ADODB.recordset")
strsql="select max([entry id])as rfp from [job log]"
rst1.Open strsql, cnn1, 0, 1, 1
rfp=rst1("rfp")

strsql = "insert rfplog ([entry id],[entry type],customer,[contact name],[requested by name],[requested by phone],[referred by],[phone number],[fax number], [floor/room], description,[Entered By],salesmanager,[current status],probability,comments,proposal,amt,estcdate,[scheduled date],mkid,proposal2,amt2)values ('" & rfp & "','" & type1 & "','" & customer & "', '" & contactname & "', '" &reqname & "', '" & reqphone & "', '" & refby & "', '" & customerphone & "','" &customerfax & "','" & floorroom & "', '" & description & "', '" & enteredby & "','" & manager& "', '" & status & "', '" & prob & "', '" & comments & "','" & jobtype & "','" & amt & "','" &enddate & "','" & schdate & "','" & mkid & "','" & jobtype2 & "','" & amt2 & "')"
cnn1.execute strsql

'response.write strsql


strsql="exec sp_newrfp"
cnn1.execute strsql

'response.write job
'response.end
set cnn1=nothing

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "rfpview.asp?rfp="& rfp & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>