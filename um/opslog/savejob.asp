<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
percentcomp=Request.Form("percentcomp")
billdate=Request.Form("billdate")
comments=Request.Form("comments")
amt=Request.Form("amt")
jobtype=Request.Form("cost")
secamt=Request.Form("secamt")
seccost=Request.Form("seccost")
email=Request.Form("email")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

strsql = "insert [job log] ([entry type],customer,[contact name],[requested by name],[requested by phone],[referred by],ChgOrderRefNum,[phone number],[fax number], [floor/room], description,[Entered By],manager,[current status],[% completed],billdate,comments,jobtype,amt,[requested target date],[scheduled date],secamt,sectype,email)values ('" & type1 & "','" & customer & "', '" & contactname & "', '" &reqname & "', '" & reqphone & "', '" & refby & "', '" & refnum & "', '" & customerphone & "','" &customerfax & "','" & floorroom & "', '" & description & "', '" & enteredby & "','" & manager& "', '" & status & "', '" & percentcomp & "', '" & billdate & "', '" & comments & "','" & jobtype & "','" & amt & "','" &enddate & "','" & schdate & "','" & secamt& "','" & seccost& "','" & email& "')"


'response.write strsql
'response.end
cnn1.execute strsql

strsql="exec sp_newjob"
cnn1.execute strsql
Set rst1 = Server.CreateObject("ADODB.recordset")
strsql="select max([entry id])as job from [job log]"
rst1.Open strsql, cnn1, 0, 1, 1
job=rst1("job")
'response.write job
'response.end
cnn1.execute strsql
set cnn1=nothing

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "opslogview.asp?job="& job & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>