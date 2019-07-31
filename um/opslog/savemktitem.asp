<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%

mktid = Request.Form("mktid")
date1=Request.Form("date")
action=Request.Form("action")
comment=Request.Form("comment")
fcomment=Request.Form("fcomment")
fdate=Request.Form("fdate")
faction=Request.Form("faction")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")

strsql = "insert mkt_progressitems (mid, date, action, comments, followupdate, followup, fcomment) values (" & mktid& ",'" &date1& "', '" & action & "', '" &comment& "', '" & fdate & "', '" & faction & "', '"&fcomment&"')"
response.write mktid
'response.end
cnn1.execute strsql

set cnn1=nothing

tmpMoveFrame =  "parent.location.href = " & Chr(34) & _
				  "mktview.asp?mkid="& mktid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>