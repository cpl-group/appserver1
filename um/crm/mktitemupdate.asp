<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%


mktid = Request.Form("mkid")
date1=Request.Form("date")
action=Request.Form("action")
comment=Request.Form("comment")
fcomment=Request.Form("fcomment")
fdate=Request.Form("fdate")
faction=Request.Form("faction")
id=Request.Form("key")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"Intranet")

strsql = "Update mkt_progressitems Set date='" & date1 &"',action='" & action &"', comments='" & comment&"', followupdate='"&fdate&"',followup='" & faction&"', fcomment='" &fcomment &"' where id='"& id&"'"
'response.write strsql
'response.end
cnn1.execute strsql
	


set cnn1=nothing
tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "mktview.asp?mkid="& mktid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>