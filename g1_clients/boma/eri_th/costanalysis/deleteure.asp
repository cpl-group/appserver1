<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
entryid=Request.QueryString("eid")
userid = Request.QueryString("userid")
bldg=Request.QueryString("bldg")
curyear=Request.QueryString("year")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
sql = "delete tblRPentries where id=" & entryid
cnn1.execute sql
set cnn1=nothing
urltemp = "unreported.asp?bldg=" & bldg & "&year=" & curyear & "&userid=" & userid &"&action=new"
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				urltemp & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>