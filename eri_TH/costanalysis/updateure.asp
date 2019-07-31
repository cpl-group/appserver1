<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
entrytype=Request.Form("type")
desc=Request.Form("description")
amt=Request.Form("amt")
period=Request.Form("period")
pid = Request.Form("pid")
bldg=Request.Form("b")
curyear=Request.Form("date1")
entryid = Request.Form("entryid")

if entrytype = 0 and amt >= 0 then 

	amt = amt * -1
	
end if


Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
sql = "update tblRPentries set description='" & desc & "',  amt= " &  clng(amt) & ", period= '" & period & "', type='" & entrytype & "' where id=" & entryid

cnn1.execute sql
set cnn1=nothing
urltemp = "unreported.asp?building=" & bldg & "&date1=" & curyear & "&pid=" & pid &"&action=new"
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				urltemp & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>