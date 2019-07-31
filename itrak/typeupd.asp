
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%

d=Request.Form("description")
manf=Request.Form("manf")
fixc=Request.Form("fixc")
b=Request.Form("ballast")
lqty=Request.Form("lqty")
zip=Request.Form("zip")
lwatts=Request.Form("lwatts")
lcnum=Request.Form("lcnum")
volts=Request.Form("volts")
estLL=Request.Form("estLL")
remarks=Request.Form("remarks")
fid=Request.Form("fid")
blife=Request.Form("blife")
bqty=Request.Form("bqty")
cid=Request.Form("cid")
action=Request.Form("submit")
bldg=Request.Form("bldg")
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")


if trim(action)="Delete" then
	strsql = "delete fixture_types where id='" & fid & "'" 
else
	strsql = "update fixture_types set description='" &d & "',manufacturer='" & manf & "',fix_catalog='" &fixc & "',ballast_type='" & b& "',lamp_qty='" &lqty & "',lamp_watts='" &lwatts & "',lamp_catalog='" & lcnum & "',volts='" & volts& "',remarks='" & remarks& "',est_lamp_life='" & estLL& "',ballast_life='" & blife& "',ballast_qty='" & bqty& "' where id='" & fid & "'"
end if
cnn1.execute strsql
tmpMoveFrame =  "location = ""fixtypesearch.asp?cid=" & cid & "&bldg=" & bldg & """"

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

set cnn1=nothing

%>