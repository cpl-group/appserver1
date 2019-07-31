
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%

dlast=Request.Form("dlast")
elect=Request.Form("elect")
fid=Request.Form("fid")
rid=Request.Form("rid")
bldg=Request.Form("bldg")
fxid=Request.Form("fxid")
comments=Request.Form("comments")
lid=Request.Form("lid")
blast=Request.Form("blast")
belect=Request.Form("belect")
fixqty=Request.Form("fixqty")
submit=Request.Form("submit")
ehw=Request.Form("ehw")
tp=Request.Form("tp")


Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

strsql="select max(datelastchanged) as date1 ,max(bdatelastchanged) as bdate1 from lamping_sch where fid='"&fxid&"'"

rst1.Open strsql, cnn1, 0, 1, 1
'date1=rst1("date1")
response.write datediff("d",rst1("date1"),dlast)
if trim(submit)="Delete" then
	strsql="DELETE FROM fixtures WHERE id="&fxid
else
	if not isnull(rst1("bdate1")) or isnull(rst1("date1")) then
		if datediff("d",rst1("date1"),dlast)<0 or datediff("d",rst1("date1"),dlast)=0 then
			cnn1.execute "DELETE FROM lamping_sch WHERE (datediff(dd, datelastchanged, '"&dlast&"')<0 or datediff(dd, datelastchanged, '"&dlast&"')=0) and fid='"&fxid&"'"
		end if
	end if
	strsql="insert lamping_sch (fid,datelastchanged,comments,electrician,bdatelastchanged,belectrician) values ('" & fxid& "', '" &dlast & "', '" &comments& "', '" &elect& "', '" &blast& "', '" &belect& "')"
end if
'response.write "<br>"&strsql
'response.end
'if trim(dlast)<>trim(rst1("date1")) then
'  response.write trim(dlast) & " " & trim(rst1("date1"))
'end if
'response.write fxid
'response.end
cnn1.execute strsql
strsql = "update fixtures set fixtureqty='"&fixqty&"' ,est_hr_wk='"&ehw&"',dlc='"&dlast&"',type='"&tp&"' where id='"&fxid&"'"

cnn1.execute strsql
'if trim(submit)="Delete" then
	tmpMoveFrame =  "document.location = ""fixsearch.asp?id="& fxid &"&bldg="&bldg&"&room="&room&"&rid="&rid&"&floor="&floor&"&fid="&fid&""""
'else
'	tmpMoveFrame =  "document.location = ""fixtureinfo.asp?id="& fxid &"&bldg="&bldg&"&room="&room&"&rid="&rid&"&floor="&floor&"&fid="&fid&""""
'end if

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
rst1.close
set cnn1=nothing
%>