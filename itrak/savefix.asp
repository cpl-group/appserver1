
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
bldg=Request.Form("bldg")
t=Request.Form("type")
floor=Request.Form("floor")
fid=Request.Form("fid")
room=Request.Form("room")
rid=Request.Form("rid")
fixqty=Request.Form("fixqty")
comments=Request.Form("comments")
esthwk=Request.Form("esthwk")
dlast=Request.Form("dlast")
ft=Request.Form("ft")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

strsql = "insert fixtures (typeid,bldgnum,room,fixtureqty,comments,est_hr_wk,dlc,type) values ('" & t& "', '" &bldg & "', '" &rid & "', '" & fixqty& "','" &comments & "','" &esthwk & "','"&dlast&"','"&ft&"')"
'response.write strsql
'response.end
cnn1.execute strsql
sqlstr = "select max(id) as id from fixtures "

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1
maxid=rst1("id")

strsql="insert lamping_sch(fid,datelastchanged,comments,bdatelastchanged) values ('"&maxid&"','"&dlast&"','original setup','"&dlast&"')"
cnn1.execute strsql

tmpMoveFrame =  "location = ""fixsearch.asp?bldg="&bldg&"&room="&room&"&rid="&rid&"&floor="&floor&"&fid="&fid&""""

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


set cnn1=nothing

%>