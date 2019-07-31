
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
floor=Request.Form("floor")
fid=Request.Form("fid")
room=Request.Form("room")
rid=Request.Form("rid")
sqft=Request.Form("sqft")
bldg=Request.Form("bldg")
submit=Request.Form("submit")


Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")


if trim(submit) = "Update" then
	strsql = "Update room set room='"&room&"', sqft='"&sqft&"' where id="&rid
elseif trim(submit) = "Delete" then
'	strsql = "DELETE FROM room where id="&rid
	set cmd  = server.createobject("ADODB.command")
	cmd.Activeconnection = cnn1
	cmd.commandType = adCmdStoredProc
	cmd.CommandText = "DELETE_ROOM"
    Set prm = cmd.CreateParameter("ROOM", adInteger, adParamInput)
    cmd.Parameters.Append prm
	cmd.Parameters("ROOM") = rid
	cmd.execute
else
	strsql = "insert room (bldg,room,floor,sqft) values ('" & bldg& "', '" & room & "', '" & fid & "', '" & sqft & "')"
end if
'response.write strsql
'response.end
if trim(submit)<>"Delete" then
	cnn1.execute strsql
end if

sqlstr = "select max(id) as id from room "

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1
maxid=rst1("id")

tmpMoveFrame =  "location = ""roomsearch.asp?bldg=" &bldg&"&floor="&floor&"&fid="&fid&""""

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

rst1.close
set cnn1=nothing

%>