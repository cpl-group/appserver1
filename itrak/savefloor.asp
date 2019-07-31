
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
floor=Request.Form("floor")
sqft=Request.Form("sqft")
bldg=Request.Form("bldg")
fid=Request.Form("fid")
Submit = Request("Submit")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getconnect(0,0,"engineering")


if trim(submit)="Update" then
	strsql = "UPDATE floor SET floor='" & floor& "', sqft='" & sqft & "' WHERE id="&fid
elseif trim(submit)="Delete" then
'	strsql = "DELETE FROM floor WHERE id="&fid
	set cmd  = server.createobject("ADODB.command")
	cmd.Activeconnection = cnn1
	cmd.commandType = adCmdStoredProc
	cmd.CommandText = "DELETE_FLOOR"
    Set prm = cmd.CreateParameter("FLOOR", adInteger, adParamInput)
    cmd.Parameters.Append prm
	cmd.Parameters("FLOOR") = fid
	cmd.execute
else
	strsql = "insert floor (bldg,floor,sqft)values ('" & bldg& "', '" & floor& "', '" & sqft & "')"
end if

'response.write strsql
'response.end
if trim(submit)<>"Delete" then
	cnn1.execute strsql
end if
tmpMoveFrame =  "location = ""floorsearch.asp?bldg=" & bldg & """"

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

set cnn1=nothing

%>