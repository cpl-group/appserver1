
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
bldg=Request.Form("bldgnum")
addr=Request.Form("address")
city=Request.Form("city")
state=Request.Form("state")
phone=Request.Form("phone")
fax=Request.Form("fax")
zip=Request.Form("zip")
sqft=Request.Form("sqft")
fl=Request.Form("fl")
c1=Request.Form("name1")
cp1=Request.Form("phone1")
c2=Request.Form("name2")
cp2=Request.Form("phone2")
c3=Request.Form("name3")
cp3=Request.Form("phone3")
bid=request.form("bid")
cid=request("cid")
submit=request.form("submit")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getconnect(0,0,"engineering")

if trim(submit)="Delete" then
'	cnn1.execute "DELETE FROM facilityinfo WHERE id="&bid
	set cmd  = server.createobject("ADODB.command")
	cmd.Activeconnection = cnn1
	cmd.commandType = adCmdStoredProc
	cmd.CommandText = "DELETE_BLDG"
    Set prm = cmd.CreateParameter("BLDG", adInteger, adParamInput)
    cmd.Parameters.Append prm
	cmd.Parameters("BLDG") = bid
	cmd.execute
	tmpMoveFrame = "document.location = ""managebldg.asp?cid=" & cid & """"
else
	strsql = "update facilityinfo set bldgname='" & bldg& "',address='" & addr & "',city='" & city & "',state='" & state & "',zip='" & zip & "',sqft='" &sqft & "' where id='" &bid&"'"
	rst1.open "SELECT labelid FROM facilityinfo f INNER JOIN nodes n ON n.nodeid=f.nodeid WHERE f.id='" &bid&"'", cnn1
	if not rst1.eof then
			cnn1.execute "UPDATE label SET name='"&addr&"' WHERE id="&trim(rst1("labelid"))
	end if
	rst1.close
	
	'response.write strsql
	'response.end
	cnn1.execute strsql

if Request.Form("srcfile") <> "" then
  tmpMoveFrame = 	"document.location = ""editbldg.asp?id="& bid & """"
else
	tmpMoveFrame =  "document.location = ""updatebldg.asp?id="& bid & "&action='updated'"""
end if
end if
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 


set cnn1=nothing

%>