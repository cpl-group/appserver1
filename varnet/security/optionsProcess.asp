<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim link, sid, target, username, Defaults(30), csid
csid = request("CSID")
sid = split(request("SID"),",")
username = request("username")

dim cmd, cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set cmd = Server.CreateObject("ADODB.command")
cnn1.Open application("cnnstr_security")

rst1.open "SELECT sid, defaultlink FROM tbladdons", cnn1
do until rst1.eof
	Defaults(cINT(rst1("SID"))) = trim(rst1("defaultlink"))
	rst1.movenext
loop
rst1.close


cmd.activeConnection = cnn1
cmd.commandText = "DELETE FROM tbladdonlinks WHERE userid='"&username&"' AND sid in (SELECT sid FROM tbladdons WHERE CSID = "&csid&")"
response.write cmd.commandText&"<BR>"
cmd.execute
dim item, templink, temptarget
for each item in sid
	templink = request("link"&trim(item))
	temptarget = request("target"&trim(item))
	if trim(templink)="" then templink = Defaults(cINT(item))
	if trim(temptarget)="" then temptarget = "self"
	cmd.commandText = "INSERT INTO tbladdonlinks (SID, userid, Link, Target, Active) VALUES ('"&trim(item)&"', '"&username&"', '"&templink&"', '"&temptarget&"', 1)"
	response.write cmd.commandText&"<BR>"
	cmd.execute
next
response.redirect "optionsList.asp?username="&username&"&csid="&csid&"&action=saved"
%>
