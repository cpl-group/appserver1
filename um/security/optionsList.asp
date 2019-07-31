<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim csid, usrname, action, label
csid = request.querystring("csid")
usrname = request.querystring("usrname")
action = request.querystring("action")
label = request.querystring("label")

dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
%>
<html>
<head><title>Options Setup</title>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>
<body onload="<%if trim(action)="saved" then response.write "window.close()"%>" link="#000099" vlink="#000099" alink="#000099">
<form name="options" method="post" action="optionsProcess.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader"><%=label%></span></td>
</tr>
</table>
<table cellpadding="3" cellspacing="0" border="0">
<%
dim UserOptions(100,3)
rst1.open "SELECT * FROM tblAddonLinks al WHERE userid='"&usrname&"'", cnn1
do until rst1.eof
	UserOptions(cINT(rst1("SID")),0) = true
	UserOptions(cINT(rst1("SID")),1) = trim(rst1("link"))
	UserOptions(cINT(rst1("SID")),2) = trim(rst1("target"))
	rst1.movenext
loop
rst1.close

rst1.open "SELECT * FROM tblAddons WHERE CSID="&csid, cnn1
do until rst1.eof
	%><tr><td valign="top"><input type="checkbox" name="SID" value="<%=rst1("SID")%>" <%if UserOptions(rst1("SID"),0) then response.write "CHECKED"%>></td><td valign="top"><a href="#" onclick="window.open('changeOptions.asp?SID=<%=rst1("SID")%>&username=<%=usrname%>&link='+document.forms.options.link<%=rst1("SID")%>.value+'&target='+document.forms.options.target<%=rst1("SID")%>.value, '<%=rst1("SID")%>','scrollbars=no, width=320, height=50, toolbar=no');"><%=rst1("label")%></a><input type="hidden" name="link<%=rst1("SID")%>" value="<%=UserOptions(cINT(rst1("SID")),1)%>"><input type="hidden" name="target<%=rst1("SID")%>" value="<%=UserOptions(cINT(rst1("SID")),2)%>"></td></tr><%
	rst1.movenext
loop
rst1.close
%></table>
<input type="hidden" name="username" value="<%=usrname%>">
<input type="hidden" name="csid" value="<%=csid%>">
<input type="submit" name="action" value="Save Changes" style="border:1px outset #ddffdd;background-color:ccf3cc;margin:3px;">
</form>
</body></html>