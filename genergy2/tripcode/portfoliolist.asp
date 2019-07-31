<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
%>
<html>
<head>
<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewbills(ypid) {
	var temp
		temp="invoicebldg.asp?ypid=" + ypid
		document.frames.admin.location=temp
} 
function loadlist(pid) {
	var temp = "buildingtc.asp?pid=" + pid
	document.frames.admin.location=temp
}
</script>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#6699CC" class="standardheader"><b>Building Trip Code Setup</b></td></tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr><td width="48%"> 
    <select name="pid" onchange="loadlist(this.value)">
    <%
    sqlstr = "select name, id from portfolio order by name"
    rst1.Open sqlstr, cnn1, 0, 1, 1
    if not rst1.eof then
    do until rst1.eof
    %><option value="<%=rst1("id")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("name")%></font></option>
    <%rst1.movenext
    loop
    end if
    rst1.close
    %>
    </select>
    <input type="button" name="Button2" value="View Building List" onClick="loadlist(pid.value)">
    </td>
    <td align="right"><input type="button" name="Submit" value="Print Trip Code List" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'></td>
</tr>
</table>
<p><IFRAME name="admin" width="100%" height="85%" src="/null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></p>
</body>
</html>