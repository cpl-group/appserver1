<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim rst1, cnn1
set cnn1 = server.createobject("ADODB.Connection")
set rst1 = server.createobject("ADODB.Recordset")
cnn1.open getconnect(0,0,"engineering")
%>
<html>
<head>
<title>Entry</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<script>

function loadClient()
{	var frm = document.forms['form1'];
	var temp = "manageaccounts.asp?cid="+frm.cid.value;
	document.frames.clients.location=temp;
}

function loadNew()
{	var temp = "newclient.asp"
	document.frames.clients.location=temp;
}

</script>
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#0099ff">
	<td><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><span class="standard">Lighting &amp; Maintenance Setup</span></font></b></td>
</tr>
<tr bgcolor="#eeeeee">
	<td>
<select name="cid">
<%rst1.open "SELECT * FROM clients", cnn1
do until rst1.eof
	response.write "<option value="""&rst1("id")&""">"&rst1("Corp_name")&"</option>"
	rst1.movenext
loop
rst1.close%>
</select>
<input type="button" value="Select Client" onclick="loadClient()" class="standard">
<input type="button" value="Create New Client" onclick="loadNew()" class="standard">
	</td>
</tr>
</table>
</form>

<IFRAME name="clients" width="100%" height="500" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>
</body>
</html>