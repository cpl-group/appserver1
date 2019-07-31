<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim date1, date2, b, utype, pid, adjtype
b = request.querystring("b")
pid = request.querystring("pid")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
utype = request.querystring("utype")
adjtype = request.querystring("adjtype")
dim title, ISeri
dim mac, plp
mac = "0"
plp = "0"
if adjtype="mac" then
	title = "Mac Adjustment"
	mac = "1"
else
	title = "Public Lighting and Power"
	plp = "1"
end if

dim cnn1, rst1, cmd, prm
Set rst1 = Server.CreateObject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rst1 = server.createobject("ADODB.Recordset")
cnn1.Open application("Cnnstr_genergy1")
cnn1.CursorLocation = adUseClient
cmd.CommandText = "sp_RevProfile"
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("util", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adVarChar, adParamInput, 10)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("eri", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("exp", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("subm", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("urar", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("urae", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("mac", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("plp", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("net", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 50)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("trev", adVarChar, adParamOutput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("texp", adVarchar, adParamOutput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("nett", adVarChar, adParamOutput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bldgnum", adVarChar, adParamOutput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("byear", adInteger, adParamOutput)
cmd.Parameters.Append prm
cmd.Name = "test"
Set cmd.ActiveConnection = cnn1

cnn1.test b, utype, date1, 0, 0, 0, 0, 0, mac, plp, 0, session("userid"), rst1

%>
<html>
<head><title></title>
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame2')" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="706" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b><%=title%></b></font></td>
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='monthlyDetails.asp?b=<%=b%>&pid=<%=pid%>&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utype='+ parent.document.forms['form1'].utype.value" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Monthly Details</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"></font></td>
	</tr>
</table>	
&nbsp;
<%if adjtype="mac" then%>
	<table border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<td>Quarter 1</td>
	<td>Quarter 2</td>
	<td>Quarter 3</td>
	<td>Quarter 4</td>
	</tr>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<%
	do until rst1.eof
		dim quarter, i
		quarter=0
		for i = 1 to 3
			if not rst1.eof then 
				quarter = quarter + rst1("mac_rev")
				rst1.movenext
			end if
		next
		response.write "<td>"&formatcurrency(quarter)&"</td>"
	loop
%>
	</table>
<%else
	dim total
	total=0%>
	<table width="350" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<%
	i = 1
	do until rst1.eof or i>12
		response.write "<tr style=""font-family: Arial, Helvetica, sans-serif; font-size: 10;"">"
		response.write "<td>"&monthname(i)&"</td>"
		response.write "<td>"&formatcurrency(rst1("plp"))&"</td>"
		total = total + cDBL(trim(rst1("plp")))
		i = i + 1
		rst1.movenext
	loop
	%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 11;font-weight:bold">
	<td>Annual Total</td>
	<td><%=formatcurrency(total)%></td>
	</tr>
	</table>

<%end if%>
