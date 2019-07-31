<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldgnum, action, pid, code
bldgnum = trim(request("bldgnum"))
action = trim(request("action"))
pid = trim(request("pid"))
code = trim(request("code"))
%>
<html>
<head>
<title>ERI Download</title>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
</head>

<body bgcolor="#eeeeee" text="#000000">
<form>
<strong>Download Buildings</strong><br>
<%

	
dim rs, cmd, prm
set cmd = server.createobject("ADODB.command")
set rs = server.createobject("ADODB.recordset")
if action<>"" and bldgnum<>"" then
	cmd.ActiveConnection = getConnect(0,bldgnum,"Billing")
	cmd.commandtext = "sp_AccountingFile_TZH"
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("BLDG", adVarChar, adParamInput, 1000)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("BY", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("BP", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("ERI", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("code", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("util", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("file", adVarChar, adParamOutput, 100)
	cmd.Parameters.Append prm
	cmd.Parameters("BLDG")= bldgnum
	cmd.Parameters("BY")= 0
	cmd.Parameters("BP")= 0
	cmd.Parameters("ERI")= 1
	cmd.Parameters("code")= code
	cmd.Parameters("util")= 2
	'cmd.Parameters("file")        = 0
	'response.write "exec "&cmd.commandtext&" '"&	cmd.Parameters("BLDG")&"', "&cmd.Parameters("BY")&", "&cmd.Parameters("BP")&" ,"&cmd.Parameters("ERI")&", '"&cmd.Parameters("code")&"', "&cmd.Parameters("util")
	'response.write "<br>"&getConnect(0,bldgnum,"Billing")
	'response.end
	'dim mt
	'for each mt in cmd.Parameters
	'	Response.Write "set @" & mt.Name & "='" & mt.Value & "'<br>"
		
	'next
	'Response.End
	cmd.execute
	'/eri_th/finance/TZH/XFA_Accounting_Files/&sub=true&prevdir=/eri_th/finance/TZH/
'	response.redirect "http://appserver1.genergy.com/eri_th/gfile_intranet.asp?clientname=gEnergy Employees&clientdir=" & replace(cmd.Parameters("file"),"\","/")
	if cmd.Parameters("file") <> "0" then
		response.redirect "http://appserver1.genergy.com/eri_th/gfile_intranet.asp?clientname=gEnergy%20Employees&clientdir="&replace(cmd.Parameters("file"),"\","/")&"&sub=true&prevdir=/eri_th/finance/TZH/"
	else
		response.write "Download was unsuccessfull."
	end if
else
	rs.open "SELECT strt, bldgnum FROM buildings WHERE bldgnum in (SELECT distinct bldg_no FROM Tenant_history th, Tenant_Info ti WHERE Owner_id='"&pid&"' and ti.tenant_no=th.tenant_no and datepart(year,date_event)>2003)ORDER BY strt",getConnect(0,0,"Engineering")
	if not rs.eof then%>
	<select name="bldgnum" size="6" multiple><%
	do until rs.eof
		%><option value="<%=rs("bldgnum")%>"><%=rs("strt")%> (<%=rs("bldgnum")%>)</option><%
		rs.movenext
	loop
	%></select><br><%
	else
		response.write "No available downloads for this portfolio."
	end if
	rs.close
	''
	rs.open "SELECT Distinct [note], code FROM Tenant_history th, Tenant_Info ti WHERE ti.tenant_no=th.tenant_no and left(code,1)='m' and datepart(year,date_event)>2003 and code<>'9999' and bldg_no in (SELECT bldgnum FROM buildings WHERE Owner_id='"&pid&"')"
	if rs.eof then%>
	<%else%>
		<select name="code">
		<%do until rs.eof%>
			<option value="<%=rs("code")%>"><%=rs("note")%> (<%=rs("code")%>)</option><%
			rs.movenext
		loop%>
		</select><br>
		<input type="submit" name="action" value="Download"><%
	end if
	rs.close
end if
%>
</form>

</body>
</html>
