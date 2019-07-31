<%option explicit
server.ScriptTimeout = 580
%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>gEnergyOne</title>
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
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
</head>
<body BGCOLOR="#eeeeee" LINK="#0000CC" VLINK="#0000CC" TEXT="#000000">
<form method="POST" name="form1" action="dataOutput.asp">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc"><td colspan="2"><span class="standardheader"><%=buildingname%> Data  Downloading</span></td></tr>
<%
dim pid, byear, bperiod, building, utilityid, buildingname, procedure, downloadlink

procedure =""

pid = request("pid")
byear = request("byear")
bperiod = request("bperiod")
building = request("building")
utilityid = request("utilityid")
procedure = request("procedure")


dim sql, rst1, cnn1, cmd, prm
set rst1 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
cnn1.open getConnect(pid,building,"billing")
cnn1.CommandTimeout = 580
cmd.ActiveConnection = cnn1

rst1.open "SELECT * FROM buildings WHERE bldgnum='"&building&"'", cnn1
if not rst1.eof then buildingname = rst1("strt")
rst1.close

if trim(procedure)<>"" then
Err.clear
on Error resume next
	if (pid <> "108") then
	
	cnn1.CursorLocation = adUseClient
	cmd.CommandText = procedure
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("util", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("file", adVarChar, adParamOutput, 80)
	cmd.Parameters.Append prm

	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("util")		= utilityid

	'response.Write(procedure + " " + building + "," + byear + "," + bperiod + "," + utilityid)
    'response.End()	
	
	else
		
	cnn1.CursorLocation = adUseClient
	cmd.CommandText = "sp_PaDtsStart"
	cmd.CommandType = adCmdStoredProc
	
	Set prm = cmd.CreateParameter("by", adVarChar, adParamInput,10)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("BP", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("BLDG", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("FLAG", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm

	Set prm = cmd.CreateParameter("util", adInteger, adParamInput)
	cmd.Parameters.Append prm
		
	Set prm = cmd.CreateParameter("lid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("datefrom", adDate, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("dateto", adDate, adParamInput, 20)
    cmd.Parameters.Append prm
	
	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("BP") = bperiod
	cmd.Parameters("by") = byear
	cmd.Parameters("BLDG") = building
	cmd.Parameters("FLAG") ="0"  'N.Ambo changed to 1 9/30/2008 then changed back to 0 10/6/2008; 0 is required for dts package to be executed so all accoutning files can be generated
	cmd.Parameters("util")		= utilityid
	cmd.Parameters("lid") = 0
	cmd.Parameters("datefrom") = null
	cmd.Parameters("dateto") = null
	
	'response.Write("sp_PaDtsStart" + " " +  bperiod +"," + building)
   ' response.End()	
	
	end if

	cmd.execute
	if err.number = 0 then 
	%>
	<tr><td colspan="2" align="center">&nbsp;<br>
			Data Files for <b><%=buildingname%></b> 
			period <%=bperiod%>, <%=byear%><br>have been created.<br>They can be accessed via your<br>gEnergyOne Data File Access Module.
		   </td></tr>
	<%
	else
	%>
	<tr><td colspan="2" align="center">&nbsp;<%=err.description%><br>
			Data Files for <b><%=buildingname%></b> 
			period <%=bperiod%>, <%=byear%><br>failed to be created.<br>Please try again or contact support@genergy.com
		   </td></tr>
	<%
	end if
else%>
<tr><td colspan="2">Data File Type For period <%=bperiod%>, <%=byear%>:</td></tr>
<tr valign="top"> 
    <% if pid <> "108" then%>
	<td width="5%">
      <select name="procedure">
        <%
          rst1.open "SELECT * FROM ADF WHERE (bldgnum='"&building&"' or bldgnum is null) and (pid="&pid&" or pid is null) ORDER BY label", getConnect(pid,building,"Billing")
          do until rst1.eof
            %><option value="<%=rst1("procname")%>"><%=rst1("label")%></option><%
            rst1.movenext
          loop
          rst1.close
        %>
      </select>
    </td>
    <%end if%>
	
	  <%if pid <> "108" then%>	
	<td width="95%"><input type="button" value="Create" onclick="document.location.href='loading.asp?url=<%=server.urlencode("/genergy2/billing/dataOutput.asp?pid="&pid&"&byear="&byear&"&bperiod="&bperiod&"&building="&building&"&utilityid="&utilityid)%>%26procedure%3D'+this.form.procedure.value"></td>
		<%else%>
   <td width="95%"><input type="button" value="Create" onclick="document.location.href='loading.asp?url=<%=server.urlencode("/genergy2/billing/dataOutput.asp?pid="&pid&"&byear="&byear&"&bperiod="&bperiod&"&building="&building&"&utilityid="&utilityid)%>%26procedure%3D'+'PA'"></td>
		<%end if%>

</tr>
<%end if%>
</table>
<input name="pid" value="<%=pid%>" type="hidden">
<input name="byear" value="<%=byear%>" type="hidden">
<input name="bperiod" value="<%=bperiod%>" type="hidden">
<input name="building" value="<%=building%>" type="hidden">
<input name="utilityid" value="<%=utilityid%>" type="hidden">
</form>
</body>
</html>
