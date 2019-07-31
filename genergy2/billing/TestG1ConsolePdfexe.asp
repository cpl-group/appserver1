<%@ Language=VBScript %>
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<%

dim building, byear,  bperiod, utilityid,pid
utilityid = request("utilityid")
building = request("building")
byear = request("byear")
bperiod = request("bperiod")

Response.Write "Bill Year:" & byear
Response.Write "Bill Period:" & bperiod
dim rst1, cnn1, sql, cmd, prm, sql2

set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")

cnn1.open getConnect(pid,building,"billing")
cnn1.commandTimeout = 0

cmd.ActiveConnection = cnn1
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "usp_TestG1ConsolePdfs"

Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
cmd.Parameters.Append prm


	cmd.parameters("bldg") = building
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.parameters("utility") = utilityid
	
cmd.execute

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>test executed!</P>

</BODY>
</HTML>
