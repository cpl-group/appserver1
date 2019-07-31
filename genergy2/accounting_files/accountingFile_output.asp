<%option explicit

%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim cnn1, coreCmd, dateFrom, dateTo, lid, util, billy, billp
dim acctid, pid, bldg, sqlStr, prm, i, rst, procName

bldg = request("bldg")
pid = request("pid")
util = request("util")
lid = request("lid")
billy = request("billyear")
billp = request("billperiod")
dateFrom = request("dateFrom")
dateTo = request("dateTo")

if (datefrom <> "" AND dateTo = "") then
   dateTo = now()
end if

if (dateFrom = "dd/mm/yy") then dateFrom = null
if (dateTo = "dd/mm/yy") then dateTo = null

if (billp = "") then billp = 0
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst = server.CreateObject("ADODB.RecordSet")
set coreCmd = server.createobject("ADODB.command")

cnn1.Open getConnect(pid,bldg,"Billing")

rst.open "Select procedureName FROM acctFile_Template a INNER JOIN bldg_acctFile_setup b ON a.templateId = b.templateId WHERE bldgNum = '" + bldg + "'", cnn1
if not rst.eof then
    procName = rst("procedureName")
else
    procName = "sp_create_acctfile"
end if

rst.close

coreCmd.ActiveConnection = cnn1
coreCmd.CommandText = procName
coreCmd.CommandType = adCmdStoredProc

Set prm = coreCmd.CreateParameter("BY", adVarChar, adParamInput, 10)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("BP", adTinyInt, adParamInput)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("bldg", adVarChar, adParamInput, 50)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("Flag", adInteger, adParamInput)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("UTIL", adInteger, adParamInput)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("lid", adInteger, adParamInput)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("dateFrom", adDate, adParamInput, 20)
coreCmd.Parameters.Append prm
Set prm = coreCmd.CreateParameter("dateTo", adDate, adParamInput, 20)
coreCmd.Parameters.Append prm

coreCmd.Parameters("bldg") = trim(bldg)
coreCmd.Parameters("lid") = CInt(lid)
coreCmd.Parameters("UTIL") = CInt(util)
coreCmd.Parameters("BY") = trim(billy)
coreCmd.Parameters("BP") = CInt(billp)
coreCmd.Parameters("dateFrom") = dateFrom
coreCmd.Parameters("dateTo") = dateTo
coreCmd.Parameters("Flag") = 0

coreCmd.execute()

if err.number = 0 then 
	%>
	<tr><td colspan="2" align="center">&nbsp;<br>
			Data Files for <b><%=bldg%></b> 
			period <%=billp%>, <%=billy%><br>have been created.<br>They can be accessed via your<br>gEnergyOne Data File Access Module.
		   </td></tr>
	<%
	else
	%>
	<tr><td colspan="2" align="center">&nbsp;<%=err.description%><br>
			Data Files for <b><%=bldg%></b> 
			period <%=billp%>, <%=billy%><br>failed to be created.<br>Please try again or contact support@genergy.com
		   </td></tr>
	<%
	end if

on error resume next
set coreCmd = nothing
set cnn1 = nothing
 %>