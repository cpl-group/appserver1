<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if
response.expires=-1
%>

<%
	dim bldg, byear, bperiod, utilid, action
	bldg = request("bldg")
	byear = request("byear")
	bperiod = request("bperiod")
	utilid = trim(request("utilid"))
	action = trim(request("action"))
	
	dim cnn1, rst1, rst2, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	cnn1.Open getLocalConnect(Replace(bldg,"+"," "))
	
	Dim cmd, prm
	set cmd = server.createobject("ADODB.Command") 
	
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	
	if action ="gen" or action ="" then
		cmd.CommandText = "usp_TestG1ConsolePdfs"
		Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
		cmd.Parameters.Append prm
			cmd.parameters("bldg") = Replace(bldg,"+"," ")
			cmd.parameters("by") = byear
			cmd.parameters("bp") = bperiod
			cmd.parameters("utility") = utilid
	else if action = "del" then
		cmd.CommandText = "sp_unpostbill_v2"
		Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("lid", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("delete", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("note", adVarChar, adParamInput, 1000)
			cmd.Parameters.Append prm
			cmd.parameters("bldg") = Replace(bldg,"+"," ")
			cmd.parameters("lid") = 0
			cmd.parameters("by") = byear
			cmd.parameters("bp") = bperiod
			cmd.parameters("utility") = utilid
			cmd.parameters("delete") = 1
			cmd.parameters("user") = "Process Bills"
			cmd.parameters("note") = ""
		
	end if
	
	cmd.execute
		
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

	<head>
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	</head>

	<body>
<% 
response.write "generating.."
'Response.Write ("<script>self.close();</script>")
'Response.End
%>
	</body>
</html>