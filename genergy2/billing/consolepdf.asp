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
	dim bldg, byear, bperiod, utilid, action, pid
	pid = request("portfolio")
	bldg = request("bldg")
	byear = request("byear")
	bperiod = request("bperiod")
	utilid = request("utilfilter")
	action = trim(request("action"))
	function toNumb(val)
		if val="" or isnull(val) then
			val = 0
		end if
		if IsNumeric(CStr(val)) then
			toNumb = cdbl(val)
		end if
	end function	
	dim cnn1, rst1, rst2, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	cnn1.Open getLocalConnect(Replace(bldg,"+"," "))
	
	Dim cmd, prm
	set cmd = server.createobject("ADODB.Command") 
	
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	
	if pid = "" then
		cmd.CommandText = "usp_TestG1ConsolePdfs"
		Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
		cmd.Parameters.Append prm
		cmd.parameters("bldg") = Replace(bldg,"+"," ")
	else
		if action = "gen" then
			cmd.CommandText = "usp_TestG1ConsolePdfs_bypid"
		end if
		if action = "zip" then
			cmd.CommandText = "ZipFilesbyPid"
		end if
		Set prm = cmd.CreateParameter("pid", adInteger, adParamInput, 20)
		cmd.Parameters.Append prm
		cmd.parameters("pid") = pid
	end if
	
		Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
		cmd.Parameters.Append prm
			cmd.parameters("by") = byear
			cmd.parameters("bp") = bperiod
			cmd.parameters("utility") = utilid
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
'response.write "generating.."
'Response.Write ("<script>self.close();</script>")
'Response.End
%>
	</body>
</html>