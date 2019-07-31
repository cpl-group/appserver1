<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if
%>

<%
	function toNumb(val)
		if val="" then
			val = 0
		end if
		toNumb = cdbl(val)
	end function
	dim bldg, byr, bp, utilid, action, pid
	pid = request("pid")
	byr = request("by")
	bp = request("bp")
	utilid = request("utilid")
	action = trim(request("action"))
	
	dim cnn1, rst1, rst2, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	cnn1.Open getLocalConnect(Replace(bldg,"+"," "))
	
	Dim cmd, prm
	set cmd = server.createobject("ADODB.Command") 
	
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "_H2O_bill_processor"

	Set prm = cmd.CreateParameter("action", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("pid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utilid", adInteger, adParamInput)
	cmd.Parameters.Append prm

		cmd.parameters("action") = Replace(action,"+"," ")
		cmd.parameters("pid") = pid
		cmd.parameters("by") = byr
		cmd.parameters("bp") = bp
		cmd.parameters("utilid") = utilid
			
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
Response.Write ("<script>self.close();</script>")
Response.End
%>
	</body>
</html>