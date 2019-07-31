<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if
response.expires=-1
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

	<head>
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	</head>

	<body>
<%

	dim bldg, byear, bperiod, utilid, action, portfolio, pdf
	portfolio = request("porftolio")
	bldg = request("bldg")
	byear = request("byear")
	bperiod = request("bperiod")
	utilid = trim(request("utilid"))
	action = trim(request("action"))
	
	dim cnn1, cnn2, rst1, rst2, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	cnn1.Open getLocalConnect(Replace(bldg,"+"," "))
	set cnn2 = server.createobject("ADODB.connection")
	cnn2.Open getLocalConnect(Replace(bldg,"+"," "))
	
	Dim cmd, prm
	set cmd = server.createobject("ADODB.Command") 

	function genpdf( bldg, byear, bperiod, utilid )
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
				cmd.parameters("bldg") = Replace(bldg,"+"," ")
				cmd.parameters("by") = byear
				cmd.parameters("bp") = bperiod
				cmd.parameters("utility") = utilid
				
		cmd.execute
	end function
	
	if portfolio = "" then
	
		pdf=genpdf(bldg, byear, bperiod, utilid)	
		
	else
		
		strsql = "select bldgnum,utilityid from DailyExportBuildings where online=1 and pid="&pid&" and utilityid="&utilid
		response.write strsql & "</br>"
		rst2.open strsql, cnn1
		do until rst2.eof 
		
			bldg = rst2("bldgnum")
			utilid = rst2("utilityid")
			response.write "bldg"&bldg&"-"&utilid  & "</br>"
			response.end
			pdf=genpdf(bldg, byear, bperiod, utilid)
			
		rst2.movenext
		loop
		rst2.close		
		
	end if

%>


<% 
'response.write "generating.."
'Response.Write ("<script>self.close();</script>")
'Response.End
%>
	</body>
</html>