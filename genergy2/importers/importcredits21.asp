<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
	<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
%>

<html>
<head>
<title>temp.Com Tutorial</title>
</head>
<body>
<%
	dim pid, bldg, customsrc, action, tid, lid, id, byear, bperiod, creditid, credit_adj, utilid
	pid = request("pid")
	bperiod = request("bperiod")
	byear = request("byear")
	bldg = request("bldg")
	tid = request("tid")
	lid = request("lid")
	utilid = request("utilid")
	Dim objFSO,oInStream,sRows,arrRows
	Dim Conn,strSQL,objExec
	Dim sFileName
	Dim mySmartUpload, rst1
	set rst1 = server.createobject("ADODB.Recordset")
	
	'*** Upload By aspSmartUpload ***'

	'*** Create Object ***'	
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

	'*** Upload Files ***'
	mySmartUpload.Upload

	'** Getfile Name ***'
	sFileName = mySmartUpload.Files("file1").FileName

	If sFileName <> "" Then
		'*** Upload **'
		mySmartUpload.Files("file1").SaveAs(Server.MapPath("temp/" & sFileName))

		'*** Create Object ***'
		Set objFSO = CreateObject("Scripting.FileSystemObject") 

		'*** Check Exist Files ***'
		If Not objFSO.FileExists(Server.MapPath("temp/" & sFileName)) Then
			Response.write("File not found.")
		Else

		'*** Open Files ***'
		Set oInStream = objFSO.OpenTextFile(Server.MapPath("temp/" & sFileName),1,False)

		'*** open Connect to Access Database ***'
		Set Conn = Server.Createobject("ADODB.Connection")
		Conn.Open getLocalConnect(bldg)				'"DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("temp/mydatabase.mdb"),"" , ""


		Do Until oInStream.AtEndOfStream 
		sRows = oInStream.readLine
		arrRows = Split(sRows,",")
			'*** Insert to table customer2 ***'
			if arrRows(0) <> "tenantname" then
				strSQL = ""
				strSQL = strSQL &"INSERT INTO misc_inv_credit_loader "
				strSQL = strSQL &"(tenantname, aptnum, balance, note, bldgnum, billyear, billperiod) "
				strSQL = strSQL &"VALUES "
				strSQL = strSQL &"('"&arrRows(0)&"','"&arrRows(1)&"','"&arrRows(2)&"' "
				strSQL = strSQL &",'"&arrRows(3)&"','"&bldg&"','"&byear&"','"&bperiod&"') "
				Set objExec = Conn.Execute(strSQL)
				Set objExec = Nothing
			end if
		Loop

		oInStream.Close()
		Conn.Close()
		Set oInStream = Nothing
		

		End If
		
		Response.write ("CSV import completed.")
	
		Conn.Open getLocalConnect(bldg)
		Dim cmd, prm
		set cmd = server.createobject("ADODB.Command") 
		
		cmd.ActiveConnection = conn
		cmd.CommandType = adCmdStoredProc
		
		
			cmd.CommandText = "MiscInvCreditLoader"
			Set prm = cmd.CreateParameter("bldgnum", adVarChar, adParamInput, 20)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("billyear", adInteger, adParamInput)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("billperiod", adInteger, adParamInput)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("utilid", adInteger, adParamInput)
			cmd.Parameters.Append prm
				cmd.parameters("bldgnum") = Replace(bldg,"+"," ")
				cmd.parameters("billyear") = byear
				cmd.parameters("billperiod") = bperiod
				cmd.parameters("utilid") = utilid
		'cmd.execute		

		strsql = "select distinct tenantname, aptnum, balance, note from misc_inv_credit_loader_passed where bldgnum='"&bldg&"' and billperiod="&bperiod&" and billyear="&byear
		rst1.open strsql, conn
		
		if not rst1.eof then %>
				<table>
					<tr>
						<td colspan 4> Successful Imports </td>
					</tr>
				<%
			do until rst1.eof
				%>
					<tr>
						<td><%= rst1("tenantname") %> </td>
						<td><%= rst1("aptnum") %> </td>
						<td><%= rst1("balance") %> </td>
						<td><%= rst1("note") %> </td>
					</tr>
				
				<%
			rst1.movenext
			loop %>
				</table>
				<%
		end if
		
		rst1.close
		
		strsql = "select distinct tenantname, aptnum, balance, note from misc_inv_credit_loader_fails where bldgnum='"&bldg&"' and billperiod="&bperiod&" and billyear="&byear
		rst1.open strsql, conn
		
		if not rst1.eof then %>
				<table>
					<tr>
						<td colspan 4> Failed Imports </td>
					</tr>
				<%
			do until rst1.eof
				%>
					<tr>
						<td><%= rst1("tenantname") %> </td>
						<td><%= rst1("aptnum") %> </td>
						<td><%= rst1("balance") %> </td>
						<td><%= rst1("note") %> </td>
					</tr>
				
				<%
			rst1.movenext
			loop %>
				</table>
				<%
		end if
		
		rst1.close
			
	conn.close()

	
	Set Conn = Nothing
	End IF
	
%>
</body>
</html>
<!--- This file download from www.temp.com -->