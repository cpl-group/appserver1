<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<!--  
    METADATA  
    TYPE="typelib"  
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library"  
-->

<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"-->
<%end if

	function getNumber(number)
	'	response.WriteText"|"&number&"|"
		if not(isNumeric(number)) then number = 0
		getNumber = number
	end function
    dim pid, byear, bperiod, building, utilityid, buildingname, procedure, downloadlink, btype
	
	Dim  Billperiod, Billyear, PortFolioId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql
	' Set Parameters
	pid = request("pid")
    byear = request("byear")
    bperiod = request("bperiod")
    building = request("building")
    utilityid = request("utilityid")
	btype = request("btype")
   
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Set Default
	if UtilityId = "" then
		Utilityid = 2
	end if
	Dim rst1, rst2, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 	
%>
<html>
<head>
<title>Export Tenant Summary</title>

<style type="text/css">
INPUT#f9 {
	font-size:9
}
</style>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
   <form name="form1" action="ExportTenantSummary.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
        
				<td>
				 
            </td>
          </tr>
        </table></td>
        </form>
	</tr>
</table>
<%	
	Dim sSql, uSql, objFSO
	dim ctime, bperiodchar, utility, crlf
	Dim csvPath,csvFile,csvColumns
	
	If bperiod <> "" then
	
		set rst1 = server.createobject("ADODB.Recordset")
		set rst2 = server.createobject("ADODB.Recordset")
		set cnn1 = server.createobject("ADODB.Connection")
		
		cnn1.open getLocalConnect(building)
   
		Select Case bperiod
			Case 1     bperiodchar = "01"
			Case 2   bperiodchar = "02"
			Case 3    bperiodchar = "03"
			Case 4      bperiodchar = "04"
			Case 5     bperiodchar = "05"
			Case 6   bperiodchar = "06"
			Case 7    bperiodchar = "07"
			Case 8     bperiodchar = "08"
			Case 9    bperiodchar = "09"
			Case 10   bperiodchar = "10"
			Case 11   bperiodchar = "11"
			Case 12      bperiodchar = "12"
		End Select
		
		ctime = byear _
        & Right("0" & bperiod, 2) '& Right("0" & DatePart("d",Date), 2)

		uSql = "select utilitydisplay from tblutility where utilityid=" & utilityid
		rst2.open uSql, cnn1, 3
		utility = rst2("utilitydisplay")	
		rst2.close
								
		sSql = "Exec sp_select_tenant_summary_export " & byear & "," & bperiod & "," & utilityid & ",'" & btype & "'"
		
		rst1.CursorLocation = 3
		rst1.open sSql , cnn1, 3 
						
		crlf = chr(13) & chr(10)

		' Create new csv file 
		csvPath = "d:\ftp\ghnet\lefrakdata\export files\"
		'csvFile = building & "_" & utility & "_" & byear & right("0" & bperiod,2) & ".txt"
		csvFile = "CPL_New_Charges_" & ctime & ".txt"
		'csvFile = "CPL_New_Charges_"&btype&"_" & ctime & ".txt"
		
		Dim UTFStream
		Set UTFStream = CreateObject("adodb.stream")
		UTFStream.Type = adTypeText
		UTFStream.Mode = adModeReadWrite
		UTFStream.Charset = "UTF-8"
		UTFStream.LineSeparator = adLF
		UTFStream.Open
		
		'If objFSO.FileExists(csvPath & csvFile) Then
		''	UTFStream.loadfromfile csvpath&csvfile
		''	UTFStream.readtext
			'dim readdata 
			'readdata = "" & UTFStream.readtext
			'UTFStream.writetext readdata
		'else
			csvColumns = "Tenant_Number,New_Charges,Billing_Period,Period_Start_Date,Period_End_Date"
			UTFStream.WriteText csvColumns 
			UTFStream.WriteText crlf
		'end if
		
		Do Until rst1.eof
			if rst1("tenantnumber") <> "" then 
				UTFStream.WriteText chr(34) & rst1("tenantnumber") & chr(34) & ","
				UTFStream.WriteText chr(34) & rst1("newcharges") & chr(34) & ","
				UTFStream.WriteText chr(34) & rst1("billperiod") & chr(34) & ","
				UTFStream.WriteText chr(34) & rst1("periodstartdate") & chr(34) & ","
				UTFStream.WriteText chr(34) & rst1("periodenddate") & chr(34)
				UTFStream.WriteText crlf
			end if
			
			rst1.movenext
		loop
		rst1.close

		UTFStream.Position = 3 'skip BOM

		Dim BinaryStream
		Set BinaryStream = CreateObject("adodb.stream")
		BinaryStream.Type = adTypeBinary
		BinaryStream.Mode = adModeReadWrite
		BinaryStream.Open

		'Strips BOM (first 3 bytes)
		UTFStream.CopyTo BinaryStream

		'UTFStream.SaveToFile "d:\temp\adodb-stream1.csv", adSaveCreateOverWrite
		UTFStream.Flush
		UTFStream.Close

		BinaryStream.SaveToFile csvPath & csvFile, adSaveCreateOverWrite
		BinaryStream.Flush
		BinaryStream.Close
		
		If objFSO.FileExists(csvPath & csvFile) Then 
		%>
		<p> Following report has been generated for all H2O based residences in LeFrak:</br>
		This file has been placed in the appropriate FTP location for pickup.</br>
		<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/ftp/lefrakdata/export files/<%=csvFile%>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=csvFile%></b></a> 
		</p>
		<%
		Else
		%>
		<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy IT department for assistance.</p>
		<%
		End IF

	End If %>

<%
	
	'set objEmail = Nothing
	'set rstMailingList = Nothing
	set objFSO = Nothing
	set rst1 = Nothing
	set rst2 = Nothing
	set cnn1 = Nothing
	
	
%>	

<%
	Dim objSWbemServices, colProcess, objProcess, resultCode
	Set objSWbemServices = GetObject ("WinMgmts:Root\Cimv2")
	Set colProcess = objSWbemServices.ExecQuery ("Select * From Win32_Process WHERE Name LIKE '%EXCEL.EXE%'")
'	For Each objProcess In colProcess
'		response.WriteText_
'		"<ul>"&_
'		"<li>Name="& objProcess.Name      &_
'		"<li>PID ="& objProcess.ProcessId &_
'		"</ul>"
'	Next
	For Each objProcess In colProcess
		resultCode = objProcess.Terminate()
	Next
'	response.end
%>
