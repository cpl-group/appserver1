<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
	'	response.write "|"&number&"|"
		if not(isNumeric(number)) then number = 0
		getNumber = number
	end function
    dim pid, byear, bperiod, building, utilityid, buildingname, procedure, downloadlink

	Dim  Billperiod, Billyear, PortFolioId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql
	' Set Parameters
	pid = request("pid")
    byear = request("byear")
    bperiod = request("bperiod")
    building = request("building")
    utilityid = request("utilityid")
   
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
    Dim objExcelReport 
    Dim objWorkBook 
    Dim objWorkSheet 
    Dim objCell 
	Dim iRow


	Dim sSql
	Dim usage, demand, utilityname

	' Total
	Dim TotalSqFt, MeterCountTotal, UsageTotal, DemandTotal, TenantChargesTotal, AdminFeesTotal
	Dim SalesTaxTotal, MiscCreditsTotal, BuildingChargesTotal


	'Initialize
	

	If bperiod <> "" then
		Set objExcelReport = CreateObject("Excel.Application")
		Set objWorkBook = objExcelReport.Workbooks.Add
	
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")
		
		cnn1.open getLocalConnect(building)
	
		
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 11

    irow = 1

    
    
	objWorkSheet.Cells(iRow,1) = "Tenant_Number"
    objWorkSheet.Cells(iRow,2) = "New_Chargers"
    objWorkSheet.Cells(iRow,3) = "Bill_Period"
    objWorkSheet.Cells(iRow,4) = "Period_Start_Date"
    objWorkSheet.Cells(iRow,5) = "Period_End_Date"
   
    		
	sSql = "Exec sp_select_tenant_summary_export " & byear & "," & bperiod & ",'" & building & "'," & utilityid
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
        objWorkSheet.Cells(iRow,1) = rst1("tenantnumber")
	    objWorkSheet.Cells(iRow,2) = rst1("newcharges")
        objWorkSheet.Cells(iRow,3) = rst1("billperiod")
        objWorkSheet.Cells(iRow,4) = rst1("periodstartdate")
        objWorkSheet.Cells(iRow,5) = rst1("periodenddate")
        
	    	    						
		rst1.movenext
	loop
	rst1.close
	
	

    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    'objWorkSheet.Rows("8:8").Select
    'objExcelReport.ActiveWindow.FreezePanes = True

	dim ctime, bperiodchar

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
    ctime = byear & bperiodchar & "01" 
    																		


	objExcelReport.DisplayAlerts = False
	objWorkBook.SaveAs "D:\FTP\ghnet\LeFrakData\Export Files\CPL_New_Charges_" & ctime & ".csv",6
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit

    set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent

    
    
    
	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "d:\ftp\ghnet\lefrakdata\export files\CPL_New_Charges_" & ctime & ".csv"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/ftp/lefrakdata/export files/CPL_New_Charges_<%=ctime%>.csv" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b>CPL_New_Charges_<%=ctime%>.csv</b></a> 
	</p>
	<%
	Else
	%>
	<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy IT department for assistance.</p>
	<%
		
	End IF


	'Dim objEmail 
	'Dim strSQL
	'Dim strMailingList
	'Dim rstMailingList
		
	'Set objEmail = Server.CreateObject("CDONTS.NewMail") 
	'Set rstMAilingList =  server.createobject("ADODB.Recordset")

	'strSQL = "SELECT email FROM contacts Where submeter_bills=1 and bldgnum ='" & building & "'"
	'strMailingList = ""
	'rstMAilingList.open strSQL , getConnect(PortFolioId,building,"Billing")
	'If not rstMailingList.EOF Then
	'	Do While not rstMailingList.EOF 
	'		if len(strMailingList) > 0 then 
	'			strMailingList = strMailingList & ";" & rstMailingList("Email")
	'		else
	'			strMailingList = rstMailingList("Email")
	'		end if
	'		rstMailingList.MoveNext 
	'	Loop 
	'End IF
	' If There is a mailing List then
	'If Len(strMailingList) > 0 then
		'objEmail.To = strMailingList
	'	objEmail.To = "AnthonyC@genergy.com; tarunskalra@hotmail.com"
	'	objEmail.From = "rb@genergy.com"
	'	objEmail.Subject = "Meter Letter for Building " & building & " , Period " & Billperiod & " " & Billyear 
	'	objEmail.AttachFile "\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls" , building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls" 
	'	objEmail.Send
		
	'	Response.Write "<P> Meter Letter Generated and sent to Building Contacts <BR>"
	'	Response.Write strMailingList 
	'	Response.Write "</P></Body></Html>"
	'Else
	'	Response.Write "<P> No Mailing List is Available for the Building <BR>"
	'	Response.Write "</P></Body></Html>"
	'End IF
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
'		response.write _
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
