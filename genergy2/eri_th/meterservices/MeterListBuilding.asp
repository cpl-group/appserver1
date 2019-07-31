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

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql
	' Set Parameters
	building = request("bldgNum")	
	Dim rst1, rst2, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 	
%>
<html>
<head>
<title>Meter Count</title>

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
   <form name="form1" action="MeterCountBuilding.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
        	<td>
				 
            </td>
          </tr>
        </table>
      </td>
     </tr>
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

    sSql = "Exec usp_TaxPrep_sierra_buildinginfo " & "'" & building & "'"
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
    
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "BuildingName:"
    objWorkSheet.Cells(iRow,2) = building

    objWorkSheet.Cells(iRow,2) = rst1("bldgname")
    objWorkSheet.Cells(iRow,2).Font.Bold = true

    rst1.close

    iRow= iRow + 2
    objWorkSheet.Cells(iRow,1) = "TenantName"
    objWorkSheet.Cells(iRow,2) = "MeterNumber"
    objWorkSheet.Cells(iRow,3) = "Utility"
	objWorkSheet.Cells(iRow,4) = "UsageMultiplier"
    objWorkSheet.Cells(iRow,5) = "DemandMultiplier"
    objWorkSheet.Cells(iRow,6) = "Location"
			
	sSql = "Exec sp_select_portbldgmeter_list_active_bldg " & "'" & building & "'"
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("tenantname")
        objWorkSheet.Cells(iRow,2) = rst1("meternum")
        objWorkSheet.Cells(iRow,3) = rst1("utility")
	    objWorkSheet.Cells(iRow,4) = rst1("manualmultiplier")
        objWorkSheet.Cells(iRow,5) = rst1("demandmultiplier")
        objWorkSheet.Cells(iRow,6) = rst1("location")
        objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
        objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	   
						
		rst1.movenext
	loop
	rst1.close
	
	

    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    'objWorkSheet.Rows("8:8").Select
    'objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "MeterList.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent

    Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "MeterList.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>MeterList.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>MeterList.xlsx</b></a> 
	</p>
	<%
	Else
	%>
	<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy IT department for assistance.</p>
	<%
		
	End IF


	%>
<Script type=text/javascript>	

</Script>
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