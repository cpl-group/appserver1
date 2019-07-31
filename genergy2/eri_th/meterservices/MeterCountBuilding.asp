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
    
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "BuildingName"
    objWorkSheet.Cells(iRow,2) = "TenantCount "
    objWorkSheet.Cells(iRow,3) = "Electric"
	objWorkSheet.Cells(iRow,4) = "ColdWater"
    objWorkSheet.Cells(iRow,5) = "Steam"
	objWorkSheet.Cells(iRow,6) = "ChilledWater"
    objWorkSheet.Cells(iRow,7) = "Gas"
    objWorkSheet.Cells(iRow,8) = "HotWater"
    objWorkSheet.Cells(iRow,9) = "Chiller"
    objWorkSheet.Cells(iRow,10) = "Condenser"
    objWorkSheet.Cells(iRow,11) = "Residential"
    objWorkSheet.Cells(iRow,12) = "Commercial"
    objWorkSheet.Cells(iRow,13) = "TotalMeterCount"
    objWorkSheet.Cells(iRow,14) = "ManualMeterCount"
    objWorkSheet.Cells(iRow,15) = "RemoteMeterCount"
		
	sSql = "Exec sp_select_portbldgmeter_counts_active_bldg " & "'" & building & "'"
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("buildingname")
        objWorkSheet.Cells(iRow,2) = rst1("tenantcount")
        objWorkSheet.Cells(iRow,3) = rst1("electricmetercount")
	    objWorkSheet.Cells(iRow,4) = rst1("coldwatermetercount")
        objWorkSheet.Cells(iRow,5) = rst1("steammetercount")
        objWorkSheet.Cells(iRow,6) = rst1("chilledwatermetercount")
        objWorkSheet.Cells(iRow,7) = rst1("gasmetercount")
	    objWorkSheet.Cells(iRow,8) = rst1("hotwatermetercount")
        objWorkSheet.Cells(iRow,9) = rst1("chillermetercount")
        objWorkSheet.Cells(iRow,10) = rst1("condensermetercount")
        objWorkSheet.Cells(iRow,11) = rst1("residentialmetercount")
        objWorkSheet.Cells(iRow,12) = rst1("commercialmetercount")
        objWorkSheet.Cells(iRow,13) = rst1("totalmetercount")
	    objWorkSheet.Cells(iRow,14) = rst1("manualmetercount")
        objWorkSheet.Cells(iRow,15) = rst1("remotemetercount")
	    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
        objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,9).Interior.ColorIndex = 40 
        objWorkSheet.Cells(iRow,10).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,11).Interior.ColorIndex = 40 
        objWorkSheet.Cells(iRow,12).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,13).Interior.ColorIndex = 40 
        objWorkSheet.Cells(iRow,14).Interior.ColorIndex = 40 
	    objWorkSheet.Cells(iRow,15).Interior.ColorIndex = 40 
						
		rst1.movenext
	loop
	rst1.close
	
	

    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    'objWorkSheet.Rows("8:8").Select
    'objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Billperiod  & UtilityId & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "MeterCount.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent

    Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "MeterCount.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>MeterCount.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>MeterCount.xlsx</b></a> 
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