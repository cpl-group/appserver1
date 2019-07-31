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
	BillYear = request("billyear")
	BillPeriod = request("billperiod")
	UtilityId = request("utilityid")
	' Set Default
	if UtilityId = "" then
		Utilityid = 2
	end if
	Dim rst1, rst2, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 	
%>
<html>
<head>
<title>Building Tax Report</title>

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
   <form name="form1" action="BuildingTaxReportNON450.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
        
				<% if trim(building)<>"" then%>
				<td> <select name="utilityid" onChange="loadutility()">
					<option value="">Select Utility</option>
						<%rst1.open "SELECT DISTINCT byp.Utility as utilityid, u.Utilitydisplay " & _
									" FROM BillYrPeriod byp inner join dbo.tblutility u " & _
									" ON byp.Utility = u.utilityid WHERE (BldgNum = '" & trim(building) &"')", getLocalConnect(building)
						do until rst1.eof   %>
						<option value="<%=rst1("utilityid")%>"<%if trim(rst1("utilityid"))=trim(utilityid) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option>
                <%      if trim(rst1("utilityid"))=trim(utilityid) then utilitydisplay = rst1("utilitydisplay")
						rst1.movenext
						loop
						rst1.close
						%>
					  </select> </td>	
				 <%end if %>
				           
			<%if trim(utilityid)<>"" then%>
            <td> <select name="billyear" onclick="loadPeriod()">
                <option value="">Select Bill Year</option>
                <%
                	sql = "SELECT distinct billyear " & _
						" FROM billyrperiod WHERE " & _
				        "billyear>=year(getdate())-1 and bldgnum='"&building&"' order by billyear desc "
				        
					rst1.open sql, getLocalConnect(building)
					do until rst1.eof%>
					<option value="<%=rst1("billyear")%>"<%if trim(rst1("billyear"))=trim(billyear) then response.write " SELECTED"%>><%=rst1("billyear")%></option>
					<%
						
							rst1.movenext
					loop
					rst1.close
					%>
					</select> </td>
					
	  			
					<td> <select name="billperiod">
					 <option value="">Select Bill Period</option>
                <%
                
				sql = "SELECT distinct billperiod " & _
						" FROM billyrperiod WHERE " & _
				        "billyear>=year(getdate())-1 and "
				sql = sql & "bldgnum='"&building&"' order by billperiod desc "
					
				rst1.open sql, getLocalConnect(building)
				do until rst1.eof
				%>
					<option value="<%=rst1("billperiod")%>" <%if trim(rst1("billperiod"))=billperiod then response.write " SELECTED"%>><%=rst1("billperiod")%></option>
                <%
				  rst1.movenext
				loop
				rst1.close
				%>
              </select> </td>
              <%end if%>
				<td>
					<input type="hidden" name="bldgNum" value="<%=Building%>"> 		
				 <input type="Submit" name="Generate Report" value="Generate Report"> 
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
	TotalSqFt = 0.0
	MeterCountTotal = 0
	UsageTotal = 0
	DemandTotal = 0 
	TenantChargesTotal = 0.0
	AdminFeesTotal = 0.0
	SalesTaxTotal = 0.0
	MiscCreditsTotal =0.0
	BuildingChargesTotal =0.0

	If billperiod <> "" then
		Set objExcelReport = CreateObject("Excel.Application")
		Set objWorkBook = objExcelReport.Workbooks.Add
	
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")
		
		cnn1.open getLocalConnect(building)
	
	sSql = "Exec usp_MaintenanceLetterBuildingInfo " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 11


' Header Columns	
	If not rst1.eof then

        		
        objWorkSheet.Cells(1,1).Font.Bold = True
        objWorkSheet.Cells(1,1).Font.Size = 19    
        objWorkSheet.Cells(1,1) = rst1("companyname") + rst1("companyname2")
        objWorkSheet.Cells(1,1).Interior.ColorIndex = 36
		
		 

		objWorkSheet.Cells(2,1).Font.Bold = False                    'need logo
		objWorkSheet.Cells(2,1) = ""
				
		objWorkSheet.Cells(3,1).Font.Bold = False
		objWorkSheet.Cells(3,1) = rst1("name")
        objWorkSheet.Cells(3,1).Interior.ColorIndex = 36 
						
		objWorkSheet.Cells(4,1).Font.Bold = False
		objWorkSheet.Cells(4,1) = rst1("strt")
        objWorkSheet.Cells(4,1).Interior.ColorIndex = 36 
						
		objWorkSheet.Cells(5,1).Font.Bold = False
		objWorkSheet.Cells(5,1) = rst1("city") & " , " & rst1("statedescr") & "," & rst1("zip")
        objWorkSheet.Cells(5,1).Interior.ColorIndex = 36 
						
		objWorkSheet.Cells(6,1).Font.Bold = False
		objWorkSheet.Cells(6,1) = "" 
		    
        objWorkSheet.Cells(7,1).Font.Bold = False
		objWorkSheet.Cells(7,1) = "Applicable Sales Tax Rate   8.875%" 
        objWorkSheet.Cells(7,1).Interior.ColorIndex = 36 
		    
        objWorkSheet.Cells(8,16).Font.Bold = False
		objWorkSheet.Cells(8,16) = "G/L Accounts (Revenue accts & Sales Tax Payable Acct)" 
						
		objWorkSheet.Cells(9,1).Font.Bold = False
		objWorkSheet.Cells(9,16) = "4061"
        objWorkSheet.Cells(9,17) = "4406"
        objWorkSheet.Cells(9,18) = "4406"
        objWorkSheet.Cells(9,19) = "2025"
        objWorkSheet.Cells(9,20) = "2025"
        objWorkSheet.Cells(9,21) = "2025"
        objWorkSheet.Cells(9,16).Interior.ColorIndex = 44
        objWorkSheet.Cells(9,17).Interior.ColorIndex = 44 
        objWorkSheet.Cells(9,18).Interior.ColorIndex = 44 
        objWorkSheet.Cells(9,19).Interior.ColorIndex = 40
        objWorkSheet.Cells(9,20).Interior.ColorIndex = 40
        objWorkSheet.Cells(9,21).Interior.ColorIndex = 40
        objWorkSheet.Cells(9,22) = ""
        objWorkSheet.Cells(9,23) = "Sales Tax Filing"
        objWorkSheet.Cells(9,24) = ""	
        objWorkSheet.Cells(9,22).Interior.ColorIndex = 43
        objWorkSheet.Cells(9,23).Interior.ColorIndex = 43
        objWorkSheet.Cells(9,24).Interior.ColorIndex = 43		
				
		objWorkSheet.Cells(10,1).Font.Bold = False
		objWorkSheet.Cells(10,1) = "Tenant" 
        objWorkSheet.Cells(10,2) = "X if Tax-Exempt  B if Base Building or ERI Load"
        objWorkSheet.Cells(10,3) = "ConEd Pre Tax"
        objWorkSheet.Cells(10,4) = "ESCO Pre Tax"
        objWorkSheet.Cells(10,5) = "Metered Energy Consumption (kWh)" 
        objWorkSheet.Cells(10,6) = "Energy Consumption (kWh) to be contrasted to Total Building Consumption for Sales Tax Reporting Purposes"
        objWorkSheet.Cells(10,7) = "Metered Consumption Allocation ($)"
        objWorkSheet.Cells(10,8) = "Admin Fees"
        objWorkSheet.Cells(10,9) = "Service Fees"
        objWorkSheet.Cells(10,10) = "Billed Consumption (Pre-Tax)" 
        objWorkSheet.Cells(10,11) = "Tax on Cost"
        objWorkSheet.Cells(10,12) = "Tax on Admin Fee"
        objWorkSheet.Cells(10,13) = "Tax on Service Fee"
        objWorkSheet.Cells(10,14) = "Total with Tax"
        objWorkSheet.Cells(10,15) = "Load Code" 
        objWorkSheet.Cells(10,16) = "Sub Meter Income (EM)"
        objWorkSheet.Cells(10,17) = "Admin Fee Income (AS)"
        objWorkSheet.Cells(10,18) = "Service Fee Income (SF)"
        objWorkSheet.Cells(10,19) = "Salex Tax Electric (SE)"
        objWorkSheet.Cells(10,20) = "Remit Tax on Admin Fee (SY)" 
        objWorkSheet.Cells(10,21) = "Remit Tax on Service Fee (ST)"
        objWorkSheet.Cells(10,22) = "Tax to be remitted"
        objWorkSheet.Cells(10,23) = "Tax not to be remitted - Recorded as JE against GL#5240 - Electric expense."
        objWorkSheet.Cells(10,24) = "Check Calc."
        objWorkSheet.Cells(10,16).Interior.ColorIndex = 44
        objWorkSheet.Cells(10,17).Interior.ColorIndex = 44
        objWorkSheet.Cells(10,18).Interior.ColorIndex = 44
        objWorkSheet.Cells(10,19).Interior.ColorIndex = 40
        objWorkSheet.Cells(10,20).Interior.ColorIndex = 40
        objWorkSheet.Cells(10,21).Interior.ColorIndex = 40
        objWorkSheet.Cells(10,22).Interior.ColorIndex = 43
        objWorkSheet.Cells(10,23).Interior.ColorIndex = 43
        objWorkSheet.Cells(10,24).Interior.ColorIndex = 43		
        				
				
	End if
	rst1.close

	iRow = 10
	sSql = "Exec usp_BuildingTaxReport_allbuildings " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("Tenantname")
        objWorkSheet.Cells(iRow,2) = rst1("xorb")
        objWorkSheet.Cells(iRow,3) = rst1("conedpretax")
        objWorkSheet.Cells(iRow,4) = rst1("escopretax")
        objWorkSheet.Cells(iRow,5) = rst1("meterkwh")
        objWorkSheet.Cells(iRow,6) = rst1("energykwh")
        objWorkSheet.Cells(iRow,7) = rst1("meterallocation")
        objWorkSheet.Cells(iRow,8) = rst1("adminfee")
        objWorkSheet.Cells(iRow,9) = rst1("servicefee")
        objWorkSheet.Cells(iRow,10) = rst1("billconsumption")
        objWorkSheet.Cells(iRow,11) = rst1("taxoncost")
        objWorkSheet.Cells(iRow,12) = rst1("taxonadminfee")
        objWorkSheet.Cells(iRow,13) = rst1("taxonservicefee")
        objWorkSheet.Cells(iRow,14) = rst1("totalwithtax")
        objWorkSheet.Cells(iRow,15) = rst1("loadcode")
        objWorkSheet.Cells(iRow,16) = rst1("RXRSubMeterIncomeEM")
        objWorkSheet.Cells(iRow,17) = rst1("RXRAdminFeeIncomeAS")
        objWorkSheet.Cells(iRow,18) = rst1("RXRServiceFeeIncomeAS")
        objWorkSheet.Cells(iRow,19) = rst1("RXRSalexTaxElectricSE")
        objWorkSheet.Cells(iRow,20) = rst1("RXRRemitTaxonAdminFeeSY")
        objWorkSheet.Cells(iRow,21) = rst1("RXRRemitTaxonServiceFeeSY")
        objWorkSheet.Cells(iRow,22) = rst1("Taxtoberemitted")
        objWorkSheet.Cells(iRow,23) = rst1("Taxnottoberemitted")
        objWorkSheet.Cells(iRow,24) = rst1("CheckCalc")
	    objWorkSheet.Cells(iRow,16).Interior.ColorIndex = 44
        objWorkSheet.Cells(iRow,17).Interior.ColorIndex = 44
        objWorkSheet.Cells(iRow,18).Interior.ColorIndex = 44
        objWorkSheet.Cells(iRow,19).Interior.ColorIndex = 40
        objWorkSheet.Cells(iRow,20).Interior.ColorIndex = 40
        objWorkSheet.Cells(iRow,21).Interior.ColorIndex = 40
        objWorkSheet.Cells(iRow,22).Interior.ColorIndex = 43
        objWorkSheet.Cells(iRow,23).Interior.ColorIndex = 43
        objWorkSheet.Cells(iRow,24).Interior.ColorIndex = 43	
						
		rst1.movenext
	loop
	rst1.close

    objWorkSheet.Cells(iRow,23).Font.ColorIndex = 3
	
	iRow= iRow + 1

    sSql = "Exec usp_BuildingTaxReportTotals_allbuildings " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("Tenantname")
        objWorkSheet.Cells(iRow,2) = rst1("xorb")
        objWorkSheet.Cells(iRow,3) = rst1("conedpretax")
        objWorkSheet.Cells(iRow,4) = rst1("escopretax")
        objWorkSheet.Cells(iRow,5) = rst1("meterkwh")
        objWorkSheet.Cells(iRow,6) = rst1("energykwh")
        objWorkSheet.Cells(iRow,7) = rst1("meterallocation")
        objWorkSheet.Cells(iRow,8) = rst1("adminfee")
        objWorkSheet.Cells(iRow,9) = rst1("servicefee")
        objWorkSheet.Cells(iRow,10) = rst1("billconsumption")
        objWorkSheet.Cells(iRow,11) = rst1("taxoncost")
        objWorkSheet.Cells(iRow,12) = rst1("taxonadminfee")
        objWorkSheet.Cells(iRow,13) = rst1("taxonservicefee")
        objWorkSheet.Cells(iRow,14) = rst1("totalwithtax")
        objWorkSheet.Cells(iRow,15) = rst1("loadcode")
        objWorkSheet.Cells(iRow,16) = rst1("RXRSubMeterIncomeEM")
        objWorkSheet.Cells(iRow,17) = rst1("RXRAdminFeeIncomeAS")
        objWorkSheet.Cells(iRow,18) = rst1("RXRServiceFeeIncomeAS")
        objWorkSheet.Cells(iRow,19) = rst1("RXRSalexTaxElectricSE")
        objWorkSheet.Cells(iRow,20) = rst1("RXRRemitTaxonAdminFeeSY")
        objWorkSheet.Cells(iRow,21) = rst1("RXRRemitTaxonServiceFeeSY")
        objWorkSheet.Cells(iRow,22) = rst1("Taxtoberemitted")
        objWorkSheet.Cells(iRow,23) = rst1("Taxnottoberemitted")
        objWorkSheet.Cells(iRow,24) = rst1("CheckCalc")
        if rst1("lineorder") = 7 then
           objWorkSheet.Cells(iRow,1).Font.ColorIndex = 8
        end if
        if rst1("lineorder") = 1 then
           objWorkSheet.Cells(iRow,14).Font.ColorIndex = 8
        end if
        if rst1("lineorder") = 8 then
           objWorkSheet.Cells(iRow,1).Font.ColorIndex = 7
        end if
        if rst1("lineorder") = 3 then
           objWorkSheet.Cells(iRow,20).Font.ColorIndex = 7
           objWorkSheet.Cells(iRow,21).Font.ColorIndex = 7
           objWorkSheet.Cells(iRow,23).Font.ColorIndex = 10
        end if
        if rst1("lineorder") = 9 then
           objWorkSheet.Cells(iRow,1).Font.ColorIndex = 3
        end if
        if rst1("lineorder") = 10 then
           objWorkSheet.Cells(iRow,1).Font.ColorIndex = 10
        end if
           
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1

    iRow= iRow + 1

    sSql = "Exec usp_BuildingTaxReportTotals2_allbuildings " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("Tenantname")
        objWorkSheet.Cells(iRow,2) = rst1("xorb")
        objWorkSheet.Cells(iRow,3) = rst1("conedpretax")
        objWorkSheet.Cells(iRow,4) = rst1("escopretax")
        objWorkSheet.Cells(iRow,5) = rst1("meterkwh")
        objWorkSheet.Cells(iRow,6) = rst1("energykwh")
        objWorkSheet.Cells(iRow,7) = rst1("meterallocation")
        objWorkSheet.Cells(iRow,8) = rst1("adminfee")
        objWorkSheet.Cells(iRow,9) = rst1("servicefee")
        objWorkSheet.Cells(iRow,10) = rst1("billconsumption")
        objWorkSheet.Cells(iRow,11) = rst1("taxoncost")
        objWorkSheet.Cells(iRow,12) = rst1("taxonadminfee")
        objWorkSheet.Cells(iRow,13) = rst1("taxonservicefee")
        objWorkSheet.Cells(iRow,14) = rst1("totalwithtax")
        objWorkSheet.Cells(iRow,15) = rst1("loadcode")
        objWorkSheet.Cells(iRow,16) = rst1("RXRSubMeterIncomeEM")
        objWorkSheet.Cells(iRow,17) = rst1("RXRAdminFeeIncomeAS")
        objWorkSheet.Cells(iRow,18) = rst1("RXRServiceFeeIncomeAS")
        objWorkSheet.Cells(iRow,19) = rst1("RXRSalexTaxElectricSE")
        objWorkSheet.Cells(iRow,20) = rst1("RXRRemitTaxonAdminFeeSY")
        objWorkSheet.Cells(iRow,21) = rst1("RXRRemitTaxonServiceFeeSY")
        objWorkSheet.Cells(iRow,22) = rst1("Taxtoberemitted")
        objWorkSheet.Cells(iRow,23) = rst1("Taxnottoberemitted")
        objWorkSheet.Cells(iRow,24) = rst1("CheckCalc")
        if rst1("lineorder") = 2 then
           objWorkSheet.Cells(iRow,9).Interior.ColorIndex = 36 
        end if
        if rst1("lineorder") = 4 then
           objWorkSheet.Cells(iRow,9).Interior.ColorIndex = 36
        end if
           
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    objWorkSheet.Rows("6:6").Select
    objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Billperiod  & UtilityId

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "BuildingTaxReport.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "BuildingTaxReport.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>BuildingTaxReport.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>BuildingTaxReport.xlsx</b></a> 
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
<Script type=text/javascript>
    function loadperiod() {
        var frm = document.forms['form1'];
        var newhref = "BuildingTaxReportNON450.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value
        document.location.href = newhref;
    }

    function loadutility() {
        var frm = document.forms['form1'];
        var newhref = "BuildingTaxReportNON450.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value;
        document.location.href = newhref;
    }
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