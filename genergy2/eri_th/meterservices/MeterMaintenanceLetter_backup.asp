<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

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
<title>Meter Maintenance Letter</title>

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
   <form name="form1" action="MeterMaintenanceLetter.asp">
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
				        "billyear>=year(getdate())-1 and bldgnum='"&building&"' and utility="&utilityid&" order by billyear desc "
				        
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
				sql = sql & "bldgnum='"&building&"' and utility="&utilityid& " order by billperiod desc "
					
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
	
	sSql = "Exec usp_MaintenanceLetterTenants " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Book Antiqua"
	objWorkSheet.Cells.Font.Size = 9

' Column constants
Const COL_FLOOR = 1
Const COL_TENANT = 2
Const COL_TENANT_SQFT = 3
Const COL_TENANT_STATUS = 4
Const COL_LEASE_START = 5
Const COL_MOVEIN_DATE = 6
Const COL_LEASEEXP_DATE = 7
Const COL_TENANT_METER_CNT = 8
Const COL_METER_ID = 9 
Const COL_METER_CURR_READ_STATUS = 10 
Const COL_METER_SERIAL_NUM = 11
Const COL_METER_MAKE = 12
Const COL_METER_MODEL = 13
Const COL_METER_FUNCTION = 14
Const COL_METER_TYPE = 15
Const COL_METER_LOCATION_FLR = 16
Const COL_METER_LOCATION_ARA = 17
Const COL_METER_STARTUP_DATE = 18 
Const COL_METER_READ_TYPE = 19
Const COL_DATA_COLL_NUM = 20
Const COL_DAT_CHANNEL_NUM = 21
Const COL_METER_READ_STATUS = 22
Const COL_METER_BILL_STATUS = 23
Const COL_METER_USAGE = 24
Const COL_METER_DEMAND = 25
Const COL_TENANT_WPSQFT_EXP = 26
Const COL_TENANT_WPSQFT_ACT = 27 
Const COL_USG3MTH_VAR_RNG = 28
Const COL_USG3MTH_VAR = 29
Const COL_DMD3MTH_VAR_RNG = 30
Const COL_DMD3MTH_VAR = 31
Const COL_USGLSTPER_VAR_RNG = 32
Const COL_USGLSTPER_VAR = 33
Const COL_DMDLSTPER_VAR_RNG = 34
Const COL_DMDLSTPER_VAR = 35
Const COL_USGLSTYR_VAR_RNG = 36
Const COL_USGLSTYR_VAR = 37
Const COL_DMDLSTYR_VAR_RNG = 38
Const COL_DMDLSTY_VAR = 39
Const COL_CUSTOMER_NOTES = 40
Const COL_TENANT_CHARGES = 41 
Const COL_ADMIN_FEES = 42
Const COL_SALES_TX = 43
Const COL_MISC_CREDITS = 44
Const COL_TENANT_TOTAL_CHGS = 45


' Header Columns	
	If not rst1.eof then
		
		objWorkSheet.Cells(1,1).Font.Bold = True
		objWorkSheet.Cells(1,1) = building & " Maintenance Letter For Period : " & Billperiod & "," & Billyear 
		objWorkSheet.Cells(1,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(1,2).Interior.ColorIndex = 40 

		objWorkSheet.Cells(2,COL_FLOOR).Font.Bold = True
		objWorkSheet.Cells(2,COL_FLOOR) ="Floor"
		objWorkSheet.Cells(2,COL_FLOOR).Interior.ColorIndex = 36 
		objWorkSheet.Cells(2,COL_FLOOR).Borders.LineStyle = 1 
		
		objWorkSheet.Cells(2,COL_TENANT ).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT) ="Tenant"
		objWorkSheet.Cells(2,COL_TENANT ).Interior.ColorIndex = 36 
		
		
		objWorkSheet.Cells(2,COL_TENANT_SQFT).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_SQFT) ="Tenant SqFt."
		objWorkSheet.Cells(2,COL_TENANT_SQFT).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_TENANT_STATUS).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_STATUS) ="Tenant Status"
		objWorkSheet.Cells(2,COL_TENANT_STATUS).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_LEASE_START ).Font.Bold = True
		objWorkSheet.Cells(2,COL_LEASE_START ) ="Lease Start Date"
		objWorkSheet.Cells(2,COL_LEASE_START).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_MOVEIN_DATE ).Font.Bold = True
		objWorkSheet.Cells(2,COL_MOVEIN_DATE ) ="Move-In Date"
		objWorkSheet.Cells(2,COL_MOVEIN_DATE).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_LEASEEXP_DATE).Font.Bold = True
		objWorkSheet.Cells(2,COL_LEASEEXP_DATE) ="Lease Expiry Date"
		objWorkSheet.Cells(2,COL_LEASEEXP_DATE).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_TENANT_METER_CNT).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_METER_CNT) ="Number Of Tenant Meters"	
		objWorkSheet.Cells(2,COL_TENANT_METER_CNT).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_ID ).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_ID) ="Meter Id #"
		objWorkSheet.Cells(2,COL_METER_ID).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_CURR_READ_STATUS ).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_CURR_READ_STATUS) ="Current Reading"
		objWorkSheet.Cells(2,COL_METER_CURR_READ_STATUS).Interior.ColorIndex = 36		
		
		objWorkSheet.Cells(2,COL_METER_SERIAL_NUM).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_SERIAL_NUM) ="Meter Serial #"
		objWorkSheet.Cells(2,COL_METER_SERIAL_NUM).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_MAKE).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_MAKE) ="Meter Make"
		objWorkSheet.Cells(2,COL_METER_MAKE).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_METER_MODEL).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_MODEL) ="Meter Model"		
		objWorkSheet.Cells(2,COL_METER_MODEL).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_FUNCTION).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_FUNCTION) ="Meter Function"		
		objWorkSheet.Cells(2,COL_METER_FUNCTION).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_METER_TYPE).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_TYPE) ="Meter Type"		
		objWorkSheet.Cells(2,COL_METER_TYPE).Interior.ColorIndex = 36				
					
		objWorkSheet.Cells(2,COL_METER_LOCATION_FLR).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_LOCATION_FLR) ="Meter Location (Floor)"
		objWorkSheet.Cells(2,COL_METER_LOCATION_FLR).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_LOCATION_ARA).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_LOCATION_ARA) ="Meter Location (Area)"
		objWorkSheet.Cells(2,COL_METER_LOCATION_ARA).Interior.ColorIndex = 36
		

		objWorkSheet.Cells(2,COL_METER_STARTUP_DATE).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_STARTUP_DATE) ="Meter StartUp Date"
		objWorkSheet.Cells(2,COL_METER_STARTUP_DATE).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_METER_READ_TYPE).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_READ_TYPE) ="Meter Read Type"
		objWorkSheet.Cells(2,COL_METER_READ_TYPE).Interior.ColorIndex = 36			

		objWorkSheet.Cells(2,COL_DATA_COLL_NUM ).Font.Bold = True
		objWorkSheet.Cells(2,COL_DATA_COLL_NUM) ="Data Coll. Number"
		objWorkSheet.Cells(2,COL_DATA_COLL_NUM).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_DAT_CHANNEL_NUM ).Font.Bold = True
		objWorkSheet.Cells(2,COL_DAT_CHANNEL_NUM) ="Communication Channel #"		
		objWorkSheet.Cells(2,COL_DAT_CHANNEL_NUM).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_READ_STATUS).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_READ_STATUS) ="Meter Reading Status"
		objWorkSheet.Cells(2,COL_METER_READ_STATUS).Interior.ColorIndex = 36		

		objWorkSheet.Cells(2,COL_METER_BILL_STATUS).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_BILL_STATUS) ="Meter Billing Status"
		objWorkSheet.Cells(2,COL_METER_BILL_STATUS).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_METER_USAGE).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_USAGE) ="Usage"
		objWorkSheet.Cells(2,COL_METER_USAGE).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_METER_DEMAND).Font.Bold = True
		objWorkSheet.Cells(2,COL_METER_DEMAND) ="Demand"
		objWorkSheet.Cells(2,COL_METER_DEMAND).Interior.ColorIndex = 36		
		
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_EXP).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_EXP) ="Watts/Square/Foot"
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_EXP).Interior.ColorIndex = 36		
		
		objWorkSheet.Cells(3,COL_TENANT_WPSQFT_EXP).Font.Bold = True
		objWorkSheet.Cells(3,COL_TENANT_WPSQFT_EXP) ="Expected Range"
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_EXP).Interior.ColorIndex = 36		

		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_ACT).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_ACT) ="Watts/Square/Foot"
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_ACT).Interior.ColorIndex = 36			

		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_ACT).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_ACT) ="Actual"			
		objWorkSheet.Cells(2,COL_TENANT_WPSQFT_ACT).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_USG3MTH_VAR_RNG ).Font.Bold = True
		objWorkSheet.Cells(2,COL_USG3MTH_VAR_RNG) ="3 Month Avg KWH "
		objWorkSheet.Cells(2,COL_USG3MTH_VAR_RNG).Interior.ColorIndex = 36

		objWorkSheet.Cells(3,COL_USG3MTH_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(3,COL_USG3MTH_VAR_RNG) ="Variance Normal Range (%)"		
		objWorkSheet.Cells(3,COL_USG3MTH_VAR_RNG).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_USG3MTH_VAR).Font.Bold = True
		objWorkSheet.Cells(2,COL_USG3MTH_VAR) ="3 Month Avg KWH "
		objWorkSheet.Cells(2,COL_USG3MTH_VAR).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(3,COL_USG3MTH_VAR).Font.Bold = True
		objWorkSheet.Cells(3,COL_USG3MTH_VAR) ="Actual Variance (%)"
		objWorkSheet.Cells(2,COL_USG3MTH_VAR).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_DMD3MTH_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(2,COL_DMD3MTH_VAR_RNG) ="3 Month Avg KW "
		objWorkSheet.Cells(2,COL_DMD3MTH_VAR_RNG).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(3,COL_DMD3MTH_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(3,COL_DMD3MTH_VAR_RNG) ="Variance Normal Range (%)"
		objWorkSheet.Cells(2,COL_DMD3MTH_VAR_RNG).Interior.ColorIndex = 36

		objWorkSheet.Cells(2,COL_DMD3MTH_VAR).Font.Bold = True
		objWorkSheet.Cells(2,COL_DMD3MTH_VAR) ="3 Month Avg KW"
		objWorkSheet.Cells(2,COL_DMD3MTH_VAR).Interior.ColorIndex = 36

		objWorkSheet.Cells(3,COL_DMD3MTH_VAR).Font.Bold = True
		objWorkSheet.Cells(3,COL_DMD3MTH_VAR) ="Actual Variance (%)"
		objWorkSheet.Cells(2,COL_DMD3MTH_VAR).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_USGLSTPER_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(2,COL_USGLSTPER_VAR_RNG) ="Last Period Avg KWH "
		objWorkSheet.Cells(2,COL_USGLSTPER_VAR_RNG).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(3,COL_USGLSTPER_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(3,COL_USGLSTPER_VAR_RNG) ="Variance Normal Range (%)"
		objWorkSheet.Cells(3,COL_USGLSTPER_VAR_RNG).Interior.ColorIndex = 36		

		objWorkSheet.Cells(2,COL_USGLSTPER_VAR).Font.Bold = True
		objWorkSheet.Cells(2,COL_USGLSTPER_VAR) ="Last Period Avg KWH "
		objWorkSheet.Cells(2,COL_USGLSTPER_VAR).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(3,COL_USGLSTPER_VAR).Font.Bold = True
		objWorkSheet.Cells(3,COL_USGLSTPER_VAR) ="Actual Variance (%)"		
		objWorkSheet.Cells(3,COL_USGLSTPER_VAR).Interior.ColorIndex = 36
		
	
		objWorkSheet.Cells(2,COL_DMDLSTPER_VAR_RNG ).Font.Bold = True
		objWorkSheet.Cells(2,COL_DMDLSTPER_VAR_RNG) ="Last period Avg KW "	
		objWorkSheet.Cells(2,COL_DMDLSTPER_VAR_RNG).Interior.ColorIndex = 36
		

		objWorkSheet.Cells(3,COL_DMDLSTPER_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(3,COL_DMDLSTPER_VAR_RNG) ="Variance Normal Range (%)"	
		objWorkSheet.Cells(3,COL_DMDLSTPER_VAR_RNG).Interior.ColorIndex = 36
				
		objWorkSheet.Cells(2,COL_DMDLSTPER_VAR).Font.Bold = True
		objWorkSheet.Cells(2,COL_DMDLSTPER_VAR) ="Last Period Avg KW "
		objWorkSheet.Cells(2,COL_DMDLSTPER_VAR).Interior.ColorIndex = 36

		objWorkSheet.Cells(3,COL_DMDLSTPER_VAR).Font.Bold = True
		objWorkSheet.Cells(3,COL_DMDLSTPER_VAR) ="Actual Variance (%)"
		objWorkSheet.Cells(3,COL_DMDLSTPER_VAR).Interior.ColorIndex = 36
							
		objWorkSheet.Cells(2,COL_USGLSTYR_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(2,COL_USGLSTYR_VAR_RNG) ="Last year Avg KWH "
		objWorkSheet.Cells(2,COL_USGLSTYR_VAR_RNG).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(3,COL_USGLSTYR_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(3,COL_USGLSTYR_VAR_RNG) ="Variance Normal Range (%)"	
		objWorkSheet.Cells(3,COL_USGLSTYR_VAR_RNG).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_USGLSTYR_VAR).Font.Bold = True
		objWorkSheet.Cells(2,COL_USGLSTYR_VAR) ="Last year Avg KWH"
		objWorkSheet.Cells(2,COL_USGLSTYR_VAR).Interior.ColorIndex = 36

		objWorkSheet.Cells(3,COL_USGLSTYR_VAR).Font.Bold = True
		objWorkSheet.Cells(3,COL_USGLSTYR_VAR) ="Actual Variance (%)"
		objWorkSheet.Cells(3,COL_USGLSTYR_VAR).Interior.ColorIndex = 36
				
		objWorkSheet.Cells(2,COL_DMDLSTYR_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(2,COL_DMDLSTYR_VAR_RNG) ="Last year Avg KW"		
		objWorkSheet.Cells(2,COL_DMDLSTYR_VAR_RNG).Interior.ColorIndex = 36

		objWorkSheet.Cells(3,COL_DMDLSTYR_VAR_RNG).Font.Bold = True
		objWorkSheet.Cells(3,COL_DMDLSTYR_VAR_RNG) ="Variance Normal Range (%)"		
		objWorkSheet.Cells(3,COL_DMDLSTYR_VAR_RNG).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_DMDLSTY_VAR).Font.Bold = True
		objWorkSheet.Cells(2,COL_DMDLSTY_VAR) ="Last year Avg KW "
		objWorkSheet.Cells(2,COL_DMDLSTY_VAR).Interior.ColorIndex = 36		

		objWorkSheet.Cells(3,COL_DMDLSTY_VAR).Font.Bold = True
		objWorkSheet.Cells(3,COL_DMDLSTY_VAR) ="Actual Variance (%)"		
		objWorkSheet.Cells(3,COL_DMDLSTY_VAR).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_CUSTOMER_NOTES ).Font.Bold = True
		objWorkSheet.Cells(2,COL_CUSTOMER_NOTES) ="Notes"
		objWorkSheet.Cells(2,COL_CUSTOMER_NOTES).Interior.ColorIndex = 36				

		objWorkSheet.Cells(2,COL_TENANT_CHARGES ).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_CHARGES) ="Tenant Charges"
		objWorkSheet.Cells(2,COL_TENANT_CHARGES).Interior.ColorIndex = 36
		
		objWorkSheet.Cells(2,COL_ADMIN_FEES).Font.Bold = True
		objWorkSheet.Cells(2,COL_ADMIN_FEES) ="Admin Fees "	
		objWorkSheet.Cells(2,COL_ADMIN_FEES).Interior.ColorIndex = 36	

		objWorkSheet.Cells(3,COL_ADMIN_FEES).Font.Bold = True
		objWorkSheet.Cells(3,COL_ADMIN_FEES) ="and Other Charges"	
		objWorkSheet.Cells(3,COL_ADMIN_FEES).Interior.ColorIndex = 36	

		objWorkSheet.Cells(2,COL_SALES_TX ).Font.Bold = True
		objWorkSheet.Cells(2,COL_SALES_TX) ="Sales Tax"
		objWorkSheet.Cells(2,COL_SALES_TX).Interior.ColorIndex = 36			

		objWorkSheet.Cells(2,COL_MISC_CREDITS ).Font.Bold = True
		objWorkSheet.Cells(2,COL_MISC_CREDITS) ="Misc. Credits"
		objWorkSheet.Cells(2,COL_MISC_CREDITS).Interior.ColorIndex = 36			

		objWorkSheet.Cells(2,COL_TENANT_TOTAL_CHGS ).Font.Bold = True
		objWorkSheet.Cells(2,COL_TENANT_TOTAL_CHGS) ="Tenant Total Charges"			
		objWorkSheet.Cells(2,COL_TENANT_TOTAL_CHGS).Interior.ColorIndex = 36
				
	End if

	iRow = 3
	Dim dUsage, dDemand
	dUsage = 0.0
	dDemand = 0.0
	
	Do Until rst1.eof
	
		iRow= iRow + 1
		
		' Tenant Header row
		objWorkSheet.Cells(iRow,COL_FLOOR) = rst1("Floor")
		objWorkSheet.Cells(iRow,COL_TENANT) = rst1("TenantName")
		objWorkSheet.Cells(iRow,COL_TENANT_SQFT) = rst1("sqft")	
		
		If not IsNull(rst1("sqft")) Then 
			TotalSqFt = TotalSqFt + rst1("Sqft")
		End If
		
		objWorkSheet.Cells(iRow,COL_TENANT_STATUS) = rst1("TenantStatus")	
		objWorkSheet.Cells(iRow,COL_LEASE_START) = rst1("LeaseStartDate")	
		objWorkSheet.Cells(iRow,COL_LEASE_START).NumberFormat = "mm/dd/yyyy"
		
		If not Isnull(rst1("LeaseExpiryDate")) then 
			objWorkSheet.Cells(iRow,COL_LEASEEXP_DATE) = rst1("LeaseExpiryDate")
			objWorkSheet.Cells(iRow,COL_LEASEEXP_DATE).NumberFormat = "mm/dd/yyyy"
		End IF
		If not Isnull(rst1("MoveInDate")) then 
			objWorkSheet.Cells(iRow,COL_MOVEIN_DATE) = rst1("MoveInDate")
			objWorkSheet.Cells(iRow,COL_MOVEIN_DATE).NumberFormat = "mm/dd/yyyy"
		End IF
		
		objWorkSheet.Cells(iRow,COL_TENANT_METER_CNT) = rst1("MeterCount")
		
		MeterCountTotal = MeterCountTotal + rst1("MeterCount")	


		if not (isNull(rst1("WattsPerSQFTLowLimit")) or isNull(rst1("WattsPerSQFTHighLimit"))) then
			objWorkSheet.Cells(iRow,COL_TENANT_WPSQFT_EXP) = rst1("WattsPerSQFTLowLimit") & " to " & rst1("WattsPerSQFTHighLimit")
		end if
			
		objWorkSheet.Cells(iRow,COL_TENANT_WPSQFT_ACT) = rst1("WattsPerSqft")
			
		If not IsNull(rst1("WattsPerSQFTLowLimit")) Then
			If CDbl(rst1("WattsPerSqft")) < Cdbl(rst1("WattsPerSQFTLowLimit")) Then
				objWorkSheet.Cells(iRow,COL_TENANT_WPSQFT_ACT).Font.ColorIndex = 3
			End If
		End If

		If not IsNull(rst1("WattsPerSQFTHighLimit")) Then
			If CDbl(rst1("WattsPerSqft")) > CDbl(rst1("WattsPerSQFTHighLimit")) Then
				objWorkSheet.Cells(iRow,COL_TENANT_WPSQFT_ACT).Font.ColorIndex = 3
			End If
		End If	

		If Not Isnull(rst1("Subtotal") ) Then 
			objWorkSheet.Cells(iRow,COL_TENANT_CHARGES) = rst1("Subtotal")	
			TenantChargesTotal = Cdbl(TenantChargesTotal) + CDbl(rst1("Subtotal"))	
			objWorkSheet.Cells(iRow,COL_TENANT_CHARGES).NumberFormat = "$#,##0.00"	
		ENd If

		If Not Isnull(rst1("AdminFee") ) Then 
			objWorkSheet.Cells(iRow,COL_ADMIN_FEES) = rst1("AdminFee")	
			AdminFeesTotal = Cdbl(AdminFeesTotal) + Cdbl(rst1("AdminFee"))	
			objWorkSheet.Cells(iRow,COL_ADMIN_FEES).NumberFormat = "$#,##0.00"		
		ENd If
		
		If Not Isnull(rst1("SalesTax") ) Then 
			objWorkSheet.Cells(iRow,COL_SALES_TX) = rst1("SalesTax")		
			SalesTaxTotal = Cdbl(SalesTaxTotal) + Cdbl(rst1("SalesTax"))
			objWorkSheet.Cells(iRow,COL_SALES_TX).NumberFormat = "$#,##0.00"
		ENd If

		If Not Isnull(rst1("Credit") ) Then 
			objWorkSheet.Cells(iRow,COL_MISC_CREDITS) = rst1("Credit")
			MiscCreditsTotal = Cdbl(MiscCreditsTotal) + Cdbl(rst1("Credit"))		
		ENd If

		If Not Isnull(rst1("TotalAmt") ) Then 
			objWorkSheet.Cells(iRow,COL_TENANT_TOTAL_CHGS) = rst1("TotalAmt")	
			BuildingChargesTotal = Cdbl(BuildingChargesTotal) + Cdbl(rst1("TotalAmt"))	
			objWorkSheet.Cells(iRow,COL_TENANT_TOTAL_CHGS).NumberFormat = "$#,##0.00"	
		ENd If
		
		set rst2 = server.createobject("ADODB.Recordset") 
		
		sSql = "Exec usp_MaintenanceLetterMeters " & rst1("LeaseUtilityId") & "," & Billyear & ",'" & BillPeriod & "'" & ", " & UtilityId 
		rst2.CursorLocation = 3
		rst2.open sSql , cnn1, 3 

		Do While Not rst2.EOF 
		
			iRow= iRow + 1
			
			objWorkSheet.Cells(iRow,COL_METER_ID) = rst2("MeterNo")	
			
			objWorkSheet.Cells(iRow,COL_METER_CURR_READ_STATUS) = rst2("ReadStatus")	
			
			If not IsNull(rst2("SerialNumber")) Then 
				objWorkSheet.Cells(iRow,COL_METER_SERIAL_NUM) = rst2("SerialNumber")	
			End If
			If not IsNull(rst2("MeterMake")	) Then 
				objWorkSheet.Cells(iRow,COL_METER_MAKE) = rst2("MeterMake")	
			End IF
			If not IsNull(rst2("MeterModel")) Then 
				objWorkSheet.Cells(iRow,COL_METER_MODEL) = rst2("MeterModel")	
			End If
			
			If not IsNull(rst2("MeterFunction")) Then 
				objWorkSheet.Cells(iRow,COL_METER_FUNCTION) = rst2("MeterFunction")	
			End If
		
			If not IsNull(rst2("MeterType")) Then 
				objWorkSheet.Cells(iRow,COL_METER_TYPE) = rst2("MeterType")	
			End If		
						
			objWorkSheet.Cells(iRow,COL_METER_LOCATION_FLR) = rst2("floor")	
			objWorkSheet.Cells(iRow,COL_METER_LOCATION_ARA) = rst2("MeterLocation")	

			If not IsNull(rst2("StartupDate")) Then 
				objWorkSheet.Cells(iRow,COL_METER_STARTUP_DATE) = rst2("StartupDate")	
				objWorkSheet.Cells(iRow,COL_METER_STARTUP_DATE).NumberFormat = "mm/dd/yyyy"
			End If
						
			objWorkSheet.Cells(iRow,COL_METER_READ_TYPE) = rst2("MeterReadType")	
			
			If not IsNull(rst2("DataCollectorNumber")) Then 
				objWorkSheet.Cells(iRow,COL_DATA_COLL_NUM) = rst2("DataCollectorNumber")	
			End If

			If not IsNull(rst2("ChannelNumber")) Then 
				objWorkSheet.Cells(iRow,COL_DAT_CHANNEL_NUM) = rst2("ChannelNumber")	
			End If
			
			objWorkSheet.Cells(iRow,COL_METER_READ_STATUS) = rst2("MeterStatus")
			
			objWorkSheet.Cells(iRow,COL_METER_BILL_STATUS) = rst2("MeterStatus")
			
			If Isnull(rst2("used")) Then
				objWorkSheet.Cells(iRow,COL_METER_USAGE)  = "-"
			Else
				objWorkSheet.Cells(iRow,COL_METER_USAGE)  = rst2("used")
				UsageTotal = Cdbl(UsageTotal) + Cdbl(rst2("used"))
			End If
			
			objWorkSheet.Cells(iRow,COL_METER_DEMAND) = " " & rst2("demand")
			
			If Not IsNull(rst2("demand")) Then
				objWorkSheet.Cells(iRow,COL_METER_DEMAND) = rst2("demand")
				DemandTotal = Cdbl(DemandTotal) + Cdbl(rst2("demand"))
			Else
				objWorkSheet.Cells(iRow,COL_METER_DEMAND) = "-"
			End If

			
			
			if not (isNull(rst2("Usage3MonthLowLimit")) or isNull(rst2("Usage3MonthHighLimit"))) then
				objWorkSheet.Cells(iRow,COL_USG3MTH_VAR_RNG) = rst2("Usage3MonthLowLimit") & " to " & rst2("Usage3MonthHighLimit")
			end if
			
			objWorkSheet.Cells(iRow,COL_USG3MTH_VAR) = rst2("3MnthAvgKWHVar")
			
			If not IsNull(rst2("Usage3MonthLowLimit")) Then
				If CDbl(rst2("3MnthAvgKWHVar")) < Cdbl(rst2("Usage3MonthLowLimit")) Then
					objWorkSheet.Cells(iRow,COL_USG3MTH_VAR).Font.ColorIndex = 3
				End If
			End If

			If not IsNull(rst2("Usage3MonthHighLimit")) Then
				If CDbl(rst2("3MnthAvgKWHVar")) > CDbl(rst2("Usage3MonthHighLimit")) Then
					objWorkSheet.Cells(iRow,COL_USG3MTH_VAR).Font.ColorIndex = 3
				End If
			End If			
			
			if not (isNull(rst2("Demand3MonthLowLimit")) or isNull(rst2("Demand3MonthHighLimit"))) then
				objWorkSheet.Cells(iRow,COL_DMD3MTH_VAR_RNG) = rst2("Demand3MonthLowLimit") & " to " & rst2("Demand3MonthHighLimit")
			end if
			
			objWorkSheet.Cells(iRow,COL_DMD3MTH_VAR) = rst2("3MnthAvgKWVar")
			
			If not IsNull(rst2("Demand3MonthLowLimit")) Then
				If CDbl(rst2("Demand3MonthLowLimit")) > CDbl(rst2("3MnthAvgKWVar")) Then
					objWorkSheet.Cells(iRow,COL_DMD3MTH_VAR).Font.ColorIndex = 3
				End If
			End If

			If not IsNull(rst2("Usage3MonthHighLimit")) Then
				If CDbl(rst2("3MnthAvgKWVar")) > CDbl(rst2("Demand3MonthHighLimit")) Then
					objWorkSheet.Cells(iRow,COL_DMD3MTH_VAR).Font.ColorIndex = 3
				End If
			End If				
		
			if not (isNull(rst2("UsageLastMonthLowLimit")) or isNull(rst2("UsageLastMonthHighLimit"))) then
				objWorkSheet.Cells(iRow,COL_USGLSTPER_VAR_RNG) = rst2("UsageLastMonthLowLimit") & " to " & rst2("UsageLastMonthHighLimit")
			end if
			
			objWorkSheet.Cells(iRow,COL_USGLSTPER_VAR) = rst2("LastBPKWHVar")

			If not IsNull(rst2("UsageLastMonthLowLimit")) Then
				If CDbl(rst2("UsageLastMonthLowLimit")) > CDbl(rst2("LastBPKWHVar")) Then
					objWorkSheet.Cells(iRow,COL_USGLSTPER_VAR).Font.ColorIndex = 3
				End If
			End If

			If not IsNull(rst2("UsageLastMonthHighLimit")) Then
				If CDbl(rst2("LastBPKWHVar")) > CDbl(rst2("UsageLastMonthHighLimit")) Then
					objWorkSheet.Cells(iRow,COL_USGLSTPER_VAR).Font.ColorIndex = 3
				End If
			End If
			
			if not (isNull(rst2("DemandLastMonthLowLimit")) or isNull(rst2("DemandLastMonthHighLimit"))) then
				objWorkSheet.Cells(iRow,COL_DMDLSTPER_VAR_RNG) = rst2("DemandLastMonthLowLimit") & " to " & rst2("DemandLastMonthHighLimit")
			end if
			
			objWorkSheet.Cells(iRow,COL_DMDLSTPER_VAR) = rst2("LastBPKWVar")				
			
			If not IsNull(rst2("DemandLastMonthLowLimit")) Then
				If CDbl(rst2("DemandLastMonthLowLimit")) > CDbl(rst2("LastBPKWVar")) Then
					objWorkSheet.Cells(iRow,COL_DMDLSTPER_VAR).Font.ColorIndex = 3
				End If
			End If

			If not IsNull(rst2("DemandLastMonthHighLimit")) Then
				If CDbl(rst2("LastBPKWVar")) > CDbl(rst2("DemandLastMonthHighLimit")) Then
					objWorkSheet.Cells(iRow,COL_DMDLSTPER_VAR).Font.ColorIndex = 3
				End If
			End If
						

			if not (isNull(rst2("UsageLastYrPeriodLowLimit")) or isNull(rst2("UsageLastYrPeriodHighLimit"))) then
				objWorkSheet.Cells(iRow,COL_USGLSTYR_VAR_RNG) = rst2("UsageLastYrPeriodLowLimit") & " to " & rst2("UsageLastYrPeriodHighLimit")
			end if
			
			objWorkSheet.Cells(iRow,COL_USGLSTYR_VAR) = rst2("LastYearKWHVar")								
			
			If not IsNull(rst2("UsageLastYrPeriodLowLimit")) Then
				If CDbl(rst2("UsageLastYrPeriodLowLimit")) > CDbl(rst2("LastYearKWHVar")) Then
					objWorkSheet.Cells(iRow,COL_USGLSTYR_VAR).Font.ColorIndex = 3
				End If
			End If

			If not IsNull(rst2("UsageLastYrPeriodHighLimit")) Then
				If CDbl(rst2("LastYearKWHVar")) > CDbl(rst2("UsageLastYrPeriodHighLimit")) Then
					objWorkSheet.Cells(iRow,COL_USGLSTYR_VAR).Font.ColorIndex = 3
				End If
			End If			

			if not (isNull(rst2("DemandLastYrPeriodLowLimit")) or isNull(rst2("DemandLastYrPeriodHighLimit"))) then
				objWorkSheet.Cells(iRow,COL_DMDLSTYR_VAR_RNG) = rst2("DemandLastYrPeriodLowLimit") & " to " & rst2("DemandLastYrPeriodHighLimit")
			end if
			objWorkSheet.Cells(iRow,COL_DMDLSTY_VAR) = rst2("LastYearKWVar")		

			If not IsNull(rst2("DemandLastYrPeriodLowLimit")) Then
				If CDbl(rst2("DemandLastYrPeriodLowLimit")) > CDbl(rst2("LastYearKWVar")) Then
					objWorkSheet.Cells(iRow,COL_DMDLSTY_VAR).Font.ColorIndex = 3
				End If
			End If

			If not IsNull(rst2("DemandLastYrPeriodHighLimit")) Then
				If CDbl(rst2("LastYearKWVar")) > CDbl(rst2("DemandLastYrPeriodHighLimit")) Then
					objWorkSheet.Cells(iRow,COL_DMDLSTY_VAR).Font.ColorIndex = 3
				End If
			End If				
			
									
			objWorkSheet.Cells(iRow,COL_CUSTOMER_NOTES) = rst2("UserNote")		
			
			rst2.MoveNext 
		Loop
		rst2.Close
		set rst2 = Nothing		
		rst1.movenext
	loop
	rst1.close

    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    objWorkSheet.Rows("4:4").Select
    objExcelReport.ActiveWindow.FreezePanes = True

	' Property Totals
	iRow = iRow + 1
	objWorkSheet.Cells(iRow,COL_FLOOR) = "Property Totals" 
	objWorkSheet.Cells(iRow,COL_TENANT_SQFT) = TotalSqFt 
	objWorkSheet.Cells(iRow,COL_TENANT_METER_CNT) = MeterCountTotal 
	objWorkSheet.Cells(iRow,COL_METER_USAGE) = UsageTotal 
	objWorkSheet.Cells(iRow,COL_METER_DEMAND) = DemandTotal 
	objWorkSheet.Cells(iRow,COL_TENANT_CHARGES) = TenantChargesTotal 
	objWorkSheet.Cells(iRow,COL_ADMIN_FEES) = AdminFeesTotal
	objWorkSheet.Cells(iRow,COL_SALES_TX) = SalesTaxTotal
	objWorkSheet.Cells(iRow,COL_MISC_CREDITS) = MiscCreditsTotal
	objWorkSheet.Cells(iRow,COL_TENANT_TOTAL_CHGS) = BuildingChargesTotal 
	
	
	' Utility Company Totals
	sSql = "SELECT UB.TotalBillAmt, UB.TotalKWH, UB.TotalKW " & _
			" FROM UtilityBill UB, BillYrPeriod BYP " & _
			" WHERE BYP.BillYear = " & Billyear & " and BYP.BillPeriod = " & Billperiod & _
			" and BYP.Utility = " & UtilityId & _
			" AND BYP.BldgNum = '" & building & "' " & _
			" AND UB.YpId = BYP.YpId" 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 	
	
	If not rst1.EOF then 
		iRow = iRow  + 1
		objWorkSheet.Cells(iRow,COL_FLOOR) = "Utility Company Totals" 
		If not Isnull(rst1("TotalKWH")) then 
			objWorkSheet.Cells(iRow,COL_METER_USAGE) = rst1("TotalKWH") 
		End If
		If not IsnUll(rst1("TotalKW")) Then 
			objWorkSheet.Cells(iRow,COL_METER_DEMAND) = rst1("TotalKW")
		End If
		If not Isnull(rst1("TotalBillAmt")) Then 
			objWorkSheet.Cells(iRow,COL_TENANT_TOTAL_CHGS) = rst1("TotalBillAmt")
			objWorkSheet.Cells(iRow,COL_TENANT_TOTAL_CHGS).NumberFormat = "$#,##0.00"
		End IF
	
	' Recovery Factors
		iRow = iRow + 1
		objWorkSheet.Cells(iRow,COL_FLOOR) = "Recovery Factors" 
		If not Isnull(rst1("TotalKWH")) then 
			objWorkSheet.Cells(iRow,COL_METER_USAGE) = (UsageTotal / Cdbl(rst1("TotalKWH"))) * 100 
		End If
		
		If not IsnUll(rst1("TotalKW")) Then 
			objWorkSheet.Cells(iRow,COL_METER_DEMAND) = (DemandTotal /  Cdbl(rst1("TotalKW"))) * 100
		End If
		If not Isnull(rst1("TotalBillAmt")) Then 
			objWorkSheet.Cells(iRow,COL_TENANT_TOTAL_CHGS) = (BuildingChargesTotal / Cdbl(rst1("TotalBillAmt"))) * 100
		End IF
		
	End IF
	
	iRow = iRow + 2
	

	objWorkSheet.Cells(iRow,COL_FLOOR) = "This maintenance report provides valuable information about your " & _
										 " property's utility submetering and overall energy cost recovery.  " & _
										 " Please review this report carefully, and pay special attention to the " & _
										 " values that fall outside expected predefined ranges (these appear in red text). " & vbCr &  _
										 "  Please refer to the notes located on the far right of the report.  " & _
										 " Your feedback is necessary to address abnormal metering results and to determine a corrective course of action ."
										 																					


	objExcelReport.DisplayAlerts = False
	objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "C:\ghnet_websites\appserver1\eri_TH\finance\" & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=building%><%=billyear%><%=billperiod%><%=UtilityId%>MeterLetter.xls" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=building%>MeterLetter.xls</b></a> 
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
function loadperiod()
{	var frm = document.forms['form1'];
	var newhref = "MeterMaintenanceLetter.asp?" + "&building="+frm.building.value+"&billyear="+frm.billyear.value
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
	var newhref = "MeterMaintenanceLetter.asp?building="+frm.building.value+"&utilityid="+frm.utilityid.value;
	document.location.href=newhref;
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
	
