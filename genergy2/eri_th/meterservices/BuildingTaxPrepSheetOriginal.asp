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

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql, pid
	' Set Parameters
	building = request("bldgNum")	
	BillYear = request("billyear")
	BillPeriod = request("billperiod")
	UtilityId = request("utilityid")
    pid = request("pid")
	' Set Default
	if UtilityId = "" then
		Utilityid = 2
	end if
	Dim rst1, rst2, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 	
%>
<html>
<head>
<title>Building Tax Prep</title>

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
   <form name="form1" action="BuildingTaxPrepSheet.asp">
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
	
	sSql = "Exec usp_TaxPrepBuildingInfo " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 8
    objWorkSheet.Columns(1).ColumnWidth = 48
    objWorkSheet.Columns(2).ColumnWidth = 7
    objWorkSheet.Columns(3).ColumnWidth = 47
    objWorkSheet.Columns(4).ColumnWidth = 15
    


' Header Columns	
	If not rst1.eof then

       	iRow = 1
    objWorkSheet.Cells(iRow,1).Font.Bold = True
    objWorkSheet.Cells(iRow,1).Font.Size = 19    
    objWorkSheet.Cells(iRow,1) = "CPLEMS"
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False                    
	objWorkSheet.Cells(iRow,1) = "Energy Management Services"
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
    objWorkSheet.Cells(iRow,1) = ""
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "29-19 30th St Long Island City, NY 11101"
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "(212) 664-7600 | cplgroupusa.com"
    iRow = iRow + 1
    iRow = iRow + 1

    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Current State Sales Tax "
    objWorkSheet.Cells(iRow,2) = rst1("currentstatesalestax")

    	iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Current City Sales Tax "
    objWorkSheet.Cells(iRow,2) = rst1("currentcitysalestax")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Current Metropolitan Commuter Transportation District (MCDT) Tax "
    objWorkSheet.Cells(iRow,2) = rst1("currentmetrotax")
    objWorkSheet.Cells(iRow,3) = rst1("bldgnumber")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total Sales Tax Rate"
    objWorkSheet.Cells(iRow,2) = rst1("totalsalestaxrate")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,3) = rst1("managerowner")
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16
    
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Is this Building Full Service or Retail Access (F/R)"
    objWorkSheet.Cells(iRow,2) = rst1("fullserviceretail")
    objWorkSheet.Cells(iRow,3) = rst1("address1")
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,3) = rst1("address2")
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = True
	objWorkSheet.Cells(iRow,1) = "Sub Meter Billing"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = rst1("monthdescr")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total Submeter Billing (Excl. Sales Tax)"
    objWorkSheet.Cells(iRow,4) = rst1("totalsubmeterbilling")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total Tenant Billing Not Subject To Sales Tax"
    objWorkSheet.Cells(iRow,4) = rst1("totaltenantbilling")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total Taxable Net of Tenant Submeter Billing"
    objWorkSheet.Cells(iRow,4) = rst1("totaltaxablenet")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total Sales Tax Charged on Submeter Billing"
    objWorkSheet.Cells(iRow,4) = rst1("totalsalestax")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = True
	objWorkSheet.Cells(iRow,1) = "Utility Billing(Delivery)"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Con-Ed Bill Net(Excluding Sales Tax)"
    objWorkSheet.Cells(iRow,4) = rst1("conedbillnet")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Con-Ed Bill Sales Tax(Paid)"
    objWorkSheet.Cells(iRow,4) = rst1("conedbillsalestaxpaid")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Con-ed Bill Sales Tax(Calculated)"
    objWorkSheet.Cells(iRow,4) = rst1("conedbillsalestaxcalc")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Con-Ed Bill Total Paid"
    objWorkSheet.Cells(iRow,4) = rst1("conedbilltotalpaid")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Con-Ed Bill Total(Calculated)"
    objWorkSheet.Cells(iRow,4) = rst1("conedbilltotalcalc")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = True
	objWorkSheet.Cells(iRow,1) = "Utility Billing(ESCO)"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "ESCO Bill Net(Excluding Sales Tax)"
    objWorkSheet.Cells(iRow,4) = rst1("escobillnet")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "ESCO Bill Sales Tax(Paid)"
    objWorkSheet.Cells(iRow,4) = rst1("escobillsalestaxpaid")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "ESCO Bill Sales Tax(Calculated)"
    objWorkSheet.Cells(iRow,4) = rst1("escobillsalestaxcalc")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "ESCO Bill Total Paid"
    objWorkSheet.Cells(iRow,4) = rst1("escobilltotalpaid")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "ESCO Bill Total(Calculated)"
    objWorkSheet.Cells(iRow,4) = rst1("escobilltotalcalc")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Combined Con-Ed & ESCO Bill Total Paid"
    objWorkSheet.Cells(iRow,4) = rst1("combinedbilltotalpaid")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Combined Con-Ed & ESCO Bill Total(Calculated)"
    objWorkSheet.Cells(iRow,4) = rst1("combinedbilltotalcalc")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Unpaid Use Tax"
    objWorkSheet.Cells(iRow,4) = rst1("unpaidusetax")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = True
	objWorkSheet.Cells(iRow,1) = "SALES & USE"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Amount of Electricity Purchased During Period(in KWH)"
    objWorkSheet.Cells(iRow,4) = rst1("electricitypurchased")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Amount of Electricity Resold During Period(in KWH)"
    objWorkSheet.Cells(iRow,4) = rst1("electricityresold")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Ratio Electricity Resold to Electricity Purchased"
    objWorkSheet.Cells(iRow,4) = rst1("ratio")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calculated Electricity Resold During Period (in Dollars, excl. sales tax - Subject to NYC locality tax)"
    objWorkSheet.Cells(iRow,1).WrapText = True
    objWorkSheet.Rows(iRow).RowHeight =  24
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,4) = rst1("calcelectricityresoldNYC")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calculated Electricity Resold During Period(in Dollars, excl. sales tax - Subject to NYC/NYS combined locality tax)"
    objWorkSheet.Cells(iRow,1).WrapText = True
    objWorkSheet.Rows(iRow).RowHeight =  24
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,4) = rst1("calcelectricityresoldNYS")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calculated Electricity Resold During Period(in Dollars, excl. sales tax)"
    objWorkSheet.Cells(iRow,4) = rst1("calcelectricityresoldPER")
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Claculated Credit For Use Tax Paid on Electricity that Was Resold"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,4) = rst1("calccredit")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Long Method of Calculating Monthly Sales Tax Due Based on ST-809 NYS Sales and Use Tax Return for Monthly Filers"

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 1 = Total Gross Sales and Services"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "Total Gross Submeter Billing(Excluding Sales Tax)"
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16	
    objWorkSheet.Cells(iRow,4) = rst1("totalgrosssubmeterbillecltax")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 2 = Gross Taxable Sales and Services"
    objWorkSheet.Cells(iRow,3) = "Total Gross Submeter Billing(Excluding Sales Tax-Tax-Exempt Tenants)"	
    objWorkSheet.Cells(iRow,4) = rst1("totalgrosssubmeterbillecltaxexempt")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 3 = Total Purchases Subject to Sales Tax"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "Calculation of purchase subject to sales tax based on unpaid use tax"	
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,4) = rst1("calcpurchases")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "4a = Total Sales Tax Billed to Submeter Tenants"	
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 8
    objWorkSheet.Cells(iRow,4) = rst1("foura")
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 8

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 4 = Sales and Use Tax"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "4b = Unpaid Use Tax Applicable to Electricity Purchases"
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 8
    objWorkSheet.Cells(iRow,4) = rst1("fourb")
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 8

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "4c = Use Tax Actually Paid on Purchases of Resold Electricity"
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 8	
    objWorkSheet.Cells(iRow,4) = rst1("fourc")
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 8

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "Total Sales Tax Billed to Submeter Tenants Plus Unpaid Use Tax"	
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16	
    objWorkSheet.Cells(iRow,4) = rst1("totalsalestaxbilled")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 5 = Credit for Prepaid Sales Tax"
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"
    objWorkSheet.Cells(iRow,4) = "$      -"

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 6 = Net Tax Due"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "Line 4 minus Line 5"	
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16	
    objWorkSheet.Cells(iRow,4) = rst1("line4minusline5")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 7 = Credits Not Identified"
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"
    objWorkSheet.Cells(iRow,4) = "$      -"

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 8 = Advance Payments"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16	
    objWorkSheet.Cells(iRow,4) = "$      -"
    	
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 9 = Total Credits"
    objWorkSheet.Cells(iRow,3) = "Line 7 plus Line 8"
     objWorkSheet.Cells(iRow,4) = "$      -"

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 10 = Sales and use Tax Due"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,3) = "Line 6 minus Line 9"		
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 16
    objWorkSheet.Cells(iRow,4) = rst1("line6minusline9")

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 11 = Penalty And Interest"
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"
    objWorkSheet.Cells(iRow,4) = "$      -"

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Line 12 = Amount Due"
    objWorkSheet.Cells(iRow,3) = "Line 10 plus Line 11"	
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 3
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 10
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 3
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 10
    objWorkSheet.Cells(iRow,4) = rst1("line10plusline11")
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 3
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 10

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Please Note: While the data included in this worksheet has been analyzed to ensure it's accuracy, CPLEMS is not an accounting firm."
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Sales & Use Tax calculations are provided as a way to assist our clients with the task of filing their monthly/quarterly tax returns."
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "However, tax calculations and applicable local tariffs should be verified by your accountant or licenced CPA."
    



    End if
      
	
	
	
	
	



    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    objWorkSheet.Rows("8:8").Select
    objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Billperiod  & UtilityId & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "TaxPrepSheet.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "TaxPrepSheet.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>TaxPrepSheet.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>TaxPrepSheet.xlsx</b></a> 
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
var newhref = "BuildingTaxPrepSheet.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value + "&pid=" + frm.pid.value;
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
var newhref = "BuildingTaxPrepSheet.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value + "&pid=" + frm.pid.value;
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