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
   <form name="form1" action="MeterMaintenanceLetterNewVersion.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
        
				<% if trim(building)<>"" then%>
				<td> <select name="utilityid" onChange="document.location='MeterMaintenanceLetterNewVersion.asp?pid=<%=pid%>&bldgnum=<%=building%>&utilityid='+this.value">
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
            <td> <select name="billyear" onChange="document.location='MeterMaintenanceLetterNewVersion.asp?pid=<%=pid%>&bldgnum=<%=building%>&utilityid=<%=utilityid%>&billyear='+this.value">
                <option value="">Select Bill Year</option>
                <%
                	sql = "SELECT distinct billyear " & _
						" FROM billyrperiod WHERE " & _
				        "billyear>=year(getdate())-1 and bldgnum='"&building&"' and utility = '"&utilityid&"' order by billyear desc "
				        
					rst1.open sql, getLocalConnect(building)
					do until rst1.eof%>
					<option value="<%=rst1("billyear")%>"<%if trim(rst1("billyear"))=trim(billyear) then response.write " SELECTED"%>><%=rst1("billyear")%></option>
					<%
						
							rst1.movenext
					loop
					rst1.close
					%>
					</select> </td>
				<%end if%>	
	  			<%if trim(billyear)<>"" then%>
					<td> <select name="billperiod">
					 <option value="">Select Bill Period</option>
                <%
                
				sql = "SELECT distinct billperiod , datestart" & _
						" FROM billyrperiod WHERE " & _
				        "billyear = " & billyear & " and "
				sql = sql & "bldgnum='"&building&"' and utility =  '"&utilityid&"' order by datestart desc "
					
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
		
		dim utilname
		rst1.open "select portfolio from portfolio where id = (select portfolioid from buildings where bldgnum = '"&building&"')", cnn1
		if not rst1.eof then portfolioid = rst1("portfolio")
		rst1.close
		rst1.open "select utility from tblutility where utilityid = " & utilityid, cnn1
		if not rst1.eof then utilname = rst1("utility")
		rst1.close

	sSql = "Exec usp_MaintenanceLetterBuildingInfo " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 7

    'objWorkSheet.PageSetup.Zoom = False
    objWorkSheet.PageSetup.FitToPagesWide = 1
    objWorkSheet.PageSetup.FitToPagesTall = 1


' Header Columns	
	If not rst1.eof then

        'objWorkSheet.Cells(1,1).Font.Bold = true
        'objWorkSheet.Cells(1,1).Font.Size = 20                    'need logo
		'objWorkSheet.Cells(1,1) = rst1("companyname")
        'objWorkSheet.Cells(1,2).Font.Bold = true
        'objWorkSheet.Cells(1,2).Font.Size = 15                    'need logo
		'objWorkSheet.Cells(1,2) = rst1("companyname2")
        'objWorkSheet.Cells(1,2).Font.ColorIndex = 40    
		'objWorkSheet.Cells(1,1).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,2).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,3).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,4).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,5).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,6).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,7).Interior.ColorIndex = 36 
		'objWorkSheet.Cells(1,8).Interior.ColorIndex = 36 		
		
        objWorkSheet.Cells(1,1).Font.Bold = True
        objWorkSheet.Cells(1,1).Font.Size = 19    
        objWorkSheet.Cells(1,1) = rst1("companyname")
        objWorkSheet.Cells(1,3).Font.Bold = true
        objWorkSheet.Cells(1,3).Font.Size = 14   
        objWorkSheet.Cells(1,3).Font.ColorIndex = 40                     'need logo
		objWorkSheet.Cells(1,3) = rst1("companyname2")
			 

		objWorkSheet.Cells(2,1).Font.Bold = False                    'need logo
		objWorkSheet.Cells(2,1) = ""
			
		
		objWorkSheet.Cells(3,1).Font.Bold = False
        objWorkSheet.Cells(3,1) = rst1("postdate")
			 
				
		objWorkSheet.Cells(4,1).Font.Bold = False
		objWorkSheet.Cells(4,1) = "" 
		
				
		objWorkSheet.Cells(5,1).Font.Bold = False
		objWorkSheet.Cells(5,1) = rst1("name")
		
				
		objWorkSheet.Cells(6,1).Font.Bold = False
		objWorkSheet.Cells(6,1) = rst1("strt")
		
				
		objWorkSheet.Cells(7,1).Font.Bold = False
		objWorkSheet.Cells(7,1) = rst1("city") & " , " & rst1("statedescr") & "," & rst1("zip")
		
				
		objWorkSheet.Cells(8,1).Font.Bold = False
		objWorkSheet.Cells(8,1) = "" 
		
				
		objWorkSheet.Cells(9,1).Font.Bold = False
		objWorkSheet.Cells(9,1) = "Attn: " & rst1("contactname")
		
				
		objWorkSheet.Cells(10,1).Font.Bold = False
		objWorkSheet.Cells(10,1) = "" 
		
				
		objWorkSheet.Cells(11,1).Font.Bold = False
		objWorkSheet.Cells(11,1) = "Re: " & rst1("utilityname") & " Meter Exception Report for " & rst1("bldgname") & " Period " & Billperiod & "," & Billyear 
			
				
		objWorkSheet.Cells(12,1).Font.Bold = False
		objWorkSheet.Cells(12,1) = ""
		
		
		objWorkSheet.Cells(13,1).Font.Bold = False
		objWorkSheet.Cells(13,1) = ""
		
		
        objWorkSheet.Cells(14,1).Font.Bold = False
		objWorkSheet.Cells(14,1) = rst1("contactname") & ","
		
				
		objWorkSheet.Cells(15,1).Font.Bold = False
		objWorkSheet.Cells(15,1) = ""
		
				
		objWorkSheet.Cells(16,1).Font.Bold = False
		objWorkSheet.Cells(16,1) = "For the service period referenced above, which covers " & rst1("startdate") & " through " & rst1("enddate") & ","
		
				
		objWorkSheet.Cells(17,1).Font.Bold = False
		objWorkSheet.Cells(17,1) = "enclosed you will find sub meter invoices suitable for distribution as well as a sub "
		
		
		objWorkSheet.Cells(18,1).Font.Bold = False
		objWorkSheet.Cells(18,1) = "meter summary report for your review."
		
				
		objWorkSheet.Cells(19,1).Font.Bold = False
		objWorkSheet.Cells(19,1) = ""
		
				
		objWorkSheet.Cells(20,1).Font.Bold = False
		objWorkSheet.Cells(20,1) = "We kindly request your feedback on the bill period processing details listed below."
		
				
		objWorkSheet.Cells(21,1).Font.Bold = False
		objWorkSheet.Cells(21,1) = ""
		
				
		objWorkSheet.Cells(22,1).Font.Bold = False
		objWorkSheet.Cells(22,1) = "The following meters reported no usage but were not estimated.  If you feel that based on "
		
				
		objWorkSheet.Cells(23,1).Font.Bold = False
		objWorkSheet.Cells(23,1) = "your knowledge of tenant activity one or more of these meters should be estimated, please "
		
				
		objWorkSheet.Cells(24,1).Font.Bold = False
		objWorkSheet.Cells(24,1) = "contact our office as soon as possible:"
		
				
		objWorkSheet.Cells(25,1).Font.Bold = False
		objWorkSheet.Cells(25,1) = ""
		
        objWorkSheet.Cells(26,1).Font.Bold = False
		objWorkSheet.Cells(26,1) = ""
        objWorkSheet.Cells(26,3) = "Tenant"
        objWorkSheet.Cells(26,4) = "Floor"
		
			
				
	End if
	rst1.close

	iRow = 26
	sSql = "Exec usp_MaintenanceLetterZeroUsage " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("meternum")
        objWorkSheet.Cells(iRow,3) = rst1("Tenantname")
        objWorkSheet.Cells(iRow,4) = rst1("floorDescr")
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	 
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "The following meters reported a usage value that is either 25% higher or 25% lower when "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "compared to the previous period.  Because the fluctuation may be part of normal tenant "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "activity, and because we have no reason to believe that the readings acquired are faulty, these "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "meters have not been estimated.   Should you feel, based on your knowledge of tenant activity, "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "that one or more of these meters should be reviewed for the purpose of estimations, please "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "contact us as soon as possible:"
	
	
	iRow= iRow + 1

    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    'objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,3) = "CurrentMonthUsage"
	objWorkSheet.Cells(iRow,4) = "PreviousMonthUsage"
    objWorkSheet.Cells(iRow,5) = "Tenant"
	objWorkSheet.Cells(iRow,6) = "Floor"
	
	
	sSql = "Exec usp_MaintenanceLetter25pctUsage " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("meternum")
        'objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
	    objWorkSheet.Cells(iRow,3) = rst1("currentmonth")
        'objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
	    objWorkSheet.Cells(iRow,4) = rst1("previousmonth")
        objWorkSheet.Cells(iRow,5) = rst1("tenant")
	    objWorkSheet.Cells(iRow,6) = rst1("floordescr")
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "The following meters have been estimated for one or more reasons. If based on your "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "knowledge of tenant activities you feel that these meters should not have been estimated, "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "please contact us as soon as possible:  "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	

    iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,3) = "Tenant"
	objWorkSheet.Cells(iRow,4) = "Floor"
	
	
	sSql = "Exec usp_MaintenanceLetterEstimated " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("meternum")
        objWorkSheet.Cells(iRow,3) = rst1("tenantname")
        objWorkSheet.Cells(iRow,4) = rst1("floordescr")
	   
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "The following tenants reported a watt per square foot value either under 2 watts per SqFt. or "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "over 6 watts per SqFt.. Unless the SqFt on file is incorrect or Management has a reason to "
	

    iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "believe that tenant's operation would constitute these values, we recommend that these "
	

    iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "meters be checked for proper operation: "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	objWorkSheet.Cells(iRow,4) = "CurrentWattsPerSqFt"
    objWorkSheet.Cells(iRow,5) = "SqFt"
	
	
	
	sSql = "Exec usp_MaintenanceTenantSqFt " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("tenant")
        'objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
	    objWorkSheet.Cells(iRow,4) = rst1("wattpersqftchar")
        objWorkSheet.Cells(iRow,5) = rst1("sqftchar")
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	
	iRow= iRow + 1

    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "The following meters reported a load factor either below 20% or above 95%.  Unless management  "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "has a reason to believe that tenant's operation would constitute these values, we recommend that "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "these meters be checked for proper operation:  "
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	

    iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,3) = "Tenant"
	objWorkSheet.Cells(iRow,4) = "Floor"
    objWorkSheet.Cells(iRow,5) = "Load Factor"
	
	
	sSql = "Exec usp_MaintenanceLetterLoadFactor " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("meternum")
        objWorkSheet.Cells(iRow,3) = rst1("tenant")
        objWorkSheet.Cells(iRow,4) = rst1("floordescr")
        objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
        objWorkSheet.Cells(iRow,5) = rst1("loadfactor")
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = "The following tenants have no reported SqFt., please provide:"
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 36  
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = ""
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 36  
	
	
	'sSql = "Exec usp_MaintenanceTenantMissingSqFt " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	'rst1.CursorLocation = 3
	'rst1.open sSql , cnn1, 3 
	'Do Until rst1.eof
	
	    'iRow= iRow + 1
	    'objWorkSheet.Cells(iRow,1) = rst1("tenantname")
	    'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 36 
	    'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 36  
						
		'rst1.movenext
	'loop
	'rst1.close
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = ""
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 36  
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = "Building-wide, the average watt per square foot value for all sub metered tenants this period  "
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	'iRow= iRow + 1
	
	'sSql = "Exec usp_MaintenanceBuildingWattSqFt " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	'rst1.CursorLocation = 3
	'rst1.open sSql , cnn1, 3 
	'Do Until rst1.eof
	
	    'objWorkSheet.Cells(iRow,1).Font.Bold = False
	    'objWorkSheet.Cells(iRow,1) = "is:  " & rst1("wattpersqftchar")
	    'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	    'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
						
		'rst1.movenext
	'loop
	'rst1.close
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = ""
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	'iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Billing services for the following tenants have ceased:"
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
		
	sSql = "Exec usp_MaintenanceTenantCeased " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("tenantname")
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Billing services for the following tenants remain on hold:"
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
		
	sSql = "Exec usp_MaintenanceTenantHold " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("tenantname")
	    
						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "The following meters were noted either during the reading or billing process"
	

    iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "and will be monitored:"
	

	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,3) = "Tenant"
    objWorkSheet.Cells(iRow,4) = "Floor"
	
	
		
	sSql = "usp_MaintenanceMeterNotes " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("meternum")
        objWorkSheet.Cells(iRow,3) = rst1("Tenantname")
        objWorkSheet.Cells(iRow,4) = rst1("floorDescr")
	    
	    						
		rst1.movenext
	loop
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = "The following meters review were noted:"
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = ""
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	
	'sSql = "usp_MaintenanceMeterReviewNotes " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	'rst1.CursorLocation = 3
	'rst1.open sSql , cnn1, 3 
	'Do Until rst1.eof
	
	    'iRow= iRow + 1
	    'objWorkSheet.Cells(iRow,1) = rst1("meternum")
	    'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	    'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	    'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	    						
		'rst1.movenext
	'loop
	'rst1.close
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = ""
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	'iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "A review of building utility costs this period vs. sub meter revenue this period revealed the"
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "following:"
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Cost"
	
	objWorkSheet.Cells(iRow,6).Font.Bold = False
	objWorkSheet.Cells(iRow,4) = "Revenue"
	
	
	iRow= iRow + 1
	
	sSql = "usp_MaintenanceMeterSummary " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total utility usage "
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4131
    objWorkSheet.Cells(iRow,3) = rst1("totalutilityusagechar")
	
	objWorkSheet.Cells(iRow,4).Font.Bold = False
	objWorkSheet.Cells(iRow,4) = "Sub meter usage "
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,5) = rst1("submeterusagechar")
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total utility demand "
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,3) = rst1("totalutilitydemandchar")
	
	objWorkSheet.Cells(iRow,4).Font.Bold = False
	objWorkSheet.Cells(iRow,4) = "Sub meter demand "
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,5) = rst1("submeterdemandchar")
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Utility delivery cost "
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,3) = "$" & rst1("utilitydeliverycost")
	
	objWorkSheet.Cells(iRow,4).Font.Bold = False
	objWorkSheet.Cells(iRow,4) = "Revenue "
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,5) = "$" & rst1("usagerevenue")
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Utility supply cost "
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,3) = "$" & rst1("utilitysupplycost")
	
	objWorkSheet.Cells(iRow,4).Font.Bold = False
	objWorkSheet.Cells(iRow,4) = "Admin fees "
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,5) = "$" & rst1("adminfees")
	
	
	'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = "Total utility tax "
	'objWorkSheet.Cells(iRow,3) = "$" & rst1("totalutilitytax")
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 36
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 36 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 36 
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total utility cost "
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,3) = "$" & rst1("totalutilitycost")
	
	objWorkSheet.Cells(iRow,4).Font.Bold = False
	objWorkSheet.Cells(iRow,4) = "Total resale "
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
	objWorkSheet.Cells(iRow,5) = "$" & rst1("totalresale")
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Usage recovery this period totaled was: " & rst1("usagerecovery") & "%"
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Demand recovery this period totaled was: " & rst1("demandrecovery") & "%"
	
	
	iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = "Tax recovery this period totaled was: " & rst1("taxrecoverypct") & "%"
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	'iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Cost recovery this period totaled was: " & rst1("costrecovery") & "%"
	
	
	rst1.close
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1

    'iRow= iRow + 1
	
	'objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,1) = ""
	'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 40
	'objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 40 
	'objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 40 
	
	'iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Meter Count Summary"
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	iRow= iRow + 1
	
		
	sSql = "usp_MaintenanceMeterCounts " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Number Of Active Meters "
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
    objWorkSheet.Cells(iRow,5) = rst1("activemeters")
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Number Of Active Meters Read And Not Billed"
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
    objWorkSheet.Cells(iRow,5) = rst1("notbilledmeters")
	
	
	iRow= iRow + 1

    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Number Of Active Meters Read And Billed"
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
    objWorkSheet.Cells(iRow,5) = rst1("billedmeters")
	
	
	iRow= iRow + 1

    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Number Of Estimated Meters"
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
    objWorkSheet.Cells(iRow,5) = rst1("estimatedmeters")
	
	
	iRow= iRow + 1

    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Number Of Estimated Meters With Estimated Charge"
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4131
    objWorkSheet.Cells(iRow,5) = rst1("chargedmeters")
	
	
		
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Please feel free to email us at RB@CPLEMS.com with any questions, comments or concerns."
	
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
	
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Billing Services"
	
	
	
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "CPLems"
	
	
	
	
	



    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    objWorkSheet.Rows("8:8").Select
    objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime, filename, root, pdfdir, pdfname
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Billperiod  & UtilityId & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	filename = Billyear  & Right("0" & billperiod, 2) &"_"& Right("0" & utilityid, 2) &"_"& building &"_"& utilname & "_MeterLetter.xlsx"
	pdfname = Billyear  & Right("0" & billperiod, 2) &"_"& Right("0" & utilityid, 2) &"_"& building &"_"& utilname & "_MeterLetter.pdf"
	root = "D:\WebSites\isabella\genergyonline.com\pdfmaker\"
	pdfdir = portfolioid &"\"& building &"\"
	objWorkBook.SaveCopyAs(root & pdfdir & filename)
	objworkbook.exportasfixedformat 0, root&pdfdir&pdfname
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = root & pdfdir & filename
	If objFSO.FileExists(strFileName) Then 
		%>
		<p> Following report has been generated :
		<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://pdfmaker.genergyonline.com/pdfMaker/<%= pdfdir %><%= pdfname %>?dt=<%=ctime%>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%= pdfname %></b></a> 
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
var newhref = "MeterMaintenanceLetterNewVersion.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value + "&pid=" + frm.pid.value;
document.location.href = newhref;
    alert("rob")
}

function loadutility()
{	var frm = document.forms['form1'];
var newhref = "MeterMaintenanceLetterNewVersion.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value + "&pid=" + frm.pid.value;
document.location.href = newhref;
alert("rob")
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