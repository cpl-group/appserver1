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

	Dim  Quarter, building, Billyear, PortFolioId, UtilityId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql, pid
    Dim  prevjan, jan, prevfeb, feb, prevmar, mar, prevapr, apr, prevmay, may
    Dim  prevjun, jun, prevjul, jul, prevaug, aug, prevsep, sep, prevoct, oct
    Dim  prevnov, nov, prevdec, dec
	' Set Parameters
	building = request("bldgNum")	
	BillYear = request("billyear")
	Quarter = request("quarter")
	UtilityId = request("utilityid")
    prevjan = request("prevjan")
    jan = request("jan")
    prevfeb = request("prevfeb")
    feb = request("feb")
    prevmar = request("prevmar")
    mar = request("mar")
    prevapr = request("prevapr")
    apr = request("apr")
    prevmay = request("prevmay")
    may = request("may")
    prevjun = request("prevjun")
    jun = request("jun")
    prevjul = request("prevjul")
    jul = request("jul")
    prevaug = request("prevaug")
    aug = request("aug")
    prevsep = request("prevsep")
    sep = request("sep")
    prevoct = request("prevoct")
    oct = request("oct")
    prevnov = request("prevnov")
    nov = request("nov")
    prevdec = request("prevdec")
    dec = request("dec")
    if prevjan = "on" then
       prevjan = 1
    else
       prevjan = 0
    end if
    if jan = "on" then
       jan = 1
    else
       jan = 0
    end if

    if prevfeb = "on" then
       prevfeb = 1
    else
       prevfeb = 0
    end if
    if feb = "on" then
       feb = 1
    else
       feb = 0
    end if

    if prevmar = "on" then
       prevmar = 1
    else
       prevmar = 0
    end if
    if mar = "on" then
       mar = 1
    else
       mar = 0
    end if

    if prevapr = "on" then
       prevapr = 1
    else
       prevapr = 0
    end if
    if apr = "on" then
       apr = 1
    else
       apr = 0
    end if

    if prevmay = "on" then
       prevmay = 1
    else
       prevmay = 0
    end if
    if may = "on" then
       may = 1
    else
       may = 0
    end if

    if prevjun = "on" then
       prevjun = 1
    else
       prevjun = 0
    end if
    if jun = "on" then
       jun = 1
    else
       jun = 0
    end if

    if prevjul = "on" then
       prevjul = 1
    else
       prevjul = 0
    end if
    if jul = "on" then
       jul = 1
    else
       jul = 0
    end if

    if prevaug = "on" then
       prevaug = 1
    else
       prevaug = 0
    end if
    if aug = "on" then
       aug = 1
    else
       aug = 0
    end if

    if prevsep = "on" then
       prevsep = 1
    else
       prevsep = 0
    end if
    if sep = "on" then
       sep = 1
    else
       sep = 0
    end if

    if prevoct = "on" then
       prevoct = 1
    else
       prevoct = 0
    end if
    if oct = "on" then
       oct = 1
    else
       oct = 0
    end if

    if prevnov = "on" then
       prevnov = 1
    else
       prevnov = 0
    end if
    if nov = "on" then
       nov = 1
    else
       nov = 0
    end if

    if prevdec = "on" then
       prevdec = 1
    else
       prevdec = 0
    end if
    if dec = "on" then
       dec = 1
    else
       dec = 0
    end if
    
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
<title>Building Tax Prep Quarterly</title>

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
   <form name="form1" action="BuildingTaxPrepSheetQuarterly.asp">
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
            <td> <select name="billyear" onclick="loadQuarter()">
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
              <td> <select name="Quarter">
					 <option value="">Select Quarter</option>
                <%
                
				sql = "SELECT distinct billperiod " & _
						" FROM billyrperiod WHERE " & _
				        "billyear>=year(getdate())-1 and "
				sql = sql & "bldgnum='450lex' and billperiod < 5 order by billperiod desc "
					
				rst1.open sql, getLocalConnect(building)
				do until rst1.eof
				%>
					<option value="<%=rst1("billperiod")%>" <%if trim(rst1("billperiod"))=quarter then response.write " SELECTED"%>><%=rst1("billperiod")%></option>
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
               <tr>	
	  			<td>Previous Year</td>
                <td>Current Year</td>
              </tr>
              <tr>	
	  			<td><input type="checkbox" name="prevjan" >&nbsp;January&nbsp;</td>
                <td><input type="checkbox" name="jan" >&nbsp;January&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevfeb" >&nbsp;February&nbsp;</td>
                <td><input type="checkbox" name="feb" >&nbsp;February&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevmar" >&nbsp;March&nbsp;</td>
                <td><input type="checkbox" name="mar" >&nbsp;March&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevapr" >&nbsp;April&nbsp;</td>
                <td><input type="checkbox" name="apr" >&nbsp;April&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevmay" >&nbsp;May&nbsp;</td>
                <td><input type="checkbox" name="may" >&nbsp;May&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevjun" >&nbsp;June&nbsp;</td>
                <td><input type="checkbox" name="jun" >&nbsp;June&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevjul" >&nbsp;July&nbsp;</td>
                <td><input type="checkbox" name="jul" >&nbsp;July&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevaug" >&nbsp;August&nbsp;</td>
                <td><input type="checkbox" name="aug" >&nbsp;August&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevsep" >&nbsp;September&nbsp;</td>
                <td><input type="checkbox" name="sep" >&nbsp;September&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevoct" >&nbsp;October&nbsp;</td>
                <td><input type="checkbox" name="oct" >&nbsp;October&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevnov" >&nbsp;November&nbsp;</td>
                <td><input type="checkbox" name="nov" >&nbsp;November&nbsp;</td>
              </tr>
              <tr>
                <td><input type="checkbox" name="prevdec" >&nbsp;December&nbsp;</td>
                <td><input type="checkbox" name="dec" >&nbsp;December&nbsp;</td>
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

	If quarter <> "" then
		Set objExcelReport = CreateObject("Excel.Application")
		Set objWorkBook = objExcelReport.Workbooks.Add
	
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")
		
		cnn1.open getLocalConnect(building)
	sSql = "Exec usp_TaxPrepBuildingInfoQuarterly " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	'sSql = "Exec usp_TaxPrepBuildingInfoQuarterly " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 12
    objWorkSheet.Columns(1).ColumnWidth = 150
    objWorkSheet.Columns(2).ColumnWidth = 0
    objWorkSheet.Columns(3).ColumnWidth = 10
    objWorkSheet.Columns(4).ColumnWidth = 10
    objWorkSheet.Columns(5).ColumnWidth = 10
    objWorkSheet.Columns(6).ColumnWidth = 10
    objWorkSheet.Columns(7).ColumnWidth = 10
    
    


' Header Columns	
	If not rst1.eof then

    
    iRow = 1
    Dim pic1
    pic1 = "https://appserver1.genergy.com/genergy2/eri_th/meterservices/fulllogo.jpg"

    
    objWorkSheet.Range("C1:E1").Select
    'objWorkSheet.Pictures.Insert(pic1).Select inserts

    'Dim opicture1
    'objWorkSheet.Pictures.Insert(pic1)
    'opicture1 = objWorkSheet.Pictures.Insert(pic1)
    objWorkSheet.Pictures.Insert("https://appserver1.genergy.com/genergy2/eri_th/meterservices/fulllogo.jpg").Select
    'objWorkSheet.Shapes.AddPicture("https://appserver1.genergy.com/genergy2/eri_th/meterservices/invoice_logo_1.jpeg", False, True, 1, 1, 1, 1)
    
    'objWorkSheet.Shapes.AddPicture "http://appserver1.genergy.com/genergy2/eri_th/meterservices/invoice_logo_1.jpeg", False, True, 0, 0, 100, 100
    iRow = iRow + 5
    'objWorkSheet.Cells(iRow,1) = pic1
   
    objWorkSheet.Cells(iRow,3) = "Account Information"
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    'objWorkSheet.Range("A1:C1").Borders.Weight = 4
    '[objWorkSheet.Range("A2:C2").Merge
    'objWorkSheet.Range("A2").MergeArea.Borders.Weight = 4
    'objWorkSheet.Range((iRow, "A1"), (iRow, "C1")).Merge
    'objWorkSheet.Range(irow,"A1").MergeArea.Borders.Weight = 4
    'objWorkSheet.Range(Cells(iRow, 1), Cells(iRow, 3)).Borders.Weight = 4
    'objWorkSheet.Range("A1").ColumnWidth
    'objWorkSheet.Cells(iRow,1).ColumnWidth = 5
    'objWorkSheet.Cells(iRow,2) = ""
    'objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    
   
    'objWorkSheet.Range("A1:B1").merge()   .HorizontalAlignment
    objWorkSheet.Cells(iRow,6) = "Current Sales Taxes"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":D" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108
    objWorkSheet.Range("F" & irow & ":G" & irow).Merge
    objWorkSheet.Range("F" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("F" & irow).HorizontalAlignment = -4108
    
    
    iRow = iRow + 1
    
    objWorkSheet.Cells(iRow,3).Font.Bold = False
	objWorkSheet.Cells(iRow,3) = "Building #"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,4) = rst1("bldgnumber")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    
	objWorkSheet.Cells(iRow,6) = "NYS"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,7) = rst1("currentstatesalestax")
    objWorkSheet.Cells(iRow,7).Font.Bold = true

    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    

    

    	iRow = iRow + 1
    objWorkSheet.Cells(iRow,3).Font.Bold = False
    objWorkSheet.Cells(iRow,3) = "Owner/Mgr"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,4) = rst1("managerowner")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
	objWorkSheet.Cells(iRow,6) = "NYC"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,7) = rst1("currentcitysalestax")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    

    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,3).Font.Bold = False
	objWorkSheet.Cells(iRow,3) = "Address"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,4) = rst1("address1")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = "*MCDT"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,7) = rst1("currentmetrotax")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,3).Font.Bold = False
    objWorkSheet.Cells(iRow,3) = "City/State/Zip"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4) = rst1("citystatezip")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
	objWorkSheet.Cells(iRow,6) = "Total"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7) = rst1("totalsalestaxrate")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,3).Font.Bold = False
    objWorkSheet.Cells(iRow,3) = "*F/R"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4) = rst1("fullserviceretail")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = "*Metro Commuter District Transportation Tax"  
    objWorkSheet.Range("F" & irow & ":G" & irow).Merge
    objWorkSheet.Range("F" & irow).MergeArea.Font.Bold = True
    objWorkSheet.Range("F" & irow).MergeArea.Font.Size = 11
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,3) = "*Full Service or Retail Access"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    

    
    iRow = iRow + 2
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,5) = "Billing Quarter"
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.ColorIndex = 2
    'objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    
    
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108
    
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	'objWorkSheet.Cells(iRow,5) = rst1("startdate") + "/" + rst1("enddate")
    objWorkSheet.Cells(iRow,5) = rst1("billingquarter")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108

    iRow = iRow + 2
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,3) = "Month1"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,4) = "Month2"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,5) = "Month3"
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,6) = "Month4"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,7) = "Quarterly Total"
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,3) = rst1("monthdescr1")
    objWorkSheet.Cells(iRow,4) = rst1("monthdescr2")
    objWorkSheet.Cells(iRow,5) = rst1("monthdescr3")
    objWorkSheet.Cells(iRow,6) = rst1("monthdescr4")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,4) = "Submeter Billing"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108

        
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Excl. Sales Tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
   
    
    objWorkSheet.Cells(iRow,3) = rst1("totalsubmeterbilling")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("totalsubmeterbilling2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("totalsubmeterbilling3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("totalsubmeterbilling4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("totalsubmeterbilling5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Untaxable"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    
    objWorkSheet.Cells(iRow,3) = rst1("totaltenantbilling")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("totaltenantbilling2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("totaltenantbilling3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("totaltenantbilling4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("totaltenantbilling5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Taxable"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    
    objWorkSheet.Cells(iRow,3) = rst1("totaltaxablenet")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("totaltaxablenet2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("totaltaxablenet3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("totaltaxablenet4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("totaltaxablenet5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    

    
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total Tax Charged"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
   
    objWorkSheet.Cells(iRow,3) = rst1("totalsalestax")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("totalsalestax2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("totalsalestax3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("totalsalestax4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("totalsalestax5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,4) = "Utility Billing - ConEd"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Total(excl. tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
	objWorkSheet.Cells(iRow,3) = rst1("conedbillnet")
    objWorkSheet.Cells(iRow,4) = rst1("conedbillnet2")
    objWorkSheet.Cells(iRow,5) = rst1("conedbillnet3")
    objWorkSheet.Cells(iRow,6) = rst1("conedbillnet4")
    objWorkSheet.Cells(iRow,7) = rst1("conedbillnet5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Sales Tax(Paid)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("conedbillsalestaxpaid")
    objWorkSheet.Cells(iRow,4) = rst1("conedbillsalestaxpaid2")
    objWorkSheet.Cells(iRow,5) = rst1("conedbillsalestaxpaid3")
    objWorkSheet.Cells(iRow,6) = rst1("conedbillsalestaxpaid4")
    objWorkSheet.Cells(iRow,7) = rst1("conedbillsalestaxpaid5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Paid)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("conedbilltotalpaid")
    objWorkSheet.Cells(iRow,4) = rst1("conedbilltotalpaid2")
    objWorkSheet.Cells(iRow,5) = rst1("conedbilltotalpaid3")
    objWorkSheet.Cells(iRow,6) = rst1("conedbilltotalpaid4")
    objWorkSheet.Cells(iRow,7) = rst1("conedbilltotalpaid5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Sales Tax(Calculated)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = rst1("conedbillsalestaxcalc")
    objWorkSheet.Cells(iRow,4) = rst1("conedbillsalestaxcalc2")
    objWorkSheet.Cells(iRow,5) = rst1("conedbillsalestaxcalc3")
    objWorkSheet.Cells(iRow,6) = rst1("conedbillsalestaxcalc4")
    objWorkSheet.Cells(iRow,7) = rst1("conedbillsalestaxcalc5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Calculated)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("conedbilltotalcalc")
    objWorkSheet.Cells(iRow,4) = rst1("conedbilltotalcalc2")
    objWorkSheet.Cells(iRow,5) = rst1("conedbilltotalcalc3")
    objWorkSheet.Cells(iRow,6) = rst1("conedbilltotalcalc4")
    objWorkSheet.Cells(iRow,7) = rst1("conedbilltotalcalc5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

   
     iRow = iRow + 1
    objWorkSheet.Cells(iRow,4) = "Utility Billing - ESCO"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Total(excl. tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
	objWorkSheet.Cells(iRow,3) = rst1("escobillnet")
    objWorkSheet.Cells(iRow,4) = rst1("escobillnet2")
    objWorkSheet.Cells(iRow,5) = rst1("escobillnet3")
    objWorkSheet.Cells(iRow,6) = rst1("escobillnet4")
    objWorkSheet.Cells(iRow,7) = rst1("escobillnet5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Sales Tax(Paid)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("escobillsalestaxpaid")
    objWorkSheet.Cells(iRow,4) = rst1("escobillsalestaxpaid2")
    objWorkSheet.Cells(iRow,5) = rst1("escobillsalestaxpaid3")
    objWorkSheet.Cells(iRow,6) = rst1("escobillsalestaxpaid4")
    objWorkSheet.Cells(iRow,7) = rst1("escobillsalestaxpaid5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Paid)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("escobilltotalpaid")
    objWorkSheet.Cells(iRow,4) = rst1("escobilltotalpaid2")
    objWorkSheet.Cells(iRow,5) = rst1("escobilltotalpaid3")
    objWorkSheet.Cells(iRow,6) = rst1("escobilltotalpaid4")
    objWorkSheet.Cells(iRow,7) = rst1("escobilltotalpaid5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Sales Tax(Calculated)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = rst1("escobillsalestaxcalc")
    objWorkSheet.Cells(iRow,4) = rst1("escobillsalestaxcalc2")
    objWorkSheet.Cells(iRow,5) = rst1("escobillsalestaxcalc3")
    objWorkSheet.Cells(iRow,6) = rst1("escobillsalestaxcalc4")
    objWorkSheet.Cells(iRow,7) = rst1("escobillsalestaxcalc5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Calculated)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("escobilltotalcalc")
    objWorkSheet.Cells(iRow,4) = rst1("escobilltotalcalc2")
    objWorkSheet.Cells(iRow,5) = rst1("escobilltotalcalc3")
    objWorkSheet.Cells(iRow,6) = rst1("escobilltotalcalc4")
    objWorkSheet.Cells(iRow,7) = rst1("escobilltotalcalc5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

   

    '-----------------------------------------------

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,4) = "Utility Billing - Combined"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Paid)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("combinedbilltotalpaid")
    objWorkSheet.Cells(iRow,4) = rst1("combinedbilltotalpaid2")
    objWorkSheet.Cells(iRow,5) = rst1("combinedbilltotalpaid3")
    objWorkSheet.Cells(iRow,6) = rst1("combinedbilltotalpaid4")
    objWorkSheet.Cells(iRow,7) = rst1("combinedbilltotalpaid5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Total(Calculated)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("combinedbilltotalcalc")
    objWorkSheet.Cells(iRow,4) = rst1("combinedbilltotalcalc2")
    objWorkSheet.Cells(iRow,5) = rst1("combinedbilltotalcalc3")
    objWorkSheet.Cells(iRow,6) = rst1("combinedbilltotalcalc4")
    objWorkSheet.Cells(iRow,7) = rst1("combinedbilltotalcalc5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Unpaid Use Tax"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = rst1("unpaidusetax")
    objWorkSheet.Cells(iRow,4) = rst1("unpaidusetax2")
    objWorkSheet.Cells(iRow,5) = rst1("unpaidusetax3")
    objWorkSheet.Cells(iRow,6) = rst1("unpaidusetax4")
    objWorkSheet.Cells(iRow,7) = rst1("unpaidusetax5")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Borders.Weight = 4
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,4) = "Electricity Sales & Use"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow).HorizontalAlignment = -4108
   
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Purchased(KWH)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
   
    objWorkSheet.Cells(iRow,3) = rst1("electricitypurchased")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("electricitypurchased2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("electricitypurchased3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("electricitypurchased4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("electricitypurchased5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Resold(KWH)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    
    objWorkSheet.Cells(iRow,3) = rst1("electricityresold")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("electricityresold2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("electricityresold3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("electricityresold4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("electricityresold5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true   
    
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Percentage Resold"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    
    objWorkSheet.Cells(iRow,3) = rst1("ratio")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("ratio2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("ratio3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("ratio4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("ratio5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    
    
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

   
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calc. Resold(excl. sales tax - Subject to NYC loc. tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = rst1("calcelectricityresoldNYC")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("calcelectricityresoldNYC2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("calcelectricityresoldNYC3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("calcelectricityresoldNYC4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("calcelectricityresoldNYC5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calc. Resold(excl. sales tax - Subject to NYC+NYS loc. tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = rst1("calcelectricityresoldNYS")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("calcelectricityresoldNYS2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("calcelectricityresoldNYS3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("calcelectricityresoldNYS4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("calcelectricityresoldNYS5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calc. Resold(excl. sales tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = rst1("calcelectricityresoldPER")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("calcelectricityresoldPER2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("calcelectricityresoldPER3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("calcelectricityresoldPER4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("calcelectricityresoldPER5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true   
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

   
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Calc. Credit For Use Tax Paid on Electricity that Was Resold"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = rst1("calccredit")
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4) = rst1("calccredit2")
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5) = rst1("calccredit3")
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6) = rst1("calccredit4")
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("calccredit5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders.Weight = 4
    objWorkSheet.Cells(iRow,5).Borders.Weight = 4
    objWorkSheet.Cells(iRow,6).Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4

    

    iRow = iRow + 5
    objWorkSheet.Cells(iRow,1) = "Long Method of Calculating Monthly Sales Tax Due Based on ST-809 NYS Sales and Use Tax Return for Monthly Filers"
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("A" & irow & ":G" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 1 = Total Gross Sales and Services"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Total Gross Submeter Billing(Excluding Sales Tax)"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("totalgrosssubmeterbillecltax5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 2 = Do I need to file additional schedules?"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"	
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("zeroamount")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "3 (Column C) = Taxable Sales(i.e Sales to Submeter Tenants, Excluding Tax Exempt)"	
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46
    objWorkSheet.Cells(iRow,7) = rst1("totalgrosssubmeterbillecltaxexempt5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = true
	objWorkSheet.Cells(iRow,1) = "Step 3 = Sales and Use Tax (Use Row For New York City/"
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "3 (Column D)= Purchases Subject to Sales Tax"	
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46
    objWorkSheet.Cells(iRow,7) = rst1("calcpurchases5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "State combined tax)"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "3 (Column E)= Applicable Tax Rate"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46
    objWorkSheet.Cells(iRow,7) = rst1("totaltaxreal")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "3 (Column F)= Sales & Use Tax"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46	
    objWorkSheet.Cells(iRow,7) = rst1("SalesanduseTax5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = true
	objWorkSheet.Cells(iRow,1) = "Step 4 = Calculate Special Taxes"
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"	
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("zeroamount")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 5 = Calculate tax credits and advance payments"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("zeroamount")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "6 (Box 14 Amount)= Sales & Use Tax"	
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("SalesanduseTax5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 6 = Calculate taxes due"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "6 (Box 15 Amount)= Calculate Special Taxes"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46
    objWorkSheet.Cells(iRow,7) = rst1("zeroamount")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "6 (Box 16 Amount)= Calculate Tax Credits"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.ColorIndex = 46
    objWorkSheet.Cells(iRow,7) = rst1("calccredit5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4
    	
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = ""
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Total Taxes Due"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("toaltaxesdue5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 46
    'objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    'objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 7 = Pay penalty and interest if you are filing late"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Typically not applicable - for building owner to fill out"		
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,7) = rst1("zeroamount")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 8 = Calculate total amount due"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 46
    objWorkSheet.Cells(iRow,3) = "Total Due"
    objWorkSheet.Cells(iRow,3).Font.Bold = true
   objWorkSheet.Cells(iRow,7) = rst1("toaltaxesdue5")
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 46
    objWorkSheet.Range("A" & irow & ":B" & irow).Merge
    objWorkSheet.Range("A" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Range("C" & irow & ":F" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Font.Bold = False
	objWorkSheet.Cells(iRow,1) = "Step 9 = Sign and Mail Your Return"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,3) = ""	
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,1).Font.ColorIndex = 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 46
    objWorkSheet.Range("C" & irow & ":G" & irow).Merge
    objWorkSheet.Range("C" & irow).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,7).Borders.Weight = 4

    iRow = iRow + 5
    
	objWorkSheet.Cells(iRow,1) = ""
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Please Note: While the data included in this worksheet has been analyzed to ensure it's accuracy, CPLEMS is not an accounting firm."
    objWorkSheet.Range("A" & irow & ":G" & irow).Merge
    objWorkSheet.Range("A" & irow).Font.Bold = True
    objWorkSheet.Range("A" & irow).Font.Size = 15
    iRow = iRow + 1    
	objWorkSheet.Cells(iRow,1) = "Sales & Use Tax calculations are provided as a way to assist our clients with the task of filing their monthly/quarterly tax returns."
    objWorkSheet.Range("A" & irow & ":G" & irow).Merge
    objWorkSheet.Range("A" & irow).Font.Bold = True
    objWorkSheet.Range("A" & irow).Font.Size = 15
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "However, tax calculations and applicable local tariffs should be verified by your accountant or licenced CPA."
    objWorkSheet.Range("A" & irow & ":G" & irow).Merge
    objWorkSheet.Range("A" & irow).Font.Bold = True
    objWorkSheet.Range("A" & irow).Font.Size = 15
    



    End if
      
	
	
	
	
	



     objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    objWorkSheet.Rows("8:8").Select
    objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Quarter  & UtilityId & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "TaxPrepSheetQtr.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "TaxPrepSheetQtr.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>TaxPrepSheetQtr.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>TaxPrepSheetQtr.xlsx</b></a> 
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
    function loadQuarter() {
        var frm = document.forms['form1'];
        var newhref = "BuildingTaxPrepSheetQuarterly.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value + "&pid=" + frm.pid.value;
        document.location.href = newhref;
    }

    function loadutility() {
        var frm = document.forms['form1'];
        var newhref = "BuildingTaxPrepSheetQuarterly.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value + "&pid=" + frm.pid.value;
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