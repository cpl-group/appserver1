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
<title>Building Tax Prep Sierra</title>

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
   <form name="form1" action="BuildingTaxPrepSheetSierra.asp">
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
	sSql = "Exec usp_TaxPrep_sierra_buildinginfo " & "'" & building & "'"
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 12
    objWorkSheet.Columns(1).ColumnWidth = 10
    objWorkSheet.Columns(2).ColumnWidth = 30
    objWorkSheet.Columns(3).ColumnWidth = 10
    objWorkSheet.Columns(4).ColumnWidth = 10
    objWorkSheet.Columns(5).ColumnWidth = 10
    objWorkSheet.Columns(6).ColumnWidth = 10
    objWorkSheet.Columns(7).ColumnWidth = 10
    objWorkSheet.Columns(8).ColumnWidth = 10
    
    


' Header Columns	
	If not rst1.eof then

    
    iRow = 1
    Dim pic1,r,img
    pic1 = "https://appserver1.genergy.com/genergy2/eri_th/meterservices/fulllogo.jpg"

    
    'objWorkSheet.Range("C1:E1").Select
    
    'objWorkSheet.Pictures.Insert("https://appserver1.genergy.com/genergy2/eri_th/meterservices/fulllogo.jpg").Select

    Set r = objWorkSheet.Range("C1:G7")
    Set img = objWorkSheet.Pictures.Insert(pic1)

    With img
    .ShapeRange.LockAspectRatio = 0
    .Top = r.Top
    .Left = r.Left
    .Width = r.Width
    .Height = r.Height
    End With
    
    iRow = iRow + 7
    
    objWorkSheet.Cells(iRow,1) = rst1("name")
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = rst1("bldgname")
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    
    
	
    

    

   	iRow = iRow + 1
   
    objWorkSheet.Cells(iRow,1) = "TAX -ID #:"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
   	objWorkSheet.Cells(iRow,4) = "SALES TAX REPORT"
     objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,8) = rst1("currentdate")
    objWorkSheet.Cells(iRow,8).Font.Bold = true
    
    rst1.close
    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,4) = "COST FROM SUBMETER BILLING"
     objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,4) = "Tenants"
    objWorkSheet.Cells(iRow,5) = "Total"
    objWorkSheet.Cells(iRow,8) = "Sales Tax"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    
   
    

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,4) = "Submetered"
    objWorkSheet.Cells(iRow,5) = "Submeter"
    objWorkSheet.Cells(iRow,6) = "Electricity Cost"
    objWorkSheet.Cells(iRow,7) = "Electricity Cost"
    objWorkSheet.Cells(iRow,8) = "Collected From"
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,1) = "Month"
    objWorkSheet.Cells(iRow,2) = "Bill Period"
    objWorkSheet.Cells(iRow,4) = "KWH"
    objWorkSheet.Cells(iRow,5) = "Charges"
    objWorkSheet.Cells(iRow,6) = "(Taxable)"
    objWorkSheet.Cells(iRow,7) = "(Tax Exempt)"
    objWorkSheet.Cells(iRow,8) = "Tenants"
   
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,2).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true

    objWorkSheet.Range("A" & irow-2 & ":A" & irow).Merge
    objWorkSheet.Range("A" & irow-2).MergeArea.Borders.Weight = 4

    objWorkSheet.Range("B" & irow-2 & ":B" & irow).Merge
    objWorkSheet.Range("B" & irow-2).MergeArea.Borders.Weight = 4

    objWorkSheet.Range("C" & irow-2 & ":C" & irow).Merge
    objWorkSheet.Range("C" & irow-2).MergeArea.Borders.Weight = 4
     objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4  'bottom
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4 'right
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    sSql = "Exec usp_TaxPrep_sierra_section1 " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("monthdescr")
        objWorkSheet.Cells(iRow,1).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4 'left
        objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,2) = rst1("billperiod")
        objWorkSheet.Cells(iRow,2).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,3).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,4) = rst1("submeteredkwh")
        objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,5) = rst1("totalsubmetercharges")
        objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,6) = rst1("electricitycosttaxable")
        objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,7) = rst1("electricitycostnontaxable")
        objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,8) = rst1("salestax")
        objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
	    
						
		rst1.movenext
	loop
	rst1.close
    

    

    
    'objWorkSheet.Range("A12:H14").Borders.Weight = 4  this put border around all cells
    
    
    iRow = iRow + 1
    sSql = "Exec usp_TaxPrep_sierra_section1total " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
    objWorkSheet.Cells(iRow,3) = "TOTAL"
    objWorkSheet.Cells(iRow,3).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,4) = rst1("submeteredkwh")
     objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,5) = rst1("totalsubmetercharges")
     objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,6) = rst1("electricitycosttaxable")
     objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,7) = rst1("electricitycostnontaxable")
     objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("salestax")
     objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    rst1.close   
    objWorkSheet.Cells(iRow,3).Font.Bold = true

    objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,2).Borders(9).Weight = 4

    objWorkSheet.Cells(iRow,2).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,3).Borders(9).Weight = 4
     objWorkSheet.Cells(iRow,3).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4

     iRow = iRow + 2
    objWorkSheet.Cells(iRow,4) = "BUILDINGS COST FOR ELECTRICITY"
    objWorkSheet.Cells(iRow,4).Font.Bold = true

    

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,5) = "Total Cost"
    objWorkSheet.Cells(iRow,6) = "Total Cost"
    objWorkSheet.Cells(iRow,7) = "Total Cost"
   
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,1) = "Month"
    objWorkSheet.Cells(iRow,2) = "Bill Period"
    objWorkSheet.Cells(iRow,3) = "Acct #"
    objWorkSheet.Cells(iRow,4) = "KWH"
    objWorkSheet.Cells(iRow,5) = "(Con Edison)"
    objWorkSheet.Cells(iRow,6) = "(Supplier)"
    objWorkSheet.Cells(iRow,7) = "(Less Sales Tax)"
    objWorkSheet.Cells(iRow,8) = "Sales Tax"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,2).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true

    objWorkSheet.Range("A" & irow-1 & ":A" & irow).Merge
    objWorkSheet.Range("A" & irow-1).MergeArea.Borders.Weight = 4

    objWorkSheet.Range("B" & irow-1 & ":B" & irow).Merge
    objWorkSheet.Range("B" & irow-1).MergeArea.Borders.Weight = 4

    objWorkSheet.Range("C" & irow-1 & ":C" & irow).Merge
    objWorkSheet.Range("C" & irow-1).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    sSql = "Exec usp_TaxPrep_sierra_section2 " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("monthdescr")
        objWorkSheet.Cells(iRow,1).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4 'left
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,2) = rst1("billperiod")
         objWorkSheet.Cells(iRow,2).Borders(10).Weight = 4
        objWorkSheet.Cells(iRow,3).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,4) = rst1("totalkwh")
        objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,5) = rst1("conedcost")
         objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,6) = rst1("suppliercost")
         objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,7) = rst1("totalcostlesstax")
         objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,8) = rst1("salestax")
       objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
	    
						
		rst1.movenext
	loop
	rst1.close
     iRow = iRow + 1
     sSql = "Exec usp_TaxPrep_sierra_section2total " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
    objWorkSheet.Cells(iRow,3) = "TOTAL"
     objWorkSheet.Cells(iRow,3).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,4) = rst1("totalkwh")
     objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,5) = rst1("conedcost")
     objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,6) = rst1("suppliercost")
     objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,7) = rst1("totalcostlesstax")
     objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("salestax")
      objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    rst1.close
    objWorkSheet.Cells(iRow,3).Font.Bold = true
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,2).Borders(9).Weight = 4

    objWorkSheet.Cells(iRow,2).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,3).Borders(9).Weight = 4
     objWorkSheet.Cells(iRow,3).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4

     iRow = iRow + 2
    objWorkSheet.Cells(iRow,4) = "BUILDINGS USE"
     objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,4).Font.Bold = true

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,5) = "Buildings Use"
    objWorkSheet.Cells(iRow,6) = "Buildings T&D"
    objWorkSheet.Cells(iRow,7) = "Buildings Supply"
    objWorkSheet.Cells(iRow,8) = "Purchases Subject"
   
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    objWorkSheet.Cells(iRow,1) = "Month"
    objWorkSheet.Cells(iRow,2) = "Bill Period"
    objWorkSheet.Cells(iRow,3) = "TOTAL KWH"
    objWorkSheet.Cells(iRow,4) = "Tenants KWH"
    objWorkSheet.Cells(iRow,5) = "KWH"
    objWorkSheet.Cells(iRow,6) = "Cost"
    objWorkSheet.Cells(iRow,7) = "Cost"
    objWorkSheet.Cells(iRow,8) = "To Sales Tax"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
    objWorkSheet.Cells(iRow,2).Font.Bold = true
    objWorkSheet.Cells(iRow,3).Font.Bold = true
    objWorkSheet.Cells(iRow,4).Font.Bold = true
    objWorkSheet.Cells(iRow,5).Font.Bold = true
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true

    objWorkSheet.Range("A" & irow-1 & ":A" & irow).Merge
    objWorkSheet.Range("A" & irow-1).MergeArea.Borders.Weight = 4

    objWorkSheet.Range("B" & irow-1 & ":B" & irow).Merge
    objWorkSheet.Range("B" & irow-1).MergeArea.Borders.Weight = 4

    objWorkSheet.Range("C" & irow-1 & ":C" & irow).Merge
    objWorkSheet.Range("C" & irow-1).MergeArea.Borders.Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    sSql = "Exec usp_TaxPrep_sierra_section3 " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("monthdescr")
     objWorkSheet.Cells(iRow,1).Borders(10).Weight = 4
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,2) = rst1("billperiod")
        objWorkSheet.Cells(iRow,2).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,3) = rst1("totalkwh")
        objWorkSheet.Cells(iRow,3).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,4) = rst1("submeteredkwh")
     objWorkSheet.Cells(iRow,4).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,5) = rst1("buildingkwh")
     objWorkSheet.Cells(iRow,5).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,6) = rst1("tdcost")
     objWorkSheet.Cells(iRow,6).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,7) = rst1("supplycost")
     objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
        objWorkSheet.Cells(iRow,8) = rst1("purchasesalestax")
     objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
	    
						
		rst1.movenext
	loop
	rst1.close
      iRow = iRow + 1
     sSql = "Exec usp_TaxPrep_sierra_section3total " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
    objWorkSheet.Cells(iRow,1) = "TOTAL"
     objWorkSheet.Cells(iRow,1).Borders(10).Weight = 4
     objWorkSheet.Cells(iRow,2).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,3) = rst1("totalkwh")
     objWorkSheet.Cells(iRow,3).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,4) = rst1("submeteredkwh")
     objWorkSheet.Cells(iRow,4).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,5) = rst1("buildingkwh")
     objWorkSheet.Cells(iRow,5).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).Borders(7).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8) = rst1("purchasesubjectsalestax")
     objWorkSheet.Cells(iRow,8).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108


    objWorkSheet.Cells(iRow,2).Font.Bold = true
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,2).Borders(9).Weight = 4

    objWorkSheet.Cells(iRow,2).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,3).Borders(9).Weight = 4
     objWorkSheet.Cells(iRow,3).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    rst1.close
     iRow = iRow + 2
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44

     objWorkSheet.Cells(iRow,1).Borders(8).Weight = 4
      objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,2).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,3).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(8).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(8).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 44
    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 44
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 44
    
    
    objWorkSheet.Cells(iRow,6) = "SUMMARY"
    objWorkSheet.Cells(iRow,7) = "Sales and Service"
    objWorkSheet.Cells(iRow,8) = "Tax"
    objWorkSheet.Cells(iRow,6).Font.Bold = true
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Font.Bold = true

    objWorkSheet.Cells(iRow,1).Borders(9).Weight = 4
      objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,2).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,3).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,2).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,5).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,6).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
      objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108

     

     iRow = iRow + 1
     sSql = "Exec usp_TaxPrep_sierra_section4 " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & Quarter & "," & prevjan & "," & jan & "," & prevfeb & "," & feb & "," & prevmar & "," & mar & "," & prevapr & "," & apr & "," & prevmay & "," & may & "," & prevjun & "," & jun & "," & prevjul & "," & jul & "," & prevaug & "," & aug & "," & prevsep & "," & sep & "," & prevoct & "," & oct & "," & prevnov & "," & nov & "," & prevdec & "," & dec
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
    objWorkSheet.Cells(iRow,1) = "Gross Sales"
     objWorkSheet.Cells(iRow,7) = rst1("grosssales")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Non Taxable Sales"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
     objWorkSheet.Cells(iRow,7) = rst1("nontaxablesales")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Taxable Sales"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7) = rst1("taxablesales")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("salestax")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

     iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "New York City/State Combined tax"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
     objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 1
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 1
     objWorkSheet.Cells(iRow,4) = "8.875%"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Credits against taxable sales and services"
    objWorkSheet.Cells(iRow,4) = "8.875%"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7) = rst1("suppliercostcredit")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("suppliercredit")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Purchases subject to tax"
    objWorkSheet.Cells(iRow,4) = "8.875%"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7) = rst1("supplierpurchasesubjectsalestax")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("supplierpurchasesubjectsalestaxtax")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "New York City - Local Tax"
    objWorkSheet.Cells(iRow,4) = "4.5%"
    objWorkSheet.Cells(iRow,1).Font.Bold = true
     objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 1
     objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 1
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Credits against taxable sales and services"
    objWorkSheet.Cells(iRow,4) = "4.5%"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7) = rst1("conedcostcredit")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("conedcredit")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Purchases subject to tax(4.5%)"
    objWorkSheet.Cells(iRow,4) = "4.5%"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7) = rst1("conedpurchasesubjectsalestax")
    objWorkSheet.Cells(iRow,7).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8) = rst1("conedpurchasesubjectsalestaxtax")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,1) = "Vendor Collection Credit"
    objWorkSheet.Cells(iRow,4) = "5.0%"
     objWorkSheet.Cells(iRow,1).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8) = rst1("vendorcredit")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
    objWorkSheet.Cells(iRow,1).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,2).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,3).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,4).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,5).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,6).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
     objWorkSheet.Cells(iRow,7).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,7).Borders(10).Weight = 4
   

    iRow = iRow + 1
    objWorkSheet.Cells(iRow,7) = "Total Amount Due"
    objWorkSheet.Cells(iRow,7).Font.Bold = true
    objWorkSheet.Cells(iRow,8).Borders(9).Weight = 4
    objWorkSheet.Cells(iRow,8) = rst1("totalamountdue")
    objWorkSheet.Cells(iRow,8).HorizontalAlignment = -4108
    objWorkSheet.Cells(iRow,8).Borders(7).Weight = 4
    objWorkSheet.Cells(iRow,8).Borders(10).Weight = 4
   
     rst1.close
   
    
   
    

   
    



    End if
      
	
	
	
	
	



     objWorkSheet.Columns("B:AP").Select
     objExcelReport.Selection.Columns.AutoFit
     
   

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Quarter  & UtilityId & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "SierraTaxPrepSheetQtr.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "SierraTaxPrepSheetQtr.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Report generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>SierraTaxPrepSheetQtr.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>SierraTaxPrepSheetQtr.xlsx</b></a> 
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
        var newhref = "BuildingTaxPrepSheetSierra.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value + "&pid=" + frm.pid.value;
        document.location.href = newhref;
    }

    function loadutility() {
        var frm = document.forms['form1'];
        var newhref = "BuildingTaxPrepSheetSierra.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value + "&pid=" + frm.pid.value;
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