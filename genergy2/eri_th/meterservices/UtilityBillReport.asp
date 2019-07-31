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
<title>Utility Bill Report</title>

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
   <form name="form1" action="UtilityBillReport.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
        
				<% if trim(building)<>"" then%>
				<td> <select name="utilityid" onChange="document.location='UtilityBillReport.asp?bldgnum=<%=building%>&utilityid='+this.value">
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
            <td> <select name="billyear" onChange="document.location='UtilityBillReport.asp?bldgnum=<%=building%>&utilityid=<%=utilityid%>&billyear='+this.value">
                <option value="">Select Bill Year</option>
                <%
                	sql = "SELECT distinct billyear " & _
						" FROM billyrperiod WHERE " & _
				        "billyear>=year(getdate())-4 and bldgnum='"&building&"' and utility = '"&utilityid&"' order by billyear desc "
				        
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
				sql = sql & "bldgnum='"&building&"' and utility = '"&utilityid&"' order by datestart desc "
					
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
    Dim i


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
	
	sSql = "Exec usp_UtilityBillReport " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Calibri"
	objWorkSheet.Cells.Font.Size = 11


' Header Columns	
	If not rst1.eof then

        i = 1
		
        objWorkSheet.Cells(i,1).Font.Bold = True
        objWorkSheet.Cells(i,1).Font.Size = 19    
        objWorkSheet.Cells(i,1) = rst1("companyname")
        objWorkSheet.Cells(i,1).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40			 
        i = i + 1
		objWorkSheet.Cells(i,1).Font.Bold = False                    'need logo
		objWorkSheet.Cells(i,1) = "Energy Management Services"
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40			
		i = i + 1
		objWorkSheet.Cells(i,1).Font.Bold = False
        objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
        i = i + 1
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "29-19 30th St Long Island City, NY 11101"
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		 		 
		i = i + 1	
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "(212) 664-7600 | cplgroupusa.com"
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40			
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1	
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "Report Date: " & rst1("reportdate")
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1	
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "" 
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
        i = i + 1
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = rst1("strt")
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "" 
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40			
        i = i + 1
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "Attn: " & rst1("contactname")
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		 	
        i = i + 1
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "" 
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40	 	
		i = i + 1				
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "Re: " & rst1("utilityname") & " Utility Bill Report " & rst1("bldgname")  & " Bill Period " & billperiod & " " & billyear 
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40	
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40	
		i = i + 1
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = rst1("contactname") & ","
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40	
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40	
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "Outlined below please find detailed information covering " & rst1("startdate") & " through " & rst1("enddate") & "."
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40	
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "Please contact us with any questions or comments."
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40			
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1		
		objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40	
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1		
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = ""
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 40 	
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 40
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 40
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 40 
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 40		
		i = i + 1		
        objWorkSheet.Cells(i,1).Font.Bold = False
		objWorkSheet.Cells(i,1) = "Bill Year"
        objWorkSheet.Cells(i,2) = "Bill Period"
        objWorkSheet.Cells(i,3) = "Utility Bill Expense"
        objWorkSheet.Cells(i,4) = "Utility Bill KWH"
        objWorkSheet.Cells(i,5) = "Sub Metered Revenue"
        objWorkSheet.Cells(i,6) = "Sub Metered KWH"
        objWorkSheet.Cells(i,7) = "Admin Fee"
        objWorkSheet.Cells(i,8) = "Service Fee"
        objWorkSheet.Cells(i,9) = "Credit"
        objWorkSheet.Cells(i,10) = "Sub Total"
        objWorkSheet.Cells(i,11) = "Tax"
        objWorkSheet.Cells(i,12) = "Grand Total"
        objWorkSheet.Cells(i,13) = "% Recoup(Revenue)"
        objWorkSheet.Cells(i,14) = "% Recoup(KWH)"
        objWorkSheet.Cells(i,15) = "ERI Revenue"
        objWorkSheet.Cells(i,16) = "Utility Cost Less SubMetered Revenue"
		objWorkSheet.Cells(i,1).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,2).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,3).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,4).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,5).Interior.ColorIndex = 36
		objWorkSheet.Cells(i,6).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,7).Interior.ColorIndex = 36
		objWorkSheet.Cells(i,8).Interior.ColorIndex = 36 
        objWorkSheet.Cells(i,9).Interior.ColorIndex = 36
		objWorkSheet.Cells(i,10).Interior.ColorIndex = 36
        objWorkSheet.Cells(i,11).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,12).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,13).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,14).Interior.ColorIndex = 36 
		objWorkSheet.Cells(i,15).Interior.ColorIndex = 36
		objWorkSheet.Cells(i,16).Interior.ColorIndex = 36 
		
			
		
				
	End if
	rst1.close
    iRow = i
	
	sSql = "Exec usp_UtilityBillReportCalc " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
	    objWorkSheet.Cells(iRow,1) = rst1("BillYear")
        objWorkSheet.Cells(iRow,2) = rst1("BillPeriod")
        objWorkSheet.Cells(iRow,3) = rst1("UtilityBillExpense")
        objWorkSheet.Cells(iRow,4) = rst1("UtilityBillKWH")
        objWorkSheet.Cells(iRow,5) = rst1("SubMeteredRevenue")
        objWorkSheet.Cells(iRow,6) = rst1("SubMeteredKWH")
        objWorkSheet.Cells(iRow,7) = rst1("AdminFee")
        objWorkSheet.Cells(iRow,8) = rst1("ServiceFee")
        objWorkSheet.Cells(iRow,9) = rst1("Credit")
        objWorkSheet.Cells(iRow,10) = rst1("Subtotal")
        objWorkSheet.Cells(iRow,11) = rst1("Tax")
        objWorkSheet.Cells(iRow,12) = rst1("GrandTotal")
        objWorkSheet.Cells(iRow,13) = rst1("SubMeteredRevenuePCT")
        objWorkSheet.Cells(iRow,14) = rst1("SubMeteredKWHPCT")
        objWorkSheet.Cells(iRow,15) = rst1("ERIRevenue")
        objWorkSheet.Cells(iRow,16) = rst1("UtilityCostLessSubMeteredRevenue")
	    objWorkSheet.Cells(iRow,1).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,2).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,3).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,4).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,5).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,6).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,7).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,8).Interior.ColorIndex = 36 
        objWorkSheet.Cells(iRow,9).Interior.ColorIndex = 36
		objWorkSheet.Cells(iRow,10).Interior.ColorIndex = 36
        objWorkSheet.Cells(iRow,11).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,12).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,13).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,14).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,15).Interior.ColorIndex = 36 
	    objWorkSheet.Cells(iRow,16).Interior.ColorIndex = 36 
	    
						
		rst1.movenext
	loop
	rst1.close
	
	
	
	
	
	sSql = "Exec usp_UtilityBillReportTotal " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	Do Until rst1.eof
	
	    iRow= iRow + 1
        objWorkSheet.Cells(iRow,2) = "Totals "
	    objWorkSheet.Cells(iRow,3) = rst1("UtilityBillExpense")
        objWorkSheet.Cells(iRow,4) = rst1("UtilityBillKWH")
        objWorkSheet.Cells(iRow,5) = rst1("SubMeteredRevenue")
        objWorkSheet.Cells(iRow,6) = rst1("SubMeteredKWH")
        objWorkSheet.Cells(iRow,7) = rst1("AdminFee")
        objWorkSheet.Cells(iRow,8) = rst1("ServiceFee")
        objWorkSheet.Cells(iRow,9) = rst1("Credit")
        objWorkSheet.Cells(iRow,10) = rst1("Subtotal")
        objWorkSheet.Cells(iRow,11) = rst1("Tax")
        objWorkSheet.Cells(iRow,12) = rst1("GrandTotal")
        objWorkSheet.Cells(iRow,13) = rst1("SubMeteredRevenuePCT")
        objWorkSheet.Cells(iRow,14) = rst1("SubMeteredKWHPCT")
        objWorkSheet.Cells(iRow,15) = rst1("ERIRevenue")
        objWorkSheet.Cells(iRow,16) = rst1("UtilityCostLessSubMeteredRevenue")
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
	    objWorkSheet.Cells(iRow,16).Interior.ColorIndex = 40 		
						
		rst1.movenext
	loop
	rst1.close
	
	

    objWorkSheet.Columns("B:AP").Select
    objExcelReport.Selection.Columns.AutoFit
     
    objWorkSheet.Rows("8:8").Select
    objExcelReport.ActiveWindow.FreezePanes = True

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Billperiod  & UtilityId & building

																				


	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\"  & ctime & "UtilityBillReport.xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "UtilityBillReport.xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>UtilityBillReport.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>UtilityBillReport.xlsx</b></a> 
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
var newhref = "UtilityBillReport.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
var newhref = "UtilityBillReport.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value;
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