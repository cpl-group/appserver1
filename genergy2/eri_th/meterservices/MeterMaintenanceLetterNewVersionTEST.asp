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
<title>Meter Maintenance Letter TEST</title>

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
   <form name="form1" action="MeterMaintenanceLetterNewVersionTEST.asp">
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
		Dim oWord 
        Dim oDoc 
        Dim oTable 
        Dim oPara1
        Dim oPara3 
        Dim oRng 
        Dim oShape 
        Dim oChart 
        Dim Pos
	     'Start Word and open the document template.
        Set oWord = CreateObject("Word.Application")
        oWord.Visible = True
        Set oDoc = oWord.Documents.Add
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")
		
		cnn1.open getLocalConnect(building)
	
	sSql = "Exec usp_MaintenanceLetterBuildingInfo " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod
	rst1.CursorLocation = 3
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	


' Header Columns	
	If not rst1.eof then
       'Insert a paragraph at the beginning of the document.
        Set oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "The following meters reported a usage value that is either 25% higher or 25% lower when compared to the previous period.  Because the fluctuation may be part of normal tenant activity, and because we have no reason to believe that the readings acquired are faulty, these meters have not been estimated.   Should you feel, based on your knowledge of tenant activity, that one or more of these meters should be reviewed for the purpose of estimations, please contact us as soon as possible:"
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 12   '24 pt spacing after paragraph.
        oPara1.Range.InsertParagraphAfter

        Dim r , c 
        Set oTable = oDoc.Tables.Add(oDoc.Bookmarks("\endofdoc").Range, 1, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 6
        For r = 1 To 1
             oTable.Cell(r, 1).Range.Text = ""
             oTable.Cell(r, 2).Range.Text = "CurrentMonthUsage"
             oTable.Cell(r, 3).Range.Text = "PreviousMonthUsage"
             oTable.Cell(r, 4).Range.Text = "Tenant"
             oTable.Cell(r, 5).Range.Text = "Floor"
        Next
        
        oTable.Cell(1, 1).Width = 10
        'oTable.Columns(2).Width = 40
        'oTable.Columns(2).Width = 10


        sSql = "Exec usp_MaintenanceLetter25pctUsage " & "'" & building & "'" & ", " & UtilityId  & "," & Billyear & "," & BillPeriod 
	    rst1.CursorLocation = 3
	    rst1.open sSql , cnn1, 3 
	    Do Until rst1.eof
	
	        iRow= iRow + 1
	        objWorkSheet.Cells(iRow,1) = rst1("meternum")
            objWorkSheet.Cells(iRow,3).HorizontalAlignment = -4108
	        objWorkSheet.Cells(iRow,3) = rst1("currentmonth")
            objWorkSheet.Cells(iRow,4).HorizontalAlignment = -4108
	        objWorkSheet.Cells(iRow,4) = rst1("previousmonth")
            objWorkSheet.Cells(iRow,5) = rst1("tenant")
	        objWorkSheet.Cells(iRow,6) = rst1("floordescr")
	    
						
		   rst1.movenext
	   loop
	   rst1.close

        
		
	
	
	end if

    'creating a word document from vb search

    

    dim ctime 
    ctime = hour(now) & minute(now) & second(now) & Billyear  & Billperiod  & UtilityId & building

																				


	oWord.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\"  & building & Billyear  & Billperiod  & UtilityId & "MeterLetter.xls")
	oDoc.SaveAs("\\2012dc\web_folders\finance\"  & ctime & "MeterLetter.docx")
	oWord.DisplayAlerts = True
	oWord.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing
	' Set up Email to be Sent


	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\"  & ctime & "MeterLetter.docx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=ctime%>MeterLetter.docx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=ctime%>MeterLetter.docx</b></a> 
	</p>
	<%
	Else
	%>
	<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy IT department for assistance.</p>
	<%
		
	End IF


	End If %>
<Script type=text/javascript>	
function loadperiod()
{	var frm = document.forms['form1'];
var newhref = "MeterMaintenanceLetterNewVersionTEST.asp?" + "&building=" + frm.building.value + "&billyear=" + frm.billyear.value + "&pid=" + frm.pid.value;
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
var newhref = "MeterMaintenanceLetterNewVersionTEST.asp?building=" + frm.building.value + "&utilityid=" + frm.utilityid.value + "&pid=" + frm.pid.value;
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