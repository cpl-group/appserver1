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
'2/11/2009 N.Ambo removed this block of code that does an allowgroups check; it sometimes causes errors on the page
if 	not(allowGroups("Genergy Users,clientOperations")) then
%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"-->
<%
end if

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
	Dim rst1, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 	
%>
<html>
<head>
<title>PA Meter Summary Report</title>

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
   <form name="form1" action="PAMeterSummary.asp">
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

	If billperiod <> "" then
		Set objExcelReport = CreateObject("Excel.Application")
		Set objWorkBook = objExcelReport.Workbooks.Add
	
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")
		
		cnn1.open getLocalConnect(building)
	
	sSql = "Exec usp_PAMeterSummary " & Billyear & "," & Billperiod & ",'" & building & "'" & ", " & UtilityId 
	rst1.CursorLocation = 3
	'Response.Write cnn1.ConnectionString 
	'Response.End 
	rst1.open sSql , cnn1, 3 
	
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
	objWorkSheet.Cells.Font.Name = "Book Antiqua"
	objWorkSheet.Cells.Font.Size = 9
     
' Header Columns	
	If not rst1.eof then

		objWorkSheet.Cells(1,1).Font.Bold = True
		objWorkSheet.Cells(1,1) = building & " Reading Data For Period : " & Billperiod & "," & Billyear 


		objWorkSheet.Cells(2,6).Font.Bold = True
		objWorkSheet.Cells(2,6) ="(KWH)"
		
		objWorkSheet.Cells(2,7).Font.Bold = True
		objWorkSheet.Cells(2,7) ="(KWH)"
		
		objWorkSheet.Cells(2,8).Font.Bold = True
		objWorkSheet.Cells(2,8) ="(KWH)"

		objWorkSheet.Cells(2,9).Font.Bold = True
		objWorkSheet.Cells(2,9) ="Multiplier"
		
		objWorkSheet.Cells(2,10).Font.Bold = True
		objWorkSheet.Cells(2,10) ="(KWH)"	
		
		objWorkSheet.Cells(2,11).Font.Bold = True
		objWorkSheet.Cells(2,11) ="(KWH)"
					
		objWorkSheet.Cells(2,12).Font.Bold = True
		objWorkSheet.Cells(2,12) ="(KW)"
		
		objWorkSheet.Cells(2,13).Font.Bold = True
		objWorkSheet.Cells(2,13) ="(KW)"


		objWorkSheet.Cells(3,1).Font.Bold = True
		objWorkSheet.Cells(3,1) ="Sequence No"
		
		objWorkSheet.Cells(3,2).Font.Bold = True
		objWorkSheet.Cells(3,2) ="Tenant Name"

		objWorkSheet.Cells(3,3).Font.Bold = True
		objWorkSheet.Cells(3,3) ="Meter #"
		
		objWorkSheet.Cells(3,4).Font.Bold = True
		objWorkSheet.Cells(3,4) ="Meter Location"

		objWorkSheet.Cells(3,5).Font.Bold = True
		objWorkSheet.Cells(3,5) ="'% Allc."
		
		objWorkSheet.Cells(3,6).Font.Bold = True
		objWorkSheet.Cells(3,6) ="Pres Rd."
		
		objWorkSheet.Cells(3,7).Font.Bold = True
		objWorkSheet.Cells(3,7) ="Prior Rd."
		
		objWorkSheet.Cells(3,8).Font.Bold = True
		objWorkSheet.Cells(3,8) ="Diff."

		objWorkSheet.Cells(3,9).Font.Bold = True
		objWorkSheet.Cells(3,9) ="K"
		
		objWorkSheet.Cells(3,10).Font.Bold = True
		objWorkSheet.Cells(3,10) ="Usage"	
		
		objWorkSheet.Cells(3,11).Font.Bold = True
		objWorkSheet.Cells(3,11) ="YTD Usage"
					
		objWorkSheet.Cells(3,12).Font.Bold = True
		objWorkSheet.Cells(3,12) ="Demand"
		
		objWorkSheet.Cells(3,13).Font.Bold = True
		objWorkSheet.Cells(3,13) ="YTD Demand"
		
		
	End if

	iRow = 3

	Do Until rst1.eof

		iRow= iRow + 1
			
		objWorkSheet.Cells(iRow,1) = rst1("SeqNumber")
		objWorkSheet.Cells(iRow,2) = rst1("TenantName")
		objWorkSheet.Cells(iRow,3) = rst1("MeterNum")	
		objWorkSheet.Cells(iRow,4) = rst1("Location")
		objWorkSheet.Cells(iRow,5) = rst1("Alloc")	
		objWorkSheet.Cells(iRow,6) = rst1("Current")	
		objWorkSheet.Cells(iRow,7) = rst1("Previous")	
		objWorkSheet.Cells(iRow,8) = rst1("Difference")	
		objWorkSheet.Cells(iRow,9) = rst1("Multipler")	
		objWorkSheet.Cells(iRow,10) = rst1("Used")	
		objWorkSheet.Cells(iRow,11) = rst1("YTDUsage")	
		objWorkSheet.Cells(iRow,12) = rst1("Demand")	
		objWorkSheet.Cells(iRow,13) = rst1("YTDDemand")	
				
		rst1.movenext
	loop
	rst1.close

	Dim strBillPeriod
	
	If Cint(Billperiod) < 10 then 
		strBillPeriod = "0" & Billperiod 	
	Else
		strBillPeriod = Billperiod		
	End If
	
	objExcelReport.DisplayAlerts = False
	'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\PA\" & building&"\" & building  & "MeterSummary" & BillYear & strBillPeriod & ".xls")
	objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\PA\" & building&"\" & building  & "MeterSummary" & BillYear & strBillPeriod & ".xlsx")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	set objWorkSheet = Nothing
	set objWorkBook = Nothing
	set objExcelReport = Nothing

	' Display link to the generated file instead of sending an email (TK : 05/16/2008)

	Dim objFSO, strFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFilename = "\\2012dc\web_folders\finance\PA\" & building & "\" & building & "MeterSummary" & BillYear & strBillPeriod & ".xlsx"
	If objFSO.FileExists(strFileName) Then 
	%>
	<p> Following report has been generated :
	<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/PA/<%=building%>\<%=building%>MeterSummary<%=Billyear%><%=strBillPeriod%>.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=building%>MeterSummary<%=Billyear%><%=strBillPeriod%>.xlsx</b></a> 
	</p>
	<%
	Else
	%>
	<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy for assistance.</p>
	<%
		
	End IF


	' Following code disabled (TK: 5/16/2008)
	' Set up Email to be Sent

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
		'Do While not rstMailingList.EOF 
			'if len(strMailingList) > 0 then 
			'	strMailingList = strMailingList & ";" & rstMailingList("Email")
			'else
			'	strMailingList = rstMailingList("Email")
			'end if
			'rstMailingList.MoveNext 
		'Loop 
	'End IF
	' If There is a mailing List then
	'If Len(strMailingList) > 0 then
		'objEmail.To = strMailingList
		'objEmail.To = "AnthonyC@genergy.com; tarunskalra@hotmail.com"
		'objEmail.From = "rb@genergy.com"
		'objEmail.Subject = "Meter Summary for Building " & building & " , Period " & Billperiod & " " & Billyear 
		'objEmail.AttachFile "\\10.0.7.21\web_folders\finance\PA\" & building&"\" & building & "MeterSummary.xls" ,building & "MeterSummary.xls"
		'objEmail.Send
		
		'Response.Write "<P> Meter Summary report Generated and sent to Building Contacts <BR>"
		'Response.Write strMailingList 
		'Response.Write "</P></Body></Html>"
	'Else
		'Response.Write "<P> No Mailing List is Available for the Building <BR>"
		'Response.Write "</P></Body></Html>"
	'End IF
	End If %>
<Script type=text/javascript>	
function loadperiod()
{	var frm = document.forms['form1'];
	var newhref = "PAMeterSummary.asp?" + "&building="+frm.building.value+"&billyear="+frm.billyear.value
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
	var newhref = "PAMeterSummary.asp?building="+frm.building.value+"&utilityid="+frm.utilityid.value;
	document.location.href=newhref;
}
</Script>
<%
	
	'set objEmail = Nothing
	'set rstMailingList = Nothing
	
	set rst1 = Nothing
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