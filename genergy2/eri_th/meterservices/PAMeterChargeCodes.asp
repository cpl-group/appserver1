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

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, rpt, pdf, Genergy_Users, demo, sql
	' Set Parameters
	building = request("bldgNum")	
	BillYear = request("billyear")
	BillPeriod = request("billperiod")

	Dim rst1, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 
	
	'if Billyear <> "" then 
		'Response.Write "Billyear :" & Billyear 
		'Response.Write "Billperiod :" & Billperiod
		'Response.End
	'End if
	'Response.Write "Building :"  & building
	'Response.End 
%>
<html>
<head>
<title>Meter Charge Codes report</title>

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
   <form name="form1" action="PAMeterChargeCodes.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
			<%if trim(building)<>"" then%>
            <td> <select name="billyear" onclick="loadPeriod()">
                <option value="">Select Bill Year</option>
                <%
                	sql = "SELECT distinct billyear " & _
						 " FROM billyrperiod WHERE " & _
				        " billyear>=year(getdate())-1 and bldgnum='" & building& "' order by billyear desc "

					rst1.open sql, getLocalConnect(building)%>
					
					<%do until rst1.eof%>
					<option value="<%=rst1("billyear")%>"<%if trim(rst1("billyear"))=trim(billyear) then response.write " SELECTED"%>><%=rst1("billyear")%></option>
					<%
						
							rst1.movenext
					loop
					rst1.close
					%>
					</select> </td>
					<%end if%>
            <td> <select name="billperiod">
                <option value="">Select Bill Period</option>
                <%
				sql = "SELECT distinct billperiod " & _
						" FROM billyrperiod WHERE " & _
				        "billyear>=year(getdate())-1 and "
				sql = sql & "bldgnum='"&building&"' "
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
    Set objExcelReport = CreateObject("Excel.Application")
    Set objWorkBook = objExcelReport.Workbooks.Add


	Dim sSql
	
	Dim usage, demand, utilityname
	
	'Response.Write "billperiod :" & Billperiod 
	'Response.End 

	If billperiod <> "" then
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")

		cnn1.open getLocalConnect(building)
	
		sSql = "SELECT m.MeterNum, m.Location, tL.BillingName, IsNull(CC.ChargeCode,'') ChargeCode , IsNULL(CN.Used,0) as Usage " & _
					"FROM Meters m, tblPAMeterChargeCodes CC, tblLeasesUtilityPrices LUP, tblLeases tL, " & _
					"	( Select MP.MeterId,  MP.BillYear, MP.BillPeriod " & _
							" From tblMetersByPeriod MP " & _
						" Where MP.Bill_Id in (Select Max(Id) from tblBillByperiod " & _
							" Where LeaseUtilityId = MP.LeaseUtilityId and YpId = MP.YpId and Reject = 0 and BillYear=" & billyear & _
								" and billperiod = " & billperiod & " ) " & _
							" and MP.YpId = (Select Max(YpId) from tblBillByperiod Where " & _
								" LeaseUtilityId = MP.LeaseUtilityId and Reject = 0 and posted=1 and BillYear=" & billyear & _
								" and billperiod = " & billperiod & " ) " & _
						"	and MP.BldgNum = '" & building & "') PM, Consumption CN " & _
				" WHERE " & _
					" m.BldgNum ='" & building & "' AND m.MeterNum *= CC.MeterNum AND m.BldgNum *= CC.BldgNum " & _
					" AND m.LeaseUtilityId = LUP.LeaseUtilityId AND LUP.BillingID = tL.BillingId " & _
					" AND M.MeterId = PM.MeterId AND PM.MeterId = CN.MeterId AND CN.BillYear = PM.BillYear " & _
					" AND CN.BillPeriod = PM.BillPeriod	" & _
				" ORDER BY m.MeterNum "
	
		'Response.Write ssql
		'Response.End 
		
		rst1.open sSql , cnn1
	
		' Select the First Worksheet
		Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
		objWorkSheet.Cells.Font.Name = "Book Antiqua"
		objWorkSheet.Cells.Font.Size = 9
     
		' Header Columns	
		If not rst1.eof then

			objWorkSheet.Cells(1,1).Font.Bold = True
			objWorkSheet.Cells(1,1) ="Meter"
			
			objWorkSheet.Cells(1,2).Font.Bold = True
			objWorkSheet.Cells(1,2) ="Location"

			objWorkSheet.Cells(1,3).Font.Bold = True
			objWorkSheet.Cells(1,3) ="Billing Name"
			
			objWorkSheet.Cells(1,4).Font.Bold = True
			objWorkSheet.Cells(1,4) ="Charge Code"

			objWorkSheet.Cells(1,5).Font.Bold = True
			objWorkSheet.Cells(1,5) ="Usage"

		End if

		iRow = 1

		Do Until rst1.eof

			iRow= iRow + 1
				
			objWorkSheet.Cells(iRow,1) = rst1("MeterNum")
			objWorkSheet.Cells(iRow,2) = rst1("Location")
			objWorkSheet.Cells(iRow,3) = rst1("BillingName")	
			objWorkSheet.Cells(iRow,4) = rst1("ChargeCode")
			objWorkSheet.Cells(iRow,5) = rst1("Usage")	
				
			rst1.movenext
		loop
		rst1.close

		objExcelReport.DisplayAlerts = False
		'objWorkBook.SaveAs("\\10.0.7.21\web_folders\finance\PA\" & building&"\" & building & "ChargeCodes.xls")
		objWorkBook.SaveCopyAs("\\2012dc\web_folders\finance\PA\" & building&"\" & building & "ChargeCodes.xlsx")
		objExcelReport.DisplayAlerts = True
		set objWorkSheet = Nothing
		set objWorkBook = Nothing
		objExcelReport.Quit
		
		set objExcelReport = Nothing
		
		' Display link to the generated file instead of sending an email
		
		Dim objFSO, strFileName
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		strFilename = "\\2012dc\web_folders\finance\PA\" & building & "\" & building & "ChargeCodes.xlsx"
		If objFSO.FileExists(strFileName) Then 
		%>
		<p> Following report has been generated :
		<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/PA/<%=building%>\<%=building%>ChargeCodes.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=building%>ChargeCodes.xlsx</b></a> 
		</p>
		<%
		Else
		%>
		<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy for assistance.</p>
		<%
		
		End IF
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
		'	objEmail.To = strMailingList
			'objEmail.To = "AnthonyC@genergy.com; tarunskalra@hotmail.com"
		'	objEmail.From = "rb@genergy.com"
		'	objEmail.Subject = "Bill Summary for Building " & building & " , Period " & Billperiod & " " & Billyear 
		'	objEmail.AttachFile "\\10.0.7.21\web_folders\finance\PA\" & building&"\" & building & "ChargeCodes.xls" ,building & "ChargeCodes.xls"
		'	objEmail.Send
			
		'	Response.Write "<P> Meter Charge Codes report Generated and sent to Building Contacts <BR>"
		'	Response.Write strMailingList 
		'	Response.Write "</P></Body></Html>"
		'Else
		'	Response.Write "<P> No Mailing List is Available for the Building <BR>"
		'	Response.Write "</P></Body></Html>"
		'End IF
	End If
	%>
<Script type=text/javascript>
	function loadperiod()
	{	var frm = document.forms['form1'];
		var newhref = "PAMeterChargeCodes.asp?" + "&building="+frm.building.value+"&billyear="+frm.billyear.value;
		document.location.href=newhref;
	}
</Script>

</body>
</html>
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