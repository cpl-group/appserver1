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


	function getFieldValue(strField, iLengthReqd)
		Dim iFieldLength, strReturn
		iFieldLength = Len(Trim(strField))
		If iFieldLength >= iLengthReqd Then
			strReturn = Left(strField, iLengthReqd)
		Else
			strReturn  = space(iLengthReqd-iFieldLength) & Trim(strField) 
		End If
		getFieldValue = strReturn 
	end Function
	
	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql
	' Set Parameters
	building = request("bldgNum")	
	BillYear = request("billyear")
	BillPeriod = request("billperiod")
	utilityId = request("UtilityId")
	' Set Default
	if UtilityId = "" then
		Utilityid = 2
	end if	
	
	Dim rst1, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 
	

%>
<html>
<head>
<title>SL Green Accounting File Export</title>

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
   <form name="form1" action="SLGreenAccountingExport.asp">
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

					rst1.open sql, getLocalConnect(building)%>
					
					<%do until rst1.eof%>
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
				sql = sql & "bldgnum='" & building & "' and utility=" & utilityid & " order by billperiod desc "
				
				rst1.open sql, getLocalConnect(building)
				do until rst1.eof
				%>
					<option value="<%=rst1("billperiod")%>" <%if trim(rst1("billperiod"))=billperiod then response.write " SELECTED"%>><%=rst1("billperiod")%></option>
                <%
				  rst1.movenext
				loop
				rst1.close
			end if
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
	Dim iRow
	
	Dim objFSO, strFileName , objFile, strRecord
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim sSql
	
	Dim usage, demand, utilityname
	Dim RecordsAffected
	'Response.Write "billperiod :" & Billperiod 
	'Response.End 
	
	RecordsAffected = 0

	If billperiod <> "" then
		set rst1 = server.createobject("ADODB.Recordset")

		set cnn1 = server.createobject("ADODB.Connection")

		cnn1.open getLocalConnect(building)
		
	
		sSql = "Exec usp_Accounting_SLG " &  Billyear  & "," & Billperiod & ", '" & building & "'" & ", "  & UtilityId  
		'Response.Write sSql
		'Response.End 
		cnn1.Execute sSql , RecordsAffected

		If RecordsAffected > 0 Then 
			sSql = "SELECT BldgStreet, BillPeriod, TenantName, [Floor], TenantNum, Counter, Subtotal " & _
					" FROM SLG_AccountingFile"

			rst1.CursorLocation = 3
			rst1.open sSql , cnn1, 3 
			
			If not rst1.eof then 
				strFilename = "C:\ghnet_websites\appserver1\eri_TH\finance\" & building & "_TenantSubTotals_" &  MonthName(Billperiod,false)  & Billyear & ".csv"
				set objFile = objFSO.CreateTextFile(strFileName, true)
				Do Until rst1.eof
				strRecord = rst1("BldgStreet") & "," & _
							rst1("BillPeriod") & "," & _
							Replace(rst1("TenantName"), ",", " ") & "," & _
							Replace(rst1("Floor"), ",", " ") & "," & _
							rst1("TenantNum") & "," & _
							rst1("Counter") & "," & _
							rst1("Subtotal")
							
							
					objFile.WriteLine(strRecord) 			
					rst1.movenext
				Loop
				objFile.Close								
		  End IF	
		 rst1.Close

		sSql = "SELECT BldgStreet, BillPeriod, TenantName, [Floor], TenantNum, Counter, Tax " & _
			" FROM SLG_AccountingFile"

		rst1.CursorLocation = 3
		rst1.open sSql , cnn1, 3 

			If not rst1.eof then 
				strFilename = "C:\ghnet_websites\appserver1\eri_TH\finance\" & building & "_TenantTax_" &  MonthName(Billperiod,false)  & Billyear & ".csv"
				set objFile = objFSO.CreateTextFile(strFileName, true)
				Do Until rst1.eof
				strRecord = rst1("BldgStreet") & "," & _
							rst1("BillPeriod") & "," & _
							rst1("TenantName") & "," & _
							rst1("Floor") & "," & _
							rst1("TenantNum") & "," & _
							rst1("Counter") & "," & _
							rst1("Tax") 
								
								
					objFile.WriteLine(strRecord) 			
					rst1.movenext
				Loop
				objFile.Close								
		  End IF	
		 rst1.Close

     End if
		' Header Columns	
		' Display link to the generated file 
		If objFSO.FileExists(strFileName) Then 
		%>
		<p> Following File(s) have been generated :
		<br>
		<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=building%>_TenantSubTotals_<%=MonthName(Billperiod)%><%=Billyear%>.csv" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=building%>_TenantSubTotals_<%=MonthName(Billperiod)%><%=Billyear%>.csv</b></a>
		<br>
		<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://appserver1.genergy.com/eri_TH/finance/<%=building%>_TenantSubTotals_<%=MonthName(Billperiod)%><%=Billyear%>.csv" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=building%>_TenantTax_<%=MonthName(Billperiod)%><%=Billyear%>.csv</b></a> 
		</p>
		<%
		Else
		%>
		<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy for assistance.</p>
		<%
		
		End IF
	End If
	%>

<Script type=text/javascript>	
function loadperiod()
{	var frm = document.forms['form1'];
	var newhref = "SLgreenAccountingExport.asp?" + "&building="+frm.building.value+"&billyear="+frm.billyear.value
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
	var newhref = "SLgreenAccountingExport.asp?building="+frm.building.value+"&utilityid="+frm.utilityid.value;
	document.location.href=newhref;
}
</Script>
</body>
</html>
<%
	
	'set objEmail = Nothing
	'set rstMailingList = Nothing
	set objFile = Nothing
	set objFSO = Nothing
	set rst1 = Nothing
	set cnn1 = Nothing
	
	
%>	
	
