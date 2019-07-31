<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	

		
		
	Dim strCompanyId, strRunDate, strMonth, strYear, strEndDate
	Dim rstJobResults, objCnn, cmdGetResults, cmdParameter 


	' Summmary Calculation

	Dim iHighRiskJobCount , nHighRiskAmt
	Dim iMediumRiskJobCount, nMediumRiskAmt
	Dim iNoRiskCount, nNoRiskAmt
	Dim sSummary


	Function getValue(rstRecs,strField) 
		If not IsNull(rstRecs(strField)) Then 
			getValue = CStr(rstRecs(strField))
		Else
			getValue = "0"
		End IF 
	End Function
	


	iHighRiskJobCount = 0
	nHighRiskAmt = 0.0
	
	iMediumRiskJobCount = 0 
	nMediumRiskAmt = 0.0

	iNoRiskCount = 0
	nNoRiskAmt = 0.0
	sSummary = ""
	strCompanyId = request("optCompanyId")
	
	strMonth = request("monthNum")
	
	strYear = request("optYear")
	
	'format rundate
	
	strRunDate = strMonth &  "/01/" & strYear
	strEndDate = DateAdd("m",1,strRunDate)
	strEndDate = DateAdd("d",-1,strEndDate)

	' Set the default value for company as Genergy
	if trim(strCompanyId) = "" then
		strCompanyId = "GE"
	end if

	set rstJobResults = server.createobject("ADODB.Recordset")
	set objCnn = server.createobject("ADODB.Connection")
	set cmdGetResults = server.CreateObject("ADODB.Command")
	set cmdParameter = server.CreateObject("ADODB.Parameter") 
	 	
	objCnn.open getConnect(0,0,"Intranet")
	objCnn.CursorLocation = 3
	cmdGetResults.ActiveConnection = objCnn
	cmdGetResults.CommandType =adCmdStoredProc 
	cmdGetResults.CommandText = "usp_JobInvoiceDueReport"
	
	set cmdParameter =   cmdGetResults.CreateParameter("@CompanyId",adChar,adParamInput,2,strCompanyId)   
	cmdGetResults.Parameters.Append(cmdParameter)
	
	set cmdParameter =  cmdGetResults.CreateParameter("@RunDate",adChar,adParamInput,10,strRunDate)
	cmdGetResults.Parameters.Append(cmdParameter)
	
	set cmdParameter =  cmdGetResults.CreateParameter("@EndDate",adChar,adParamInput,10,strEndDate)
	cmdGetResults.Parameters.Append(cmdParameter)

	set cmdParameter =  cmdGetResults.CreateParameter("@HighRisk",adNumeric,adParamInput)
	cmdParameter.Precision = 5
	cmdParameter.NumericScale = 2
	cmdParameter.Value = 0.5
	cmdGetResults.Parameters.Append(cmdParameter)

	set cmdParameter =  cmdGetResults.CreateParameter("@MediumRisk",adNumeric,adParamInput)
	cmdParameter.Precision = 5
	cmdParameter.NumericScale = 2
	cmdParameter.Value = 0.75
	cmdGetResults.Parameters.Append(cmdParameter)

	set rstJobResults = cmdGetResults.Execute 
	

	if not rstJobResults is Nothing	then
		Do While not rstJobResults.EOF 
			If getValue(rstJobResults,"RiskFactor") = "2" then
				iHighRiskJobCount = iHighRiskJobCount + 1
				nHighRiskAmt = nHighRiskAmt + CDbl(getValue(rstJobResults,"TotalEstimate"))
			Else
				If 	getValue(rstJobResults,"RiskFactor") = "1"	then
					iMediumRiskJobCount  = iMediumRiskJobCount + 1
					nMediumRiskAmt = nMediumRiskAmt + CDbl(getValue(rstJobResults,"TotalEstimate"))	
				Else 
					If 	getValue(rstJobResults,"RiskFactor") = "0"	then
						iNoRiskCount   = iNoRiskCount + 1
						nNoRiskAmt = nNoRiskAmt + CDbl(getValue(rstJobResults,"TotalEstimate"))	
					End If
				End If
			End If
			rstJobResults.MoveNext 
		Loop
		rstJobResults.MoveFirst 
	End If
	
%>
<script language="JavaScript" type="text/javascript">
	if (screen.width > 1024) {
	document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
	} else {
	document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
	}
</script>
</head>

	<table width="100%" border="1" cellpadding="2" cellspacing="0" ID="tblResults">
		<% if iHighRiskJobCount > 0 then %>
		<tr>
			<td  style="width:6%;font-weight:bold">&nbsp;</td>	
			<td  style="width:14%; font-weight:bold;color:red">High Risk Job Count</td>
			<td  style="width:10%;font-weight:bold;color:red" align=right ><%=iHighRiskJobCount%></td>
			<td  style="width:6%;font-weight:bold;color:red" >High Risk$</td>
			<td  style="width:6%;font-weight:bold;color:red" align=right><%=nHighRiskAmt%></td>
			<td  colspan=7>&nbsp;</td>	
		</tr>
		<% End If	%>	
		<% if iMediumRiskJobCount > 0 then %>
		<tr>
			<td  style="width:6%;font-weight:bold">&nbsp;</td>	
			<td  style="width:14%; font-weight:bold;color:red">Medium Risk Job Count</td>
			<td  style="width:10%;font-weight:bold;color:red" align=right><%=iMediumRiskJobCount%></td>
			<td  style="width:6%;font-weight:bold;color:red">Medium Risk $</td>
			<td  style="width:6%;font-weight:bold;color:red" align=right><%=nMediumRiskAmt%></td>
			<td  colspan=7>&nbsp;</td>	
		</tr>
		<% End If	%>	
		<% if iNoRiskCount > 0 then %>
		<tr>
			<td  style="width:6%;font-weight:bold">&nbsp;</td>	
			<td  style="width:14%; font-weight:bold;color:green">No Risk Job Count</td>
			<td  style="width:10%;font-weight:bold;color:green" align=right><%=iNoRiskCount%></td>
			<td  style="width:6%;font-weight:bold;color:green">No Risk $</td>
			<td  style="width:6%;font-weight:bold;color:green" align=right><%=nNoRiskAmt%></td>
			<td  colspan=7>&nbsp;</td>	
		</tr>
		<% End If	%>	
		<tr>
			<td  style="width:6%; font-weight:bold">Job#</td>
			<td  style="width:14%;font-weight:bold">Description</td>
			<td  style="width:10%;font-weight:bold">Customer</td>
			<td  style="width:6%;font-weight:bold">Start Date</td>
			<td  style="width:6%;font-weight:bold">Completion Date</td>	
			<td  style="width:8%;font-weight:bold">% Complete</td>
			<td  style="width:8%;font-weight:bold">Contract Amt</td>
			<td  style="width:8%;font-weight:bold">WIP</td>
			<td  style="width:8%;font-weight:bold">Amt Invoiced</td>
			<td  style="width:8%;font-weight:bold">Amt Paid</td>
			<td  style="width:8%;font-weight:bold">Project Manager</td>
			<td  style="width:6%;font-weight:bold">Status</td>
		</tr>
<%  	if not rstJobResults is Nothing	then
			Do While not rstJobResults.EOF 
		
	%>

		<tr bgcolor=#ffffff >
		<td style="width:6%;"><% Response.Write getValue(rstJobResults,"Job") %></td>
		<td style="width:14%;"><% Response.Write getValue(rstJobResults,"Description") %></td>
		<td style="width:10%;"><% Response.Write getValue(rstJobResults,"Customer") %></td>
		<td style="width:6%;"><% Response.Write FormatDateTime(getValue(rstJobResults,"Actual_Start_Date"),2) %></td>
		<td style="width:6%;"><% Response.Write FormatDateTime(getValue(rstJobResults,"Actual_Complete_Date"),2) %></td>
		<td style="width:8%;"><% Response.Write getValue(rstJobResults,"Percent_Complete") %></td>
		<td style="width:8%;"><% Response.Write getValue(rstJobResults,"TotalEstimate") %></td>
		<td style="width:8%;"><% Response.Write getValue(rstJobResults,"ASOFDATE_DUE_ESTIMATE") %></td>
		<td style="width:8%;"><% Response.Write getValue(rstJobResults,"INVOICED_AMT") %></td>
		<td style="width:8%;"><% Response.Write getValue(rstJobResults,"PAID_INVOICE") %></td>
		<td style="width:8%;"><% Response.Write getValue(rstJobResults,"ProjectManager") %></td>
		<%	
					If getValue(rstJobResults,"RiskFactor") = "2" then
						iHighRiskJobCount = iHighRiskJobCount + 1
						nHighRiskAmt = nHighRiskAmt + CDbl(getValue(rstJobResults,"TotalEstimate"))
				%>
		<td style="width:6%;" bgcolor="Red">High Risk</td>
		<%	
					Else
						If 	getValue(rstJobResults,"RiskFactor") = "1"	then
						iMediumRiskJobCount  = iMediumRiskJobCount + 1
						nMediumRiskAmt = nMediumRiskAmt + CDbl(getValue(rstJobResults,"TotalEstimate"))						
				%>

							<td style="width:6%;" bgcolor="#ffff99">Medium Risk</td>
		<%
						Else%>
							<td style="width:6%;" bgcolor=#ccff66>No risk</td>
					<%	End If
					End If
				%>
		</tr>
		<%
		rstJobResults.MoveNext 
		Loop 

	%>
</table> 
	
<% 	Else%>
<table width="100%" border="1" cellpadding="3" cellspacing="0" ID="tblResults">
<tr>
<td>No Records found!</td>
</tr>
</table> 
<% End If
	
	If rstJobResults.State = 1 then
		rstJobResults.Close 
	End If
	objCnn.Close
	set rstJobResults = Nothing
	set cmdGetResults = Nothing
	set objCnn = Nothing
 %>
