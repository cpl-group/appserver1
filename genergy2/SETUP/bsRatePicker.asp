<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim buildingNum, portfolioid, view1970, buildingName
buildingNum = request("buildingNum")
if isempty(buildingNum) then
	response.write("no building was passed as a parameter")
	response.end
end if
portfolioid = request("pid")
if isempty(portfolioid) then
	portfolioid = request("portfolioid")
end if

dim title



dim cnnLocal, cnnMainModule, sqlName, rstName
set cnnMainModule = server.createobject("ADODB.connection")
cnnMainModule.open getConnect(portfolioid,buildingNum,"billing")
set cnnLocal = server.createobject("ADODB.connection")
set rstName = server.createobject("ADODB.recordset")
cnnLocal.open getLocalConnect(buildingNum)

sqlName = "select bldgname from buildings where bldgnum = '" & buildingNum & "'"
rstName.open sqlName, cnnLocal
if not rstName.eof then
	buildingName = rstName("bldgname")
end if
rstName.close

view1970=request("view1970")
if lcase(view1970) = "true" then
	view1970 = true
	title = "1970 Rate Parameters for " & buildingName
else
	view1970 = false
	title = "Building Specific Invoice Amount for " & buildingName
end if
%>	
<link rel="Stylesheet" href="setup.css" type="text/css">
<title><%=title%></title>
<script>
function closeWinda(){
if (confirm('Are you sure you would like to cancel?') == true){window.close()}
}
</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td>
			<font color='white'><%=title%></font>
		</td>
		<td align = "right">
			<input name="close"  style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" type="button" value="Close Window" onclick="javascript:closeWinda();">
			<%if not(isBuildingOff(buildingnum)) then%>
			<input name="new" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" type="button" value= "New Entry" onclick="javascript:window.location='bsrateView.asp?buildingNum=<%=buildingNum%>&pid=<%=portfolioId%>&view1970=<%=view1970%>'">
			<%end if%>
		</td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#cccccc">
		<td align="center">Utility:</td>
		<td align="center">Season:</td>
		<%if view1970=true then%>
			
		<%else%>
			<td align="center">Bill Year:</td>
			<td align="center">Bill Period:</td>
			<td align="center">Rate:</td>
		<%end if%>
		<td align="center">Dates:</td>
		<td align="center">Amount:</td>
	</tr>
	<%
	dim sqlMiscCredits, rstMiscCredits
	sqlMiscCredits = "select mic.id, mic.datefrom, mic.dateto, mic.billyear, mic.billperiod, mic.amt, mic.utilityid, mic.seasonid, mic.rateid from misc_inv_credit mic where bldgnum = '" & buildingNum & "'"
	if view1970=true then
		sqlMiscCredits = sqlMiscCredits & " AND mic.rateid = '16'"
	else
		sqlMiscCredits = sqlMiscCredits & " AND mic.rateid != '16'"
	end if
	set rstMiscCredits = server.createobject("ADODB.recordset")
	'response.write sqlMiscCredits
	'response.end
	rstMiscCredits.open sqlMiscCredits, cnnLocal
	if not rstMiscCredits.eof then
		do while not rstMiscCredits.eof
			%>
			<tr bgColor="#FFFFFF" onClick="javascript:window.location='bsrateView.asp?buildingNum=<%=buildingNum%>&pid=<%=portfolioId%>&view1970=<%=view1970%>&micID=<%=rstMiscCredits("id")%>'" onMouseOver="this.bgColor = 'lightgreen';" onMouseOut="this.bgColor = '#FFFFFF';">
				<td align="center">
					<%
					dim rstUtility
					set rstUtility = server.createobject("ADODB.recordset")
					rstUtility.open "select utility from tblutility where utilityid = '" & rstMiscCredits("utilityid") & "'", cnnMainModule
					if not rstUtility.eof then 
						response.write(rstUtility("utility"))
					else
						response.write "All Utilities"
					end if
					%>
				</td>
				<td align="center">
					<%
					dim rstSeason, sneezin
					set rstSeason = server.createobject("ADODB.recordset")
					rstSeason.open "select season from rateSeasons where id = '" & rstMiscCredits("seasonid") & "'", cnnMainModule
					if not rstSeason.eof then
						sneezin = rstSeason("season")
						if sneezin = "0" then sneezin = "-" end if
					else
						sneezin = "-"
					end if
					response.write sneezin
					%>
				</td>
				<%if view1970=false then%>
					<td align="center">
						<%dim billYear
						billYear = rstMiscCredits("billYear")
						if billYear = "0" then
							response.write("-")
						else
							response.write(billYear)
						end if%>
					</td>
					<td align="center">
						<%dim billPeriod
						billPeriod = rstMiscCredits("billPeriod")
						if billPeriod= "0" then
							response.write("-")
						else
							response.write(billPeriod)
						end if%>					
					</td>
					<td align="center">
						<%dim rstRate, sqlRate
						set rstRate = server.createobject("ADODB.recordset")
						sqlRate = "select type from rateTypes where id = '" & rstMiscCredits("rateid") & "'"
						rstRate.open sqlRate,cnnMainModule
						if not rstRate.eof then response.write rstRate("type") else response.write "All Rates" end if
						rstRate.close
						%>
					</td>
				<%end if%>
				<td align="center"><%=rstMiscCredits("dateFrom")%> - <%=rstMiscCredits("dateTo")%></td>
				<td align="center"><%=rstMiscCredits("amt")%></td>
			</tr>
		<%
		rstMiscCredits.movenext
		loop
	end if
	%>
	<tr bgcolor="#3399cc" height="10%">
		<td colspan="7">
		</td>
	</tr>	
	
</table>

