<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim buildingNum, portfolioid, view1970, buildingName
buildingNum = request("buildingNum")
if isempty(buildingNum) then buildingNum = request("bldg")
if isempty(buildingNum) then
	response.write("no building was passed as a parameter")
	response.end
end if
portfolioid = request("pid")
if isempty(portfolioid) then
	portfolioid = request("portfolioid")
end if

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

dim buildingCPFrom
buildingCPFrom = request("bldgCPFrom")
%>	
<link rel="Stylesheet" href="setup.css" type="text/css">
<script>
function copy(){
	var sql = "exec sp_copy_billperiods @bldg varchar(20),@obldg varchar(20),@by int,@bp int,@utility smallint"
}
function hideall(){
	document.all["pageform"].style.display='none';
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<div id="pageform">
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td>
			<span class="standardheader"><font color='white'>Copy Bill Periods to <%=buildingName%></font></span>
		</td>
	</tr>
</table>
<form name="bldgCPFromForm" action = "copyBPFromBuilding.asp" method="get" onsubmit="hideall();">
<input type="hidden" name="buildingNum" value="<%=buildingNum%>">
<table>
	<tr>
		<td>Copy from building:</td>
		<td>
			<select name="bldgCPFrom" onChange="hideall();document.forms.bldgCPFromForm.billyear.value='0';submit()">
				<%
				dim rstAllBldg, sqlAllBldg
				set rstAllBldg = server.createobject("ADODB.recordset")
				sqlAllBldg = "select strt, b.bldgnum,tripcode from buildings b left join ["&application("coreip")&"].dbcore.dbo.super_tripcodes s on b.bldgnum=s.bldgnum order by tripcode, strt"
				'response.write ("</select>" & sqlAllBldg)
				'response.end
				rstAllBldg.open sqlAllBldg, cnnLocal
				if not rstAllBldg.eof then
					do while not rstAllBldg.eof
						dim bldgNameF, bldgNumF, epi
						if len(rstAllBldg("strt")) > 25 then epi = "..." else epi = ""
						bldgNameF = left(rstAllBldg("strt"),25) & epi & "  (" & rstAllBldg("bldgnum") & ")"
						bldgNumF = rstAllBldg("bldgnum")
						%><option value="<%=bldgNumF%>" <%if buildingCPFrom = bldgNumF then response.write "selected" end if%>>[Trip <%=rstAllBldg("tripcode")%>] <%=bldgNameF%></option><%
						rstAllBldg.movenext
					loop
				end if
				%>
			</select>
		</td>		
	</tr>
	<tr>
		<td>Select utility to copy from</td>
		<td>
			<select name = "utility" onChange="hideall();submit()"><%
				if buildingCPFrom <> "" then
					if buildingCPFrom <> buildingNum then%>
						<option value="0">Copy all utilities</option><%
					end if
					dim sqlUtility, rstUtility
					set rstUtility = server.createobject("ADODB.recordset")
					sqlUtility = "select distinct bybp.utility as utilityId, tu.utilityDisplay as utilityName from billyrperiod bybp inner join tblutility tu on bybp.utility = tu.UtilityId where bybp.bldgnum = '" & buildingCPFrom & "'"
					'response.write "<!--" & sqlUtility & "-->"
					rstUtility.open sqlUtility, getLocalConnect(buildingCPFrom)
					if rstUtility.eof then
						%><script>alert("there are no utilities currently set up in this building.")</script><%
						'response.End
					else
						do while not rstUtility.eof
							dim utilityId, utilityName
							utilityId = rstUtility("utilityID")
							utilityName = rstUtility("utilityName")
							%><option value="<%=utilityId%>" <%if cint(utilityId) = cint(request("utility")) then response.write("selected") end if%>><%=utilityName%></option><%
							rstUtility.movenext
						loop
						rstUtility.close
					end if
				end if%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Select utility to copy to</td>
		<td>
			<select name = "utilityto">
				<%
				if buildingCPFrom <> "" then
					if cint(request("utility"))="0" then
						%><option value="0">Copy to all utilities</option><%
					end if
					if cint(request("utility"))<>"0" then
						sqlUtility = "select distinct bybp.utility as utilityId, tu.utilityDisplay as utilityName from billyrperiod bybp inner join tblutility tu on bybp.utility = tu.UtilityId where bybp.bldgnum = '" & buildingNum & "'"
						sqlUtility = "SELECT u.utilityid as utilityId, u.utilityDisplay as utilityName FROM tblutility u WHERE u.utilityid in (SELECT utility FROM tblleasesutilityprices lup INNER JOIN tblleases l ON lup.billingid=l.billingid and l.bldgnum='" & buildingNum & "')"
						response.write sqlUtility
						rstUtility.open sqlUtility, getLocalConnect(buildingNum)
						if rstUtility.eof then
							%><script>alert("there are no utilities currently set up in the destination building.")</script><%
							'response.End
						else
							do while not rstUtility.eof
								utilityId = rstUtility("utilityID")
								utilityName = rstUtility("utilityName")
								if not(buildingCPFrom = buildingNum and trim(request("utility")) = trim(utilityId)) then
									%><option value="<%=utilityId%>" <%if trim(utilityId) = trim(request("utilityto")) then response.write("selected") end if%>><%=utilityName%></option><%
								end if
								rstUtility.movenext
							loop
							rstUtility.close
						end if
					end if
				end if
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Bill Year:</td>
		<td>
			<select name = "billyear" onChange="hideall();submit()">
				<%
				%><option value="0" onClick="document.forms.bldgCPFromForm.billperiod.value='0'">Copy All Bill Years</option><%
				if buildingCPFrom <> "" then
					
					dim sqlBillYear, rstBillYear, billYear
					sqlBillYear = "select distinct BillYear FROM BillYrPeriod WHERE (BldgNum = '" & buildingCPFrom & "')"
					set rstBillYear = server.createObject("ADODB.recordset")
					rstBillYear.open sqlBillYear, getLocalConnect(buildingCPFrom)
					if not rstBillYear.eof then
						do while not rstBillYear.eof
							billYear = rstBillYear("BillYear")
							%><option value="<%=billYear%>" <%if request("billyear") = billYear then response.write("selected") end if%>>
								<%=billYear%>
							</option><%
							rstBillYear.movenext
						loop
					end if
					rstBillYear.close
				end if
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Bill Period:</td>
		<td>
			<select name="billperiod">
				
				<%
				if buildingCPFrom <> "" and request("billyear") <> "" then
					if request("billyear") = "0" then
						%><option value="0">Copy All Bill Periods</option><%
					end if
					dim sqlBillPeriod, rstBillPeriod, billPeriod
					sqlBillPeriod = "select distinct BillPeriod FROM BillYrPeriod WHERE (BldgNum = '" & buildingCPFrom & "') AND (BillYear = '" & request("billyear") & "') order by BillPeriod"
					set rstBillPeriod = server.createObject("ADODB.recordset")
					rstBillPeriod.open sqlBillPeriod, getLocalConnect(buildingCPFrom)
					if not rstBillPeriod.eof then
						do while not rstBillPeriod.eof
							billPeriod = rstBillPeriod("BillPeriod")
							%><option value="<%=billPeriod%>" 
								<%if request("billperiod") = billPeriod then response.write("selected") end if%>
								><%=billPeriod%>
							</option><%
							rstBillPeriod.movenext
						loop
					end if
					rstBillPeriod.close
				end if
				%>
			</select>
		</td>
	</tr>
</table>					
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td align="center">
			<input name="ok" type="button" value="Copy" onClick="document.forms.bldgCPFromForm.action='copyBPFromBuildingSave.asp';submit()">
			<input name="cancel" value="Cancel" type="button" onClick="window.close()">
		</td>
	</tr>
</table>
</div>
</body></html>