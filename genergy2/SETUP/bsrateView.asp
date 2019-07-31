<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim buildingNum, buildingName, portfolioid, view1970, micID
buildingNum = request("buildingNum")
if isempty(buildingNum) then
	response.write("no building was passed as a parameter")
	response.end
end if
portfolioid = request("pid")
if isempty(portfolioid) then
	portfolioid = request("portfolioid")
end if

micID = request("micID")

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

dim title, helpSource
if lcase(trim(request("view1970"))) = "true" then
	view1970 = true
	title = "1970 Rate Parameters for " & buildingName
	helpSource = "1970rate"
else
	view1970 = false
	title = "Building Specific Invoice Amount for " & buildingName
	helpSource = "bsrate"
end if


dim utility, amount, season, dateFrom, dateTo, by, bp, notes, rate, desc
'response.write(":" & trim(request("reload")) & ":")
if lcase(trim(request("reload"))) = "true" then
	utility = request("utility")
	amount = request("amount")
	season = request("season")
	dateFrom = request("dateFrom")
	dateTo = request("dateTo")
	by = request("by")
	bp = request("bp")
	notes = request("notes")
	rate = request("rateid")
	desc = request("description")
else
	dim rst, sql
	set rst = server.createobject("ADODB.recordset")
	sql = "select amt as amount, seasonid as season, [description], rateid, dateFrom, dateTo, billYear as billy, billPeriod as billp, note as notes, utilityid as utility from misc_inv_credit where id = '" & micID & "'"
	rst.open sql, cnnLocal,1
	if not rst.eof then
		rate = rst("rateid")
		if isnull(rate) or isempty(rate) then
			rate= -1
		end if
		utility = rst("utility")
		amount = rst("amount")
		season = rst("season")
		dateFrom = rst("dateFrom")
		dateTo = rst("dateTo")
		by = rst("billy")
		bp = rst("billp")
		notes = rst("notes")

		desc = rst("description")
		'if isnull(desc) then
		'	desc = 0
		'end if
	else
		rate = -1
	end if
	rst.close
end if
	
'response.write("BY:" & by & "<br>")
'response.write("BP:" & bp & "<br>")

%>
<html>
<head>
<title><%=title%></title>
<link rel="Stylesheet" href="../../styles.css" type="text/css">
<script>
function openCustomWin(clink, cname, cspec)
{	customlink = open(clink, cname, cspec)
	customlink.focus()
}

function validateInput(){
	var amount = document.forms.bsrateform.amount.value;
	var utility = document.forms.bsrateform.utility.value;
	var note = document.forms.bsrateform.notes.value;
	
	var dateFrom = document.forms.bsrateform.dateFrom.value;
	var dateTo = document.forms.bsrateform.dateTo.value;

	var billYear = document.forms.bsrateform.by.value;
	var billPeriod = document.forms.bsrateform.bp.value;
	var submitOk = true;
	var dateEmpty = false;
	if ((dateFrom == null) || (dateFrom == "") || (dateTo == "") || (dateTo == null)){		
		dateEmpty = true;
	}
	else if (  (isGoodDate(dateTo) == false)  ||  (isGoodDate(dateFrom) == false)){
		alert("Invalid Dates entered.");
		submitOk = false;
	}
	
	var bybpEmpty = false;
	if ((billYear == null) || (billPeriod == null) || (billYear == "") || (billPeriod == "") || (billYear == 0) || (billPeriod == 0)){
		bybpEmpty = true;
	}
	
	if ((utility == null) || (utility == "")){
		alert("Utility is a required field.");
		submitOk = false;
	}
	if ((note == null) || (note == "")){
		alert("A note is required.");
		submitOk = false;
	}
	if ((isNaN(amount)) || (amount == null) || (amount == "")){
		alert("Amount must be a numeric value.");
		submitOk = false;
	}
	
	if ((bybpEmpty==true) && (dateEmpty==true)){
		alert("You must provide a date or a bill year and bill period.")
		submitOk = false;
	}
	if (submitOk == true){
		document.forms.bsrateform.action = "bsrateSave.asp";
		document.forms.bsrateform.submit();
	}
}
function cancelButtonPooshed(){
	if (confirm("Are you sure you would like to cancel?") == true){
		window.history.back()
	}
}
function deleteButtonPooshed(micid){
	if (confirm("Are you sure you would like to delete this rate?") == true){
		window.document.location = "bsRateDelete.asp?micid=" + micid
	}
}
function clearDates(){
	document.forms.bsrateform.dateFrom.value = "";
	document.forms.bsrateform.dateTo.value = "";
}
function byChanged(view1970){
	if (view1970 == false){
		document.forms.bsrateform.season.value='';
		clearDates();
	}
	
	document.forms.bsrateform.reload.value='True';
	document.forms.bsrateform.submit()
}
function bpChanged(view1970){
	if (view1970 == false){
		document.forms.bsrateform.season.value='';
		clearDates();
	}
}

function seasonChanged(view1970){
	if (view1970 == false){
		document.forms.bsrateform.bp.value='0';
		document.forms.bsrateform.by.value='0'; 
	}
}
	

function isGoodDate(date){
	var valid = true;
	var month;
	var day;
	var year;
	var dateArr = new Array(3);
	dateArr = date.split("/")
	month = dateArr[0];
	//month = date.substr(0,2);
	day = dateArr[1];
	year = dateArr[2];
	//alert("month:" + month);
	//alert("day:" + day);
	//alert("year:" + year);
	
	if ((isNaN(month)) || (month < 1) || (month > 12)){
		valid = false;
	}
	if ((isNaN(day)) || (day < 1) || (day > 31)){
		valid = false;	
	}
	if (isNaN(year)){
		valid = false;
	}
	return valid;
}
	
	
</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="bsrateform" method="post" action="bsrateView.asp">
<input type="hidden" name="user" value="<%=getXMLUserName()%>">
<input type="hidden" name="buildingNum" value="<%=buildingNum%>">
<input type="hidden" name="portfolioid" value="<%=portfolioid%>">
<input type="hidden" name="reload" value="False">
<input type="hidden" name="micID" value="<%=micID%>">
<input type="hidden" name="view1970" value="<%=view1970%>">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr bgcolor="#3399cc">
		<td>
			<font color='white'><%=title%></font>
		</td>
		   <td align="right"><button id="qmark2" onclick="openCustomWin('help.asp?page=<%=helpSource%>','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
	</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr>
		<td>Utility:</td>
		<td>
			<select name="utility">
			<option value="0">All Utilities</option>
				<%
				dim sqlUtility, rstUtility
				set rstUtility = server.createobject("ADODB.recordset")
				sqlUtility = "select distinct lup.utility as utilityId, tu.utilityDisplay as utilityName from tblleasesutilityprices lup inner join tblutility tu on lup.utility = tu.UtilityId where lup.billingId in (select billingid from tblleases where bldgnum = '" & buildingNum & "')"
				rstUtility.open sqlUtility, cnnLocal
				if rstUtility.eof then
					%><script>alert("there are no utilities currently set up in this building.")</script><%
					response.End
				end if
				do while not rstUtility.eof
					dim utilityId, utilityName
					utilityId = rstUtility("utilityID")
					utilityName = rstUtility("utilityName")
					%><option value="<%=utilityId%>" <%if utility = utilityId then response.write("selected") end if%>><%=utilityName%></option><%
					rstUtility.movenext
				loop
				rstUtility.close
				%>
			</select>		
		</td>
		
	</tr>
	<tr>
		<td>Season:</td>
		<td>
			<select name = "season" onChange="javascript:seasonChanged(<%=lcase(view1970)%>);">
				<option value="0" <%if season = "0" then response.write("selected") end if%>>N/A</option>
				<%
				dim sqlSeason, rstSeason
				set rstSeason = server.createobject("ADODB.recordset")
				sqlSeason = "select sea.id, sea.season from rateseasons sea inner join buildings bill on sea.regionid = bill.region where bill.bldgNum = '" & buildingNum & "'"
				'sqlSeason = "select * from [" & application("superip") & "].mainModule.dbo.rateseasons"
				'response.write sqlSeason
				'response.end
				rstSeason.open sqlSeason, cnnLocal
				if not rstSeason.eof then
					do while not rstSeason.eof
						dim seasonId, seasonName
						seasonId = rstSeason("id")
						seasonName = rstSeason("season")
						%><option value="<%=seasonId%>" <%if cint(season) = cint(seasonId) then response.write("selected") end if%>><%=seasonName%></option><%							rstSeason.moveNext
					loop
					rstSeason.close
				end if
				%>
			</select>
		</td>
	</tr>
	<%if not view1970 then		'1970 view doesnt get dates%>
		
		<tr>
			<td>Date From: (MM/DD/YYYY)</td>
			<td>
				<input type="text" maxlength="10" size="7" name="dateFrom" value="<%=dateFrom%>" onchange="javaScript:forms.bsrateform.bp.value='0';forms.bsrateform.by.value='0'; ">
			</td>
		</tr>
		<tr>
			<td>Date To: (MM/DD/YYYY)</td>
			<td>
				
				<%
				'response.write("dateTo:" & dateTo & ":")
				'response.End
				%>
				<input type="text" maxlength="10" size="7" name="dateTo" value="<%=dateTo%>" onchange="javaScript:forms.bsrateform.bp.value='0';forms.bsrateform.by.value='0'; ">
			</td>
		</tr>
		<tr>
			<td>
				Rate that this entry applies to:
			</td>
			<td>
				<select name="rateid">
				<option value="0">All Rates</option>
					<%
					dim sqlRate, rstRate
					sqlRate =  "select type, rateTenant from ratetypes rt inner join (select distinct rateTenant from tblLeasesUtilityPrices inner join tblLeases on tblLeasesUtilityPrices.billingId = tblLeases.billingId where tblLeases.bldgNum = '" & buildingNum & "') x on rt.id = x.rateTenant"
					'response.write sqlRate
					'response.end
					set rstRate = server.createobject("ADODB.recordset")
					rstRate.open sqlRate, cnnLocal
					if not rstRate.eof then
						do while not rstRate.eof
							%>
							<script>//alert("<%'=rstRate("rateTenant")%>:::<%'=rate%>")</script>
							<option value="<%=rstRate("rateTenant")%>" <%if cint((trim(rate))) = cint(trim(rstRate("rateTenant"))) then %>selected<%end if%>><%=rstRate("type")%></option>
							<%
							rstRate.movenext
						loop
					end if
					%>
				</select>
			</td>
		</tr>

	<%else		'view1970 is true, cori will automatically create dates with a trigger%>		
		<input type="hidden" name="dateFrom" value="">
		<input type="hidden" name="dateTo" value="">
		<input name="rateid" type="hidden" value="16">

		<tr>
			<td colspan = "2">
				Enter the bill year and bill period that the 1970 rate increase is based on:
			</td>
		</tr>
	<%end if%>
		<tr>
		<td>Bill Year:</td>
		<td>
			<select name="by" onChange="javascript:byChanged(<%=lcase(view1970)%>);">
				<option value="0" <%if by = "" then response.write("selected") end if%>>N/A</option>
				<%
				dim sqlBillYear, rstBillYear, billYear
				sqlBillYear = "select distinct BillYear FROM BillYrPeriod WHERE (BldgNum = '" & buildingNum & "')"
				set rstBillYear = server.createObject("ADODB.recordset")
				rstBillYear.open sqlBillYear, cnnLocal
				if not rstBillYear.eof then
					do while not rstBillYear.eof
						billYear = rstBillYear("BillYear")
						%><option value="<%=billYear%>" <%if cint(by) = cint(billYear) then response.write("selected") end if%>>
							<%=billYear%>
						</option><%
						rstBillYear.movenext
					loop
				end if
				rstBillYear.close
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Bill Period:</td>
		<td>
			<select name="bp" onChange="javascript:bpChanged(<%=lcase(view1970)%>);">
				<option value="0" <%if bp = "" then response.write ("selected") end if%>>N/A</option>
				<%
				dim sqlBillPeriod, rstBillPeriod, billPeriod
				if by <> "" then
					sqlBillPeriod = "select BillPeriod FROM BillYrPeriod WHERE (BldgNum = '" & buildingNum & "') AND (BillYear = '" & by & "')"
					set rstBillPeriod = server.createObject("ADODB.recordset")
					rstBillPeriod.open sqlBillPeriod, cnnLocal
					if not rstBillPeriod.eof then
						do while not rstBillPeriod.eof
							billPeriod = rstBillPeriod("BillPeriod")
							%>
							<script>//alert("BillPeriod:<%=billPeriod%>:  bp:<%=bp%>:")</script>
							<option value="<%=billPeriod%>" <%if cint(bp) = cint(billPeriod) then response.write("selected") end if%>><%=billPeriod%></option><%
							rstBillPeriod.movenext
						loop
					end if
					rstBillPeriod.close
				end if
				%>
			</select>
		</td>
	<tr>
		<td>Amount:</td>
		<td>
			<input type="text" name="amount"  size="5"value="<%=amount%>">(Enter percentages as decimals)
		</td>
	</tr>
	<% if not view1970 then %>
		<tr>
			<td>Description:</td>
			<td>
				<select name="description">
					<option value=0>N/A</option>
					<%dim rstDesc
					set rstDesc = server.createobject("ADODB.recordset")
					rstDesc.open "select * from rateDescription order by description", getConnect(portfolioid,buildingNum,"billing")
					while not rstDesc.eof
						'response.write ("rst:"&rstDesc("id")&":desc:"&desc&":")
						%><option value=<%=rstDesc("id")%><%if cint(rstDesc("id")) = cint(desc) then%> selected<%end if%>><%=rstDesc("description")%></option><%
						rstDesc.moveNext
					wend%>
				</select>
			</td>
		</tr>
	<% end if%>
		
	<tr>
		<td height = "20" valign="bottom">Notes:</td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<textarea rows="3" cols="50" maxlength="250" name="notes"><%=notes%></textarea>
		</td>
	</tr>
		<tr>
		<td colspan="2" align="center" height="5"></td>
	</tr>
</table
><table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr bgcolor="#3399cc">
		<td align="center">
		<%if not(isBuildingOff(buildingnum)) then%>
			<input type="button" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="javaScript:validateInput()" name="saveButton" value="Save">&nbsp;
			<%if allowgroups("Genergy_Corp") then%>
				<input type="button" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onClick="deleteButtonPooshed('<%=micid%>')" name="deleteButton" value="Delete">
			<%end if%>
		<%end if%>
			<input type="button" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onClick="cancelButtonPooshed()" name="cancelButton" value="Cancel">&nbsp;
		</td>
	</tr>
</table>
