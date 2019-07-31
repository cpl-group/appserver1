<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim pid, building, byear
pid = request.querystring("pid")
building = request.querystring("building")
byear = request.querystring("byear")

dim rst1, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getConnect(pid,building,"billing")
%>
<html>
<head>
<title>Bill Validation</title>
</head>
<script>
function loadportfolio()
{	var frm = document.forms['form1'];
	var newhref = "billsummary_select.asp?pid="+frm.pid.value;
	document.location.href=newhref;
}

function loadbuilding()
{	var frm = document.forms['form1'];
	var newhref = "billsummary_select.asp?pid="+frm.pid.value+"&building="+frm.building.value;
	document.location.href=newhref;
}

function loadyear()
{	var frm = document.forms['form1'];
	var newhref = "billsummary_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value;
	document.location.href=newhref;
}

function loadperiod()
{	var frm = document.forms['form1'];
	if((frm.building.value!='')&&(frm.byear.value!='')&&(frm.bperiod.value!=''))
	{	var newhref = "bill_summary.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&bperiod="+frm.bperiod.value;
		document.frames['mainval'].location=newhref;
	}
}

</script>
<body>
<table width="100%" border="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#3399CC" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">View Bill Summary</font></b></td></tr>
</table>
<form name="form1">
<select name="pid" onchange="loadportfolio()">
<option value="">Select Portfolio</option>
<%rst1.open "SELECT distinct portfolioid FROM buildings ORDER BY portfolioid", cnn1
do until rst1.eof%>
	<option value="<%=trim(rst1("portfolioid"))%>"<%if trim(rst1("portfolioid"))=trim(pid) then response.write " SELECTED"%>><%=rst1("portfolioid")%></option>
<%	rst1.movenext
loop
rst1.close%>
</select>
<%if trim(pid)<>"" then%>
<select name="building" onchange="loadbuilding()">
<option value="">Select Building</option>
<%
rst1.open "SELECT BldgNum, strt FROM buildings WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
do until rst1.eof%>
	<option value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>><%=rst1("strt")%></option>
<%	rst1.movenext
loop
rst1.close
%>
</select>
<%end if
if trim(building)<>"" then%>
	<select name="byear" onchange="loadyear()">
	<%rst1.open "SELECT Distinct BillYear FROM BillYrPeriod WHERE BldgNum='"&building&"'", cnn1
	if rst1.eof then
		response.write "<option value="""">No Billing Years</option>"
	else
		response.write "<option value="""">Select Bill Year</option>"
	end if
	do until rst1.eof
		%><option value="<%=rst1("Billyear")%>"<%if trim(rst1("billyear"))=trim(byear) then response.write " SELECTED"%>><%=rst1("Billyear")%></option><%
		rst1.movenext
	loop
	rst1.close
	%>
	</select>
<%end if
if trim(byear)<>"" then%>
	<select name="bperiod">
	<option value="">Select Bill Period</option>
	<%rst1.open "SELECT Distinct BillPeriod FROM BillYrPeriod WHERE BldgNum='"&building&"' and BillYear="&byear, cnn1
	do until rst1.eof
		response.write "<option value="""&rst1("BillPeriod")&""">"&rst1("BillPeriod")&"</option>"
		rst1.movenext
	loop
	rst1.close
	%>
	</select>
	<input type="button" name="action" value="View" onclick="loadperiod()">
<%end if%>
</form>
<iframe src="/null.htm" name="mainval" id="mainval" width="100%" height="100%" marginwidth="0" marginheight="0" style="background-color: White;"></iframe>
</body>
</html>
