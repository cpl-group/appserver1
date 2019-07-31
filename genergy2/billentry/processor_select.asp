<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim pid, building, byear, ypid
pid = request.querystring("pid")
building = request.querystring("building")
byear = request.querystring("byear")
ypid = request.querystring("ypid")

dim rst1, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open application("cnnstr_genergy2")
%>
<html>
<head>
<title>Bill Validation</title>
</head>
<script>
function loadportfolio()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value;
	document.location.href=newhref;
}

function loadbuilding()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value+"&building="+frm.building.value;
	document.location.href=newhref;
}

function loadyear()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value;
	document.location.href=newhref;
}

function loadperiod()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&ypid="+frm.ypid.value;
	document.location.href=newhref;
}

</script>
<body>
<table width="100%" border="0" bgcolor="#FFFFFF">
<tr>
    <td bgcolor="#3399CC" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Review/Edit+</font></b></td>
  </tr>
</table>
<form name="form1">
<select name="pid" onchange="loadportfolio()">
<option value="">Select Portfolio</option>
<%rst1.open "SELECT distinct b.portfolioid, name FROM buildings b, portfolio p WHERE p.id=b.portfolioid ORDER BY portfolioid", cnn1
do until rst1.eof%>
	<option value="<%=trim(rst1("portfolioid"))%>"<%if trim(rst1("portfolioid"))=trim(pid) then response.write " SELECTED"%>><%=rst1("name")%></option>
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
	<option value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>><%=rst1("strt")%>, <%=trim(rst1("Bldgnum"))%></option>
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
if trim(byear)<>"" and trim(building)<>"" then%>
	<select name="ypid" onchange="loadperiod()">
	<option value="">Select Bill Period</option>
	<%rst1.open "SELECT * FROM billyrperiod WHERE bldgnum='"&building&"' and billyear="&byear&" order by billperiod", cnn1
	do until rst1.eof
		%><option value="<%=rst1("ypid")%>"<%if trim(rst1("ypid"))=trim(ypid) then response.write " SELECTED"%>><%=rst1("BillPeriod")%></option><%
		rst1.movenext
	loop
	rst1.close
	%>
	</select>
<%end if%>
</form>
<%if trim(ypid)<>"" then%>
<%
rst1.open "SELECT *, u.utilitydisplay as utilitytype FROM tblbillbyperiod b, tblleasesutilityprices lup, tblleases l, tblutility u WHERE lup.leaseutilityid=b.leaseutilityid and lup.billingid=l.billingid and u.utilityid=lup.utility and ypid="&ypid, cnn1
do until rst1.eof
	response.write rst1("billingname")&"|"
	response.write rst1("utilitytype")&"|"
	response.write rst1("ypid")&"<br>"
	rst1.movenext
loop
%>

<%end if%>
</body>
</html>
