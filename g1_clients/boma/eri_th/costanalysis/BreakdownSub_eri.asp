<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim date1, date2, b, utype, pid, adjtype
b = request.querystring("b")
pid = request.querystring("pid")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
utype = request.querystring("utype")
adjtype = request.querystring("adjtype")
dim title, ISeri
if adjtype="sub" then
	title = "Submeter"
	ISeri = "0"
else
	title = "ERI"
	ISeri = "1"
end if

dim rst1, cnn1, sql, prm, cmd
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.Open application("Cnnstr_genergy1")
cnn1.CursorLocation = adUseClient
cmd.CommandText = "sp_eri_subm"
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 10)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adChar, adParamInput, 4)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("eri", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("tn", adVarChar, adParamOutput, 10)
cmd.Parameters.Append prm
cmd.Name = "sp_eri_subm"

cnn1.sp_eri_subm b, date1, ISeri,1, rst1
'response.write b&","&date1&","&ISeri&",1"
'response.end
%>
<html>
<head><title></title>
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame2')" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b><%=title%> Breakdown</b></font></td>
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='monthlyDetails.asp?b=<%=b%>&pid=<%=pid%>&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utype='+ parent.document.forms['form1'].utype.value" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Monthly Details</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"></font></td>
	</tr>
</table>

<%
dim i, tempname
if adjtype="eri" then
'##############################################################################################
'# ERI
'##############################################################################################
	%>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td valign="top">
	<table width="705" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<td width="65" align="center" width="">Tenant&nbsp;#</td>
	<td width="185" align="center">Tenant&nbsp;Name</td>
	<td width="65" align="center">Sqft</td>
	<td width="65" align="center">Monthly&nbsp;Charge</td>
	<td width="65" align="center">Yearly&nbsp;Charge</td>
	<td width="65" align="center">$&nbsp;/&nbsp;Sqft</td>
	<td width="65" align="center">Lease&nbsp;Exp.</td>
	<td width="65" align="center">Date&nbsp;Move</td>
	<td width="65" align="center">Out&nbsp;Date</td>
	</tr></table>
	</td></tr>
	<tr><td><div style="overflow:auto;height:150">
	<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<%
	dim totalTenants, totalSqft, totalMonthly, totalYear, averageSqft, averageYear
	totalTenants = 0
	totalSqft = 0
	do until rst1.eof
		tempname = rst1("tenantname")
		if len(tempname)>30 then tempname = left(tempname, 30)&"..."
		response.write "<tr style=""font-family: Arial, Helvetica, sans-serif; font-size: 10;"">"
		response.write "<td width=""65""><nobr>"& rst1("tenant_no") &"</nobr></td>"
		response.write "<td width=""185""><nobr>***</nobr></td>"
		response.write "<td width=""65""><nobr>"& formatnumber(rst1("sqft"),0) &"</nobr></td>"
		response.write "<td width=""65""><nobr>"& formatcurrency(rst1("Monthly_Cost")) &"</nobr></td>"
		response.write "<td width=""65""><nobr>"& formatcurrency(rst1("Yearly_Cost")) &"</nobr></td>"
		response.write "<td width=""65""><nobr>"& formatcurrency(rst1("cost_sqft")) &"</nobr></td>"
		response.write "<td width=""65""><nobr>"& rst1("Lease_Exp_Date") &"</nobr></td>"
		response.write "<td width=""65""><nobr>"& rst1("move_out_date") &"&nbsp;</nobr></td>"
		response.write "<td width=""65""><nobr>"& rst1("move_out_date") &"&nbsp;</nobr></td>"
		response.write "</tr>"
		totalTenants = totalTenants+1
		totalSqft = totalSqft + cDBL(rst1("sqft"))
		totalMonthly = totalMonthly + cDBL(rst1("Monthly_Cost"))
		totalYear = totalYear + cDBL(rst1("Yearly_Cost"))
		averageYear = averageYear + rst1("cost_sqft")
		rst1.movenext
	loop
	if AverageYear > 0 and TotalTenants > 0 then 
	averageYear = averageYear / totalTenants 
	end if %>
	</table></div>
	</td></tr></table>&nbsp;
<table border="0" cellspacing="0" cellpadding="0" align="center">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Total no. of ERI Tenants:</td><td>&nbsp;&nbsp;<td/><td><%=totaltenants%></td></tr>
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Total of ERI sqft:</td><td>&nbsp;&nbsp;<td/><td><%=formatnumber(totalSqft,0)%></td></tr>
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Total of ERI Monthly Charge:</td><td>&nbsp;&nbsp;<td/><td><%=formatcurrency(totalMonthly)%></td></tr>
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Total of ERI Yearly Charge:</td><td>&nbsp;&nbsp;<td/><td><%=formatcurrency(totalYear)%></td></tr>
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>AVG of ERI $/sqft:</td><td>&nbsp;&nbsp;<td/><td><%=formatcurrency(averageYear)%></td></tr>
</table>
	
<%else
'##############################################################################################
'# SUBMETER
'##############################################################################################
%>
	<table border="0" cellspacing="0" cellpadding="0" align="center"><tr><td valign="top">
	<table width="350" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<td align="center">Tenant&nbsp;#</td>
	<td align="center">Tenant Name</td>
	<td align="center">Square FT</td>
	</tr><%
	i = 12
	dim numberoftenants
	numberoftenants = 0
	do until rst1.eof
		if i > cint(trim(rst1("BillPeriod"))) or i = cint(trim(rst1("BillPeriod"))) then
			numberoftenants = numberoftenants +1
			tempname = rst1("billingname")
			if len(tempname)>30 then tempname = left(tempname, 30)&"..."
			response.write "<tr style=""font-family: Arial, Helvetica, sans-serif; font-size: 10;"">"
			response.write "<td><NOBR>"& rst1("tenantnum") &"</NOBR></td>"
			response.write "<td><NOBR>***</NOBR></td>"
			if adjtype<>"eri" then response.write "<td align=""right"">"& formatnumber(rst1("sqft")) &"</td>"
			response.write "</tr>"
			i = cint(trim(rst1("BillPeriod")))
		else
			i = i + 1
		end if
		rst1.movenext
	loop
%>	
	</table>
	</td><td valign="top">
	<div style="width:338; overflow:auto; height: <%=numberoftenants*15+33%>;"><!-- cint(trim(cmd.parameters("tn"))) -->
	<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<%for i=1 to 12
		response.write "<td align=""center"">"& left(monthname(i),3) &"</td>"
	next
	response.write "</tr>"
	if numberoftenants<>0 then rst1.movefirst
	do until rst1.eof
		response.write "<tr style=""font-family: Arial, Helvetica, sans-serif; font-size: 10;"">"
		for i = 1 to 12
			if not(rst1.eof) then
				if i=cint(trim(rst1("BillPeriod"))) then
					response.write "<td align=""right"">&nbsp;"& formatcurrency(rst1("amt")) &"&nbsp;</td>"
					rst1.movenext
				else
					response.write "<td align=""right"">&nbsp;"& formatcurrency(0) &"&nbsp;</td>"
				end if
			else
				response.write "<td align=""right"">&nbsp;"& formatcurrency(0) &"&nbsp;</td>"
			end if
		next
		response.write "</tr>"
	loop%>
	</table>
	</div>
	</td></tr></table>
<%end if%>
