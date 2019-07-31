<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim startdate, enddate, asatdate, groupname, bldg, datasource,pid, sqlstr, daterange, submitted
pid = request("pid") 
startdate = request("startdate")
enddate = request("enddate")
asatdate = request("asatdate")
groupname = request("groupname")
bldg = request("bldg")
daterange = request("daterange")
submitted = request("formsubmitted")

'response.write(bldg)
'response.Write(submitted)
'response.Write(pid)

if trim(bldg)="" then bldg=""

dim rst1, cnn1 

set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")

if bldg<>"" then cnn1.open getLocalConnect(bldg)

'If using a date range then the default end date is the current date and the default start date is 30 days prior
if daterange then
	if not(isdate(startdate)) then
	startdate = dateadd("d",-30,date())
	end if
	if not(isdate(enddate)) then enddate = date()
'If using a specific date then the default si teh current date
else
	if not(isdate(asatdate)) then 
		asatdate = date()
	end if
end if



dim portfolio
%>
<script>

function showdate() {	
		document.form2.formsubmitted.value = "1"
		document.form2.submit()			
	}

</script>
<html>
<head>
	<title>Meter Data Source Checking</title>
	<link rel="stylesheet" type="text/css" href="StylesReports.css" />
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>


<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr>
  <td bgcolor="#3399CC">
    <span class="standardheader">Meter Datasource Data Check</span></td>
  </tr>
<tr>
  <td bgcolor="#eeeeee">
<form name="form2" action="DataSourceCheck2.asp" method="get">
<input name="formsubmitted" type="hidden" value="2" >
<select name="bldg" onchange="document.forms[0].submit()">
<%
'Get the list of all buildings per peorfilio for the drop down
  if trim(pid) <> "" then 
  	sqlstr = "SELECT * FROM buildings b, portfolio p WHERE p.id=b.portfolioid and p.id = "&trim(pid)& " ORDER BY name, strt"
  else
	sqlstr = "SELECT * FROM buildings b, portfolio p WHERE p.id=b.portfolioid ORDER BY name, strt"
  end if 
  rst1.open sqlstr, getConnect(0,0,"dbCore")
  do until rst1.eof
    if portfolio<>rst1("name") then
      portfolio = rst1("name")
      %><optgroup label="<%=portfolio%>"><%
    end if
    'Now gray out all offline buildings
    %><option <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=rst1("bldgnum")%>" <%if bldg=rst1("bldgnum") then response.write "SELECTED"%>><%=rst1("strt")%>, (<%=rst1("bldgnum")%>)</option><%
    rst1.movenext
  loop
  rst1.close
%>
</select>
 <% ' the date range check box is unchecked then this signifies use of a specific date and the start date and end date fields will be hidden
 if daterange=false then %>
	Check Dates as at <input type="text" name="asatdate" value="<%=asatdate%>" size="9" maxlength="10" > OR
<%end if%>
	<input type=checkbox  name="daterange" value="1" onclick="showdate();" <%if daterange then %>checked<%end if%>> Check Date Range 
 <% 'if the dat range check box is slected then show the startdate and end date fields
 if daterange then %>
          <input type="text" name="startdate" value="<%=startdate%>" size="9" maxlength="10" >&nbsp;through&nbsp;<input type="text" name="enddate" value="<%=enddate%>" size="9" maxlength="10" >
 <%end if%>
&nbsp;<input type="submit" value="reload" ID="Submit1" NAME="Submit1">
</td></tr>
<tr><td>
<%
'This section calls the stroed procedure which returns the results
if bldg<>"" then	
	
	dim cmd, p, prm1, prm2, prm

	cnn1.CursorLocation = adUseClient
	
	set cmd = server.createobject("ADODB.Command") 'connect the command
	cnn1.CursorLocation = adUseClient
		
	cmd.CommandType = adCmdStoredProc
	cmd.CommandTimeout = 0	
	
	if daterange then
	
		cmd.CommandText = "sp_datacheck_period"
		Set prm1 = cmd.CreateParameter("bldgnum", advarchar, adParamInput, 20)
		cmd.Parameters.Append prm1
		Set prm1 = cmd.CreateParameter("date1", advarchar, adParamInput, 12)
		cmd.Parameters.Append prm1
		Set prm1 = cmd.CreateParameter("date2", advarchar, adParamInput, 12)
		cmd.Parameters.Append prm1
		
		cmd.Parameters("bldgnum")	= bldg
		cmd.Parameters("date1")		= startdate
		cmd.Parameters("date2")		= enddate
	
	else
	
		cmd.CommandText = "sp_datacheck"
		Set prm1 = cmd.CreateParameter("bldgnum", advarchar, adParamInput, 20)
		cmd.Parameters.Append prm1
		Set prm1 = cmd.CreateParameter("date1", advarchar, adParamInput, 12)
		cmd.Parameters.Append prm1
		
		cmd.Parameters("bldgnum")	= bldg
		cmd.Parameters("date1")		= cstr(asatdate)
		
	end if
	
	cmd.Name = "test"
	Set cmd.ActiveConnection = cnn1
	
	cnn1.test  rst1
	
	
	
%>
</form>
<P>&nbsp;</P>
<TABLE id="Table2" height=30 cellSpacing=0 cellPadding=0 width=168 border=1>
  <TBODY>
  <TR>
    <TD width=50 bgColor=#ff0000 height=1><FONT size=1></FONT></FONT></FONT></TD>
    <TD height=1><FONT face=Verdana><FONT size=1><STRONG>Manually Read</STRONG> </FONT></FONT></TD></TR>
  <TR>
    <TD width=50 bgColor=#ffbc33 height=17><FONT face=Verdana size=1></FONT></TD>
    <TD height=17><STRONG><FONT face=Verdana size=1>Offline and Remote</FONT></STRONG></TD></TR>
  <TR>
    <TD width=50 bgColor=#85e300><FONT face=Verdana size=1></FONT></TD>
    <TD><STRONG><FONT face=Verdana size=1>Online and Remote</FONT></STRONG></TD></TR></TBODY>
</TABLE>

</td></tr>
</table>
<%
if not rst1.EOF then
	'bldgname = rst1.bldgname
%>
<div>
<table border="0" cellpadding="0" cellspacing="0" width="1238" class="xl7015389" style='border-collapse:collapse;table-layout:fixed;width:931pt' ID="Table1">
	<col class="xl8215389" width="169" style='mso-width-source:userset;mso-width-alt: 6180;width:127pt'>
	<col class="xl8315389" width="89" style='mso-width-source:userset;mso-width-alt: 3254;width:67pt'>
	<col class="xl8415389" width="89" style='mso-width-source:userset; mso-width-alt:2413'>
	<col class="xl8515389" width="89" style='mso-width-source:userset;mso-width-alt: 3254;width:67pt'>
	<col class="xl8315389" width="122" style='mso-width-source:userset;mso-width-alt: 4461;width:92pt'>
	<col class="xl8415389" width="128" style='mso-width-source:userset;mso-width-alt: 4681;width:96pt'>
	<col class="xl8515389" width="89" style='mso-width-source:userset;mso-width-alt: 3254;width:67pt'>
	<col class="xl8315389" width="89" span="2" style='mso-width-source:userset; mso-width-alt:3254;width:67pt'>
	<col class="xl8415389" width="89" style='mso-width-source:userset;mso-width-alt: 3254;width:67pt'>
	<col class="xl8515389" width="94" style='mso-width-source:userset;mso-width-alt: 3437;width:71pt'>
	<col class="xl8315389" width="83" style='mso-width-source:userset;mso-width-alt: 3035;width:62pt'>
	<col class="xl8615389" width="108" style='mso-width-source:userset;mso-width-alt: 3949;width:81pt'>
<tr height="33" style='mso-height-source:userset;height:24.95pt'>
	<td colspan="3" height="33" class="xl6915389" width="347" style='height:24.95pt;  width:194pt'>Portfolio</td>
	<td colspan="3" class="xl6915389" width="339" style='border-left:none;width:255pt'>Building</td>
	<%if daterange then %>
		<td colspan="3" class="xl6915389" width="267" style='border-left:none;width:201pt'>Date Range</td>
	<%else%>
		<td colspan="3" class="xl6915389" width="267" style='border-left:none;width:201pt'>As At</td>
	<%end if%>
	<td colspan="2" class="xl6915389" width="183" style='border-left:none;width:138pt'>Biller</td>
	<td colspan="2" class="xl6915389" width="191" style='border-left:none;width:143pt'>Meter Total</td>
</tr>
<tr height="33" style='mso-height-source:userset;height:24.95pt'>
	<td colspan="3" height="33" class="xl7115389" style='height:24.95pt'><%=portfolio%></td>
	<td colspan="3" class="xl7115389" style='border-left:none'><%=rst1("bldgnum")%> , <%=rst1("bldgname")%></td>
	<%if daterange then %>
		<td colspan="3" class="xl7115389" style='border-left:none'><%=startdate%> To <%=enddate%></td>
	<%else %>
		<td colspan="3" class="xl7115389" style='border-left:none'><%=asatdate%></td>
	<%end if%>
	<td colspan="2" class="xl7115389" style='border-left:none'></td>
	<td colspan="2" class="xl7115389" style='border-left:none'><%=rst1.RecordCount%></td>
</tr>
<tr height="19" style='height:14.25pt'>
	<td height="19" class="xl7215389" style='height:14.25pt'></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
</tr>
<tr height="33" style='mso-height-source:userset;height:24.95pt'>
	<td height="33" class="xl7215389" style='height:24.95pt'></td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
	<td colspan="3" class="xl7315389">Meter Information</td>
	<td colspan="4" class="xl7415389" style='border-right:1.5pt solid black;border-left:none'>Interval %</td>
	<td class="xl7615389">&nbsp;</td>
	<td class="xl7215389"></td>
	<td class="xl7215389"></td>
</tr>
<tr class="xl6815389" height="26" style='mso-height-source:userset;height:20.1pt'>
	<td height="26" class="xl6315389" style='height:20.1pt'>Tenant Name</td>
	<td class="xl6415389" style='border-left:none'>Floor</td>
	<td class="xl6515389" style='border-left:none'>Online Flag</td>
	<td class="xl6615389" style='border-top:none'>Meter ID</td>
	<td class="xl6415389" style='border-top:none;border-left:none'>Meter Type</td>
	<td class="xl6515389" style='border-top:none;border-left:none'>Read Method</td>
	<td class="xl6615389" style='border-top:none'>Actual</td>
	<td class="xl6415389" style='border-top:none;border-left:none'>Daily</td>
	<td class="xl6415389" style='border-top:none;border-left:none'>Weekly</td>
	<%if daterange then %>
		<td class="xl6515389" style='border-top:none;border-left:none'>Date Range</td>
	<%else%>
		<td class="xl6515389" style='border-top:none;border-left:none'>Monthly</td>
	<%end if %>	
	<td class="xl6615389" style='border-top:none'>LM#</td>
	<td class="xl6415389" style='border-left:none'>LM CC</td>
	<td class="xl6715389" style='border-left:none'>Modbus ID</td>
</tr>
<%
 Do While Not rst1.EOF
	'This section chooses the rigt flag colors depending on the read and remote stauts of the meter
	'The styles used will change 
	Dim styling, stylingnum, off
 
	if rst1("remote_stat")= "Manual" then 
		styling = "xl8423594"
		'stylingnum = "xl8423595"
	elseif rst1("meter_stat")=  "Offline" then
			styling = "xl8423596" 'set flag  orange when remote meter is offline
	else
		styling = "xl8423595"
	end if
%>
<tr height="18" style='height:13.5pt' class="xl6323594"   >
	<td height="18" class="xl7815389" style='height:13.5pt'><%=rst1("tname")%></td>
	<td class="xl7815389" style='border-left:none'><%=rst1("flr")%></td>
	<td class=<%=styling%>  style='border-left:none'><%=rst1("meter_stat")%></td>
	<td class=<%=styling%>> <%=rst1("meterid")%></td>
	<td class=<%=styling%> style='border-left:none'><%=rst1("metertype")%></td>
	<td class=<%=styling%>  style='border-left:none'><%=rst1("remote_stat")%></td>
	<%if daterange then %>
		<td class=<%=styling%>></td>
		<td class=<%=styling%>  style='border-left:none'></td>
		<td class=<%=styling%> style='border-left:none'></td>
		<td class=<%=styling%> style='border-left:none'><%=rst1("pratio")%></td>
	<%else%>
		<td class=<%=styling%>><%=rst1("cratio")%></td>
		<td class=<%=styling%> style='border-left:none'><%=rst1("dratio")%></td>
		<td class=<%=styling%> style='border-left:none'><%=rst1("wratio")%></td>
		<td class=<%=styling%> style='border-left:none'><%=rst1("mratio")%></td>
	<%end if%>	
	<td class="xl8015389"><%=rst1("lmnum")%></td>
	<td class="xl7815389" style='border-left:none'><%=rst1("lmchannel")%></td>
	<td class="xl7815389" style='border-left:none'><%=rst1("modbusid")%></td>
</tr>
	<% 
		rst1.MoveNext
	Loop
	rst1.Close
	cnn1.Close
	%>
<![if supportMisalignedColumns]>
<tr height="0" style='display:none'>
	<td width="169" style='width:127pt'></td>
	<td width="89" style='width:67pt'></td>
	<td width="0"></td>
	<td width="89" style='width:67pt'></td>
	<td width="122" style='width:92pt'></td>
	<td width="128" style='width:96pt'></td>
	<td width="89" style='width:67pt'></td>
	<td width="89" style='width:67pt'></td>
	<td width="89" style='width:67pt'></td>
	<td width="89" style='width:67pt'></td>
	<td width="94" style='width:71pt'></td>
	<td width="83" style='width:62pt'></td>
	<td width="108" style='width:81pt'></td>
</tr>
<![endif]>
</table>
</div>
<%
end if 
end if
%>
</body>
</html>
