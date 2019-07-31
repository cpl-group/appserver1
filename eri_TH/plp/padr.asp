<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Option Explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<title>Power Availability Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF">

		
<%
'connect to tlbdg to get a building name using its bldgid
dim cnn, rs, sql, bldgnum
Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cnn.open getConnect(0,0,"engineering")
sql = "SELECT address from tlbldg WHERE bldgnum = '" & Request("bldgid") & "'"
bldgnum = Request("bldgid")
rs.Open sql, cnn

dim buildingName

if rs.eof then
	Response.write("Building ID not found.")
	Response.end
else
	buildingName = rs("address")
end if
rs.close	

dim printableView
printableView = Request("prview")
if printableView = "" then
	printableView = 0
end if


if printableView = 0 then		'printableView = 0
	dim link
	link = "http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?devIP="&request.servervariables("SERVER_NAME")&"&sn=/eri_TH/plp/padr.asp&qs="&server.urlencode("bldgid="&Request("bldgID")&"&prview=1")
	'Response.Write(link)
	%>
	
	<table width="100%" border="0" bgcolor="#6699cc" cellpadding="0" cellspacing="0">
		<tr height>
			<td><font color="#FFFFFF" size = "4"><span class="standardheader"> Power Availability Data Report for <%=buildingName%></font></span></td>
			<td align = "right">
				<a href = "<%=link%>"><img src = "/images/print_pdf.gif" border=0></a>
			</td>
		</tr>
	</table>
	</span>

	<%
else
	dim i
	for i = 1 to 27 step 1%>
		<br><%
	next
	%>
	<center><h2>Power Availability Data Report
	
	<br><br><br><br><br>
	
	<%=buildingName%></h2></center>
	<%
	for i = 1 to 36 step 1%>
		<br><%
	next
	
end if
%>	
<!-- legend table -->
<b>
<table width="100%" border="0" bgcolor="#CCCCCC">
	<tr> 
		<td width="4%"></td>
		<td width="7%">Legend</td>
		<td width="11%"></td>
		<td width="39%"></td>
		<td width="5%">Notes:</td>
		<td width="24%">sw = switch/breaker</td>
	</tr>
	<tr> 
		<td></td>
		<td></td>
		<td><font color="#FF0000">High Voltage Risers</font></td>
		<td></td>
		<td></td>
		<td># See Voltage Drop Detail Report</td>
	</tr>
	<tr> 
		<td></td>
		<td></td>
		<td><font color="#0000FF">Low Voltage Risers</font></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<!--<tr> 
		<td width = "100%" colspan =5 bgcolor="#6699cc" height=3></td>
	</tr>-->
</table>
</b>

<%
if printableView = 1 then
	%>
	<WxPrinter PageBreak>
	<%
end if
%>	
<br><br><br>
<%
sql = "exec sp_pwr_by_flr2_rob_test '" + Request("bldgid") + "'"
rs.Open sql, cnn

if not (rs.eof) then
	dim lineCount
	lineCount = 0		' this is for page breaks in the pdf
	do while not (rs.eof)
	
		dim thisFloor, lastFloor, totalPowerAvailable, totalAmpsAvailable
		thisFloor	= rs("flr")	' is updated in the inner loop and compared to last floor
		lastFloor	= rs("flr")	' is updated in the outer loop
		totalPowerAvailable = 0		' per floor
		totalAmpsAvailable = 0		' per floor
		if ((lineCount > 16) AND (printableView = 1)) then
			lineCount = 0
			%>
			<WxPrinter PageBreak>
			<br><br><br>
			<%
		end if
		%>
		
		<br>
		<!-- top most table, the floor number --> 
		<table width="100%"   cellspacing=0 >
			<tr >
				<td  width="4%"  height="3">
				<td colspan=3 width="96%" bgcolor="#6699cc" height="3">
			</tr>
			<tr> 
				<td width="4%"></td>
    			<td width="18%" align = "right" bgcolor="#000000"><h3><font color="#FFFFFF"><%= rs("flr") %></font></h3></td>
				<td width="10%" bgcolor="#6699cc"><h3>Floor</h3></td>
				<td width="68%"></td>
			</tr>
		</table>
	
		<!-- second table, usable floor area -->
		<table width="40%" cellspacing=0>
			<tr> 
			<td width="10%"></td>
			<td width="52%"><h4><strong>Floor area usable</strong></h4></td>
			<td width="38%"><h4><%= formatnumber(rs("sqft"),0) %> </h4></td>
			</tr>
		</table>

		<!-- the main table, top row is the label for each column -->
		<table width="100%" cellspacing=7 cellpadding=0>
  			<tr>
  			<td width="4%" ></td>
    		<td width="9%" ><strong>Riser Or Switch</strong></td>
    		<td width="3%"  align="center"><strong>Set</strong></td>
    		<td width="6%"  align="center"><strong>Size(MCM)</strong></td>
    		<td width="6%"  align="center"><strong>Volts(V)</strong></td>
    		<td width="9%"  align="center"><strong>Wire Capacity (A)</strong></td>
    		<td width="8%"  align="center"><strong>SF_Frame (A)</strong></td>
    		<td width="7%"  align="center"><strong>SW_Fuse (A)</strong></td>
    		<% if bldgnum = "10016thAVE" OR bldgnum = "044SW" then%>
    		<td width="7%"  align="center"><strong>(*)Power Factor</strong></td>
    		<td width="7%"  align="center"><strong>(*)Safety Factor</strong></td>
    		<%else%>
    		<td width="7%"  align="center"><strong>Power Factor</strong></td>
    		<td width="7%"  align="center"><strong>Safety Factor</strong></td>
    		<%end if%>
    		<td width="7%"  align="center"><strong>Floor served</strong></td>
    		<td width="10%" align="center"><strong>Area served (usable Sq. ft.)</strong></td>
    		<td width="7%" align="center"><strong>Amps Available</strong></td>
    		<td width="10%" align="center"><strong>Power available (W/sq.ft)</strong></td>
			</tr>
			
			<tr>
				<td width="4%" cellpadding=0  height="2"></td>
				<td width="96%" cellpadding=0 colspan=13 bgcolor="#6699cc" height="2"></td>
			</tr>
  		<%
		lineCount = lineCount + 5	
		do until ((rs.eof) OR (lastFloor <> thisFloor))
			totalPowerAvailable = cdbl(totalPowerAvailable) + cdbl(rs("wsqft"))
			totalAmpsAvailable = cdbl(totalAmpsAvailable) + cdbl(rs("amps_available"))
			%>
			<tr> 
			<!-- this row of the main table is created in the inner loop, it is the actual data for each floor-->
			<td align="center"></td>
    		<td><strong>
				<%if (cdbl(rs("volts")) > 210) then %>
					<font color="#FF0000"> 
				<%else%> 
					<font color="#0000FF">
				<%end if%>
			<%= rs("risername") %>
			</font></strong>
			</td>
			<td align="center"><%= rs("sets") %></td>
			<td align="center"><%= rs("size") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= rs("metal") %></td>
			<td align="center"><%= rs("volts") %></td>
			<td align="center"><%= rs("wire_capacity") %></td>
			<td align="center"><%= rs("sw_frame") %></td>
			<td align="center"><%= rs("sw_fuse") %></td>
			<td align="center"><%= formatnumber(rs("power_factor"),2) %></td>
			<td align="center"><%= formatnumber(rs("safety_factor"),2) %></td>
			<td align="center"><%= rs("maxm") %></td>
			<td align="center"><%= formatnumber(rs("area"),0) %></td>
			<td align="center"><%= formatnumber(rs("amps_available"),2) %></td>
			<td align="center"><%= formatnumber(rs("wsqft"),2) %></td>
			</tr>

			<%
			lineCount = lineCount + 1
  			rs.movenext
			if (not (rs.eof)) then
				thisFloor = rs("flr")
			end if
		loop 'end of inner loop that steps through all risers on a floor
		%>
		<!-- end of the main table, this row sums up total power -->
		<tr>
			<td colspan=7></td>
			<td colspan=7 bgcolor="#6699cc" height=2></td>
		</tr>
		<tr>
			<td width colspan=7></td>
			<td width colspan=5 align="RIGHT"><strong>Total</strong></td>
			<td width align = "center"><%= formatnumber(totalAmpsAvailable,1) %></td>
			<td width align = "center"><%= formatnumber(totalPowerAvailable,1) %></td>
		</tr>
	</table>
	<br>
	<br>
	<br>
	<%
		
		
	loop 'end out outer loop that steps through all floors
end if
if bldgnum = "10016thAVE" OR bldgnum = "044SW" then
%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="4%">
			</td>
			<td width="96%">
			<strong><font size="-2" face="Arial, Helvetica, sans-serif">Power availability calculations 
			are based on power distribution system information gathered by Genergy or made available to Genergy by the property’s management.  
			Genergy normally applies a 15-20% safety factor reduction when calculating power availability figures.  
			For more information on the use of safety factors or to modify the safety factor calculation criteria 
			we encourage you to contact our office. </font></strong>
			</td>
		</tr>
		<tr>
			<td width="4%">
			</td>
			<td width="96%">
			<strong><font size="-2" face="Arial, Helvetica, sans-serif">* Power Factor and Safety Factor have been changed to Unity
			at customers request. </font></strong>
			</td>
		</tr>
	</table>
<%
else
%>
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="4%">
			</td>
			<td width="96%">
			<strong><font size="-2" face="Arial, Helvetica, sans-serif">Power availability calculations 
			are based on power distribution system information gathered by Genergy or made available to Genergy by the property’s management.  
			Genergy normally applies a 15-20% safety factor reduction when calculating power availability figures.  
			For more information on the use of safety factors or to modify the safety factor calculation criteria 
			we encourage you to contact our office. </font></strong>
			</td>
		</tr>
	</table>
<%
end if
%>
<br>

</body>
</html>
