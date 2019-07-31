<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'2/25/2008 N.Ambo amended sql statements for printing "TID" and "METERID" type tickets

dim cnn,rs,rstNote, sql, GWK
set cnn = server.createobject("ADODB.Connection")
set rs = server.createobject("ADODB.Recordset")
set rstNote = server.createObject("ADODB.Recordset")

' open connection
cnn.open getConnect(0,0,"dbCore")'application("cnnstr_itservices")
dim ticketId
ticketId = request("ticketID")
if isempty(ticketID) then
	response.write("invalid ticket id: " & ticketId)
	response.End
end if

sql = "select * from tickets where id = '" & ticketId & "'"

rs.open sql, cnn,1
if rs.EOF then
	response.write("invalid ticket id: " & ticketId)
	response.End
end if
%>
<head>
<title>Ticket Report: <%=ticketID%></title>
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" onload="window.print()">
<center>
<table border=0 cellpadding="3" cellspacing="1" width="80%" bgcolor="#dddddd"><%
	sql = "select * from ttnotes where ticketid =" & rs("id") &" order by date"
	rstNote.open sql, cnn,1
	Dim notecount 
	notecount = rstNote.recordcount%>
	<tr bgcolor="#ffffff" valign="top"> 
		<td width="2%" align="left" nowrap style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;border-top:1px solid #cccccc;">
			<b>Ticket # <%=rs("id")%>, Opened by <%=rs("Requester")%> on <%=FormatDateTime(trim(rs("date")),2)%>.
			Requested for <%=FormatDateTime(rs("duedate"),2)%></b>
		</td>
		<td width="2%" align="left" style="border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;border-top:1px solid #cccccc;">
			<div align="right"><b><%=rs("userid")%></b></div>
		</td>
	</tr>
  <%if rs("ticketfortype")<>"" then%>
    <tr bgcolor="#ffffff" valign="top"> 
    <td colspan="2" align="left">
	<%
  				dim rstTicketFor
				set rstTicketFor = server.CreateObject("adodb.recordset")
				select case rs("ticketfortype")	
					case "PID"
						sql = "select name from portfolio where id = '" & rs("ticketfor") & "'"
						rstTicketFor.open sql, getconnect(0,0,"dbcore")'application("cnnstr_SuperMod")
						%>Portfolio: <u><%
						if not rstTicketFor.eof then
							response.write (rstTicketFor("name") & " (PID: ")
						end if
						rstTicketFor.close
						%><%=rs("ticketfor")%>)</u><%
						
					case "BLDGNUM"
						dim bldgSql
						bldgSql = "select portfolioid, strt from buildings where bldgNum = '" & rs("ticketfor") & "'"
						rstTicketFor.open bldgSql, getLocalConnect(rs("ticketfor"))
						if not rstTicketFor.eof then
							%>Building: <u><%=rstTicketFor("strt")%></u>
						<%end if
						
					case "TID"
						dim tempPid, tempBldg, tenantName, sqlTid
						sqlTid = "select billingname, bldgnum from tblLeases tl where tl.billingId = '" & split(rs("ticketfor"),"-")(1) & "'"
						rstTicketFor.open sqlTid, getConnect(0,rs("ticketfor"),"Billing")'application("cnnstr_SuperMod")
						if not rstTicketFor.eof then
							tempBldg = rstTicketFor("bldgnum")
							tenantName = rstTicketFor("billingname")
							rstTicketFor.close
							rstTicketFor.open "select portfolioid from buildings where bldgnum = '" & tempBldg & "'", getLocalConnect(tempBldg)
							if not rstTicketFor.eof then
								tempPid = rstTicketFor("portfolioid")
							end if
							rstTicketFor.close
						end if
						%>
						Tenant: <u><%=tenantName%> (BillingID <%=rs("ticketfor")%>)</u><%
					case "METERID"
						dim meterNum, tempLid, sqlMeter, tempTid
						sqlMeter = "select meternum, bldgnum, leaseutilityid as lid from Meters m where m.meterID = '" & split(rs("ticketfor"),"-")(1) & "'"
						rstTicketFor.open sqlMeter,getConnect(0,rs("ticketfor"),"billing")'application("cnnstr_SuperMod")
						if not rstTicketFor.eof then
							tempLid = rstTicketFor("lid")
							meterNum = rstTicketFor("meterNum")
							tempBldg = rstTicketFor("bldgnum")
							rstTicketFor.close
							sqlMeter = "select tl.billingid as tid, b.portfolioid as pid from tblleasesutilityprices tlup inner join tblleases tl on tl.billingid = tlup.billingid inner join buildings b on tl.bldgnum = b.bldgnum where tlup.leaseutilityid = '" & tempLid & "'"
							rstTicketFor.open sqlMeter,  getLocalConnect(tempBldg)
							if not rstTicketFor.eof then
								tempTid = rstTicketFor("tid")
								tempPid = rstTicketFor("pid")
							end if
						end if
						%>
						Meter: <a href="javascript:viewMeter('<%=tempPid%>', '<%=tempBldg%>', '<%=tempTid%>', '<%=tempLID%>', '<%=rs("ticketfor")%>')">
							<%=meterNum%> (MeterID <%=rs("ticketfor")%>)
						</a>
						<%
				end select
				%>
							&nbsp;<%if not(isnull(rs("billyear"))) then%>bill year <%=rs("billyear")%><%end if%>
							<%if not(isnull(rs("billperiod"))) then%> period <%=rs("billperiod")%><%end if%>
 		</td>
  </tr>
  <%end if%>
	<tr bgcolor="#ffffff" valign="top"> 
		<td colspan="2" align="left" style="border-right:1px solid #cccccc;border-left:1px solid #cccccc;"><b>Ticket Details:</b><%=rs("initial_trouble")%></td>
	</tr>
</table>
<br><Br>
<table width="80%" border=0 cellpadding="3" cellspacing="0" bordercolor="#cccccc">
	<tr> 
		<td align="left" nowrap style="border-top:1px solid #cccccc;border-left:1px solid #cccccc;"><b>Notes:</b></td>
		<td style="border-top:1px solid #cccccc;">&nbsp;</td>
		<td style="border-top:1px solid #cccccc;">&nbsp;</td>
		<td style="border-top:1px solid #cccccc;border-right:1px solid #cccccc;">&nbsp;</td>
	</tr>
	
	<%if not rstNote.eof then%>
		<tr> 
			<td width="15%" style="border-left:1px solid #cccccc;">&nbsp;</td>
			<td width="15%" align="left" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;">Date</td>
			<td width="5%" align="left" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;">Hours</td>
			<td width="70%" style="border-right:1px solid #cccccc;border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;">Note</td>
		</tr>
		<%while not rstNote.eof%>
			<tr valign="top"> 
				<td width="15%" style="border-left:1px solid #cccccc;">&nbsp;</td>
				<td width="15%" nowrap style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"><%=rstNote("date")%>, 
				<%=rstNote("uid")%></td>
				<td width="5%" nowrap style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"><%=rstNote("time")%></td>
				<td width="70%" style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;border-right:1px solid #cccccc;"><%=rstNote("note")%></td>
			</tr>
			<%rstNote.movenext
		wend
	else%>
		<td colspan="4" align="center" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;">
			NO TROUBLE NOTES FOUND FOR THIS TICKET
		</td>
	<%end if 
	
	rstNote.close%>
	<tr>
		<td colspan="4" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;">&nbsp;</td>
	</tr>
</table> 
<br>
</center>
</body>
</html>
<%rs.close%>