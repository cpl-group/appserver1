<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		dim cnn, rs,rst2, sql,totaltickets,L24,LWK,GWK,RT, listlength, currentprocess, currentstatus, userlist
		
		set cnn = server.createobject("ADODB.Connection")
		set rs = server.createobject("ADODB.Recordset")
		set rst2 = server.createobject("ADODB.Recordset")
		' open connection
		cnn.open getConnect(0,0,"dbCore")
	
	if request("status") = 0 then 
		currentstatus = "Opened"
		listlength = "1/1/2002"
	else
		currentstatus = "Closed"
		if request("listlength") = "" then
			listlength = date() - 7
		else 
			if ucase(request("listlength")) = "YTD" then
				listlength = "1/1/" & cstr(year(date()))
			else
				listlength = date() - cint(request("listlength"))
			end if
		end if
	end if

		userlist = request("userlist") 

		sql = "select * from tickets where closed = " & request("status") & " and [date] >= '" & listlength & "' and userid in ("&userlist&") order by userid, date"
		
		rs.open sql, cnn,1
		totaltickets = rs.recordcount
		L24	=0
		LWK	=0
		GWK	=0
		RT	=0
		if not rs.EOF then 
			%>
			<title><%=currentstatus%> Tickets Report for <%=replace(userlist,"'","")%></title>
	<link rel="Stylesheet" href="../../styles.css" type="text/css">
			<style type="text/css">
			<!--
			BODY {
			SCROLLBAR-FACE-COLOR: #dddddd;
			SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;
			SCROLLBAR-SHADOW-COLOR: #eeeeee;
			SCROLLBAR-3DLIGHT-COLOR: #999999;
			SCROLLBAR-ARROW-COLOR: #000000;
			SCROLLBAR-TRACK-COLOR: #336699;
			SCROLLBAR-DARKSHADOW-COLOR: #333333;
			}
			
			td.red {color: red}
			-->
			</style>
			</head>
<script>
function showall(){
	var func = eval('document.all.allnotes')
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
	var displaytype = (func.innerHTML != '[-]' ? 'none':'block');
	var tag = document.all//('note162');
	for (i = 0; i < tag.length; i++){
		if (tag[i].name == 'noteset') tag[i].style.display = displaytype
		if (tag[i].name == 'notefunc') tag[i].innerHTML = func.innerHTML
	} 
}
function note(id){
	var tag = document.getElementById('note'+id) 
	tag.style.display = (tag.style.display == "block" ? "none" : "block");
	var func = eval('document.all.func'+id)
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
}
</script>			
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" onload="window.print()">
<div width:100%> 
    <% 
		  	Dim tcolor, age	
			while not rs.EOF 
			if request("status") = 0 then 	  
			age = date() - cdate(trim(rs("duedate")))
			else 
			age = cdate(trim(rs("fixdate"))) - cdate(trim(rs("date")))	
			end if 
		  if rs("runticket") then 
    	  	   tcolor = "#cccccc"
			   currentprocess = "rt"
			   RT = RT + 1	
		  else
		   if age > 7 then 
			   tcolor = "#cc0033"
			   GWK = GWK + 1
			   currentprocess = "gwk"  
		   else 
		   		if age < 2 then 
				tcolor = "#66ff66"
				L24 = L24 + 1
 			    currentprocess = "l24"
				else
				tcolor = "#FFcc00"
				LWK = LWK + 1
			   	currentprocess = "lwk"
				end if
			end if
		   end if
		if request("showage") = "" or request("showage") = currentprocess then
 
%>  
<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
  <%
			  sql = "select * from ttnotes where ticketid =" & rs("id") &" order by date"
			  rst2.open sql, cnn,1
			  Dim notecount 
			  notecount = rst2.recordcount
%>
  <tr bgcolor="#ffffff" valign="top"> 
    <td width="2%" align="left" nowrap><strong>Ticket # <%=rs("id")%>, Opened 
      by <%=rs("Requester")%> on <%=FormatDateTime(trim(rs("date")),2)%>. Requested 
      for <%=FormatDateTime(rs("duedate"),2)%></strong></td>
    <td width="2%" align="left" nowrap><div align="right"><strong><%=rs("userid")%></strong></div></td>
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
						rstTicketFor.open sql, getConnect(rs("ticketfor"),0,"Billing")
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
						rstTicketFor.open sqlTid, getConnect(rs("ticketfor"),0,"Billing")
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
						rstTicketFor.open sqlMeter, getConnect(rs("ticketfor"),0,"Billing")
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
    <td colspan="2" align="left"><strong>Ticket Details:</strong><%=rs("initial_trouble")%></td>
  </tr>
</table>

<div style="border-top:1px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000;width:100%"> 
  <table  width="100%" border=0 cellpadding="3" cellspacing="0">
    <tr > 
      <td align="left" nowrap ><strong>Notes:</strong></td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
    <%
  			  if not rst2.eof then 
  %>
    <tr > 
      <td width="15%" >&nbsp;</td>
      <td width="15%" align="left">Date</td>
		<td width="5%">Time</td>
      <td width="70%" >Note</td>
    </tr>
    <%
			  while not rst2.eof
			  %>
    <tr valign="top"> 
      <td width="15%" >&nbsp;</td>
      <td width="15%" nowrap style="border-bottom:1px solid #000000;"><%=rst2("date")%>, 
        <%=rst2("uid")%></td>
		<td width="5%" style="border-bottom:1px solid #000000;"><%=rst2("time")%></td>
      <td width="70%" style="border-bottom:1px solid #000000;"><%=rst2("note")%></td>
    </tr>
    <% 
			  rst2.movenext
			 wend
			  else %>
    <td colspan="5" align="center">NO TROUBLE NOTES FOUND FOR THIS TICKET</td>
    <% end if 
			 rst2.close%>
    <tr valign="top">
      <td colspan=4><strong>Comments:</strong></td>
    </tr>
    <tr valign="top">
      <td colspan=4>&nbsp;</td>
    </tr>
    <tr valign="top">
      <td colspan=4>&nbsp;</td>
    </tr>
    <tr valign="top">
      <td colspan=4>&nbsp;</td>
    </tr>

  </table> 
</div>
<br>
<% 
		end if
		  rs.movenext
		  wend
		  %>
</body>
</div>
</html>

<%
			rs.close
		else
		response.write "<font size=1 face='Arial, Helvetica, sans-serif'> No tickets found</font>"
		end if 
%>