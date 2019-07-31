<% 
'2/14/2008 N.AMBO AMENDED FRO USE OF NEW MASTER_NOTES TABLE
class tickets
	public label, note, requester, department, userid, ccuid, client 'these are the only variable that should need be changed
	private bperiod, byear
	private cnn, rst, sql, jobnum, masterticketid, ticketfor, ticketfortype, masternotecount, totaltickets, opentickets, criticalopentickets, pid, bldg
	'ccuid = users that get updated when ticket is updated
	Private Sub Class_Initialize()
		set cnn = server.createobject("ADODB.connection")
		set rst = server.createobject("ADODB.recordset")
		cnn.open getConnect(0,0,"dbCore")
		requester = "RBAdmin"
		department = "RB"
		userid = "RBADMIN"
		client = 0
		pid = request("pid")
		if request("bldg")<>"" then
			bldg = request("bldg")
		elseif request("bldgnum")<>"" then
			bldg = request("bldgnum")
		elseif request("building")<>"" then
			bldg = request("building")
		end if
		
	end sub
	
	function findtickets(ttype, id) 'Check if Master Ticket is Open for account
		ticketfortype=ttype
		ticketfor = id
				
		if ticketfortype="joblog" then jobnum = id
		note = note & id
		do
			sql = "select isnull(a.tcount,0) as tcount from (select count(*) as tcount from master_notes where notefortype = '"&ticketfortype&"' and notefor = '"&id&"') a" 
			rst.open sql, cnn
			if not rst.eof then 
				'masterticketid = rst("id")
				masternotecount = rst("tcount")
				rst.close
				exit do
			else
				'2/11/2008 N.AMBO REMOVED THIS SECTION BECAUSE MASTER TICKETS WILL NO LONGER BE USER FOR NOTES; WE HAVE SEGREGATED TE NOTES AND THE TICKETS
				'NOTES ARE NO LONGER ATTACHED TO A MASTER TICKET AND ARE STORED IN A SEPARATE TABLE CALLED MASTER_NOTES
				'No ticket found for account, initialize ticket in tickets table		
				'sql = "insert into tickets ("
				'if ticketfortype="joblog" then sql = sql & "jobnum,"
				'sql = sql & " initial_trouble, requester,department,userid, client,ccuid, runticket, ticketfor, ticketfortype, master) values ('"
				'if ticketfortype="joblog" then sql = sql &trim(ticketfor)&"','"
				'sql = sql & trim(note)&"','"&trim(requester)&"','"&trim(department)&"','"& trim(userid)& "','"&client&"','" & trim(ccuid) &"','1','" &trim(ticketfor)&"','" &trim(ticketfortype)& "','1')"
				'cnn.execute sql
				rst.close
			end if 
		loop until false
		dim byfilter
		byfilter = "(billyear <> '"&byear&"' and billperiod <> '"&bperiod&"')"
		sql = "select a.tixcount as opentickets, b.tixcount as criticalopentickets, c.tixcount as totaltickets from "&_
		"(select count(*) as tixcount from tickets where ((jobnum='"&id&"' and 'joblog'='"&ticketfortype&"') or (ticketfortype='"&ticketfortype&"' and ticketfor='"&id&"')) and closed = 0  and master <> 1) a, "&_
		"(select count(*) as tixcount from tickets where ((jobnum='"&id&"' and 'joblog'='"&ticketfortype&"') or (ticketfortype='"&ticketfortype&"' and ticketfor='"&id&"')) and closed = 0 and "&byfilter&" and master <> 1) b, "&_
		"(select count(*) as tixcount from tickets where ((jobnum='"&id&"' and 'joblog'='"&ticketfortype&"') or (ticketfortype='"&ticketfortype&"' and ticketfor='"&id&"')) and master <> 1) c"
		rst.open sql, cnn
		if not rst.eof then 
			opentickets			=	rst("opentickets")
			criticalopentickets = 	rst("criticalopentickets")
			totaltickets		= 	rst("totaltickets")
		end if 
		rst.close
	end function 

	sub Display(pid,notes, totals, critical)%>
		<table width="100" border="0" cellspacing="0" cellpadding="0"><tr> 
		<%if totals then%>
			<td width="20" align="center" style="background-color:#336699; color:White;">&nbsp;<span id="notecount"><%=masternotecount%></span>&nbsp;</td>
			<td  nowrap>&nbsp;<a href="#" onclick="javascript:window.open('/genergy2/notes/notes.asp?notefortype=<%=ticketfortype%>&hlabel=<%=Label%>&notefor=<%=ticketfor%>','NoteManager','width=800,height=400, scrollbars=no')" class="mgmtlink"><%=Label%> Notes</a>&nbsp;</td>
		<%end if%>
		<%if totals or critical then%>
			<td align="center" nowrap>&nbsp;Trouble Tickets:</td>
		<%end if%>
		<%if totals then%>
			<td width=20 bgcolor="#00FF00" align="center">&nbsp;<%=totaltickets%>&nbsp;</td>
			<td nowrap>&nbsp;<a href="#" onclick="window.open('/genergy2_intranet/itservices/ttracker/troublesearch.asp?pid=<%=pid%>&searchstring=<%=ticketfor%>&ticketfortype=<%=ticketfortype%>&ticketfor=<%=ticketfor%>&action=Search&searchbox=false&allclosed=true','SearchNotes','width=800,height=400, scrollbars=no')">Total</a>&nbsp;</td>
			<td width=20 bgcolor="#FFCC00" align="center">&nbsp;<%=opentickets%>&nbsp;</td>
			<td nowrap>&nbsp;<a href="#" onclick="window.open('/genergy2_intranet/itservices/ttracker/troublesearch.asp?pid=<%=pid%>&searchstring=<%=ticketfor%>&ticketfortype=<%=ticketfortype%>&ticketfor=<%=ticketfor%>&action=Search&searchbox=false&status=0','SearchNotes','width=800,height=400, scrollbars=no')">Open</a>&nbsp;</td>
		<%end if%>
		<%if critical then%>
			<td width=20 bgcolor="#FF0000" align="center"><font color="#FFFF00">&nbsp;<%=criticalopentickets%>&nbsp;</font></td>
			<td nowrap>&nbsp;Critical&nbsp;</td>
		<%end if%>
		</tr></table><%
	end sub
	public sub endscript()
		response.write "buttoninsert"
		response.end
	end sub
	public sub MakeButton()%>
		<input type="button" value="Open a trouble ticket" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="javascript:window.open('/genergy2_intranet/itservices/ttracker/ticket.asp?mode=new&<%=ticketfortype%>=<%=ticketfor%>&ticketfortype=<%=ticketfortype%>&bldg=<%=bldg%>&pid=<%=pid%>&ticketfor=<%=ticketfor%>&child=1','<%=replace(label," ","_")%>TroubleTicket','width=680,height=325')">
	<%end sub
end class%>
