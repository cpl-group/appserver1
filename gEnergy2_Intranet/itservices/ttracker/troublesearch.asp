<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'2/19/2008 N.Ambo made modifications to design of search screen; added additional search options to limit results to the user the task was assigned to, requested by, or copied to

dim searchbox, searchstring, status, searchtime,pid,pidIP
searchtime = now
if lcase(request("searchbox"))="true" then searchbox=true else searchbox=false
searchstring = secureRequest("searchstring")
if request("status")="1" then status = 1 else status = 0
dim cnn, rs,rst2, sql, orderby, totaltickets, totalticketsShowing,L24,LWK,GWK,RT, listlength, currentprocess, currentstatus, userlist, typex
set cnn = server.createobject("ADODB.Connection")
set rs = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")

typex = trim(request("searchtype"))

' open connection
cnn.open getConnect(0,0,"dbCore")
cnn.CommandTimeout = 60*5
'response.write request("ticketfor")
'response.end
pid = secureRequest("pid")
pidIP = getPidIPMain(pid)
pidIP = "[" & pidIP & "].dbBilling.dbo."

if status = 0 then 
	currentstatus = "Opened"
	listlength = "1/1/2002"
else
	currentstatus = "Closed"
	if request("listlength") = "" then
		listlength = date() - 7
	else 
		if request("listlength") = "YTD" then
			listlength = "1/1/" & cstr(year(date()))
		else
			listlength = date() - cint(request("listlength"))
		end if
	end if
end if

sql = "SELECT top 800 t.id as tid, t.ticketfortype, ticketfor, t.[date] as [tdate], duedate, t.fixdate, runticket, client, t.userid as tuserid, Requester, initial_trouble, tn.note as tnnote, tn.date as tndate, tn.uid as tnuid, tn.time as tntime, isnull(tn.id,-1) as hasnote, closed as status, master FROM tickets t LEFT JOIN ttnotes tn ON t.id=tn.ticketid "

if searchstring<>"" then 'add search clause 
	if request("buildings") = "True" then sql = sql & " LEFT JOIN "&pidIP&"buildings ipb ON ticketfor=ipb.bldgnum"
	if request("accounts") = "True" then sql = sql & " LEFT JOIN "&pidIP&"tblleases ten ON ticketfor=cast(ten.billingid as varchar)"
	if request("meters") = "True" then sql = sql & " LEFT JOIN "&pidIP&"meters m ON ticketfor=cast(m.meterid as varchar)"
	sql = sql & " WHERE t.[date] >= '" & listlength & "' "
	sql = sql & " and (jobnum like '%" & searchstring & "%' or t.id like '%" & searchstring & "%' or ticketfor like '%" & searchstring & "%' or userid like '%" & searchstring & "%' or requester like '%" & searchstring & "%' or ccuid like '%" & searchstring & "%' or initial_trouble like '%" & searchstring & "%' or t.id in (select ticketid from ttnotes where note like '%" & searchstring & "%') or department in (select depid from departments where department like '%" & searchstring & "%') "
	if request("portfolios") = "True" then sql = sql & " or (ticketfor in (select cast(id as varchar) from "&pidIP&"portfolio where [name] like '%" & searchstring & "%') AND ticketfortype = 'PID') "
	if request("buildings") = "True" then sql = sql & " or ((ipb.strt like '%" & searchstring & "%' or ipb.bldgname like '%" & searchstring & "%' or ipb.btbldgname like '%" & searchstring & "%') AND ticketfortype='BLDGNUM') "
	if request("accounts") = "True" then sql = sql & " or ((ten.billingname like '%" & searchstring & "%' or ten.tname like '%" & searchstring & "%' or ten.tenantnum like '%" & searchstring & "%') AND ticketfortype='TID') "
	if request("meters") = "True" then sql = sql & " or ((m.meternum like '%" & searchstring & "%' or m.location like '%" & searchstring & "%') and ticketfortype='METERID') "
	sql = sql & " ) and (1<>1"
	if request("internalops") = "True" then sql = sql & " OR ticketfortype = '' OR ticketfortype='joblog' OR ticketfortype='vendorid'"
	if request("ticketfortype")="" then 'it is not coming directly from a UM section
		if request("portfolios") = "True" then sql = sql & " OR ticketfortype = 'PID'"
		if request("buildings") = "True" then sql = sql & " OR ticketfortype = 'BLDGNUM'"
		if request("accounts") = "True" then sql = sql & " OR ticketfortype = 'TID'"
		if request("meters") = "True" then sql = sql & " OR ticketfortype = 'METERID'"
	end if
	if request("ticketfortype") <> "" and request("ticketfor") <> "" then sql = sql & " OR (ticketfortype = '"&request("ticketfortype")&"' and ticketfor = '"&request("ticketfor")&"')"
	if isnumeric(searchstring) and (request("ticketfortype")="" or request("ticketfortype")="joblog") then sql = sql & " or (jobnum = '" & searchstring & "') "
	sql = sql & " ) "
else 
	sql = sql & " WHERE t.[date] >= '" & listlength & "' "
	if isnumeric(searchstring) then sql = sql & " or (jobnum = '" & searchstring & "') "
end if

if lcase(request("allclosed")) = "true" then
	'sql = sql & "and (closed = 0 or closed = 1) "
else
	sql = sql & " and closed = " & status & " "
end if


if (typex="Userid" or typex="Requester" or typex ="ccuid") and searchstring <> "" then
		sql = sql & " and " &typex & " like '%" & searchstring & "%'"
	end if 
	
'response.write sql
'response.end
orderby = "order by client desc, t.date , t.id, tn.date desc"
sql = sql & orderby

if not(searchbox) or searchstring<>"" then rs.open sql, cnn
totaltickets = 0
totalticketsShowing = 0
L24	=0
LWK	=0
GWK	=0
RT	=0
%>
<title>Trouble Ticket Search</title>
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
		if (tag[i].name == 'noteset') tag[i].style.display = displaytype;
		if (tag[i].name == 'notefunc') tag[i].innerHTML = func.innerHTML;
	} 
}
function note(id){
	var tag = document.getElementById('note'+id) 
	tag.style.display = (tag.style.display == "block" ? "none" : "block");
	var func = eval('document.all.func'+id)
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
}

function viewPortfolio(pid){
	window.open('/genergy2/setup/portfolioedit.asp?pid=' + pid, 'ViewPortfolio','width=900,height=700,resizable=yes,scrollbars=yes')
}

function viewTenant(pid, bldg, tid){
	window.open('/genergy2/setup/tenantedit.asp?pid=' +pid + '&bldg=' +bldg + '&tid=' +tid, 'ViewTenant','width=900,height=700,resizable=yes,scrollbars=yes')
}

function viewMeter(pid, bldg, tid, lid, mid){
	window.open('/genergy2/setup/contentfrm.asp?action=meteredit&pid=' + pid+ '&bldg=' + bldg+ '&tid=' + tid+ '&lid=' + lid+ '&meterid=' + mid, 'ViewMeter','width=900,height=700,resizable=yes,scrollbars=yes')
}

function viewBuilding(bldg, pid){
	window.open('/genergy2/setup/buildingedit.asp?pid=' + pid + "&bldg=" + bldg, 'ViewBuilding','width=900,height=700,resizable=yes,scrollbars=yes')
}
</script>			
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#eeeeee">
<%if searchbox then%>
	<form name="form2" method="get" action="troublesearch.asp">
		<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
			<tr bgcolor="#eeeeee">	
				<td style="border-right:1px solid #dddddd">
					<input type="text" name="searchstring" value="<%=searchstring%>">
				<select name="searchtype">
				    <option value="All" <%if typex ="All" then %> selected <%end if%>>All</option>
				    <option value="Userid" <%if typex ="Userid" then %> selected <%end if%>>Assigned To</option>
				    <option value="Requester" <%if typex ="Requester" then %> selected <%end if%>>Requested By</option>
				    <option value="ccuid" <%if typex ="ccuid" then %> selected <%end if%>>Copied To</option>
			    </select>
					<input type="submit" name="action" value="Search" onclick="javascript:document.all.searching.innerHTML = '<font size=2>Searching...Please Wait</font>'">
				</td>	
				<td style="border-right:1px solid #dddddd">
					<input type="checkbox" name="allclosed"
						<%if request("allclosed") = "True" then%>checked<%end if%> value="True">Include closed tickets				
				</td>
				<td>
					Include tickets opened for: 
				</td>
				<td>
					<input type="checkbox" name="internalops"
						<%if (isempty(request("action")) and searchstring="") or request("internalops") = "True" then %>checked<%end if%> value="True">Internal Operations
					<input type="checkbox" name="portfolios"
						<%if request("portfolios") = "True" then %>checked<%end if%> value="True">Portfolios
					<input type="checkbox" name="buildings"	
						<%if request("buildings") = "True" then %>checked<%end if%> value="True">Buildings
					<input type="checkbox" name="accounts"
						<%if request("accounts") = "True" then %>checked<%end if%> value="True">Accounts
					<input type="checkbox" name="meters"
						<%if request("meters") = "True" then %>checked<%end if%> value="True">Meters
				</td>
			</tr>
		<input type="hidden" name="status" value="<%=status%>">
		<input type="hidden" name="searchbox" value="<%=searchbox%>">
	</form>
	</table>
<%end if%>
<%if rs.state = 1 then%>
	<%if not rs.EOF then%>
		<table border=0 cellpadding="3" cellspacing="1" width="100%">
			<tr bgcolor="#6699cc"> 
				<td width="2%" align="center" valign="middle" nowrap><span id="allnotes" style="cursor:hand;text-decoration:none;" onclick="javascript:showall()">[+]</span></td>
				<td width="3%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td bgcolor="#66ff66" title="Tickets Open For Less Then 2 Days">&nbsp;</td>
				<td bgcolor="#FFcc00" title="Tickets Open For Less Then 1 Week">&nbsp;</td>
				<td bgcolor="#cc0033" title="Tickets Open For More Then 1 Week">&nbsp;</td>
			</tr>
		</table>
		</td>
		<td width="4%" align="center" class="standardheader">ID</td>
		<td width="9%" align="center" class="standardheader">Assigned To</td>
		<td width="10%" align="center" class="standardheader">Date Opened</td>
		<td width="10%" align="center" class="standardheader">Requester</td>
		<td width="39%" align="center" class="standardheader">Description</td>
		<%if status = 1 or lcase(request("allclosed")) = "true" then %>
			<td width="10%" align="center" class="standardheader">Closed/Due</td>
		<%else%>
			<td width="10%" align="center" class="standardheader">Due Date</td>
		<%end if%>
		<td width="18%" align="center" class="standardheader">Opened For</td>
		<td width="15" align="center" class="standardheader">&nbsp;&nbsp;</td>
		</tr>
		</table>
	<%end if%>
	<div style="width:100%; backgroundColor:white; overflow:auto; height:70%;border-bottom:1px solid #000000;" id="searching">
	<%if not rs.EOF then%>
		<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
		<% 
		Dim tcolor, age, tid
		while not rs.EOF
			if tid <> rs("tid") then 
				totaltickets = totaltickets+1
				if rs("status") = 0 then 	  
					age = date() - cdate(trim(rs("duedate")))
				else 
					if isnull(rs("fixdate")) then 
					age = cdate(trim(rs("tdate")))
					else
					age = cdate(trim(rs("fixdate"))) - cdate(trim(rs("tdate")))	
					end if
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
			end if
			tid = rs("tid")
			if request("showage") = "" or request("showage") = lcase(currentprocess) then
				totalticketsShowing=totalticketsShowing+1%>
				<tr bgcolor="#ffffff" valign="top"> 
				<td width="2%" align="center" valign="middle" nowrap>
				<%if rs("hasnote") <> -1 then%>
					<span id="func<%=rs("tid")%>" name = "notefunc" style="cursor:hand;text-decoration:none;" onclick="note('<%=rs("tid")%>')">[+]</span>
				<%else%>
					<span id="func<%=rs("tid")%>" name = "empty">[o]</span>
				<%end if%>
				</td>
				<td bgcolor="<%=tcolor%>" width="3%" align="center" valign="middle"><%IF rs("client") then %> <img src="../images/critical.gif" width="20" height="20" align="center"><%end if%></td>
				<td width="4%" align="center"  onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:document.location='ticket.asp?pid=<%=pid%>&mode=update&tid=<%=tid%>&ticketfor=<%=rs("ticketfor")%>&ticketfortype=<%=rs("ticketfortype")%>';try{parent.document.all.Function.innerHTML = 'Viewing Ticket <%=tid%>'}catch(exception){};" ><%=tid%></td>
				<td width="9%"><a href="userreports.asp?status=<%=status%>&userlist='<%=rs("tuserid")%>'&listlength=<%=request("listlength")%>" target="_blank"><%=rs("tuserid")%></a></td>
				<td width="10%"><%=FormatDateTime(trim(rs("tdate")),2)%></td>
				<td width="10%"><%=rs("Requester")%></td>
				<td width="39%"><%=left(rs("initial_trouble"),50)%><%if Len(trim(rs("initial_trouble"))) > 50 then%>...<%end if%></td>
				<%if rs("status") then %>
					<td width="10%" align="right">Closed <%=rs("fixdate")%></td>
				<%elseif rs("master") then %>
					<td width="10%" align="right">N/A</td>
				<%else%>
					<td width="10%" align="right">Due <%=FormatDateTime(rs("duedate"),2)%></td>
				<%end if %>
				<td width="18%" align="right">
					<%
					dim cnnMainModule, rstTicketFor
					set rstTicketFor = server.CreateObject("adodb.recordset")
					'set cnnMainModule = server.createObject("adodb.connection")
					'cnnMainModule.open getConnect(0,bldgnum,"dbBilling")
					select case ucase(rs("ticketfortype"))
					case "PID"
						%>Portfolio <%=rs("ticketfor")%><%
					case "BLDGNUM"
						%>Building <%=rs("ticketfor")%><%
					case "TID"
						%>Tenant <%=rs("ticketfor")%><%
					case "METERID"
						%>Meter <%=rs("ticketfor")%><%
					case else
						if rs("ticketfortype")<>"" then
							%><%=rs("ticketfortype")%>&nbsp;<%=rs("ticketfor")%><%
						else
							%>Internal Operations<%
						end if
					end select
					%>
					</td>
					</tr>
					<tr id="note<%=tid%>" name="noteset" style="height:100;border-bottom:1px solid #cccccc;display:none;background-color :#ffffcc;"><td colspan="9">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
						<td width="15%" >Date</td>
						<td width="70%" >Note</td>
						<td width="10%" >uid</td>
						<td width="5%">Time</td>
						<td></td>
						</tr>
						</table>
					<div style="width:100%; overflow:auto; height:95;border-bottom:1px solid #cccccc;background-color :#ffffcc;">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
					<%if rs("hasnote") <> -1 then 
						do until rs.eof%>
							<tr valign="top"> 
							<td width="15%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rs("tndate")%></td>
							<td width="70%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=replace(rs("tnnote"),vbcrlf,"<br>")%></td>
							<td width="10%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rs("tnuid")%></td>
							<td width="5%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rs("tntime")%></td>
							</tr>
							<%
							rs.movenext
							if not rs.eof then if tid <> rs("tid") then exit do
						loop
					else
						%><td colspan="3" align="center">NO TROUBLE NOTES FOUND FOR THIS TICKET</td><%
						rs.movenext
					end if
					%>
					</table> 		 
					</div>
					</td></tr>
<!-- 					</table> -->
				<%
				else
					rs.movenext
				end if
			wend
		else
			%><font size=1 face='Arial, Helvetica, sans-serif'> No tickets found</font><%
		end if%>
	</table>
		</div>
		<table border="0" align="center" cellpadding="3" cellspacing="2" style="border:1px solid #dddddd;margin:3px;">
		<%	if status = 1 then %>
		<tr > 
		<td colspan="8"><p align="center"><a href="./troublesearch.asp?status=<%=status&"&dept="&request("dept") & "&listlength=7"%>" onclick="javascript:parent.document.all.Function.innerHTML = '<%=currentstatus%> Tickets : Last 7 Days'">last 
		7 days</a> | <a href="./troublesearch.asp?status=<%=status&"&dept="&request("dept") & "&listlength=30"%>" onclick="javascript:parent.document.all.Function.innerHTML = '<%=currentstatus%> Tickets : Last 30 Days'">last 
		30 days</a> | <a href="./troublesearch.asp?status=<%=status&"&dept="&request("dept") & "&listlength=60"%>" onclick="javascript:parent.document.all.Function.innerHTML = '<%=currentstatus%> Tickets : Last 60 Days'">last 
		60 days</a> | <a href="./troublesearch.asp?status=<%=status&"&dept="&request("dept") & "&listlength=120"%>" onclick="javascript:parent.document.all.Function.innerHTML = '<%=currentstatus%> Tickets : Last 120 Days'">last 
		120 days</a> | <a href="./troublesearch.asp?status=<%=status&"&dept="&request("dept") & "&listlength=YTD"%>" onclick="javascript:parent.document.all.Function.innerHTML = '<%=currentstatus%> Tickets : Year To Date'">Year 
		to Date</a></p></td>
		</tr>
	<%end if %>
	<tr>
	<%
	dim radiolink
	radiolink = "troublesearch.asp?searchbox="&searchbox&"&allclosed="&request("allclosed")&"&searchstring="&searchstring&"&ticketfortype="&request("ticketfortype")&"&ticketfor="&request("ticketfor")&"&status="&status&"&dept="&request("dept")&"&listlength="&request("listlength")&"&internalops="&request("internalops")&"&portfolios="&request("portfolios")&"&buildings="&request("buildings")&"&accounts="&request("accounts")&"&meters="&request("meters")&"&showage="
	%>
	<td><div style="position:inline;width:18px;height:12px;background:#66ff66;border:1px solid #999999;"> 
	<input name="age" type="radio" value="l24" onclick="javascript:document.location='<%=radiolink%>'+this.value" <%if request("showage")="l24" then%> checked <%end if%>>
	</div></td>
	<td><a href="<%=radiolink%>l24">Less than 24 hours <b>( <%if totaltickets>0 then%><%=formatpercent(l24/totaltickets)%><%else%><%=formatpercent(0)%><%end if%> )</b></a></td>
	<td><div style="position:inline;width:18px;height:12px;background:#ffcc00;border:1px solid #999999;"> 
	<input name="age" type="radio" value="lwk" onclick="javascript:document.location='<%=radiolink%>'+this.value" <%if request("showage")="lwk" then%> checked <%end if%>>
	</div></td>
	<td><a href="<%=radiolink%>lwk">Less than 1 week <b>( <%if totaltickets>0 then%><%=Formatpercent(LWK/totaltickets)%><%else%><%=formatpercent(0)%><%end if%> )</b></a></td>
	<td><div style="position:inline;width:18px;height:12px;background:#cc0033;border:1px solid #999999;"> 
	<input name="age" type="radio" value="gwk" onclick="javascript:document.location='<%=radiolink%>'+this.value" <%if request("showage")="gwk" then%> checked <%end if%>>
	</div></td>
	<td><a href="<%=radiolink%>gwk">Greater than 1 week <b>( <%if totaltickets>0 then%><%=Formatpercent(GWK/totaltickets)%><%else%><%=formatpercent(0)%><%end if%> )</b></a></td>
	<td><div style="position:inline;width:18px;height:12px;background:#cccccc;border:1px solid #999999;"> 
	<input name="age" type="radio" value="rt" onclick="javascript:document.location='<%=radiolink%>'+this.value" <%if request("showage")="rt" then%> checked <%end if%>>
	</div></td>
	<td><a href="<%=radiolink%>rt">Running Tickets <b>( <%if totaltickets>0 then%><%=Formatpercent(RT/totaltickets)%><%else%><%=formatpercent(0)%><%end if%> )</b></a></td>
	<%if request("showage")<>"" then%>
	<td><div style="position:inline;width:18px;height:12px;background:#cccccc;border:1px solid #999999;"> 
	<input name="age" type="radio" value="" onclick="javascript:document.location='<%=radiolink%>'">
	</div></td>
	<td><a href="<%=radiolink%>">Clear</a></td>
	<%end if%>
	</tr>
	<tr> 
	<td><img src="../images/critical.gif" width="20" height="20" align="center"></td>
	<td>= Client Related Ticket</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	</table>
	<%
	rs.close
else
	%><div id="searching"></div><%
end if 'beginning--> if rs.state = 1 then (open)
searchtime = datediff("s",searchtime,now)
%>

<script> 
if (parent.document.all.count){
parent.document.all.count.innerHTML = '<%if totalticketsShowing<totaltickets then response.write "Showing "&totalticketsShowing&" out of "%><%=totaltickets%> Tickets Found </b><span style="font-size: 10;font-style:none;">(database processed in <%=searchtime%> second<%if searchtime>1 then%>s<%end if%>)</span><%if totalticketsShowing>=800 then%><br>search has reached max limit of 800<%end if%>'
}
</script>
<%'=totalticketsShowing%>
</body>
</html>

