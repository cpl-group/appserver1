<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<%if not(allowGroups("Genergy Users,clientOperations")) then
	%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if



dim action, searchstring
action = trim(request("action"))
searchstring = trim(request("searchstring"))
%>
<html>
<head>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<title>Utility Manager Search</title>
<script>
function openCustomWin(clink, cname, cspec){
	open(clink, cname, cspec)
}
</script>
</head>

<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="get" action="searchresult.asp">
	<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
		<tr bgcolor="#eeeeee">	
			<td style="border-right:1px solid #dddddd">
				<input type="text" name="searchstring" value="<%=searchstring%>">
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
					<%if isempty(request("action")) or request("internalops") = "True" then %>checked<%end if%> value="True">Internal Operations
				<input type="checkbox" name="portfolios"
					<%if isempty(request("action")) or request("portfolios") = "True" then %>checked<%end if%> value="True">Portfolios
				<input type="checkbox" name="buildings"	
					<%if isempty(request("action")) or request("buildings") = "True" then %>checked<%end if%> value="True">Buildings
				<input type="checkbox" name="accounts"
					<%if isempty(request("action")) or request("accounts") = "True" then %>checked<%end if%> value="True">Accounts
				<input type="checkbox" name="meters"
					<%if isempty(request("action")) or request("meters") = "True" then %>checked<%end if%> value="True">Meters
			</td>
		</tr>
	</table>
</form>
<div align="center" id="searching">
<%
if action = "Search" then

	dim cnnIT, rst, sql
	set cnnIT = server.createobject("ADODB.connection")
	set rst = server.createobject("ADODB.recordset")
	cnnIT.open application("cnnstr_itservices")
	
	sql = "SELECT *, tickets.id as ticketid FROM tickets WHERE (tickets.id like '%" & searchstring & "%' or ticketfor like '%" & searchstring & "%' or userid like '%" & searchstring 
	sql = sql & "%' or requester like '%" & searchstring & "%' or ccuid like '%" & searchstring & "%' or initial_trouble like '%" & searchstring
	sql = sql & "%' or tickets.id in (select ticketid from ttnotes where note like '%" & searchstring 
	sql = sql & "%') or department in (select depid from departments where department like '%" & searchstring & "%')"
	sql = sql & "or (ticketfor in (select id from [" & application("superip") & "].mainmodule.dbo.portfolio where [name] like '%" 
	sql = sql & searchstring & "%') AND ticketfortype = 'PID') or (ticketfor in (select bldgnum from " & makeIPUnion("buildings","") 
	sql = sql & " ipb where ipb.strt like '%" & searchstring & "%' or ipb.bldgname like '%" & searchstring & "%' or ipb.btbldgname like '%" & searchstring 
	sql = sql & "%') AND ticketfortype = 'BLDGNUM') or (ticketfor in (select meterid from " & makeIPUnion("meters","") & " m where m.meternum like '%"
	sql = sql & searchstring & "%' or m.location like '%" & searchstring & "%') and ticketfortype='METERID') or (ticketfor in (select billingid from " & makeIpUnion("tblleases","") & " ten where ten.billingname like '%" & searchstring & "%' or ten.tname like '%" & searchstring & "%' or ten.tenantnum like '%" & searchstring & "%') and ticketfortype='TID'))"
	
	if request("internalops") <> "True" then
		sql = sql & " AND ticketfortype <> ''"
	end if 
	if request("portfolios") <> "True" then
		sql = sql & " AND ticketfortype <> 'PID'"
	end if
	if request("buildings") <> "True" then
		sql = sql & " AND ticketfortype <> 'BLDGNUM'"
	end if
	if request("accounts") <> "True" then
		sql = sql & " AND ticketfortype <> 'TID'"
	end if
	if request("meters") <> "True" then
		sql = sql & " AND ticketfortype <> 'METERID'"
	end if
	if request("allclosed") <> "True" then
		sql = sql & " And closed <> 1"
	end if
	
	
	sql = sql & " order by client desc, date"
	response.write sql
	
	rst.open sql, cnnIT
	if not rst.eof then
		%>
		<script>
		function viewPortfolio(pid){
		//	if (top.frames[1]){
		//		//alert("blatt");
		//		top.frames[1].location = '/genergy2/setup/portfolioedit.asp?pid=' + pid
		//	} else{
		//		//alert("blort")
		//		window.opener.location = '/genergy2/setup/portfolioedit.asp?pid=' + pid
		//	}
			window.open('/genergy2/setup/portfolioedit.asp?pid=' + pid, 'ViewPortfolio','width=900,height=700,resizable=yes,scrollbars=yes')
		}
		
		function viewTenant(pid, bldg, tid){
		//	if (top.frames[1]){
		//		top.frames[1].location = '/genergy2/setup/tenantedit.asp?pid=' +pid + '&bldg=' +bldg + '&tid=' +tid;
		//	} else {
		//		window.opener.location = '/genergy2/setup/tenantedit.asp?pid=' +pid + '&bldg=' +bldg + '&tid=' +tid;
		//	}
			window.open('/genergy2/setup/tenantedit.asp?pid=' +pid + '&bldg=' +bldg + '&tid=' +tid, 'ViewTenant','width=900,height=700,resizable=yes,scrollbars=yes')
		}
		
		function viewMeter(pid, bldg, tid, lid, mid){
		//	if (top.frames[1]){
		//		top.frames[1].location = '/genergy2/setup/contentfrm.asp?action=meteredit&pid=' + pid+ '&bldg=' + bldg+ '&tid=' + tid+ '&lid=' + lid+ '&meterid=' + mid;
		//	} else {
		//		window.opener.location = '/genergy2/setup/contentfrm.asp?action=meteredit&pid=' + pid+ '&bldg=' + bldg+ '&tid=' + tid+ '&lid=' + lid+ '&meterid=' + mid;
		//	}
			window.open('/genergy2/setup/contentfrm.asp?action=meteredit&pid=' + pid+ '&bldg=' + bldg+ '&tid=' + tid+ '&lid=' + lid+ '&meterid=' + mid, 'ViewMeter','width=900,height=700,resizable=yes,scrollbars=yes')
		}
		
		function viewBuilding(bldg, pid){
		//	if (top.frames[1]){
		//		top.frames[1].location = '/genergy2/setup/buildingedit.asp?pid=' + pid + "&bldg=" + bldg
		//	} else {
		//		window.opener.location = '/genergy2/setup/buildingedit.asp?pid=' + pid + "&bldg=" + bldg
		//	}
			window.open('/genergy2/setup/buildingedit.asp?pid=' + pid + "&bldg=" + bldg, 'ViewBuilding','width=900,height=700,resizable=yes,scrollbars=yes')
		}
				
			
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
		<table border=0 cellpadding="3" cellspacing="1" width="100%">
			<tr bgcolor="#6699cc"> 
				<td width="2%" align="center" valign="middle" >
					<span id="allnotes" style="cursor:hand;text-decoration:none;" onclick="javascript:showall()">[+]</span>
				</td>
				<td width="3%">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td bgcolor="#66ff66" title="Tickets Open For Less Then 2 Days">&nbsp;</td>
							<td bgcolor="#FFcc00" title="Tickets Open For Less Then 1 Week">&nbsp;</td>
							<td bgcolor="#cc0033" title="Tickets Open For More Then 1 Week">&nbsp;</td>
						</tr>
					</table>
				</td>
				<td width="4%" align="center"><span class="standardheader">ID</span></td>
				<td width="9%" align="center"><span class="standardheader">Assigned To</span></td>
				<td width="10%" align="center"><span class="standardheader">Date Opened</span></td>
				<td width="10%" align="center"><span class="standardheader">Requester</span></td>
				<td width="30%" align="center"><span class="standardheader">Description</span></td>
				<td width="10%" align="center"><span class="standardheader">Due Date</span></td>
				<td width="19%" align="center"><span class="standardheader">Opened For</span></td>
			</tr>
		</table>
		<div style="width:100%; overflow:auto; height:80%;border-bottom:1px solid #cccccc;">
			<%
			while not rst.eof
				Dim tcolor, age	
				if rst("closed") = 0 then 	  
					age = date() - cdate(trim(rst("duedate")))
				else 
					age = cdate(trim(rst("fixdate"))) - cdate(trim(rst("date")))	
				end if 
				
				if rst("runticket") then 
					tcolor = "#cccccc"	
				else
					if age > 7 then 
						tcolor = "#cc0033"
					else 
						if age < 2 then 
							tcolor = "#66ff66"
						else
							tcolor = "#FFcc00"
						end if
					end if
				end if
					
				%>
				<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
					<%
					sql = "select * from ttnotes where ticketid =" & rst("ticketid") &" order by date desc"
					dim rstNotes
					set rstNotes = server.createobject("ADODB.recordset")
					rstNotes.open sql, cnnIT,1
					Dim notecount 
					notecount = rstNotes.recordcount
					%>
					<tr bgcolor="#ffffff" valign="top"> 
						<td width="2%" align="center" valign="middle" >
							<%if not rstNotes.eof then %>
								<span id="func<%=rst("ticketid")%>" name = "notefunc" style="cursor:hand;text-decoration:none;" onclick="note('<%=rst("ticketid")%>')">[+]</span>
							<%else %>
								<span id="func<%=rst("ticketid")%>" name = "empty">[o]</span>
							<%end if %>
						</td>
						<td bgcolor="<%=tcolor%>" width="3%" align="center" valign="middle">
							<%IF rst("client") then %> 
								<img src="../images/critical.gif" width="20" height="20" align="center">
							<%end if%>
						</td>
						<td width="4%" align="center" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" 							onClick="javascript:document.location='ticket.asp?mode=update&tid=<%=rst("ticketid")%>';parent.document.all.Function.innerHTML = 'Viewing Ticket <%=rst("ticketid")%>'" ><%=rst("ticketid")%>
						</td>
						<td width="9%">
							<a href="userreports.asp?status=0&userlist='<%=rst("userid")%>'&listlength=<%=request("listlength")%>" target="_blank">
								<%=rst("userid")%>
							</a>
						</td>
						<td width="10%"><%=FormatDateTime(trim(rst("date")),2)%></td>
						<td width="10%"><%=rst("Requester")%></td>
						<td width="30%">
							<%=left(rst("initial_trouble"),50)%>
							<%if Len(trim(rst("initial_trouble"))) > 50 then%>...<%end if%>
						</td>
						<td width="10%" align="right">
							<%if rst("closed") = 1 then %>
								<%=rst("fixdate")%>
							<%else%>
								<%=FormatDateTime(rst("duedate"),2)%>
							<%end if %>
						</td>
						<td width="19%" align="right">
							<%
							dim cnnMainModule, rstTicketFor
							set rstTicketFor = server.CreateObject("adodb.recordset")
							set cnnMainModule = server.createObject("adodb.connection")
							cnnMainModule.open Application("cnnstr_SuperMod")
							select case rst("ticketfortype")	
								case "PID"
									rstTicketFor.open "select name from portfolio where id = '" & rst("ticketfor") & "'", cnnMainModule
									%><div align="left">Portfolio: </div>
									<a href="javascript:viewPortfolio('<%=rst("ticketfor")%>');"><%
									if not rstTicketFor.eof then
										response.write (rstTicketFor("name") & "(PID:  " & rst("ticketfor")) & ")"%></a><%
									end if
									
								case "BLDGNUM"
									dim bldgSql
									bldgSql = "select portfolioid, strt from buildings where bldgNum = '" & rst("ticketfor") & "'"
									rstTicketFor.open bldgSql, getLocalConnect(rst("ticketfor"))
									if not rstTicketFor.eof then
										%><div align="left">Building: </div>
										<a href="javascript:viewBuilding('<%=rst("ticketfor")%>', '<%=rstTicketFor("portfolioid")%>')">
										<%=rstTicketFor("strt")%>  (BldgNum:  <%=rst("ticketfor")%>)</a>
									<%end if
									
								case "TID"
									dim tempPid, tempBldg, tenantName, sqlTid
									sqlTid = "select billingname, bldgnum from " & makeIPUnion("tblLeases", "") & " tl where tl.billingId = '" & rst("ticketfor") & "'"
									rstTicketFor.open sqlTid, cnnMainModule
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
									%><div align="left">Tenant: </div>
									<a href="javascript:viewTenant('<%=tempPid%>','<%=tempBldg%>', '<%=rst("ticketfor")%>')">
										<%=tenantName%>  (Tenant ID:  <%=rst("ticketfor")%>)
									</a>
									<%
					
								case "METERID"
									dim meterNum, tempLid, sqlMeter, tempTid
									sqlMeter = "select meternum, bldgnum, leaseutilityid as lid from " & makeIpUnion("meters", "") 
										sqlMeter = sqlMeter & " m where m.meterID = '" & rst("ticketfor") & "'"
										
									rstTicketFor.open sqlMeter, cnnMainModule
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
									%><div align="left">Meter:</div> 
									<a href="javascript:viewMeter('<%=tempPid%>', '<%=tempBldg%>', '<%=tempTid%>', '<%=tempLID%>', '<%=rst("ticketfor")%>')">
										<%=meterNum%>  (Meter ID:  <%=rst("ticketfor")%>)
									</a>
									<%
								case else
									%>Internal Operations<%
							end select
							%>
						</td>	
					</tr>
				</table>
			
				<div id="note<%=rst("ticketid")%>" name="noteset" style="width:100%; height:100;border-bottom:1px solid #cccccc;display:none;background-color :#ffffcc;">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
							<td width="15%" >Date</td>							
							<td width="70%" >Notewell </td>							
							<td width="10%" >uid</td>
							<td width="5%">Time</td>
							<td></td>
						</tr>
					</table>
					<div style="width:100%; overflow:auto; height:95;border-bottom:1px solid #cccccc;background-color :#ffffcc;">
						<table width="100%" border="0" cellspacing="0" cellpadding="3">
							<%
							if not rstNotes.eof then 
								while not rstNotes.eof
									%>
									<tr valign="top"> 
										<td width="15%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rstNotes("date")%></td>
										<td width="70%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rstNotes("note")%></td>
										<td width="10%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rstNotes("uid")%></td>
										<td width="5%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rstNotes("time")%></td>
									</tr>
									<% 
									rstNotes.movenext
								wend
							else
								 %><td colspan="3" align="center">NO TROUBLE NOTES FOUND FOR THIS TICKET</td><%		  
							end if 
							rstNotes.close%>
						</table> 		 
					</div>
				</div>
			<%rst.movenext%>
		<%wend%>
		</div> 
	<%else		' no tickets found%>
		No tickets found matching criteria.
	<%end if
end if%>
</div>	
</body>
</html>