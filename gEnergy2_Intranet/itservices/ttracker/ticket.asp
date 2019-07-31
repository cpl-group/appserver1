<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'COMMENTS
'1/16/2007 N.Ambo
'This page is used for initally creating, saving,updating, and closing trouble tickets; 
'it has code for modes: NEW, UPDATE, SAVE, CLOSE
'2/15/2008 N.Ambo chaged size of window for transfer ticket screen from 300,100 to 300,200

if 	not(allowgroups("Genergy Users")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim lid, bldgnum, byear, bperiod, period, UMinfo,pid
byear = 0
bperiod = 0

lid = trim(request("tid"))

if trim(request("bldgnum"))="" then bldgnum = trim(request("bldg")) else bldgnum = trim(request("bldgnum"))
if bldgnum="" then bldgnum = trim(request("building"))
if pid = "" then pid = trim(request("pid"))

'get bill year and bill period if these values have been passed from another screen
if instr(trim(request("period")),"|")>0 then period = trim(request("period")) else period = ""
if period<>"" then
	if split(period,"|")(0)<>"" then byear = split(period,"|")(0) else byear=0
	if split(period,"|")(1)<>"" then bperiod = split(period,"|")(1) else bperiod=1
end if

dim cnn, rst, strsql,UsersRS,SubjectInfo
dim jid,Desc , company ,job ,cStatus ,projmanager,comppercent ,jobnotes,primarybilling,secondarybilling,primary_amt,secondary_amt, tcolor, customer, address_street, jcity,jstate,jzip,jfloor, cust_name,projid, address_floor, jtype,jtypeid,objUser,strDomain, strGroup,oGroup, y,cnnMainModule, ticketfortype, ticketfor
ticketfortype = request("ticketfortype")
ticketfor = request("ticketfor")
SubjectInfo =request("info")
set cnnMainModule = server.createobject("ADODB.connection")
set cnn 	= server.createobject("ADODB.connection")
set rst 	= server.createobject("ADODB.recordset")


cnnMainModule.open getConnect(pid,bldgnum,"Billing")
cnn.open getConnect(pid,bldgnum,"dbCore")

'if a billingid (lid) is passed then get building number based on that billingid
if bldgnum = "0" and lid <> "" then
	rst.open "SELECT bldgnum FROM tblleases l WHERE billingid='"&lid&"'", cnnMainModule
	if not rst.eof then bldgnum = rst("bldgnum")
	rst.close
end if

select case request("mode")
	'This code is followed when the trouble ticket is first opened; the mode will then be
	'changed to 'save' when the 'Trouble Tickets' page opens so that when the user clicks the save button the code 
	'under the mode 'save' will be carried out
	case "new"
		dim userlist, usernamelist
		set UsersRS = server.createobject("ADODB.recordset")
		UsersRS.open "select * from adusers_genergyusers where email is not null order by company,department,fullname", getConnect(0,0,"dbCore")
		UsersRS.MoveFirst
		%>
    	<html>
    	<head>
    		<title>Trouble Tickets</title>
			<link rel="Stylesheet" href="../../styles.css" type="text/css">
			<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		</head>
		<script>	
		function closeticket(){
			if (confirm("Cancel ticket?")){
				document.location.href="troublesearch.asp?status=0&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>" 
			}
		}
		function jobPicked(job){
			document.form1.jobnumber.value = job
			
		}
		
		</script>
		<body bgcolor="#dddddd">
			<form name="form1" method="post" action="ticket.asp">
			<input type="hidden" name="mode" value="save">
				
				
			
  <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#6699cc"> 
      <td colspan="3"><span class="standardheader">New Trouble Ticket</span></td>
    </tr>
	<% 'If the ticket is being opened for a job then the job field is not shown on screen
	if request("jobid") = "" then %>
    <tr>
      <td colspan="3">Job Number 
        <input name="jobnumber" type="text"  size="10" maxlength="10">
        &nbsp;(optional)&nbsp;<a href="#" onClick="javascript:window.open('/um/opslog/timesheet-beta/joblist.asp','QuickSearch','toolbars=no, height=300, width=400');">Quick 
        job search</a>&nbsp;</td>
    </tr>
	<% end if %>
    <tr> 
      <td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"></td>
    </tr>
    <tr> 
      <td colspan="3" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"></td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <strong>Trouble Ticket Details</strong> </td>
      <td width="50%" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <strong>Initial Trouble Report</strong> </td>
    </tr>
    <tr valign="top"> 
      <td width="26%" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <p>Requester</p></td>
      <td width="19%" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
       
	  <% 
			
	  GenerateUserList "requester",UsersRS,"","" ,getKeyValue("user")
	  %>
		
		 </td>
      <td width="50%" rowspan="8"
						style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <textarea name="note" cols="70%" rows="10"></textarea> </td>
    </tr>
    <tr valign="top"> 
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        Assign Ticket to </td>
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
      <% GenerateUserList "uid",UsersRS,"","",getKeyValue("user") %>
	  </td>
    </tr>
    <tr valign="top"> 
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        Due Date (xx/xx/xxxx) </td>
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <input name="duedate" type="text" value="<%=date()+2%>" size="10" maxlength="10"> 
      </td>
    </tr>
    <tr valign="top"> 
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        Copy User </td>
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <% GenerateUserList "ccuid",UsersRS,"multiple","4",getKeyValue("user") %>
	 </td>
    </tr>
    <tr valign="top"> 
      <td height="23" colspan="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        Client Related Ticket 
        <input type="checkbox" name="client" value="1">
        Running Ticket 
        <input type="checkbox" name="runticket" value="1" onclick="duedate.value='NA';"> 
      </td>
    </tr>
    <tr valign="top"> 
      <td height="23" colspan="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;">
<%
dim tempRST
set tempRST = server.createObject("adodb.recordset")
select case lcase(ticketfortype)
case "portfolioid"
	tempRST.open "select name from portfolio where id = '" & ticketfor & "'", getConnect(ticketfor,0,"Billing")%>
	This ticket will be opened for <%=tempRST("name")%>, Portfolio ID <%=ticketfor%> 
	<input type = "hidden" name="pid" value="<%=ticketfor%>"> 
   <input type = "hidden" name="info" value="<%=tempRST("name")%>"> 
	<%tempRST.close
case "bldgnum"
	tempRST.open "select portfolioid, strt from buildings where bldgNum = '" & ticketfor & "'", getLocalConnect(ticketfor)%>
	This ticket will be opened for <%=tempRST("strt")%> :Building Number <%=ticketfor%> 
	<input type = "hidden" name="bldgnum" value="<%=ticketfor%>"> 
    <input type = "hidden" name="info" value="<%=tempRST("strt")%>"> 
	<%tempRST.close
case "tid"
	tempRST.open "select billingname, bldgnum from tblLeases tl where '"&split(getPIDIP(pid),"\")(1)&"-' + ltrim(convert(varchar,tl.billingId)) = '" & ticketfor & "'", getLocalConnect(bldgnum)%>
	This ticket will be opened for <%=tempRST("billingname")%> : Tenant ID 
	<%=ticketfor%> <input type = "hidden" name="tid" value="<%=ticketfor%>"> 
     <input type = "hidden" name="info" value="<%=tempRST("billingname")%>"> 
	<%tempRST.close
case "meterid"
	tempRST.open "select meternum,meterid from Meters m where '"&split(getPIDIP(pid),"\")(1)&"-' + ltrim(convert(varchar, m.meterID)) = '" & ticketfor & "'", getLocalConnect(bldgnum)
	if not tempRST.eof then%>
		This ticket will be opened for Meter: <%=tempRST("meternum")%> : ID <%=ticketfor%> 
	<%else%>
		This ticket will be opened for Meter ID <%=ticketfor%> 
	<%end if%>
	<input type = "hidden" name="meterid" value="<%=ticketfor%>"> 
    <input type = "hidden" name="info" value="<%=tempRST("meterid")%>"> 
	<%tempRST.close
case "joblog"
	
	tempRST.open "SELECT * FROM dbo.MASTER_JOB WHERE id='" & request("jobid") & "'", getConnect(0,0,"intranet")
    if not tempRST.eof then%>
		This ticket will be opened for Job Number: <%=tempRST("id")%> 
		<%
		Dim showjobselect
		showjobselect=true
	end if%>
	<input type = "hidden" name="jobid" value="<%=request("jobid")%>"> 
	<input type = "hidden" name="jobnumber" value="<%=request("jobid")%>"> 
	<%tempRST.close
end select
if bldgnum <> "" and bldgnum<>"0" then%>
Billperiod&nbsp; <select name="period">
<option value="0">NA</option>
<%
tempRST.open "SELECT distinct billyear, billperiod FROM billyrperiod WHERE bldgnum='"&bldgnum&"' and dateend>dateadd(m,-6,getdate()) and dateend<dateadd(m,6,getdate()) ORDER BY billyear desc, billperiod desc", getLocalConnect(bldgnum)
do until tempRST.eof
%>
<option value="<%=tempRST("billyear")&"|"&tempRST("billperiod")%>"><%=tempRST("billperiod")&", "&tempRST("billyear")%></option>
<%
tempRST.movenext
loop
tempRST.close
%>
</select> 
<%end if
set tempRST=nothing
%>
      </td>
    </tr>
    <tr bgcolor="#dddddd"> 
      <td colspan="3"> <div style="margin-left:1px;"> 
          <input name="child" type="hidden" value="<%=request("child")%>">
          <input name="bldg" type="hidden" value="<%=bldgnum%>">
          <a href="javascript:form1.submit();">save</a> | <a href="javascript:<%if request("child") = 1 then %> window.close()<%else%>closeticket()<%end if%>;">cancel</a> 
        </div></td>
    </tr>
  </table>
        <input type = "hidden" name="ticketfor" value="<%=request("ticketfor")%>"> 
		<input type = "hidden" name="ticketfortype" value="<%=request("ticketfortype")%>"> 
		</form>
	</body>
	</html>
	<%
    case "update"
	dim tid, dueText
	tid = request("tid")
	strsql = "select * from tickets where id = " & tid 
	rst.open strsql, cnn
	if not rst.eof then 
  %>
    <html>
    <head>
    <title>Trouble Tickets</title>
    <script>
	function closeticket(ticketid)
	{
		if (confirm("Close ticket " + ticketid +"?")){
		  document.location.href="ticket.asp?mode=close&tid="+ticketid+"&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>"
		}
	}
	function openwin(url,mwidth,mheight){
	window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
	}
	</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#dddddd">
<form name="form1" method="post" action="ticket.asp">
<input type="hidden" name="mode" value="saveupdate">
<input type = "hidden" name="ticketfor" value="<%=rst("ticketfor")%>"> 
<input type = "hidden" name="ticketfortype" value="<%=rst("ticketfortype")%>"> 

    
  <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#6699cc" valign="center"> 
      	<td nowrap>
	  		<span class="standardheader"><u><% if rst("closed") then response.write "Closed" else response.write "Open" end if%></u>
				Trouble Ticket <u><%=rst("id")%></u> <%if isnumeric(rst("jobnum")) and rst("jobnum") <> "0" then %> associated with job number <a href="#" onclick="openwin('https://appserver1.genergy.com/gEnergy2_Intranet/opsmanager/joblog/viewjob.asp?jid=<%=rst("jobnum")%>',800,500)"><%=rst("jobnum")%></a><%end if%></span>
		</td>
		<td align="right">
	  		<img src="/images/print.gif" style="cursor:hand" onclick="javascript:openwin('singleticketprint.asp?ticketid=<%=rst("id")%>', 'SingleTicketPrint','width=500,height=500,resizable=yes,scrollbars=yes')"><a href="javascript:openwin('singleticketprint.asp?ticketid=<%=rst("id")%>', 'SingleTicketPrint','width=500,height=500,resizable=yes,scrollbars=yes')" style="text-decoration:none;decoration:none">&nbsp;&nbsp;<span class="standardheader">Print Ticket</span></a>
        </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><br>
        Ticket opened for <u><%=rst("Requester")%></u> in 
		 <% 
		  Dim rst2
		  set rst2 = server.createobject("ADODB.recordset")
		  strsql = "select * from departments order by department"
		  rst2.open strsql,cnn 
		  while not rst2.eof 
		   if trim(rst("department"))= trim(rst2("depid")) then
		   response.write "<u>" & rst2("department") & "</u>"
		   end if
		   
		  rst2.movenext
		  wend
		  rst2.close
		  if rst("runticket") then 
		 		dueText = "is an ongoing issue"
		  else
		  		dueText = "is due on "& FormatDateTime(trim(rst("duedate")),2)
		  end if 
		  %> at <%=rst("date")%>. This ticket is assigned to <u><%=rst("userid")%></u> and <%=dueText%>. This ticket is CC to <u><%=rst("ccuid")%></u>.
        </td>
    </tr>
	<%if rst("client") then %><tr> <td><b>This is a client related ticket</b>.</td></tr><%end if%>
	<%if rst("ticketfortype") <> "" then
		%>
		<script>
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
				
		<tr>
			<td colspan="3" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;">
				This ticket has been opened for
				<%
			
			
				dim rstTicketFor
				set rstTicketFor = server.CreateObject("adodb.recordset")

				select case ucase(rst("ticketfortype"))
					case "PORTFOLIOID"
						dim sql
						sql = "select name from portfolio where id = '" & rst("ticketfor") & "'"
						rstTicketFor.open sql, getConnect(pid,bldgnum,"dbcore")
						%>Portfolio: <a href="javascript:viewPortfolio('<%=rst("ticketfor")%>');"><%
						if not rstTicketFor.eof then
							response.write (rstTicketFor("name") & " (PID: ")
						end if
						rstTicketFor.close
						%><%=rst("ticketfor")%>)</a><%
						
					case "BLDGNUM"
						dim bldgSql
						bldgSql = "select portfolioid, strt from buildings where bldgNum = '" & rst("ticketfor") & "'"
						rstTicketFor.open bldgSql, getLocalConnect(rst("ticketfor"))
						if not rstTicketFor.eof then
							%>Building: <a href="javascript:viewBuilding('<%=rst("ticketfor")%>', '<%=rstTicketFor("portfolioid")%>')"><%=rstTicketFor("strt")%></a>
						<%end if
						
					case "TID"
						dim tempPid, tempBldg, tenantName, sqlTid
						
						sqlTid = "select billingname, bldgnum from tblLeases tl where '"&split(getPIDIP(pid),"\")(1)&"-' + ltrim(convert(varchar,tl.billingId))= '" & rst("ticketfor") & "'"

						rstTicketFor.open sqlTid, getConnect(pid,bldgnum,"billing")
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
						Tenant: <a href="javascript:viewTenant('<%=tempPid%>','<%=tempBldg%>', '<%=split(rst("ticketfor"),"-")(1)%>')">
							<%=tenantName%> (BillingID <%=rst("ticketfor")%>)
						</a>
						<%
		
					case "METERID"
						dim meterNum, tempLid, sqlMeter, tempTid
						sqlMeter = "select meternum, bldgnum, leaseutilityid as lid from Meters m where'"&split(getPIDIP(pid),"\")(1)&"-' + ltrim(convert(varchar,m.meterID))  = '" & rst("ticketfor") & "'"
						rstTicketFor.open sqlMeter, cnnMainModule
						if not rstTicketFor.eof then
							tempLid = rstTicketFor("lid")
							meterNum = rstTicketFor("meterNum")
							tempBldg = rstTicketFor("bldgnum")
							rstTicketFor.close
							sqlMeter = "select tl.billingid as tid, b.portfolioid as pid from tblleasesutilityprices tlup inner join tblleases tl on tl.billingid = tlup.billingid inner join buildings b on tl.bldgnum = b.bldgnum where tlup.leaseutilityid = '" & tempLid & "'"
							rstTicketFor.open sqlMeter,  cnnMainModule
							if not rstTicketFor.eof then
								tempTid = rstTicketFor("tid")
								tempPid = rstTicketFor("pid")
							end if
						end if
						%>
						Meter: <a href="javascript:viewMeter('<%=tempPid%>', '<%=tempBldg%>', '<%=tempTid%>', '<%=tempLID%>', '<%=split(rst("ticketfor"),"-")(1)%>')">
							<%=meterNum%> (MeterID <%=rst("ticketfor")%>)
						</a>
						<%
				end select
				%>
							&nbsp;<%if not(isnull(rst("billyear"))) then%>bill year <%=rst("billyear")%><%end if%>
							<%if not(isnull(rst("billperiod"))) then%> period <%=rst("billperiod")%><%end if%>
			</td>
		</tr>
	<%end if %>
			
    <tr valign="top">
      <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;">Initial 
        Trouble Report</td>
      <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        Trouble Notes</td>
    </tr>
    <tr valign="top"> 
      <td width="38%" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;">
	  <div style="width:100%; overflow:auto; height:100;">
	  <table>
	  <tr>
	  <td>
	  <%=rst("initial_trouble")%>
	  </td></tr>
	  </table>
	  </div>	  
	  </td>
      <td width="62%" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;">		  
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
				<td width="34%" >Date</td>
				<td width="55%" >Note</td>
				<td width="10%">uid</td>
				<td width="13%">time</td>
				<td></td>
          </tr>
 </table>
	<div style="width:100%; overflow:auto; height:100;border-bottom:1px solid #cccccc;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<%
			strsql = "select * from ttnotes where ticketid =" & rst("id") &" order by date desc"
			rst2.open strsql, cnn,1
			Dim notecount 
			notecount = rst2.recordcount
			
			if not rst2.eof then 
				while not rst2.eof
					%>
					<tr bgcolor="#cccccc" valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" 
						onMouseOut="this.style.backgroundColor = '#cccccc'" onClick="javascript:openwin('./note.asp?mode=view&nid=<%=rst2("id")%>',400,200)" > 
						
						<td width="35%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst2("date")%></td>
						<td width="55%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=left(rst2("note"),36)%><%if len(rst2("note"))>36 then%>...<%end if%>&nbsp;</td>
						<td width="12%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst2("uid")%></td>
						<td width="13%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst2("time")%></td>
					</tr>
					<% 
					rst2.movenext
				wend
			else%>
				<td colspan="3" align="center">NO TROUBLE NOTES FOUND FOR THIS TICKET</td>		  
			<%end if %>
		</table> 		 
	</div>
 <% if rst("closed") then
 response.write "No New Notes May Be Added, Ticket is Closed"
 else
	%>	
	<script>
	function closeWindow(status){
		try{
			if (parent.document.all.Function.innerHTML){
				parent.document.all.Function.innerHTML = '';
				window.back();
			}
		}catch(exception){};
	}

	</script>	
	<a href="javascript:openwin('./note.asp?mode=new&ticketid=<%=rst("id")%>&child=<%=request("child")%>&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>',300,200)">new note</a> | 
	<a href="javascript:openwin('./note.asp?mode=transfer&ticketid=<%=rst("id")%>&uid=<%=rst("userid")%>',300,200)">transfer ticket</a> | 
	<a href="javascript:openwin('./note.asp?child=<%=request("child")%>&mode=requestupdate&ticketid=<%=rst("id")%>&mailto=<%=rst("userid")%>',300,200)">request update</a></td>
<%end if%>
    </tr>
    <tr bgcolor="#dddddd"> 
      <td colspan="2"> <div style="margin-left:1px;"> 
	  <%
	  dim exithref
	  if not rst("closed") then %>
	  	<%if notecount > 0 then %>
           <a href="javascript:closeticket('<%=rst("id")%>');<%if request("child") = 1 then%>opener.document.location = opener.document.location; window.close()<%end if%>">Close Ticket</a> | 
		 <%end if%>
		 <a href="<%if request("child") = 1 then%> javascript:window.close() <%else%> javascript:history.back()<%end if%>">Exit</a>
		<%else%>
		 <a href="<%if request("child") = 1 then%> javascript:window.close() <%else%>javascript:history.back()<%end if%>">Exit</a>  
		<%end if %> 
          
        </div></td>
    </tr>
  </table>

    </form>
    </body>
    </html>
<%
  else
  %>
  <script>
		try{
			parent.document.all.Function.innerHTML = "All Open Tickets : Ticket <%=tid%> Not Found"
		}catch(exception){}
	document.location = "./troublesearch.asp?status=0&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>"
  </script>
  <%
  end if 
  
  case "save"
  	dim note, Requester, department,userid,clientticket,childstatus,ccuid, duedate,runticket, jobid
  
    note 			= replace(request("note"),"'","''")
    requester 		= request("requester")
    department 		= request("department")
	userid 			= request("uid") 
	ccuid 			= request("ccuid")
	clientticket 	= request("client")
	childstatus 	= request("child")
	duedate 		= request("duedate")
	ccuid 			= replace(ccuid,",",";")
	runticket 		= trim(request("runticket"))
	jobid 			= request("jobnumber")
	
	if request("portfolioid") <> "" then
		ticketfor = request("portfolioid")
		ticketfortype = "PORTFOLIOID"
	elseif request("tid") <> "" then
		ticketfor = request("tid")
		ticketfortype = "TID"
	elseif request("meterid") <> "" then
		ticketfor = request("meterid")
		ticketfortype = "METERID"
	elseif request("bldgnum") <> "" then
		ticketfor=  request("bldgnum")
		ticketfortype = "BLDGNUM"
	elseif request("jobid") <> "" then
		ticketfor=  request("jobid")
		ticketfortype = "joblog"
	elseif request("ticketfortype") <> "" and request("ticketfor") <> "" then
		ticketfor=  request("ticketfor")
		ticketfortype = request("ticketfortype")
	else
		ticketfor = ""
		ticketfortype = ""
	end if
	note = getUMInfo(ticketfortype, ticketfor, bldgnum) & vbcrlf & note
	
	if len(clientticket) = 0 then 
		clientticket = 0
	end if
		
    set cnn = server.createobject("ADODB.connection")
    set rst = server.createobject("ADODB.recordset")
    cnn.open getConnect(pid,bldgnum,"dbCore")
	
	if trim(note) = "" or trim(Requester) ="" or trim(department) = "-1" then 
	%>
	<script>
	<% if childstatus = "1" then %>
		alert("Ticket not created; Required Information Missing")
		window.close()
	<% else%>
	try{
		parent.document.all.Function.innerHTML = "All Open Tickets : Last Ticket not created; Required Information Missing"
	}catch(exception){}
	document.location = "./index.asp?status=0&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>"
	<% end if %>
	</script>
	<%	
	else
	if runticket <> "" and runticket <> "0"   then 
		strsql = "insert into tickets (jobnum,initial_trouble, requester,department,userid, client,ccuid, runticket, ticketfor, ticketfortype, billyear, billperiod) values ('"&trim(jobid)&"','"&trim(note)&"','"&trim(requester)&"','"&trim(department)&"','"& trim(userid)& "','" & trim(clientticket) & "','" & trim(ccuid) &"','" &trim(runticket)&"','" &trim(ticketfor)&"','" &trim(ticketfortype)& "', '"&byear&"', '"&bperiod&"')"
	else
		strsql = "insert into tickets (jobnum,initial_trouble, requester,department,userid, client,ccuid, duedate, ticketfor, ticketfortype, billyear, billperiod) values ('"&trim(jobid)&"','"&trim(note)&"','"&trim(requester)&"','"&trim(department)&"','"& trim(userid)& "','" & trim(clientticket) & "','" & trim(ccuid) &"','" &trim(duedate)&"','" &trim(ticketfor)&"','" &trim(ticketfortype)& "', '"&byear&"', '"&bperiod&"')"
	end if 
	cnn.Execute strsql
	
	sendupdate "new"
   %>
   <script>
   	<% if childstatus = "1" then %>
		try{
			opener.document.location = opener.document.location
		}catch(exception){}
		window.close()
	<% else%>
	   document.location = "./troublesearch.asp?status=0&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>&searchstring=<%if request("ticketfortype")<>"" then%>+<%else%><%=getxmlusername()%>&internalops=True<%end if%>" 
		try{
		   parent.document.all.Function.innerHTML = 'All Open Tickets'
		}catch(exception){}
   <%end if %>
   </script>	
<%
	end if
   
 
   case "close"
  	tid = request("tid")
	childstatus = request("child")
    set cnn = server.createobject("ADODB.connection")
    set rst = server.createobject("ADODB.recordset")
    cnn.open getConnect(pid,bldgnum,"dbCore")

    strsql = "update tickets set closed = 1, fixdate='" & date() & "' where id=" & tid
	cnn.Execute strsql
	sendupdate tid
%>
	<script>
   	<% if childstatus = "1" then %>
		opener.document.location = opener.document.location
		window.close()
	<% else %>
		document.location = "./troublesearch.asp?status=0&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>&searchstring=<%if request("ticketfortype")<>"" then%>+<%else%><%=getxmlusername()%>&internalops=True<%end if%>" 
		try{
			parent.document.all.Function.innerHTML = 'All Open Tickets'
		}catch(exception){}
	<% end if %>
   	</script>	
<%
  case "enterid" 
  %>
    <html>
    <head>
    <title>Trouble Tickets : Enterid</title>
	<link rel="Stylesheet" href="../../styles.css" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
	<script>
	function openticket(){
	if (document.form1.tid.value == ""){
		document.location = "./troublesearch.asp?status=0&ticketfortype=<%=request("ticketfortype")%>&ticketfor=<%=request("ticketfor")%>"
		try{
			parent.document.all.Function.innerHTML = "All Open Tickets"
		}catch(exception){}
	}else{
		document.location = "ticket.asp?mode=update&tid=" + document.form1.tid.value
	}
	}
	</script>
	<body bgcolor="#dddddd" >
	<form name="form1" method="post" action="ticket.asp">
	<input type="hidden" name="mode" value="update">    
	  <table border=0 cellpadding="3" cellspacing="0" width="100%" height="100%" bgcolor="#eeeeee">
		<tr bgcolor="#6699cc"> 
		  <td colspan="2" align="center"><span class="standardheader">Enter ID: <input name="tid" type="text" size="10" ><input name="open" type="button" value="OPEN" onclick="openticket()"> </span>
			</td>
		</tr>		
	  </table>
    </form>
    </body>
    </html>
  <%
   case else
end select
%>
<%
function sendupdate(ticket)
	
	dim cnn2,rs,requestor,assignedto,userarray, emailarray,dateopen,masternote,tixstatus,subject
    set cnn2 = server.createobject("ADODB.connection")
    set rs 	 = server.createobject("ADODB.recordset")
    cnn2.open getConnect(pid,bldgnum,"intranet")

	if ticket = "new" then 
		sql = "select top 1 * from tickets order by id desc" 
	else
		sql = "select * from tickets where id =" & ticket
	end if
	
	
	rs.open sql, cnn
	if not rs.eof then 
	'UMinfo = getUMInfo(rs("ticketfortype"), rs("ticketfor"))
	requestor  = rs("requester")
	assignedto = rs("userid")
	userarray  = rs("ccuid")
	dateopen   = rs("date")
	duedate    = rs("duedate")
	masternote = "[Original Note: "&rs("initial_trouble") &" submitted by "&requestor&" and currently assigned to "&assignedto&"]"
	ticket 	   = rs("id")
	
	if rs("closed") then 
		tixstatus = "Ticket Closed"
	else
		tixstatus = "Open Ticket"
	end if
	
	end if
	rs.close
	
	userarray  = userarray & ";" & requestor &";"&assignedto
	userarray  = split(userarray,";")
	
	for each x in userarray 
		rs.open "select email from ADusers_GenergyUsers where username = '" & trim(x) & "'", getConnect(0,0,"dbCore") 
		if not rs.eof then 
			if emailarray <> "" then 
			emailarray = emailarray & ";" & rs("email")
			else
			emailarray = rs("email")			
			end if 
		end if
		rs.close		
	next
	sql = "select * from ttnotes where ticketid = " & ticket & " order by date desc"
	rs.open sql, cnn
	if not rs.eof then 
		
		while not rs.eof
			masternote = masternote & vbCrLf & vbCrLf & " On "& rs("date") & " User " & rs("uid") &" added: " & rs("note") & vbCrLf 
		rs.movenext
		wend
	
	end if
	rs.close
	dim addon,adonsql
	'for ticket subject change added by Ando 10/16/06
	select case lcase(ticketfortype)
	case "portfolioid"
	addon ="For Portfolio: " & SubjectInfo
	case "bldgnum"
	addon= "For  Building: " & SubjectInfo
	case "tid"
    addon ="For Tenant: " & SubjectInfo
    case  "meterid"
	 adonsql ="SELECT BLDGNAME,m.bldgnum,m.meternum, m.meterid,L.TENANTNUM,TNAME FROM METERS m,tblLeasesUtilityPrices LUP,TBLLEASES L,BUILDINGS B WHERE METERID ='"&SubjectInfo&"' AND m.leaseutilityid = LUP.leaseutilityid AND LUP.BILLINGID = L.BILLINGID AND M.BLDGNUM = B.BLDGNUM"
	   
      rs.open adonsql, getConnect(pid,bldgnum,"Billing")
 	if not rs.eof then
		addon ="For Meternum: " & rs("meternum") &" , Tenant ("& rs("TENANTNUM") &"): " &rs("TNAME")  &",Building: " & rs("BLDGNAME") 
	end if
	rs.close
    end select
	
    
	subject = tixstatus & ":" & ticket &" (Opened on:"&dateopen&") " & addon
	sendmail emailarray,"GSA",subject, masternote

end function

%>
<!--#INCLUDE VIRTUAL="/genergy2_intranet/itservices/ttracker/UMInfo.asp"-->
