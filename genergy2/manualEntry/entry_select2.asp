<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'1/18/2008 N.Ambo modified to allow meter results to show as far back as bill year 2000
dim pid, building, byear, bperiod, meterid, utilityid, scroll
pid = request("pid")
building = request("building")
byear = request("byear")
bperiod = request("bperiod")
meterid = request("meterid")
utilityid = request("utilityid")
scroll = request("scroll")


if trim(utilityid) = "" then utilityid = 0
if trim(byear) = "" then byear = 0
if trim(scroll)="" then scroll = 0
if trim(bperiod) = "" then bperiod = 0

dim rst1, rst2, cnn1, rst5
set rst1 = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
set rst5 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
if trim(building)<>"" then
  cnn1.open getLocalConnect(building)
end if

%>
<html>
<head>
<title>Meter Manual Entry</title>
<script>
function showextended()
{
	for (i=0;i<extendedUsage.length;i++){
    	extendedUsage[i].style.display=(extendedUsage[i].style.display=='inline' ? 'none' : 'inline')
     }
}
function loadportfolio()
{	var frm = document.forms['form1'];
	var newhref = "entry_select.asp?pid="+frm.pid.value+"&scroll="+getScroll();
	document.location.href=newhref;
}

function loadbuilding()
{	var frm = document.forms['form1'];
	var newhref = "entry_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&scroll="+getScroll();
	document.location.href=newhref;
}

function loadyear()
{	var frm = document.forms['form1'];
	var newhref = "entry_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&scroll="+getScroll();
	document.location.href=newhref;
}

function loadperiod(bperiod)
{	var frm = document.forms['form1'];
	if((frm.building.value!='')&&(frm.byear.value!='')&&(bperiod!=''))
	{	var newhref = "entry_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&bperiod="+bperiod+"&scroll="+getScroll();
		document.location.href=newhref;
	}
}

function loadmeter(meterid)
{	var frm = document.forms['form1'];
	var newhref = "entry_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&bperiod=<%=bperiod%>&meterid="+meterid+"&scroll="+getScroll();
	document.location.href=newhref;
}

function checkPeriod(){
  if (document.forms['form1'].addperiod.selectedIndex > 0) { 
    document.form1.addentry.style.display = "inline";
    document.all.addentrymessage.innerHTML = "";
  } else { 
    document.form1.addentry.style.display = "none";
    document.all.addentrymessage.innerHTML = "Enter readings for this period in the fields below";
  }
}

function editReadings(bp,by){
  temp = 'entry_select.asp?pid=<%=pid%>&building=<%=building%>&byear='+by+'&bperiod='+bp+'&meterid=<%=meterid%>'+"&scroll="+getScroll();
  document.location = temp;
}

function getScroll(){
  try{return(document.all['meterlist'].scrollTop);}catch(exception){}
}
function JumpTo(url){
	var frm = document.forms['form1'];
	var url = url + "?pid=<%=pid%>&bldg=<%=building%>&building=<%=building%>&utilityid=<%=utilityid%>&byear=<%=byear%>&bperiod=" + frm.bperiod.value;
	window.document.location=url;
}
</script>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" onLoad="try{document.all['meterlist'].scrollTop=<%=scroll%>}catch(exception){}">
<form name="form1" action="entrysave.asp" method="post" onSubmit="document.forms[0].scroll.value = document.all['meterlist'].scrollTop">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr> 
      <td bgcolor="#6699cc"><span class="standardheader">Manual Entry</span></td>
      <td align="right" bgcolor="#6699cc"><% if building <> "" then %><select name="select" onChange="JumpTo(this.value)">
          <option value="#" selected>Jump to...</option>
          <option value="/genergy2/billing/processor_select.asp">Bill Processor</option>
          <option value="../validation/re_index.asp">Review Edit</option>
        <option value="/genergy2/setup/buildingedit.asp">Building Setup</option>
          <option value="/genergy2/billentry/entry.asp">Utility Bill Entry</option>
          <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem 
          Report</option>
            <option value="/genergy2/manualentry/entry_select_new.asp">Manual Entry v.2 Test</option>
            <option value="/genergy2/accounting_files/historic_acctFile.asp">Accounting Transactions</option>
      </select><% end if %></td>
    </tr>
    <tr> 
     <td colspan="2"  bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0">
        <tr valign="bottom"> 
          <td> 
            <%
						if allowGroups("Genergy Users") then%>
            <select name="pid" onChange="loadportfolio()">
              <option value="">Select Portfolio</option>
              <%
									dim sqlCommand
					sqlCommand = "(SELECT distinct pt.id, pt.name FROM buildings bd, portfolio pt WHERE pt.id=bd.portfolioid)"
								rst1.open sqlCommand & " UNION (SELECT distinct p.id, p.name FROM buildings_dbbilling b, portfolio p WHERE p.id=b.portfolioid) ORDER BY name", getConnect(pid,building,"dbCore")
								do until rst1.eof%>
              <option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then %> SELECTED <%end if%>><%=rst1("name")%></option>
              <%	rst1.movenext
					loop
					rst1.close
					%>
            </select> 
            <%		
						elseif isnumeric(pid) then
							rst1.open "SELECT name FROM portfolio WHERE id="&pid, getConnect(pid,building,"billing")
							if not rst1.eof then response.write rst1("name")
							rst1.close%>
            <input type="hidden" name="pid" value="<%=pid%>"> 
            <%
						end if			%>
          </td>
          <%
					
					
					if trim(pid)<>"" then%>
          <td> <select name="building" onChange="loadbuilding()">
              <option>Select Building</option>
              <%
								rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", getConnect(pid,building,"billing")
								if rst1.eof then
									rst1.close
									rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", getConnect(pid,building,"dbCore")
								end if
								do until rst1.eof%>
              <option <%if isBuildingOff(rst1("Bldgnum")) then%>class="grayout"<%end if%> value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then%> SELECTED<%end if%>> 
              <%=rst1("strt")%>, <%=trim(rst1("Bldgnum"))%></option>
              <%
									rst1.movenext
								loop
								rst1.close
	
								%>
            </select> </td>
          <%
					end if
					
					
					if trim(building)<>"" then%>
          <td width="6"></td>
          <td style="border:1px solid #cccccc;"> Only list meters with no readings 
            for:<br> <select name="byear" onChange="javascript:loadyear()">
              <%
								rst1.open "SELECT Distinct BillYear FROM BillYrPeriod WHERE BldgNum='"&building&"'", cnn1
								if rst1.eof then	%>
              <option value="">No Billing Years</option>
              <%
								else	%>
              <option value="">Select Bill Year</option>
              <%
								end if
								do until rst1.eof			%>
              <option value="<%=rst1("Billyear")%>"<%if trim(rst1("billyear"))=trim(byear) then%> SELECTED<%end if%>> 
              <%=rst1("Billyear")%> </option>
              <%
									rst1.movenext
								loop
								rst1.close		%>
            </select> 
            <%
							if trim(byear)="" then	%>
          </td>
          <%
							end if 
					end if
			if  cint(byear) <> 0 then%>
          <select name="bperiod" onChange="javascript:loadperiod(this.value)">
            <option value="0">All Online Meters</option>
            <%
			rst1.open "SELECT Distinct BillPeriod, billYear FROM BillYrPeriod WHERE BldgNum='"&building&"' and BillYear="&byear, cnn1
			do until rst1.eof								%>
            <option value="<%=rst1("BillPeriod")%>" <%if cint(bperiod)=cint(rst1("billperiod")) then%> selected <%end if%>> 
            <%=rst1("BillPeriod")%> </option>
            <%
			rst1.movenext
		loop
		rst1.close
		%>
          </select>
		  <% else %> <input name="bperiod" type="hidden" value="0">
		  <% end if%>
		  </td>
        <%
						dim Wmeternum, mmultiplier, dmultiplier, meterUtype, billingname, extusg, mCumulativeDem
						extusg = false
						if trim(meterid)<>"" then
							rst1.open "SELECT * FROM meters m, tblleasesutilityprices lup, tblleases l, tblutility u WHERE l.billingid=lup.billingid and u.utilityid=lup.utility  and m.leaseutilityid=lup.leaseutilityid and meterid="&meterid, cnn1
							
							if not rst1.eof then 
								Wmeternum = rst1("meternum")
								mmultiplier = rst1("manualmultiplier")
								dmultiplier = rst1("demandmultiplier")
								meterUtype = rst1("utilitydisplay")
								utilityid = rst1("utilityid")
								billingname = rst1("billingname")
								mCumulativeDem = rst1("Cumulative")
								if rst1("extusg") then extusg = true
								%>
        <%
							end if
							rst1.close
						end if
		%>
                <td>&nbsp;&nbsp;<input type="button" value="Refresh" onClick="loadperiod(bperiod.value)"></td>
</tr>
      </table></td>
    </tr>
  </table>
<br>

<%
if trim(building)<>"" then%>
	<%
		'if building = "RNEEMBKA" then
			dim mSql, vSql, pdate
			if byear and bperiod then
				pdate =  byear &"/"& bperiod &"/1"
			else
				pdate = dateadd("m", -1, now())
			end if
			vSql = "exec [dbo].[AutoProcessingVacateList] '"&pdate&"', '"&building&"'"
			mSql = "exec [dbo].AutoProcessingMoveInList '"&pdate&"', '"&building&"'"
			'response.write vsql & "</br>"
			'response.write msql & "</br>"
			'response.end
			rst5.open vSql, cnn1
			%> <table width="100%"><tr><td width="50%"> <%
			if not rst5.eof then
				%>
				Moved Out:
				<table><tr><td>Tenant</td><td>Vacate Date</td><td>Reading for the Day</td></tr>
				<%
				
				do until rst5.eof
					%>
					<tr><td><%= rst5("tenant_name") %></td><td><%= rst5("vacate_date") %></td><td><%= formatnumber(rst5("reading"),2) %></td></tr>
					<%
					rst5.movenext
				loop
				%> </table></br> <%
			end if
			rst5.close
			%> </td> <td width="50%"> <%
			rst5.open mSql, cnn1
			if not rst5.eof then
				%>
				Moved In:
				<table><tr><td>Tenant</td><td>Move In Date</td><td>Reading for the Day</td></tr>
				<%
				
				do until rst5.eof
					%>
					<tr><td><%= rst5("tenant_name") %></td><td><%= rst5("move_in_date") %></td><td><%= formatnumber(rst5("reading"),2) %></td></tr>
					<%
					rst5.movenext
				loop
				%> </table> <%
				
			end if
			rst5.close
			%> </td></tr></table> <%
		'end if
	%>
	
	<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#ffffff"><%
		if trim(bperiod)<>"0" and trim(bperiod)<>"" then%>
			<tr>
				<td width="12">&nbsp;</td>
				<td class="standard" colspan="4">Unentered meters for period <%=bperiod%> are listed below</td>
			</tr>	<%
		else%>
			<tr>
				<td width="12">&nbsp;</td>
				<td class="standard" colspan="4">All online meters are listed below
				</td>
			</tr><%
		end if%>
	</table>

	<div id="meterlist" style="overflow:auto; width:97%;height:200px;border:1px solid #cccccc;">
		<table border=0 cellpadding="3" cellspacing="1" width="100%">
			<tr bgcolor="#dddddd">
				<td width="120"><span class="standard"><b>Meter</b></span></td>
				<td width="220"><span class="standard"><b>Location</b></span></td>
				<td width="120"><span class="standard"><b>Tenant Number</b></span></td>
				<td><span class="standard"><b>Tenant Name</b></span></td>
				<td><span class="standard"><b>Date Last Read</b></span></td>
				
					<td><span class="standard"><b>MeterNum</b></span></td>
				
				<td width="18">&nbsp;</td>
			</tr>
			<%
			dim meterfilter
			if not(trim(bperiod)="0" or trim(bperiod)="") then 
				meterfilter = " and meterid not in (SELECT meterid FROM consumption WHERE ((BillYear="&byear&" and Billperiod="&bperiod&"))) and meterid not in (SELECT meterid FROM peakdemand WHERE BillYear="&byear&" and Billperiod="&bperiod&")"
			end if
			sql = "SELECT meternum, location, tenantnum, billingname, manualmultiplier, demandmultiplier, meterid, datelastread FROM meters m INNER JOIN tblleasesutilityprices lup on lup.leaseutilityid=m.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid WHERE m.online=1 and  m.BldgNum='"&building&"' "&meterFilter&" ORDER BY billingname, meternum"
             
			rst1.open sql, cnn1
			do until rst1.eof	%>
				<tr valign="top" onMouseOver="this.style.backgroundColor='lightgreen'" onMouseOut="this.style.backgroundColor='white'" onClick="javascript:loadmeter('<%=rst1("meterid")%>');">
					<td width="120"><%=rst1("meternum")%></td>
					<td width="220"><%=rst1("location")%></td><%
					'response.write "<td width=""10%"">"&rst1("manualmultiplier")&"</td>"
					'response.write "<td width=""10%"">"&rst1("demandmultiplier")&"</td>"%>
					<td width="120"><%=rst1("tenantnum")%></td>
					<td><%=rst1("billingname")%></td>
					<td><%=rst1("datelastread")%></td>
					
					<td><span class="standard"><b><%= rst1("meterid") %></b></span></td>
					
				</tr><%
				rst1.movenext
			loop
			rst1.close
			%>
		</table>
	</div>		<%
end if

if trim(meterid)<>"" then%>
</td>
</table>
<br>
<table border=0 cellspacing="0" cellpadding="3" width="97%">
	<tr valign="top">
		<td width="12">&nbsp;</td>
		<td>
			<table border=0 cellpadding="3" cellspacing="0" width="100%">
				<tr bgcolor="#6699cc">
					<td><span class="standardheader"><b>Meter: <%=Wmeternum%>&nbsp;(<%=meterUtype%>)&nbsp;in&nbsp;<%=billingname%></b></span></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td style="padding-right:4px;padding-top:2px;">Consumption Multiplier:</td><td><%if mmultiplier<>"" then%><%=mmultiplier%><%else%>1.0<%end if%></td>
				</tr>
				<tr>
					<td style="padding-right:4px;padding-top:2px;">Demand Multiplier:</td><td><%if dmultiplier<>"" then%><%=dmultiplier%><%else%>1.0<%end if%></td>
				</tr>
				<tr>
					<td style="padding-right:4px;padding-top:2px;">Date Last Read:</td>
					<%
					dim someDate
					rst1.open "SELECT datelastread FROM meters WHERE meterid="&meterid, cnn1
					if not rst1.eof then someDate = rst1("datelastread")
					rst1.close
					%>
					<td>
						<input name="datelastread" type="text" size="9" value="<%=someDate%>">
						<%if not(isBuildingOff(building)) then%>
						<input type="submit" name="action" value="Update Date">
						<input type="submit" name="action" value="Update All Dates">
						<%end if%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top">
		<td></td>
		<td>
<%
printHeaderRow()
dim where, showExtended
if trim(bperiod)<>"" and trim(byear)<>"" and trim(bperiod)<>"0" then where = ""'" and (cc.billyear<"&byear&" or (cc.billyear="&byear&" and cc.billperiod<="&bperiod&")) "
dim sql
'1/18/2008 N.Ambo removed "top 60" in statement so that all revords can show as far back as 2000 since insertions still need to be made
'sql = "SELECT top 60 isnull(posted,0) as posted, rawprev, c.estimated as estimatedc, c.usernote as usernotec, p.estimated as estimatedp, p.usernote as usernotep, b.billperiod, b.billyear,c.rawcurrent, c.rawprevious, c.rawused,isnull(c.rawcurrentoff,0) as rawcurrentoff, isnull(c.rawpreviousoff,0) as rawpreviousoff, isnull(c.rawusedoff,0) as rawusedoff,isnull(c.rawcurrentint,0) as rawcurrentint, isnull(c.rawpreviousint,0) as rawpreviousint, isnull(c.rawusedint,0) as rawusedint, p.rawdemand, c.rawonpeak, c.rawoffpeak, c.rawintpeak, p.datepeak, b.datestart, datepeak_off, p.datepeak_off, p.rawdemand_off, p.rawprev_off, p.datepeak_int, p.rawdemand_int, p.rawprev_int FROM billyrperiod b LEFT JOIN consumption c ON b.billyear=c.billyear and b.billperiod=c.billperiod and c.meterid="&meterid&" LEFT JOIN peakdemand p ON p.billyear=c.billyear and p.billperiod=c.billperiod and p.meterid=c.meterid LEFT JOIN meters m ON m.meterid=c.meterid LEFT JOIN tblleasesutilityprices lup ON lup.leaseutilityid=m.leaseutilityid LEFT JOIN tblbillbyperiod bbp ON b.ypid=bbp.ypid and bbp.leaseutilityid=lup.leaseutilityid and bbp.reject=0 WHERE b.bldgnum='"&building&"' and b.utility=(SELECT lup.utility FROM tblleasesutilityprices lup, meters m WHERE m.leaseutilityid=lup.leaseutilityid and m.meterid="&meterid&") and b.datestart<'"&dateadd("m",1,Date())&"'  ORDER BY b.billyear desc, b.billperiod desc"
sql = "SELECT isnull(posted,0) as posted, rawprev, c.estimated as estimatedc, c.usernote as usernotec, p.estimated as estimatedp, p.usernote as usernotep, b.billperiod, b.billyear,c.rawcurrent, c.rawprevious, c.rawused,isnull(c.rawcurrentoff,0) as rawcurrentoff, isnull(c.rawpreviousoff,0) as rawpreviousoff, isnull(c.rawusedoff,0) as rawusedoff,isnull(c.rawcurrentint,0) as rawcurrentint, isnull(c.rawpreviousint,0) as rawpreviousint, isnull(c.rawusedint,0) as rawusedint, p.rawdemand, c.rawonpeak, c.rawoffpeak, c.rawintpeak, p.datepeak, b.datestart, datepeak_off, p.datepeak_off, p.rawdemand_off, p.rawprev_off, p.datepeak_int, p.rawdemand_int, p.rawprev_int FROM billyrperiod b LEFT JOIN consumption c ON b.billyear=c.billyear and b.billperiod=c.billperiod and c.meterid="&meterid&" LEFT JOIN peakdemand p ON p.billyear=c.billyear and p.billperiod=c.billperiod and p.meterid=c.meterid LEFT JOIN meters m ON m.meterid=c.meterid LEFT JOIN tblleasesutilityprices lup ON lup.leaseutilityid=m.leaseutilityid LEFT JOIN tblbillbyperiod bbp ON b.ypid=bbp.ypid and bbp.leaseutilityid=lup.leaseutilityid and bbp.reject=0 WHERE b.bldgnum='"&building&"' and b.utility=(SELECT lup.utility FROM tblleasesutilityprices lup, meters m WHERE m.leaseutilityid=lup.leaseutilityid and m.meterid="&meterid&") and b.datestart<'"&dateadd("m",1,Date())&"'  ORDER BY b.billyear desc, b.billperiod desc"
rst1.open sql, cnn1
'response.write sql
'response.end
if rst1.eof then response.write "No entries found."
dim latest, rowcolor
do until rst1.eof
'response.write "fddgdg" 'bperiod &"," & rst1("billperiod")
'response.end
if trim(bperiod)=trim(rst1("billperiod")) and trim(byear)=trim(rst1("billyear")) then 'latest = 1 else latest=0
latest = 1
editPane()
else
rowcolor = "FFFFFF"
latest = 0
%>
<tr bgcolor="#<%=rowcolor%>" valign="top">
<td><%=rst1("billyear")%></td>
<td><%=rst1("billperiod")%></td>
<%if isnull(rst1("rawcurrent")) then %>
<td colspan="8">No Readings Inserted</td>
<td><%if not(isbuildingOff(building)) then%><input type="button" value="Insert" onClick="editReadings(<%=rst1("billperiod")%>,<%=rst1("billyear")%>)" style="width:35px"><%end if%></td>
<%else%>
<td><%=makeInput("rawcurrent", rst1("rawcurrent"))%><%if cdbl(rst1("rawcurrentoff"))<>0 then%>/<%=makeInput("rawcurrentoff", rst1("rawcurrentoff"))%><%end if%><%if cdbl(rst1("rawcurrentint"))<>0 then%>/<%=makeInput("rawcurrentint", rst1("rawcurrentint"))%><%end if%></td>
<td><%=makeInput("rawprevious", rst1("rawprevious"))%><%if cdbl(rst1("rawpreviousoff"))<>0 then%>/<%=makeInput("rawpreviousoff", rst1("rawpreviousoff"))%><%end if%><%if cdbl(rst1("rawpreviousint"))<>0 then%>/<%=makeInput("rawpreviousint", rst1("rawpreviousint"))%><%end if%></td>
<td><%=makeInput("rawused", rst1("rawused"))%><%if cdbl(rst1("rawusedoff"))<>0 then%>/<%=makeInput("rawusedoff", rst1("rawusedoff"))%><%end if%><%if cdbl(rst1("rawusedint"))<>0 then%>/<%=makeInput("rawusedint", rst1("rawusedint"))%><%end if%></td>
<td><%=makeInput("estimatedc", rst1("estimatedc"))%></td>
<td><%=makeInput("usernotec", rst1("usernotec"))%></td>
<td><%=makeInput("rawdemand", rst1("rawdemand"))%></td>
<td><%=makeInput("estimatedp", rst1("estimatedp"))%></td>
<td><%=makeInput("usernotep", rst1("usernotep"))%></td>
<td align="right"><%if lcase(rst1("posted"))="false" then%><%if not(isbuildingOff(building)) then%><input type="button" value="Edit" onClick="editReadings(<%=rst1("billperiod")%>,<%=rst1("billyear")%>)" style="width:35px"><%end if%><%else%>Posted<%end if%></td>
<%end if%>
</tr>
<%
end if
rst1.movenext
loop
rst1.close
%>
</table>

</td></tr>

</table>
<%end if%>
<input type="hidden" name="meterid" value="<%=meterid%>">
<input type="hidden" name="scroll" value="0">
<br><br>
</form>
</body>
</html>
<%
function makeInput(name, value)'
	if isnull(value) then value = ""
	if latest=1 then
		if name="usernotec" or name="usernotep" then
			makeInput = "<textarea name="""&name&""" cols="
			if name="usernotec" then makeInput = makeInput & "90" else makeInput = makeInput & "40"
			makeInput = makeInput & " rows=4>"&value&"</textarea><br>"
		elseif name="estimatedc" or name="estimatedp" then
			dim checked
			if value="True" then checked = " CHECKED"
			makeInput = "<input name="""&name&""" type=""checkbox"" value=""1"""&checked&">"
		else
			makeInput = "<input name="""&name&""" type=""text"" size=""10"" value="""&value&""">"
		end if
	elseif name="estimatedc" or name="estimatedp" then
		if value="True" then 
			makeInput = "Yes" 
		else 
			makeInput = "No" 
		end if
	else
		makeInput = value
	end if
end function

function printHeaderRow()
  response.write "<br>"
  response.write "<table border=0 cellpadding=""4"" cellspacing=""1"" bgcolor=""#cccccc"" width=""100%"">"
  response.write "<tr bgcolor=""#eeeeee"" valign=""top"" style=""font-weight:bold;"">"
  response.write "  <td width=""5%"" style=""border-top:1px solid #eeeeee;"">Year</td>"
  response.write "  <td width=""5%"" style=""border-top:1px solid #eeeeee;"">Period</td>"
  response.write "  <td width=""5%"" style=""border-top:1px solid #eeeeee;"">Reading</td>"
  response.write "  <td width=""5%"" style=""border-top:1px solid #eeeeee;"">Previous</td>"
  response.write "  <td width=""5%"" style=""border-top:1px solid #eeeeee;"">Usage</td>"
  response.write "  <td width=""4%"" style=""border-top:1px solid #eeeeee;"">Est.</td>"
  response.write "  <td width=""31%"" style=""border-top:1px solid #eeeeee;"">Note</td>"
  response.write "  <td width=""5%"" style=""border-top:1px solid #eeeeee;"">Raw Demand</td>"
  response.write "  <td width=""4%"" style=""border-top:1px solid #eeeeee;"">Est.</td>"
  response.write "  <td width=""31%"" style=""border-top:1px solid #eeeeee;"">Note</td>"
  response.write "  <td style=""border-top:1px solid #eeeeee;"">&nbsp;</td>"
  response.write "</tr>"
end function


sub editPane()
dim   notelenth, isEntered, estimatedc, rawonpeak, rawcurrent, usernotec, rawoffpeak, rawprevious, rawintpeak, rawused, estimatedp, datepeak, rawdemand, usernotep, rawprev, rawcurrentoff, rawcurrentint, rawpreviousoff, rawpreviousint, rawusedoff, rawusedint, datepeak_off, rawdemand_off, rawprev_off, datepeak_int, rawdemand_int, rawprev_int
if isnull(rst1("rawprevious")) then 
  isEntered = false
  dim sql2
  sql2 = "SELECT top 1 p.rawprev, c.estimated as estimatedc, c.usernote as usernotec, p.estimated as estimatedp, p.usernote as usernotep, c.rawcurrent, c.rawprevious, c.rawused,c.rawcurrentoff, c.rawpreviousoff, c.rawusedoff,c.rawcurrentint, c.rawpreviousint, c.rawusedint, p.rawdemand, c.rawonpeak, c.rawoffpeak, c.rawintpeak, p.datepeak, m.extusg FROM consumption c, peakdemand p, meters m, tblleasesutilityprices lup, billyrperiod b WHERE b.billperiod=c.billperiod and b.billyear=c.billyear and lup.leaseutilityid=m.leaseutilityid and c.meterid=m.meterid and c.billperiod=p.billperiod and c.billyear=p.billyear  and b.bldgnum=m.bldgnum and b.utility=lup.utility and c.meterid=p.meterid and c.meterid="&meterid&" and b.datestart<'"&rst1("datestart")&"' ORDER BY b.datestart desc"
 ' response.write sql2
 ' response.end
  rst2.open sql2, cnn1
  if not rst2.eof then
    estimatedc 		= "0"
    rawonpeak 		= "0"
    rawcurrent 		= rst2("rawcurrent")
    rawcurrentoff 	= rst2("rawcurrentoff")
    rawcurrentint 	= rst2("rawcurrentint")		
    usernotec 		= ""
    rawoffpeak 		= "0"
    rawprevious 	= rst2("rawcurrent")
    rawpreviousoff 	= rst2("rawcurrentoff")
    rawpreviousint 	= rst2("rawcurrentint")
    rawintpeak 		= "0"
    rawused 		= "0"
    estimatedp 		= "0"
    datepeak 		= rst1("datestart")
    rawdemand 		= "0"
    usernotep 		= ""
    rawprev 		= rst2("rawdemand")
	datepeak_off	= rst1("datestart")
	rawdemand_off	= "0"
	rawprev_off		= ""
	datepeak_int	= rst1("datestart")
	rawdemand_int	= "0"
	rawprev_int		= ""
  end if
  rst2.close
else 

  isEntered 	= true
  estimatedc 	= rst1("estimatedc")
  rawonpeak 	= rst1("rawonpeak")
  rawcurrent 	= rst1("rawcurrent")
  rawcurrentoff = rst1("rawcurrentoff")
  rawcurrentint = rst1("rawcurrentint")
  usernotec 	= rst1("usernotec")
  rawoffpeak 	= rst1("rawoffpeak")
  rawprevious 	= rst1("rawprevious")
  rawpreviousoff= rst1("rawpreviousoff")
  rawpreviousint= rst1("rawpreviousint")
  rawintpeak 	= rst1("rawintpeak")
  rawused 		= rst1("rawused")
  rawusedoff 	= rst1("rawusedoff")
  rawusedint 	= rst1("rawusedint")
  estimatedp 	= rst1("estimatedp")
  datepeak 		= rst1("datepeak")
  rawdemand 	= rst1("rawdemand")
  rawprev 		= rst1("rawprev")
  usernotep 	= rst1("usernotep")
  datepeak_off	= rst1("datepeak_off")
  rawdemand_off	= rst1("rawdemand_off")
  rawprev_off	= rst1("rawprev_off")
  datepeak_int	= rst1("datepeak_int")
  rawdemand_int	= rst1("rawdemand_int")
  rawprev_int	= rst1("rawprev_int")
end if

dim EUshow
if extusg then EUshow = "inline" else EUshow = "none"
%>
<tr><td colspan="11" bgcolor="white">
<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#cccccc">
<tr bgcolor="#eeeeee">
  <td colspan="2" style="padding:3px;"><b>Period <%=bperiod%> of <%=byear%></b></td>
  <td colspan="2" style="padding:3px;" align="right">
  <%if isEntered then%>
  <input name="action" value="Update" type="submit">&nbsp;<input name="action" value="Delete" type="submit">
  <%else%>
  <input name="action" value="Save" type="submit">
  

  <%end if%>
  &nbsp;<input name="action" value="Cancel" type="button" onClick="editReadings(0,<%=byear%>);"><input type="hidden" name="workingperiod" value="<%=bperiod%>"><input type="hidden" name="workingyear" value="<%=byear%>"></td></tr>
<tr valign="top">
<td>
<!-- begin consumption -->
  <table border=0 cellpadding="3" cellspacing="0">
  <tr><td colspan="7"><b>Consumption</b><%=makeInput("estimatedc", estimatedc)%>&nbsp;Estimated <br><font size=1><label onClick="showextended()" style="cursor:hand;">[<font color="#0000FF"><u>show/hide extended entry</u></font>]</label></font></td></tr>
  <tr valign="top" style="font-weight:bold;">
      <td colspan="2">Peak<span id="extendedUsage" style="display:<%=EUshow%>;"><br><font size=1>(only ON PEAK usage below)</font></span></td>
      <td rowspan="4" width="12">&nbsp;</td>
      <td colspan="2">Usage Reading <span id="extendedUsage" style="display:<%=EUshow%>;"><br>(ON PEAK)</span></td>
      <td colspan="2" valign="bottom" id="extendedUsage" style="display:<%=EUshow%>;">(OFF PEAK)</td>
      <td colspan="2" valign="bottom" id="extendedUsage" style="display:<%=EUshow%>;">(INT PEAK)</td> 
  </tr>
  <tr><td>Raw On Peak:</td>
      <td><%=makeInput("rawonpeak", rawonpeak)%></td>
      <td>Current:</td>
      <td><%=makeInput("rawcurrent", rawcurrent)%></td>
      <td id="extendedUsage" style="display:<%=EUshow%>;">Current:</td>
      <td id="extendedUsage" style="display:<%=EUshow%>;"><%=makeInput("rawcurrentoff", rawcurrentoff)%></td>
      <td id="extendedUsage" style="display:<%=EUshow%>;">Current:</td>
      <td id="extendedUsage" style="display:<%=EUshow%>;"><%=makeInput("rawcurrentint", rawcurrentint)%></td>
  </tr>
  <tr><td>Raw Off Peak:</td>
      <td><%=makeInput("rawoffpeak", rawoffpeak)%></td>
      <td>Previous:</td>
      <td><%=makeInput("rawprevious", rawprevious)%></td>
      <td id="extendedUsage" style="display:<%=EUshow%>;">Previous:</td>
      <td id="extendedUsage" style="display:<%=EUshow%>;"><%=makeInput("rawpreviousoff", rawpreviousoff)%></td>
      <td id="extendedUsage" style="display:<%=EUshow%>;">Previous:</td>
      <td id="extendedUsage" style="display:<%=EUshow%>;"><%=makeInput("rawpreviousint", rawpreviousint)%></td>	  
  </tr>
  <tr><td>Raw Intermediate:</td>
      <td><%=makeInput("rawintpeak", rawintpeak)%></td>
      <td>Usage:</td>
      <td><%=makeInput("rawused", rawused)%></td>
      <td id="extendedUsage" style="display:<%=EUshow%>;">Usage:</td>
      <td id="extendedUsage" style="display:<%=EUshow%>;"><%=makeInput("rawusedoff", rawusedoff)%></td>
      <td id="extendedUsage" style="display:<%=EUshow%>;">Usage:</td>
      <td id="extendedUsage" style="display:<%=EUshow%>;"><%=makeInput("rawusedint", rawusedint)%></td>
  </tr>
  <tr style="font-weight:bold;">
  	   <td Colspan=7>Note</td>
  </tr>
  <tr>
  	   <td Colspan=7><%=makeInput("usernotec", usernotec)%>
  	    </td>
  </tr>
  </table>

  
      
 <%
 if not isnull(rst1("rawpreviousoff")) and not isnull(rst1("rawpreviousint")) then 
	 if cdbl(rst1("rawcurrentoff")) = 0 and cdbl(rst1("rawcurrentint"))=0 then
	 	showExtended = false
	 else
	 	showExtended = true
	 end if
 else
 	showExtended = false
 end if
 
 if showExtended then 
	 %>
 <script>//showextended()</script>
 <%end if%>
<!-- end consumption -->
</td>
<td width="20" style="border-right:1px solid #999999;">&nbsp;</td>
<td width="20">&nbsp;</td>
<td>
<!-- begin demand -->
  <table border="0" cellpadding="3" cellspacing="0">
  <tr><td colspan="2"><b>Demand</b><%=makeInput("estimatedp", estimatedp)%>&nbsp;Estimated</td></tr>
  <tr><td>Peak Date</td>
      <td><%=makeInput("datepeak", datepeak)%></td>
  </tr>
  <tr><td>Current:</td>
      <td><%=makeInput("rawdemand", rawdemand)%></td>
  </tr>
  <tr><td>Previous:</td>
      <td><%=makeInput("rawprev", rawprev)%></td>
  </tr>
  <tr id="extendedUsage" style="display:<%=EUshow%>;">
  	<td>Off Peak Date</td>
  	<td><%=makeInput("datepeak_off", datepeak_off)%></td>
  </tr>
  <tr id="extendedUsage" style="display:<%=EUshow%>;">
  	<td>Off Current:</td>
    <td><%=makeInput("rawdemand_off", rawdemand_off)%></td>
  </tr>
  <tr id="extendedUsage" style="display:<%=EUshow%>;">
  	<td>Off Previous:</td>
    <td><%=makeInput("rawprev_off", rawprev_off)%></td>
  </tr>
  <tr id="extendedUsage" style="display:<%=EUshow%>;">
  	<td>Int Peak Date</td>
    <td><%=makeInput("datepeak_int", datepeak_int)%></td>
  </tr>
  <tr id="extendedUsage" style="display:<%=EUshow%>;">
  	<td>Int Current:</td>
    <td><%=makeInput("rawdemand_int", rawdemand_int)%></td>
  </tr>
  <tr id="extendedUsage" style="display:<%=EUshow%>;">
  	<td>Int Previous:</td>
    <td><%=makeInput("rawprev_int", rawprev_int)%></td>
  </tr>
  <tr><td>Difference:</td>
      <td><%if rawdemand="" or rawprev="" then response.write "N/A" else response.write cdbl(rawdemand) - cdbl(rawprev) end if%></td>
  </tr>  
  <tr><td>Note</td></tr>
  <tr>
      <td colspan="3"><%=makeInput("usernotep", usernotep)%><br></td>
  </tr>
  <tr><td colspan="3"><%if mCumulativeDem="True" then response.write "<em>Cumulative Demand Meter</em>"%></td></tr>
  </table>
<!-- end demand -->
</td>
</tr>
</table>
</td></tr>
<%end sub%>
