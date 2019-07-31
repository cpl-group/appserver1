
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'12/13/2007 N.Ambo Added condition for portfolio 108; if the portfolio is 108 then unpost bills should only be available
'to users in the IT Services group; search 'unpost' (lines 213-217, 316-319)

dim pid, building, byear, ypid, lid, bperiod, utilityid, utilitydisplay, historic,billurl,logo
dim billCount,pdfLinker, logoh, logow
'pdfLinker = "209.213.207.24"
pdfLinker = "pdfmaker.genergyonline.com"
'pdfLinker = "10.0.7.78"
billCount = -1
pid = request("pid")
building = request("building")
if instr(request("bperiod"),"/")>0 then
	byear = split(request("bperiod"),"/")(1)
	bperiod = split(request("bperiod"),"/")(0)
else
	byear = request("byear")
	bperiod = request("bperiod")
end if
lid = request("lid")
utilityid = request("utilityid")
if lcase(request("historic"))="true" then historic=true else historic=false

if utilityid = "" then utilityid = 0
if byear = "" then byear = 0
if bperiod = "" then bperiod = 0

dim rst1, cnn1, posted, sql
set rst1 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getConnect(pid,building,"Billing")

dim super
if allowgroups("GY_Supervisors_ES") then
	super=true
else
	super=false
end if

dim billlink
dim maxPageCount

if trim(building)<>"" then
	rst1.open "SELECT location, b.bldgnum,billurl,logo, logoh, logow FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&building&"'", cnn1
	'response.Write("SELECT location, b.bldgnum,billurl FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&building&"'")
	'response.Write("<br>" & cnn1)
	'response.End()
	if not rst1.eof then 
		billlink = rst1("location")
		billurl = rst1("billurl")
		logo = rst1("logo")
		logoh = rst1("logoh")
		logow = rst1("logow")		
	end if 
	rst1.close

	if pid = "108" then
		sql = "SELECT b.id as billid FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
		if isnumeric(utilityid) then sql = sql & " and utility="&utilityid
		
		dim billid
			
		dim blgdrset
		Set blgdrset = Server.CreateObject("ADODB.recordset")
		blgdrset.open sql, cnn1
		maxPageCount = 0
		do until blgdrset.eof
			Dim meterinfo
			Set meterinfo = Server.CreateObject("ADODB.recordset")
			billid = blgdrset("billid")
			meterinfo.open "select count(*) as metercount from tblmetersbyperiod tm,buildings b,meters m where tm.bldgnum =b.bldgnum and tm.meternum=m.meternum and b.bldgnum = m.bldgnum and bill_id="&billid, cnn1
			
			dim tempMaxPageCount
			tempMaxPageCount = meterinfo("metercount") \ 40 + 1
			
			if meterinfo("metercount") > 5 then
				tempMaxPageCount = tempMaxPageCount + 1
			end if
			
			if tempMaxPageCount > maxPageCount then
				maxPageCount = tempMaxPageCount
			end if
		
			meterinfo.Close()
			blgdrset.movenext
		loop
		blgdrset.Close()
	end if
end if
%>
<html>
<head>
<title>Bill Validation</title>

<style type="text/css">
INPUT#f9 {
	font-size:9
}
</style>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td width="48%" bgcolor="#6699CC"><span class="standardheader">Bill Processor</span></td>
    <td width="52%" align="right" bgcolor="#6699CC">
	<%if building<> "" then %><select name="select" onChange="JumpTo(this.value)">
        <option value="#" selected>Jump to...</option>
        <option value="../validation/re_index.asp">Review Edit</option>
        <option value="/genergy2/setup/buildingedit.asp">Building Setup</option>
        <option value="/genergy2/manualentry/entry_select.asp">Manual Entry</option>
        <option value="/genergy2/billentry/entry.asp">Utility Bill Entry</option>
        <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem 
        Report</option>
      </select>
	  <%end if%>
    </td>
  </tr>
  <form name="form1" action="billprocess.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td> 
              <%if allowGroups("Genergy Users") then%>
              <select name="pid" onChange="loadportfolio()">
                <option value="">Select Portfolio</option>
                <%rst1.open "SELECT distinct id, name FROM portfolio p ORDER BY name", getConnect(0,0,"dbCore")
            do until rst1.eof%>
                <option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then response.write " SELECTED"%>><%=rst1("name")%></option>
                <%	rst1.movenext
            loop
            rst1.close%>
              </select> 
              <%elseif isnumeric(pid) then%>
              <input type="hidden" name="pid" value="<%=pid%>"> 
              <%end if%>
            </td>
            <%if trim(pid)<>"" then%>
            <td> <select name="building" onChange="loadbuilding()">
                <option value="">Select Building</option>
                <%
      rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
      do until rst1.eof%>
                <option <%if isBuildingOff(rst1("Bldgnum")) then%>class="grayout"<%end if%> value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>><%=rst1("strt")%>, 
                <%=trim(rst1("Bldgnum"))%></option>
                <%	rst1.movenext
      loop
      rst1.close
      %>
              </select> </td>
            <%end if
    if trim(building)<>"" then%>
            <td> <select name="utilityid" onChange="loadutility()">
                <option value="">Select Utility</option>
                <%rst1.open "SELECT DISTINCT byp.Utility as utilityid, u.Utilitydisplay FROM BillYrPeriod byp inner join dbo.tblutility u ON byp.Utility = u.utilityid WHERE (BldgNum = '" & trim(building) &"')", getLocalConnect(building)
      do until rst1.eof
        %>
                <option value="<%=rst1("utilityid")%>"<%if trim(rst1("utilityid"))=trim(utilityid) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option>
                <%
        if trim(rst1("utilityid"))=trim(utilityid) then utilitydisplay = rst1("utilitydisplay")
        rst1.movenext
      loop
      rst1.close
      %>
              </select> </td>
            <%end if
    if trim(utilityid)<>0 then%>
            <td> <select name="bperiod">
                <option value="">Select Bill Period</option>
                <%
				sql = "SELECT distinct cast(billperiod as varchar)+'/'+billyear as periodyear, billyear, billperiod FROM billyrperiod WHERE "
				if not(historic) then sql = sql & "billyear>=year(getdate())-1 and "
				sql = sql & "bldgnum='"&building&"' and utility="&utilityid&" order by billyear, billperiod"
				rst1.open sql, getLocalConnect(building)
      do until rst1.eof
        %>
                <option value="<%=rst1("periodyear")%>"<%if trim(rst1("periodyear"))=trim(bperiod&"/"&byear) or (bperiod="0" and month(dateadd("m",-1,now))&"/"&year(dateadd("m",-1,now))=rst1("periodyear")) then response.write " SELECTED"%>><%=rst1("periodyear")%></option>
                <%
        rst1.movenext
      loop
      rst1.close
      %>
              </select> </td>
            <td> <input type="button" name="view" value="View" onClick="loadperiod()"> 
            </td>

		<% else %>
		<td><input name="bperiod" type="hidden" value="0/0"></td>
        <%end if%>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td colspan=3><%if trim(utilityid)<>0 then%><label style="border:1px solid #eeeeee; color:black; font-weight: bold; border-bottom-style: solid;cursor:hand;font-size:xx-small" onClick="
	document.location='processor_select.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&bperiod=<%=bperiod%>&historic=<%if historic then%>false<%else%>true<%end if%>'
	" onMouseOver="this.style.borderColor='black';" onMouseOut="this.style.borderColor='#eeeeee';" type="" src="" value="New Job">&nbsp;Click to <%if historic then%>Hide<%else%>Show<%end if%>&nbsp;Historical&nbsp;Periods&nbsp;</label><%end if%></td>
          </tr>
        </table></td>
    </tr>
    <%if trim(bperiod)<>0 then%>
    <tr bgcolor="#dddddd"> 
      <td colspan="2"> 
	  	<select name="actions" onChange="warningMsg(this);">
			<%if not(isBuildingOff(building)) then%>
				<option value="Produce Bills For Current Period">Produce Bills For Current Period</option>
				<option value="Produce Partial Bills">Produce Partial Bills</option>
				<option value="Post Bills">Post Bills</option>
		         <%if pid <> "108" and allowGroups("IT Services,GenergyCorporateExec,GY_Supervisors_ES") Then%>
					<option value="Unpost Bills">Unpost Bills</option>
	        	<%elseif pid = "108" and allowGroups("IT Services")Then%>  
					<option value="Unpost Bills">Unpost Bills</option>				
	        	<%end if%>
			<%end if%>
			<option value="View All Detailed Bills">View All Detailed Bills</option>
			<option value="View [utility only] Bills">View <%=utilitydisplay%> Bills</option>
			<option value="Bill Summary">Bill Summary</option>
			<option value="Excel Bill Summary">Excel Bill Summary</option>
			<% if cint(pid) = 6 Then %>
				<option value="EmailBills">Email Tenant Bills</option>
			<% end if %>
			<option value="View All Tenant Summaries">View All Tenant Summaries</option>
			<%if not(isBuildingOff(building)) then%>
				<option value="Data Files">Data Files</option>
			<%end if%>
		</select>
		<input type="button" onClick="billAction()"  value="Run" name="actionbutton">
		<div id="noteDiv" style="display:none;"><span id="noteSpan" value="empty"></span></div>
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
	  &nbsp;
		  <table cellpadding="0" cellspacing="0" width="100%"><tr><td>
        <%
        ' Added by Tarun : 06.25.2008 (ERI Bill Processing)
        'if pid = "122" then
			sql = "SELECT " & _
					"(SELECT count(distinct lup.leaseutilityid) FROM tblleasesutilityprices lup, tblleases l, meters m " & _
					" WHERE l.billingid=lup.billingid 	and lup.leaseutilityid=m.leaseutilityid " & _
					" and m.nobill=0 and online=1  and l.leaseexpired=0 and l.bldgnum='"&building& "'" & _
					" and lup.utility=2) as billsneeded, (SELECT count(*) FROM tblEriBills WHERE reject=0 " & _
					" and leaseutilityid in (SELECT leaseutilityid 	FROM tblleasesutilityprices lup, tblleases l " & _
					" WHERE leaseexpired=0 and l.billingid=lup.billingid and l.bldgnum='"&building& "')" & _ 
					" and totalamount is not null and billperiod="&bperiod&"and billyear="&byear&"and utility="&utilityid&") as billsprocessed," & _ 
					"(SELECT count(*) FROM tblEriBills WHERE reject=0 and totalamount is null and buildingnumber='"&building& "'" & _ 
					" and billperiod="&bperiod&" and billyear="&byear&" and utility="&utilityid&") as billserrored "
		'else
			sql = _
			"SELECT "& _
			"(SELECT count(distinct lup.leaseutilityid) FROM tblleasesutilityprices lup, tblleases l, meters m WHERE l.billingid=lup.billingid and lup.leaseutilityid=m.leaseutilityid and m.nobill=0 and online=1 and l.leaseexpired=0 and l.bldgnum='"&building&"' and lup.utility="&utilityid&") as billsneeded, "&_
			"(SELECT count(*) FROM tblbillbyperiod WHERE reject=0 and leaseutilityid in (SELECT leaseutilityid FROM tblleasesutilityprices lup, tblleases l WHERE leaseexpired=0 and l.billingid=lup.billingid and l.bldgnum='"&building&"') and totalamt is not null and billperiod="&bperiod&" and billyear="&byear&" and utility="&utilityid&") as billsprocessed, "&_
			"(SELECT count(*) FROM tblbillbyperiod WHERE reject=0 and totalamt is null and bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear&" and utility="&utilityid&") as billserrored"
		'end if
		'response.write sql
    rst1.open sql, getLocalConnect(building)
    if not rst1.eof then
	billCount = rst1("billsneeded")	%>
        <span class="notetext"><%=rst1("billsprocessed")%> bill<%if cint(rst1("billsprocessed"))>1 then response.write "s"%> processed out of <%=rst1("billsneeded")%> bill<%if cint(rst1("billsneeded"))>1 then response.write "s"%> needed.
		<%if cint(rst1("billserrored"))>0 then%>Found <%=rst1("billserrored")%> errored bill<%if cint(rst1("billserrored"))>1 then response.write "s"%>!<%end if%>
		</span><br> 
	<%end if
    rst1.close    
    %></td><td align="right">
		<a href="#" onClick="window.open('viewDeletedBills.asp?<%=request.querystring%>','','width=600,height=300,toolbox=no,scrollbars=yes')">View Deleted Bills</a>
		</td></tr></table>
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <%
    dim someSql
    ' Added by Tarun : 06.25.2008 (ERI Bill Processing)
    'If pid = "122" then 
		somesql = "SELECT case when b.partial=1 then left(datestart,11)+' - '+left(dateend,11) else 'NA' end as partialdates, " & _
					"b.partial, left(b.postdate,11) as postdate, b.ypid, isnull(ma.metersaccept,0) as maccept, " & _
					"isnull(md.metersdata,0) as mdata, isnull(bm.billedmeters,0) as bmeter, b.TenantNumber as TenantNumber, " & _
					"b.posted as post, 	u.utilitydisplay as utilitytype, lup.leaseutilityid, b.billperiod, b.billingname, " & _
					" lup.utility, lup.billingid, b.totalamount, CASE WHEN b.totalamount is null THEN 1 ELSE 0 END as errored, " & _
					" tx.tcount FROM tblERIBills b " & _
				" INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=b.leaseutilityid " & _
				" INNER JOIN tblleases l ON lup.billingid=l.billingid " & _
				" INNER JOIN tblutility u ON u.utilityid=lup.utility  " & _
				" LEFT JOIN (SELECT count(*) as metersaccept, leaseutilityid, billperiod, billyear FROM consumption cc, meters mm " & _
				" WHERE mm.meterid=cc.meterid and svalidate=1 GROUP BY leaseutilityid, billperiod, billyear) ma " & _
			    " ON lup.leaseutilityid=ma.leaseutilityid and b.billperiod=ma.billperiod and b.billyear=ma.billyear " & _
				" LEFT JOIN (SELECT count(*) as metersdata, leaseutilityid, billperiod, billyear FROM consumption cc, meters mm " & _
				" WHERE mm.meterid=cc.meterid and online=1 	GROUP BY leaseutilityid, billperiod, billyear) md " & _
				" ON lup.leaseutilityid=md.leaseutilityid and b.billperiod=md.billperiod and b.billyear=md.billyear " & _
				" LEFT JOIN (SELECT count(*) as billedmeters, ypid, leaseutilityid, bill_id FROM tblmetersbyperiod " & _
				" GROUP BY ypid, leaseutilityid, bill_id) bm ON b.EriBillid=bm.bill_id left join (select ticketfor as billingid, isnull(count(*),0) as tcount " & _
				" from ["& Application("CoreIP") & "].dbCore.dbo.tickets where ticketfortype = 'tid' and closed=0 and billyear=" & byear & _
				" and billperiod ="&bperiod&" group by ticketfor) tx on tx.billingid= ltrim(convert(varchar(10),lup.billingid)) " & _
				" WHERE b.reject = 0 and b.buildingnumber='"&building&"' and b.billperiod="&bperiod&" and b.billyear="&byear  & _
				" and lup.utility="&utilityid&" ORDER BY tenantname "
    'Else
		somesql = "SELECT case when b.partial=1 then left(datestart,11)+' - '+left(dateend,11) else 'NA' end as partialdates, b.partial, left(b.postdate,11) as postdate, b.ypid, isnull(ma.metersaccept,0) as maccept, isnull(md.metersdata,0) as mdata, isnull(bm.billedmeters,0) as bmeter, b.TenantNum as TenantNumber, b.posted as post, u.utilitydisplay as utilitytype, lup.leaseutilityid, b.billperiod, b.billingname, lup.utility, lup.billingid, b.totalamt, CASE WHEN b.totalamt is null THEN 1 ELSE 0 END as errored, tx.tcount FROM tblbillbyperiod b INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=b.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid INNER JOIN tblutility u ON u.utilityid=lup.utility LEFT JOIN (SELECT count(*) as metersaccept, leaseutilityid, billperiod, billyear FROM consumption cc, meters mm WHERE mm.meterid=cc.meterid and svalidate=1 GROUP BY leaseutilityid, billperiod, billyear) ma ON lup.leaseutilityid=ma.leaseutilityid and b.billperiod=ma.billperiod and b.billyear=ma.billyear LEFT JOIN (SELECT count(*) as metersdata, leaseutilityid, billperiod, billyear FROM consumption cc, meters mm WHERE mm.meterid=cc.meterid and online=1 GROUP BY leaseutilityid, billperiod, billyear) md ON lup.leaseutilityid=md.leaseutilityid and b.billperiod=md.billperiod and b.billyear=md.billyear LEFT JOIN (SELECT count(*) as billedmeters, ypid, leaseutilityid, bill_id FROM tblmetersbyperiod GROUP BY ypid, leaseutilityid, bill_id) bm ON b.id=bm.bill_id left join (select ticketfor as billingid, isnull(count(*),0) as tcount from ["&Application("CoreIP")&"].dbCore.dbo.tickets where ticketfortype = 'tid' and closed=0 and billyear="&byear&" and billperiod ="&bperiod&" group by ticketfor) tx on tx.billingid='"&split(getBuildingIP(building),"\")(1)&"-' + ltrim(convert(varchar(10),lup.billingid)) WHERE b.reject = 0 and b.bldgnum='"&building&"' and b.billperiod="&bperiod&" and b.billyear="&byear&" and lup.utility="&utilityid&" ORDER BY tenantname"
	'end if
    posted = false
    rst1.open someSql, getLocalConnect(building)
    if not rst1.eof then bperiod=rst1("billperiod")
    %>
        <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
          <tr valign="bottom" bgcolor="#dddddd" style="font-weight:bold;"> 
            <td>View Bill</td>
            <td>View Detail Bill</td>
            <td>Tenant Summary</td>
            <td width="34%">Tenant Name</td>
            <td>Tenant Number</td>
            <td>Meters W/Data</td>
            <td>Meters Accepted</td>
            <td>Meters Billed</td>
            <td>Partial Bill Dates</td>
            <td>Post Date</td>
            <%if not(isBuildingOff(building)) then%><td>Delete<a href="#note">*</a></td><%end if%>
          </tr>
          <%
  if rst1.eof then posted = True
  Dim rowcolor
  do until rst1.eof
  	if rst1("errored") = "0" then rowcolor = "#ffffff" else rowcolor = "#FFDDDD"
    if rst1("post")="True" then posted=true%>
          <tr bgcolor="<%=rowcolor%>"> 
		  <%if rst1("errored") = "0" then%>
            <td align="center"><a href="javascript:viewBill('<%=rst1("leaseutilityid")%>', '<%=rst1("ypid")%>','false', <%=rst1("utility")%>, true);"><img src="images/pdf_bill.gif" width="21" height="22" border="0"></a><br>
              <%if allowGroups("Genergy Users") Then%>
              <a href="javascript:viewBill('<%=rst1("leaseutilityid")%>', '<%=rst1("ypid")%>','false', <%=rst1("utility")%>, false);">html</a>
              <%end if%>
            </td>
            <td align="center"><a href="javascript:viewBill('<%=rst1("leaseutilityid")%>', '<%=rst1("ypid")%>','true', <%=rst1("utility")%>, true);"><img src="images/pdf_bill.gif" width="21" height="22" border="0"></a><br>
              <%if allowGroups("Genergy Users") Then%>
              <a href="javascript:viewBill('<%=rst1("leaseutilityid")%>', '<%=rst1("ypid")%>','true', <%=rst1("utility")%>, false);">html</a>
              <%end if%>
            </td>
            <td align="center"><a href="#" onClick="window.open('/genergy2/UMreports/tenantBillSummary.asp?pid=<%=pid%>&building=<%=building%>&billingid=<%=rst1("billingid")%>&leaseutilityid=<%=rst1("leaseutilityid")%>&syear=<%=byear%>&utilityid=<%=utilityid%>','','width=800,height=600');"><img src="images/pdf_bill.gif" width="21" height="22" border="0"></a></td>
		  <%else%>
		  	<td colspan="3" align="center">Errored Bill</td>
		  <%end if%>
            <td><%=rst1("billingname")%><%if rst1("tcount") <> "0" then %> <font color="#FF0000">[<a href="#" onClick="window.open('/genergy2_intranet/itservices/ttracker/troublesearch.asp?pid=<%=pid%>&searchstring=<%=rst1("billingid")%>&action=Search&searchbox=false&accounts=True','SearchNotes','width=800,height=400, scrollbars=no')"><%=rst1("tcount")%> critical ticket(s) still open</a>]</font><%end if%></td>
            <td><%=rst1("TenantNumber")%></td>
            <td><%=rst1("mdata")%></td>
            <td><%=rst1("maccept")%></td>
            <td><%=rst1("bmeter")%></td>
			<td><%=rst1("partialdates")%></td>
			<td nowrap><%=rst1("postdate")%></td>
			<%if not(isBuildingOff(building)) then%>
	            <td>
	              <%if trim(rst1("post"))="True" then%>
		              <%if pid <> "108" and allowGroups("IT Services,GenergyCorporateExec,GY_Supervisors_ES") Then%>
		              <input type="button" name="action" value="Unpost" onClick="unpostBill('<%=rst1("leaseutilityid")%>');" id="f9">
						<%elseif pid = "108" and allowGroups("IT Services") Then%>
						<input type="button" name="action" value="Unpost" onClick="unpostBill('<%=rst1("leaseutilityid")%>');" id="f9">
		              <%else%>
		              Posted
		              <%end if%>
	              <%else%>
		              <input type="button" name="action" value="Delete Bill" onClick="deleteBill('<%=rst1("leaseutilityid")%>');" id="f9">
		              &nbsp;
		              <%if allowGroups("IT Services,GenergyCorporateExec,GY_Supervisors_ES") Then%>
		              <input type="button" name="action" value="Post" onClick="postBill('<%=rst1("leaseutilityid")%>');" id="f9">
		              <%end if%>
	              <%end if%>
	            </td>
			<%end if%>
          </tr>
          <%
    rst1.movenext
  loop
  rst1.close
  %>
        </table>
        <%if not(posted) and not(isbuildingOff(building)) then%>
        <div align="right">
          <input type="submit" value="Delete All Bills" name="action" id="f9">
        </div>
        <%end if%>
        <p align="right"><a name="note">* Only bills not yet posted can be deleted</a></p>
        <input type="hidden" name="lid" value=""> </td>
    </tr>
    <% end if %>
	<input type="hidden" name="historic" value="<%=historic%>">
  </form>
</table>
</body>
</html>
<script>

function warningMsg(tag)
{
var func = eval('document.all.noteDiv');
var note = eval('document.all.noteSpan')
var billCount = <%=billCount%>;
billCount = Math.ceil(Math.ceil(billCount/25)*3);

          if(tag.value=="View All Detailed Bills") 
     { 
             note.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Note this will approximately take " + billCount + " minutes.";
			 func.style.display = "inline"; 
           
     } 
     else if(tag.value=="View [utility only] Bills")
	 { 
         note.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Note this will approximately take " + billCount + " minutes.";
         func.style.display = "inline"; 
     } 
     else 
     { 
     func.style.display = "none"; 
     } 
 } 
     

function loadportfolio()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value;
	document.location.href=newhref;
}

function loadbuilding()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value+"&building="+frm.building.value;
	document.location.href=newhref;
}

function loadutility()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&utilityid="+frm.utilityid.value;
	document.location.href=newhref;
}

function loadperiod()
{	var frm = document.forms['form1'];
	var newhref = "processor_select.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&utilityid="+frm.utilityid.value+"&bperiod="+frm.bperiod.value+"&historic="+frm.historic.value
	document.location.href=newhref;
}

function deleteBill(itemlid){
  document.form1.lid.value=itemlid;
  document.location = "billprocess.asp?pid=<%=pid%>&building=<%=building%>&bperiod=<%=bperiod%>/<%=byear%>&utilityid=<%=utilityid%>&lid=" + itemlid +"&historic="+document.form1.historic.value+"&action=Delete Bill";
}

function postBill(itemlid){
  document.form1.lid.value=itemlid;
  document.location = "billprocess.asp?pid=<%=pid%>&building=<%=building%>&bperiod=<%=bperiod%>/<%=byear%>&utilityid=<%=utilityid%>&lid=" + itemlid +"&historic="+document.form1.historic.value+"&action=Post";
}

function unpostBill(itemlid){
  document.form1.lid.value=itemlid;
  document.location = "billprocess.asp?pid=<%=pid%>&building=<%=building%>&bperiod=<%=bperiod%>/<%=byear%>&utilityid=<%=utilityid%>&lid=" + itemlid +"&historic="+document.form1.historic.value+"&action=Unpost";
}

function viewBill(rstlid, ypid, detailed, utility, buildpdf){
	buildpdf = (buildpdf!=false?true:false);
  var url = 'loading.asp?url=<%=server.urlencode("http://pdfmaker.genergyonline.com/"&billlink&"genergy2=true&devIP="&request.ServerVariables("SERVER_NAME")&"&billurl="&billurl&"&building="&building&"&logo="&logo&"&logoh="&logoh&"&logow="&logow&"&pid="&pid&"&lid=")%>'+rstlid+'%26bperiod%3D<%=bperiod%>%26byear%3D<%=byear%>%26y%3D'+ypid+'%26ypid%3D'+ypid+'%26l%3D'+rstlid+'%26detailed%3D'+detailed+'%26utilityid%3D'+utility+'%26buildpdf%3D'+buildpdf;
  //alert(url);
  billpdf = window.open(url,'','width=600,height=500,resizable=yes');
}
function JumpTo(url){
	var frm = document.forms['form1'];
	var url = url + "?pid=<%=pid%>&bldg=<%=building%>&building=<%=building%>&utilityid=<%=utilityid%>&byear="+frm.bperiod.value.split("/")[1]+"&bperiod=" + frm.bperiod.value.split("/")[0];
	window.document.location=url;
}


function billAction(){
	var frm = document.forms[0];
	if(frm.actions.value=="View All Detailed Bills"){
		window.open('loading.asp?url=<%=server.urlencode("http://"& pdfLinker &"/"&billlink&"genergy2=true&devIP="&request.ServerVariables("SERVER_NAME")&"&billurl="&billurl&"&pid="&pid&"&logo="&logo&"&logoh="&logoh&"&logow="&logow&"&byear="&byear&"&bperiod="&bperiod&"&y=&building="&building&"&bldg="&building&"&b="&building&"&maxPageCount="&maxPageCount&"&detailed=true&billCount="&billCount)%>','','width=600,height=500,resizable=yes,scrollbars=yes');
	}else if(frm.actions.value=="View [utility only] Bills"){
		window.open('loading.asp?url=<%=server.urlencode("PdfLinks.asp?pid="&pid&"&byear="&byear&"&bperiod="&bperiod&"&utilityid="&utilityid&"&building="&building)%>','','width=600,height=500,resizable=yes,scrollbars=yes');
	//}else if(frm.actions.value=="View [utility only] Bills"){
	//	window.open('loading.asp?url=<%=server.urlencode("http://"& pdfLinker &"/"&billlink&"genergy2=true&devIP="&request.ServerVariables("SERVER_NAME")&"&billurl="&billurl&"&pid="&pid&"&logo="&logo&"&logoh="&logoh&"&logow="&logow&"&byear="&byear&"&bperiod="&bperiod&"&y=&building="&building&"&bldg="&building&"&b="&building&"&utilityid="&utilityid&"&maxPageCount="&maxPageCount&"&billCount="&billCount)%>','','width=600,height=500,resizable=yes,scrollbars=yes');
	}else if(frm.actions.value=="Bill Summary"){
		window.open('loading.asp?url=<%=server.urlencode("bill_summary.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"&y=&l="&lid&"&utilityid="&utilityid)%>','','width=800,height=500,resizable=yes,scrollbars=yes');
	}else if(frm.actions.value=="Excel Bill Summary"){
		window.open('loading.asp?url=<%=server.urlencode("excelBillSummary.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"&y=&l="&lid&"&utilityid="&utilityid)%>','','width=800,height=500,resizable=yes,scrollbars=yes');
	}else if(frm.actions.value=="EmailBills"){
		window.open('loading.asp?url=<%=server.urlencode("emailTenantBills.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"&y=&l="&lid&"&utilityid="&utilityid)%>','','width=800,height=500,resizable=yes,scrollbars=yes');
	}else if(frm.actions.value=="View All Tenant Summaries"){
		window.open('loading.asp?url=<%=server.urlencode("/genergy2/UMreports/tenantBillSummary.asp?pid="&pid&"&building="&building&"&syear="&byear&"&utilityid="&utilityid)%>','','width=800,height=600,scrollbars=yes');
	}else if(frm.actions.value=="Data Files"){
		window.open('dataOutput.asp?action=IBS&pid=<%=pid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&y=&building=<%=building%>&utilityid=<%=utilityid%>','','width=300,height=100,resizable=no,scrollbars=no');
	//}else if(frm.actions.value=="Invoice Notes"){
	//	window.open('invoice_notes.asp?pid=<%=pid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&building=<%=building%>&utilityid=<%=utilityid%>','','width=300,height=100,resizable=no,scrollbars=no');
	}else{
		frm.submit();
	}
	

	
}
</script>
