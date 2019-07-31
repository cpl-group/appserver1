<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, ypid, building, pid, byear, bperiod, utilityid, detailed, meterbreakdown, calcintpeak, SJPproperties, ShowUsageDetails, ShowDemandDetails,onlinereview,totdeliverychrgs, maxmeters, textheader, masterTotal, utilityname

leaseid = trim(Request("l"))
ypid = trim(request("y"))
utilityid = trim(request("utilityid"))
building = trim(request("building"))
pid = trim(request("pid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
logo = trim(request("logo"))
detailed = trim(request("detailed"))
textheader = trim(request("textheader"))

if trim(request("SJPproperties"))="true" then SJPproperties = true else SJPproperties = false
if trim(request("meterbreakdown"))="no" then meterbreakdown = false else meterbreakdown = true
if trim(request("summaryusage"))="true" then showusagedetails= true else showusagedetails = false
if trim(request("summarydemand"))="true" then showdemanddetails = true else showdemanddetails = false
dim pdfsession
pdfsession = request("pdf")

if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) or ( pdfsession ="yes" ) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if

dim cnn1, rst1, rst2, bldgrs, usagelabel, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

'if trim(getbuildingIP(building))="0" then response.redirect "/eri_th/pdfmaker/genergyInvoice.asp?"&request.servervariables("QUERY_STRING") else cnn1.Open getLocalConnect(building)
cnn1.Open getLocalConnect(building)
usagelabel = "used"
dim DBmainIP
DBmainIP = ""
if utilityid="" and leaseid<>"" then
  rst1.open "SELECT utility FROM tblleasesutilityprices WHERE leaseutilityid="&leaseid, cnn1
  if not rst1.eof then utilityid = cint(rst1("utility"))
  rst1.close
  set rst1 = nothing
end if


dim templid, tempypid, temputility, logo, extusage
set bldgrs = Server.CreateObject("ADODB.Recordset")

if logo = "" then logo = "invoice_logo_1.jpg"
response.write "<html><head><title></title></head><body bgcolor=""#FFFFFF"">"
if trim(request("billid"))<>"" and trim(leaseid)<>"" then
	showtenantbill leaseid, trim(request("billid")), utilityid,extusage
elseif leaseid<>"" and ypid<>"" then
	sql = "SELECT bp.id as billid FROM tblbillbyPeriod bp WHERE bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid
  bldgrs.open sql, cnn1
  if not bldgrs.eof then showtenantbill leaseid, bldgrs("billid"), utilityid,extusage
	bldgrs.close
elseif building<>"" then
	if ypid<>"" then
		bldgrs.open "select b.id as billid, b.leaseutilityid, b.utility, extusg FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and ypid="&ypid&" ORDER BY TenantName", cnn1
		do until bldgrs.eof
			templid = trim(bldgrs("leaseutilityid"))
		    temputility = trim(bldgrs("utility"))
			Extusage = trim(bldgrs("extusg"))
			showtenantbill templid, bldgrs("billid"), temputility,extusage
			bldgrs.movenext
		loop
	elseif byear<>"" and bperiod<>"" then
    sql = "SELECT b.id as billid, b.leaseutilityid, ypid, b.utility, extusg FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
    if isnumeric(utilityid) then sql = sql & " and utility="&utilityid
    sql = sql & "  ORDER BY TenantName"
		bldgrs.open sql, cnn1
		do until bldgrs.eof
			templid = trim(bldgrs("leaseutilityid"))
			tempypid = trim(bldgrs("ypid"))
			temputility = trim(bldgrs("utility"))
			Extusage = trim(bldgrs("extusg"))
			showtenantbill templid, bldgrs("billid"), temputility, extusage
			bldgrs.movenext
		loop
	end if
end if
set cnn1 = nothing
response.write "</body></html>"

'response.write Totalout




'### begin of showtenantbill, is rest of file ###

sub showtenantbill(leaseid, billid, utilityid, extusage)

dim loopgroup, metercount, tot_onpeak, tot_offpeak, tot_kwhused,tot_kwhusedoff,tot_kwhusedon,tot_kwhusedint, tot_demand_p, tot_demand_c, coincidentflag, usagedivisor, unittype, tot_intpeak, tot_demandoff_p, tot_demandint_p,currentdemandP,currentUsage, serviceclass,serviceclassid
usagedivisor = 1
Set rst2 = Server.CreateObject("ADODB.recordset")
'get utility name
rst2.open "SELECT utility FROM tblutility WHERE utilityid="&utilityid, getConnect(pid,building,"billing")
if not rst2.eof then
	utilityname = rst2("utility")
end if
rst2.close

if extusage = "" then 
	sql = "select extusg from tblbillbyperiod where id = " & billid
	rst2.open sql, cnn1, 2
	
	if not rst2.eof and trim(rst2("extusg")) <> "" then 
		extusage = rst2("extusg")
	else 
		extusage = false
	end if 
	rst2.close	
end if

select case utilityid
case 3, 10,6,1
	select case utilityid
	case 6,1
		usagedivisor = 1
	case else
		usagedivisor = 100
  end select 
	sql = "SELECT b.portfolioid, isnull(r.btstrt,'') as billto, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, r.*, l.onlinebill FROM tblbillbyperiod r, tblleases l, buildings b, "&DBmainIP&"ratetypes rt WHERE r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&billid
case else
  	sql = "SELECT b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(tunitcostkw,6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, datestart,dateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_bill_summary r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
end select
'response.write sql
'response.end
rst2.open sql, cnn1, 2

'rstsubTotal, rsttaxnonull, rst2Totalamtnonull
if not rst2.eof then
pid = rst2("portfolioid")
unittype = rst2("unittype")
serviceclass = rst2("rt")
serviceclassid = rst2("rtid")
'rst2subTotal = rst2("subTotal")
'rst2taxnonull = rst2("taxnonull")
'rst2Totalamtnonull = rst2("Totalamtnonull")
select case serviceclassid 
	case 36
		maxmeters = 3
	case else	
		maxmeters = 9
end select
%>

<!-- header -->
<table width="100%" border="0" bgcolor="#FFFFFF"><tr><td height="68"><img src="http://appserver1.genergy.com/eri_th/pdfMaker/<%=logo%>" hspace="40" width="202" height="143"></td><td align="right">
<table cellspacing="0" cellpadding="1" border="0" bgcolor="#CCCCCC"><tr><td>
<table border="0" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF"><tr><td></td></tr>
<%
dim subTotal
select case utilityid
case 3, 10, 6		'water utilities%>
  <%
  subTotal = tot_demand_P
  subTotal = subTotal + subTotal * Formatnumber(rst2("adminfee"),2)
  subTotal = subTotal + cdbl(rst2("serviceFeenonull"))
	%>
  <tr bgcolor="#FFFFFF">
  	<td align="right">Sub Total</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("subtotal"))%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  	<td align="right">Sales Tax</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("taxnonull"))%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  	<td align="right"><font size="+1"><b>Total&nbsp;Due</b></font></td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><font size="+1"><b><%=FormatCurrency(rst2("Totalamtnonull"))%></b></font></td>
  </tr>
  <%masterTotal = rst2("Totalamtnonull")%>
<%case 1 'Steam
	Dim SteamBreakdown, i,details
	steamBreakdown = split(cstr(rst2("energydetail")),"|")
    miscCredits 2,leaseid
		%>
    <tr bgcolor="#FFFFFF"> 
      <td align="right" nowrap>Sub Total </td>
      <td align="right" ><hr noshade size="1"></td>
      <td align="right" nowrap><%=formatcurrency(rst2("subtotal"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="right" nowrap>Sales Tax</td>
      <td align="right" nowrap>&nbsp;</td>
      <td align="right" nowrap><%=FormatCurrency(rst2("taxnonull"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="right" nowrap><strong>Total&nbsp;Due</strong></td>
      <td align="right" ><hr noshade size="1"></td>
      <td align="right"  nowrap><strong><%=formatcurrency(rst2("totalamt"))%></strong></td>
    </tr>
	<%masterTotal = rst2("Totalamt")%>
<%case else%>
  <tr bgcolor="#FFFFFF">
  	<td align="right" bgcolor="#FFFFFF">Sub Total</td>
  	<td width="20" bgcolor="#FFFFFF">&nbsp;</td>
  	<td align="right" bgcolor="#FFFFFF"><%=FormatCurrency(rst2("subTotal"),2)%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  	<td align="right">Sales Tax</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("taxnonull"),2)%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
  	<td align="right"><font size="+1"><b>Total&nbsp;Due</b></font></td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><font size="+1"><b><%=FormatCurrency(rst2("Totalamtnonull"),2)%></b></font></td>
  </tr>
  <%masterTotal = rst2("Totalamtnonull")%>
<%end select%>
  </table></td></tr>
</table>
</td></tr></table>
</td></tr></table>

<!-- header end -->
<table width="102%" cellpadding="2" cellspacing="0" border="0" height="85%">
<%if textheader="advRealty.asp" then%>
<tr><td align="center"><!--#INCLUDE FILE="advRealty.asp" --></td></tr>
<%end if%>
<tr><td <%if (lcase(trim(serviceclass))="lpls2" and pid <> "15") or serviceclassid= 36 then%><%else%>height="380"<%end if%> valign="top">
<!-- meterlisting -->
<table width="640" cellpadding="5" cellspacing="2" border="0" align="center" bgcolor="#CCCCCC">
<tr><td width="384" rowspan="2" bgcolor="#FFFFFF" colspan=<% if showusagedetails = false then %>"6"<%else%>"4"<%end if %>><font size="+1"><strong><%=utilityname%> Bill</strong></font></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Number</td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Date</td>
</tr>
<tr><td colspan="2" width="256" bgcolor="#eeeeee" align="center">EL.<%=rst2("billperiod") & Right(rst2("billyear"),2) &"."& rst2("tenantnum") %></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center"><%if not(isnull(rst2("postdate"))) then response.write formatdatetime(rst2("postdate"),2)%></td>
</tr>
<%
select case utilityid
case 3, 10,6,1
  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility")
case else
  if rst2("calcintpeak") then calcintpeak = true else calcintpeak = false
	  if extUsage then 
		  metertableheadsExtUsage rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend")
	  else
		  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), calcintpeak
	  end if
end select

Set rst1 = Server.CreateObject("ADODB.recordset")
rst1.open "select * from tblmetersbyperiod where bill_id="&billid, cnn1
'response.write "select * from tblmetersbyperiod where bill_id="&billid
'response.end
tot_onpeak 		=0
tot_offpeak		=0
tot_intpeak		=0
tot_demandint_p	=0
tot_demandoff_p	=0
tot_kwhused		=0
tot_kwhusedoff	=0
tot_kwhusedint	=0
tot_demand_p	=0
tot_demand_c	=0
coincidentflag 	= false
loopgroup 		= 1
while not rst1.eof
	metercount = metercount + 1
	if rst1("coincident")="True" then coincidentflag=True
		select case utilityid
			case 3, 10, 6
				if meterbreakdown then
					%>
					<tr bgcolor="#FFFFFF">
						<td align="center" colspan="3"><%=rst1("Meternum")%></td>
						<td align="right"><%=Formatnumber(rst1("manualmultiplier"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<%currentUsage = Formatnumber(((cdbl(rst1(usagelabel))/usagedivisor)),2)%>
						<td align="right" colspan="3"><%=currentUsage%></td>
						<%currentDemandP = formatcurrency((currentUsage*formatnumber(rst2("demanddetail"),6))+(currentUsage*formatnumber(rst2("energydetail"),6)))%>
						<td align="right"><%=currentDemandP%></td>
						<%tot_demand_p = tot_demand_p + currentDemandP%>
					</tr>
					<%
				end if
			case 1
				if meterbreakdown then
					%>
					<tr bgcolor="#FFFFFF">
						<td align="center" colspan="3"><%=rst1("Meternum")%></td>
						<td align="right"><%=Formatnumber(rst1("manualmultiplier"),3)%></td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<%currentUsage = Formatnumber(((cdbl(rst1(usagelabel))/usagedivisor)),2)%>
						<td align="right" colspan="3"><%=currentUsage%></td>
					</tr>
					<%
				end if
			case else
				if meterbreakdown then
					if extusage then 
					%>
					<tr bgcolor="#FFFFFF">
						<td align="center" colspan="2"><%=rst1("Meternum")%></td>
						<td align="left"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<td align="right">On Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("used"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("demand_P"),2)%></td>
					</tr>
					<tr bgcolor="#FFFFFF">
						<td></td>
						<td align="right" Colspan=3>Off Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawpreviousoff"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrentoff"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("usedoff"),2)%></td>
						<td align="right" colspan=3>&nbsp;</td>
					</tr>
					<tr bgcolor="#FFFFFF">
						<td></td>
						<td align="right" Colspan=3>Mid Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawpreviousint"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrentint"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("usedint"),2)%></td>
						<td align="right" colspan=3>&nbsp;</td>
					</tr>
					<%
					else
					%>
					<tr bgcolor="#FFFFFF">
						<td align="center" colspan="3"><%=rst1("Meternum")%></td>
						<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<td align="center"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="center"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("onpeak"),2)%></td>
						<%if calcintpeak then%>
							<td align="right"><%=Formatnumber(rst1("intpeak"),2)%></td>
						<%end if%>
						<td align="right"><%=Formatnumber(rst1("offpeak"),2)%></td>
						<td align="right"><%=Formatnumber(rst1(usagelabel),2)%></td>
						<td align="right"><%=Formatnumber(rst1("demand_P"),2)%></td>
					</tr>
					<%
					end if
				end if
		end select
	if calcintpeak then
		tot_intpeak= tot_intpeak+ formatnumber(cdbl(rst1("intpeak")),2)
		tot_demandoff_p= tot_demandoff_p + formatnumber(cdbl(rst1("demand_off")),2)
		tot_demandint_p= tot_demandint_p + formatnumber(cdbl(rst1("demand_int")),2)
	end if 
	tot_offpeak= tot_offpeak + formatnumber(cdbl(rst1("offpeak")),2)
	tot_onpeak = tot_onpeak + cdbl(rst1("onpeak"))
	if utilityid = 3 or utilityid = 10 then
		tot_kwhused= tot_kwhused + (cdbl(rst1(usagelabel))/usagedivisor)
	else
		if extusage then 
			tot_kwhusedon 	= tot_kwhusedon + (cdbl(rst1(usagelabel))/usagedivisor)
			tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(rst1("usedoff"))/usagedivisor)
			tot_kwhusedint 	= tot_kwhusedint + (cdbl(rst1("usedint"))/usagedivisor)
			tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
		else
			tot_kwhused = tot_kwhused + (cdbl(rst1(usagelabel))/usagedivisor)
		end if
	end if
	if isnumeric(rst1("demand_C")) then tot_demand_c=cdbl(rst1("demand_C"))
		tot_demand_p= tot_demand_p + cdbl(rst1("demand_P"))
		rst1.movenext
	if loopgroup>maxmeters and not(rst1.eof) then
		loopgroup = 0%>
		<tr>
			<td align="right" colspan="10">Continues On The Next Page...</td>
		</tr>
		</table></td></tr>
		</table>
		
		<WxPrinter PageBreak>
		
		<table width="100%" cellpadding="2" cellspacing="0" border="0" height="100%">
		<tr><td <%if (lcase(trim(serviceclass))="lpls2" and pid <> "15") or serviceclassid= 36 then%><%else%>height="380"<%end if%> valign="top">
		<table width="640" border="0" align="center" cellpadding="5" cellspacing="2" bgcolor="#CCCCCC">
		<%select case utilityid
			case 3, 10, 6,1
			  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility")
			case else
			  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), calcintpeak
		end select
	end if
	loopgroup = loopgroup + 1
wend%>
<%select case utilityid
case 3, 10,6,1 'Water bill%>
<tr bgcolor="#FFFFFF">
	<td colspan="5"></td>
	<td align="center">Totals</td>
	<td align="right" colspan="3"><b><%=Formatnumber(tot_kwhused,2)%></b></td>
	<%if utilityid <> 1 then%>
	<td align="right"><b><%if coincidentflag then response.write formatcurrency(FormatNumber(tot_demand_C,2)) else response.write formatcurrency(FormatNumber(tot_demand_P,2))%></b></td>
   <%end if%>
</tr>
<%case else
	if extusage then
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (KWH)</td>
		<td align="right">On</td>
		<td align="right">Off</td>
		<td align="right">Mid</td>
		<td align="right">Total</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right">&nbsp;</td>
		<td align="right"><%=Formatnumber(tot_kwhusedon,2)%></td>
		<td align="right"><%=Formatnumber(tot_kwhusedoff,2)%></td>
		<td align="right"><%=Formatnumber(tot_kwhusedint,2)%></td>
		<td align="right"><%=Formatnumber(tot_kwhused,2)%></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (KW)</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right"><%=Formatnumber(tot_demand_p,2)%></td>
	</tr>
	<%else%>
	<tr bgcolor="#FFFFFF">
		<% if showusagedetails = false and showDemanddetails = false then %>
		<td colspan="5"></td>
		<%else%>
		<td colspan="3"></td>
		<%end if %>
		<td align="center">Totals</td>
	   <% if showusagedetails = false then %>
		<td align="right"><%if meterbreakdown then%><%=Formatnumber(tot_onpeak,2)%><%end if%></td>
	  <%if calcintpeak then%><td align="right"><%if meterbreakdown then%><%=Formatnumber(tot_intpeak,2)%><%end if%></td><%end if%>
		<td align="right"><%if meterbreakdown then%><%=Formatnumber(tot_offpeak,2)%><%end if%></td>
	  <% end if %>
		<td align="right"><%=Formatnumber(tot_kwhused,2)%></td>
		<% if showdemanddetails = false then %>
		<td align="right"><%if coincidentflag then response.write FormatNumber(tot_demand_C,2) else response.write FormatNumber(tot_demand_P,2)%></b></td>
			<%if calcintpeak then%>
		<td align="right"><%if coincidentflag then response.write FormatNumber(tot_demandint_p,2) else response.write FormatNumber(tot_demandint_p,2)%></td>
		<td align="right"><%if coincidentflag then response.write FormatNumber(tot_demandoff_p,2) else response.write FormatNumber(tot_demandoff_p,2)%></td>
			<%end if%>
		<%end if%>
	</tr>
	<%
	end if
end select%>
</table>
<%if detailed="true" and utilityid<>3 then%>
<table width="640" align="center" border="0">
<%if trim(rst2("energydetail"))<>"" then%>
<tr><td>Calculation of Consumption Charge:</td></tr>
<tr><td><%=replace(replace(rst2("energydetail"),"|","<br>")," ","&nbsp;")%></td></tr>
<%end if%>
<%if trim(rst2("demanddetail"))<>"" then%>
<tr><td>Calculation of Demand Charge:</td></tr>
<tr><td><%=replace(replace(rst2("demanddetail"),"|","<br>")," ","&nbsp;")%></td></tr>
<%end if%>
</table>
<%end if%>
<%if rst2("invoiceNote")<>"" then%>
<table align="center" cellpadding="0" cellspacing="0" width="640"><tr><td width="320">
<table width="320" bgcolor="black" cellpadding="3" cellspacing="1">
<tr><td bgcolor="white"><%=rst2("invoiceNote")%></td></tr>
</table>
<td width="320">&nbsp;</td>
</td></tr></table>
<%end if%>
<!-- end meter listing -->
</td></tr><tr><td height="105">
<!-- Totaling section -->
<%if trim(serviceclass) <> "LPLS2" or (pid = "15" and detailed<>"true") then%>
<table border=0 cellpadding="0" cellspacing="0" width="640" align="center">
<tr>
          <td width="50%" valign="top"> 
					
  </td>
  <td width="50%" align="right" valign="top">
<!-- start Totals section -->
<%
select case utilityid
case 3, 10, 6		'water utilities%>
  <table border="0" cellpadding="3" cellspacing="1">
  <%
  subTotal = tot_demand_P
  subTotal = subTotal + subTotal * Formatnumber(rst2("adminfee"),2)
  subTotal = subTotal + cdbl(rst2("serviceFeenonull"))
	%>
  <tr>
  	<td align="right" colspan=2>Sub Total</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("subtotal"))%></td>
  </tr>
  <tr>
  	<td align="right" colspan=2>Sales Tax</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("taxnonull"))%></td>
  </tr>
  <tr>
  	<td align="right" colspan=2><font size="+1"><b>Total&nbsp;Due</b></font></td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><font size="+1"><b><%=FormatCurrency(rst2("Totalamtnonull"))%></b></font></td>
  </tr>
  </table>
  <%masterTotal = rst2("Totalamtnonull")%>
<%case 1 'Steam
	steamBreakdown = split(cstr(rst2("energydetail")),"|")
	%>
	<table border="0" cellpadding="3" cellspacing="1">
    <%
    miscCredits 2,leaseid
		%>
    <tr> 
      <td align="right" nowrap>Sub Total </td>
      <td align="right" ><hr noshade size="1"></td>
      <td align="right" nowrap><%=formatcurrency(rst2("subtotal"))%></td>
    </tr>
    <tr> 
      <td align="right" nowrap>Sales Tax</td>
      <td align="right" nowrap>&nbsp;</td>
      <td align="right" nowrap><%=FormatCurrency(rst2("taxnonull"))%></td>
    </tr>
    <tr> 
      <td align="right" nowrap><strong>Total&nbsp;Due</strong></td>
      <td align="right" ><hr noshade size="1"></td>
      <td align="right"  nowrap><strong><%=formatcurrency(rst2("totalamt"))%></strong></td>
    </tr>
  </table>
	<%masterTotal = rst2("Totalamt")%>
<%case else%>
  <table border="0" cellpadding="3" cellspacing="1">
  <tr>
  	<td align="right" colspan=2>Sub Total</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("subTotal"),2)%></td>
  </tr>
  <tr>
  	<td align="right" colspan=2>Sales Tax</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(rst2("taxnonull"),2)%></td>
  </tr>
  <tr>
  	<td align="right" colspan=2><font size="+1"><b>Total&nbsp;Due</b></font></td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><font size="+1"><b><%=FormatCurrency(rst2("Totalamtnonull"),2)%></b></font></td>
  </tr>
  </table>
  <%masterTotal = rst2("Totalamtnonull")%>
<%end select%>
  </td>
</tr>
</table>
<%else
end if 
%>
<!-- end Totals section -->
</td>
</tr>
<tr><td valign="bottom" height="175">
<%
dim hidedemand
if trim(ucase(serviceclass))="AVG COST 1" or coincidentflag or utilityid=3 or utilityid = 10 or utilityid = 1 then hidedemand="true" else hidedemand = ""
if showusagedetails = false then %>
<table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">
<tr><td width="10%" align="center"><img src="http://<%=request.servervariables("SERVER_NAME")%>/genergy2/invoices/MakeChartyrly.asp?genergy2=<%=trim(request("genergy2"))%>&lid=<%=leaseid%>&by=<%=rst2("billyear")%>&bp=<%=rst2("billperiod")%>&billid=<%=billid%>&hidedemand=<%=hidedemand%>&building=<%=building%>&unittype=<%=unittype%><%if extusage then %>&includepeaks=false&extusg=true<%else%>&includepeaks=<%=meterbreakdown%><%end if%>&calcintpeak=<%=calcintpeak%>&isGVA=true" width="600" height="175"></td></tr>
</table>
<%end if%>
</td>
</tr>
<tr><td valign="top">
<%footer rst2("tenantname"), rst2("tenantnum"), rst2("billingaddress"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("billingaddress"), rst2("btcity"), rst2("btstate"), rst2("btzip"), rst2("onlinebill"), textheader%>
</td></tr>
</table>
<WxPrinter PageBreak>
<%end if%>

<!-- </div> -->
<%

rst2.close
set rst2 = nothing
end sub

'###########################################################################################
sub footer(tenantname, tenantnum, tstrt, tcity, tstate, tzip, btbldgname, billingaddress, btcity, btstate, btzip, isonlinebill,textheader)%>
	<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr><td colspan="3" align="center"><font size="2"><% if textheader="advRealty.asp" then %>KEEP ABOVE PART FOR YOUR RECORDS<%end if%>&nbsp;</font></td></tr>
  <tr><td colspan="3" bgcolor="black" height="1"><img src="/images/spacer.gif" width="1" height="1"></td></tr>
  <tr><td colspan="3" align="center"><font size="2"><% if textheader="advRealty.asp" then %>DETACH HERE AND RETURN WITH PAYMENT<%end if%></font>&nbsp;<br></td></tr>
	<tr>
		<td width="50%" valign="top"><font size="1">Tenant Name and Address:</FONT><br>
	  <font size="3"><b><%=tenantname%> (<%=tenantnum%>)<br>
	  <%=replace(tstrt,vbNewLine,"<br>")%><br>
	  <%=tcity%>, <%=tstate%>&nbsp;<%=tzip%></b></font>
  </td>
		<td width="40%" valign="top"><font size="1">Make Check Payable To:</font><br>
		<font size="3"><b><%=btbldgname%><br>
    <%
    %>
		<%=replace(rst2("billto"),vbNewLine,"<br>")%><br>
		<%=rst2("btcity")%>, <%=rst2("btstate")%>&nbsp;<%=rst2("btzip")%><br>
	  *Payment due upon receipt</b><br></font>
	  </td>
		<% if textheader="advRealty.asp" then %><td width="10%" valign="top" align="right" nowrap><br><strong><font size="3"><u>TOTAL DUE: <%=masterTotal%></u><font></strong><br></td><% end if %>
	</tr>
  <tr>
  <% if textheader<>"advRealty.asp" then %><td colspan="3">&nbsp;<br><center>If you have any questions about your bill, please call Genergy @ 212-664-7600 ext 105.<br>
      <%if trim(isonlinebill)="True" then%>To view online bill, login to www.genergyonline.com with access code <b><%=tenantnum%>.<%=building%></b>.<%end if%></center></td><%end if%></tr>
	</table>
<%end sub

sub metertableheads(billyear, billperiod, datestart, dateend, showintpeaks)%>
    <tr bgcolor="#eeeeee">
    	<td align="center">Period</td>
    	<td align="center">From</td>
    	<td align="center">To</td>
    	<td align="center">No.&nbsp;Days</td>
		<%if showUsagedetails = false then %><td colspan="2" width="20%" align="center">&nbsp;</td><%end if%>
    	<td <%if showusagedetails = false then %>colspan=<%if showintpeaks then%>"4"<%else%>"3"<%end if%> <%end if%> width="30%" align="center">CONSUMPTION</td>
    	<%if showDemanddetails = false then %> <td <%if showintpeaks then%>colspan="3"<%end if%> align="center" rowspan="2">KW of DEMAND</td><%end if%>
    </tr>
    <tr>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=billyear%>/<%=billperiod%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=datestart%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=dateend%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=dateend-datestart%></td>
		<%if showUsagedetails = false then %>
    	<td bgcolor="#eeeeee" width="64" align="center">Previous&nbsp;Reading</td>
    	<td bgcolor="#eeeeee" width="64" align="center">Current&nbsp;Reading</td>
    	<td bgcolor="#eeeeee" width="64" align="center">On&nbsp;Peak</td>
      <%if showintpeaks then%>
      	<td bgcolor="#eeeeee" width="64" align="center">Int&nbsp;Peak</td>
      <%end if%>
    	<td bgcolor="#eeeeee" width="64" align="center">Off&nbsp;Peak</td>
		<%end if %>
    	<td bgcolor="#eeeeee" width="64" align="center">Total&nbsp;Usage</td>
	 <%if showDemanddetails = false then %>
      <%if showintpeaks then%>
        <td bgcolor="#eeeeee" width="64" align="center">Int</td>
        <td bgcolor="#eeeeee" width="64" align="center">Off</td>
      <%end if%>
	 <%end if %>
    </tr>
    <tr>
    	<td colspan="3" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="center" bgcolor="#eeeeee">Multiplier</td>
    	<td colspan="6" align="center" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
<%
end sub

sub metertableheadsExtUsage(billyear, billperiod, datestart, dateend)%>
    <tr bgcolor="#eeeeee">
    	<td align="center">Period</td>
    	<td align="center">From</td>
    	<td align="center">To</td>
    	<td align="center">No.&nbsp;Days</td>
		<td colspan="2" width="20%" align="center">&nbsp;</td>
		<td colspan="3" width="30%" align="center">CONSUMPTION</td>
    	<td colspan="3" align="center">DEMAND</td></tr>
    <tr>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=billyear%>/<%=billperiod%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=datestart%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=dateend%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=dateend-datestart%></td>
    	<td bgcolor="#eeeeee" width="64" align="center">Previous Reading</td>
    	<td bgcolor="#eeeeee" width="64" align="center">Current Reading</td>
    	<td bgcolor="#eeeeee" width="64" align="center" colspan=3>&nbsp;KWH</td>
    	<td bgcolor="#eeeeee" width="64" align="right" colspan=3>KW</td>
    </tr>
    <tr>
    	<td colspan="2" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="left" bgcolor="#eeeeee">Multiplier</td>
    	<td colspan="8" align="center" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
<%end sub

sub miscCredits(numcells, tmpleaseid)
	if cint(rst2("creditnonull")) <> 0 then
		dim rstMiscCred, credSql
		credSql = "select isnull(description,'Misc Credit') as [desc], credit from tblcreditbyperiod where billyear = "&byear&" and billperiod = "&bperiod&" and leaseutilityid = "&tmpleaseid&" and not credit = 0"
		set rstMiscCred = server.createobject("adodb.recordset")
		rstMiscCred.open credSql, getLocalConnect(building)
		'response.write credSql
		if not rstMiscCred.eof then
			do while not rstMiscCred.eof
				dim desc
				desc = rstMiscCred("desc")	
				if numcells = 3 then		%>
					<tr>
						<td align="right" width="15%">Credit:</td>
						<td align="right"><%=desc%></td>
						<td>&nbsp;</td>
						<td align="right">$<%=formatnumber(rstMiscCred("credit"),2)%></td>
					</tr>		<%
				else		%>
					<tr>
						<td align="right" width="10%">Credit:</td>
						<td align="right"><%=desc%></td>
						<td width="15%" align="right">$<%=formatnumber(rstMiscCred("credit"),2)%></td>
					</tr>		<%
				end if
				rstMiscCred.movenext
			loop
		end if
	end if
end sub

sub metertableheadsWater(billyear, billperiod, datestart, dateend, utilityid)
		dim usageHeader
		Select Case Utilityid
				case 3,10
					 	usageHeader = "CCF Used"
				case 6
						usageHeader = "Ton/Hrs"
				case 1 
						usageHeader = "M/LBS Used"
				end Select 
		%>
    <tr bgcolor="#eeeeee">
    	<td align="center">Period</td>
    	<td align="center">From</td>
    	<td align="center">To</td>
    	<td align="center">No.&nbsp;Days</td>
    	<td colspan="2" width="20%" align="center">&nbsp;</td>
		
    	<td colspan="<%if utilityid <> 1 then%>3<%else%>4<%end if%>"  width="30%" align="center">CONSUMPTION</td>
    	<%if utilityid <> 1 then %><td align="center" rowspan="2">Sub&nbsp;Total</td><%end if%>
    </tr>
    <tr>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=billyear%>/<%=billperiod%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=datestart%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=dateend%></td>
    	<td width="64" align="center" bgcolor="#FFFFFF"><%=dateend-datestart%></td>
    	<td bgcolor="#eeeeee" width="64" align="center">Previous Reading</td>
    	<td bgcolor="#eeeeee" width="64" align="center">Current Reading</td>
    	<td bgcolor="#eeeeee" align="right" colspan="<%if utilityid <> 1 then%>3<%else%>4<%end if%>"><%=usageHeader%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td colspan="3" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="center" bgcolor="#eeeeee">Multiplier</td>
    	<td colspan="6" align="center" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
<%end sub

sub meterlist()
				
end sub
%>