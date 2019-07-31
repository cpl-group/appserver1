<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, ypid, building, pid, byear, bperiod, utilityid, detailed, meterbreakdown, calcintpeak, SJPproperties, ShowUsageDetails, ShowDemandDetails,onlinereview,totdeliverychrgs, maxmeters, textheader, masterTotal,demo, reject

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
demo = request("demo")
if demo = "" then demo = false end if 
if trim(request("reject"))="" then reject = 0 else reject = 1

if trim(request("meterbreakdown"))="no" then meterbreakdown = false else meterbreakdown = true
if trim(request("summaryusage"))="true" then showusagedetails= true else showusagedetails = false
if trim(request("summarydemand"))="true" then showdemanddetails = true else showdemanddetails = false

'if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
dim pdfsession,startCount,endCount
pdfsession = request("fdp")
startCount = request("s")
endCount = request("e")
if ((request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj"))) OR (pdfsession = "pdffdp"))  then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if

dim cnn1, rst1, rst2, rst3, bldgrs, usagelabel, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
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
	sql = "SELECT bp.id as billid FROM tblbillbyPeriod bp WHERE reject=0 and bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid
	
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
		if not bldgrs.eof then
			utilityid = trim(bldgrs("utility"))
			showsummary building, byear, bperiod,utilityid
			'response.End()
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
end if
set cnn1 = nothing
response.write "</body></html>"

'response.write Totalout




'### begin of showtenantbill, is rest of file ###

function showsummary(building, byear, bperiod, utilityid)

Dim firstLoop
firstLoop = 0
dim loopgroup, metercount, tot_onpeak, tot_offpeak, tot_kwhused,tot_kwhusedoff,tot_kwhusedon,tot_kwhusedint, tot_demand_p, tot_demand_c, coincidentflag, usagedivisor, unittype, tot_intpeak, tot_demandoff_p, tot_demandint_p,currentdemandP,currentUsage, serviceclass,serviceclassid, utilityname, paymentterm
dim abs_onpeak,abs_offpeak,abs_intpeak,abs_demandint_p,abs_demandoff_p,abs_kwhused,abs_kwhusedoff,abs_kwhusedint,abs_demand_p,abs_demand_c
Dim monthly_charge,industry_settlment,vat,day_units,night_units,max_demand,ccl_tot_units,tot_invoices
usagedivisor = 1
Set rst2 = Server.CreateObject("ADODB.recordset")
'get utility name
'response.Write("SELECT utility FROM tblutility WHERE utilityid="&utilityid)
'response.End()
rst3.open "SELECT utility FROM tblutility WHERE utilityid="&utilityid, getConnect(pid,building,"billing")
if not rst3.eof then
	utilityname = rst3("utility")
end if
rst3.close

'get payment text
rst3.open "SELECT isnull(paymentterm,'') as paymentterm FROM portfolio WHERE id="&pid, getConnect(pid,building,"billing")
if not rst3.eof then
	paymentterm = rst3("paymentterm")
	if trim(paymentterm)="" then paymentterm = "*Payment due upon receipt"
end if
rst3.close


Dim billsql
billsql = "SELECT b.id as billid, b.leaseutilityid, ypid, b.utility, extusg FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
    if isnumeric(utilityid) then billsql = billsql & " and utility="&utilityid
    billsql = billsql & "  ORDER BY TenantName"
Dim billrs
Set billrs = Server.CreateObject("ADODB.recordset")
billrs.open billsql,getConnect(pid,building,"billing"),adOpenStatic
Dim billid,extusage
'Main LOOP

'Absolute Totals
abs_onpeak 		=0
abs_offpeak		=0
abs_intpeak		=0
abs_demandint_p	=0
abs_demandoff_p	=0
abs_kwhused		=0
abs_kwhusedoff	=0
abs_kwhusedint	=0
abs_demand_p	=0
abs_demand_c	=0
'----------
metercount=0
monthly_charge = 0
industry_settlment = 0
vat = 0
day_units = 0
night_units = 0
max_demand = 0
ccl_tot_units = 0
tot_invoices = 0



Set rst1 = Server.CreateObject("ADODB.recordset")

do until billrs.eof

Dim recordnum,percent
recordnum= billrs.recordcount
percent = "100%"
if recordnum > 6 then percent = "0%" else percent = "100%"




billid = trim(billrs("billid"))
extusage = trim(billrs("extusg"))


if extusage = "" then 
	sql = "select extusg from tblbillbyperiod where id = " & billid
	rst3.open sql, cnn1, 2
	
	if not rst3.eof and trim(rst3("extusg")) <> "" then 
		extusage = rst3("extusg")
	else 
		extusage = false
	end if 
	rst3.close	
end if
select case utilityid
case 3, 10,6,1,4
	select case utilityid
	case 6,1,4
		usagedivisor = 1
	case else
		usagedivisor = 100
  end select 
	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto, r.addonfee as myaddonfee, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, (select case count(distinct isnull(addonfee,0)) when 0 then 0 else 1 end as aoncnt from tblmetersbyperiod where bill_id=r.id group by leaseutilityid,ypid,bill_id) as showaddonfee, r.invoice_note as invoiceNote, isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.*, l.onlinebill,r.penalty FROM tblbillbyperiod r, tblleases l, buildings b, "&DBmainIP&"ratetypes rt WHERE r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&billid
case else
	if SJPproperties then
  	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_Bill_summary_NoBill r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	else
  	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_bill_summary_NoBill r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	end if
end select

rst2.open sql, cnn1, 2
if not rst2.eof then
pid = rst2("portfolioid")
unittype = rst2("unittype")
serviceclass = rst2("rt")
serviceclassid = rst2("rtid")
select case serviceclassid 
	case 36
		maxmeters = 3
	case else	
		maxmeters = 5
end select



if firstLoop = 0 then
'Heading Columns%>
<!-- header -->
<table width="102%" cellpadding="2" cellspacing="0" border="0" height="<%=percent%>" id=mainTable align="center" >
<tr><td align="center"><br><font size="10"><strong><%=utilityname%> Bill Summary</strong></font><br><br></td></tr>
<tr><td valign="top" align="center">
<%mainheader rst2("tenantname"), rst2("tenantnum"), rst2("billingaddress"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("billingaddress"), rst2("btcity"), rst2("btstate"), rst2("billzip"), rst2("onlinebill"), textheader,demo, paymentterm%>
</td></tr>
<tr><td>

	<table width="640" cellpadding="5" cellspacing="2" border="0" align="center">
<tr><td width="384" rowspan="2" colspan=<% if showusagedetails = false then %>"6"<%else%>"4"<%end if %>><font size="4"><u>Billing Summary</u></font></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Number</td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Date</td>
</tr>
<tr>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">PGE/<%=rst2("billperiod") &"/" & Right(rst2("billyear"),2) &"/"& getPropNumber(building) %></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center"><%if isdate(rst2("dateend")) then response.write formatdatetime(rst2("dateend"),2)%></td>
</tr>
<%

select case utilityid
case 3, 10,6,1,4
  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility")
case else
  if rst2("calcintpeak") then calcintpeak = true else calcintpeak = false
	  if extUsage then 
		  metertableheadsExtUsage rst2("billyear"), rst2("billperiod"), rst2("datestart"), rst2("dateend")
	  else
		  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart"), rst2("dateend"), calcintpeak
	  end if
end select
firstLoop = 1
end if

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
'



while not rst1.eof
	metercount = metercount + 1
	if rst1("coincident")="True" then coincidentflag=True
		select case utilityid
			case 3, 10, 6
				if meterbreakdown then
					%>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<td align="center" colspan="1"><%=rst1("Meternum")%></td>
						<td align="right"><%=Formatnumber(rst1("manualmultiplier"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<%currentUsage = Formatnumber(((cdbl(rst1(usagelabel))/usagedivisor)),2)%>
						<td align="right" colspan="3"><%=currentUsage%></td>
						<%currentDemandP = cdbl(formatnumber(currentUsage*formatnumber(rst2("demanddetail"),6),2))+formatnumber(currentUsage*formatnumber(rst2("energydetail"),6),2)%>
						<td align="right"><%=currentDemandP%></td>
						<%tot_demand_p = tot_demand_p + currentDemandP%>
					</tr>
					<%
				end if
			case 1, 4
				if meterbreakdown then
					%>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<td align="center" colspan="2"><%=rst1("Meternum")%></td>
						<td align="right"><%=Formatnumber(rst1("manualmultiplier"),3)%></td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<%currentUsage = Formatnumber(((cdbl(rst1(usagelabel))/usagedivisor)),2)%>
						<td align="right" colspan="3"><%=currentUsage%></td>
					</tr>
					<%
				end if
			case else 'electricity
				if meterbreakdown then
					dim hasShownMulti_and_Num
					hasShownMulti_and_Num = false
					if (extusage and not(isnull(rst1("used")) or isnull(rst1("usedoff")) or isnull(rst1("usedint")))) and rst1("mextusg") then 
					if rst1("rawcurrent") <> 0 or (rst1("rawcurrentoff") = 0 and rst1("rawcurrentint") = 0)then 
						'in other words if on peak is zero don't show it unless off peak and int peak are 0 too (for these crazy instances of zeroed out usages that still get a charge.)%>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<%if not hasShownMulti_and_Num then%>
							<td align="center" colspan="1"><%=rst1("Meternum")%></td>
							<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<%end if%>
						<td align="right">On Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("used"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("demand_P"),2)%></td>
					</tr>
					<%hasShownMulti_and_Num=true
					end if 
					if rst1("rawcurrentoff") <> 0 then %>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<%if not hasShownMulti_and_Num then%>
							<td align="center" colspan="1"><%=rst1("Meternum")%></td>
							<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<%end if%>
						<td align="right" <%if hasShownMulti_and_Num then%>Colspan=3<%end if%>>Off Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawpreviousoff"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrentoff"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("usedoff"),2)%></td>
						<td align="right" colspan=3><%if rst1("demand_off")<>"" then%><%=Formatnumber(rst1("demand_off"),2)%><%end if%></td>
					</tr>
					<%hasShownMulti_and_Num=true
						loopgroup = loopgroup +1
					end if 
					if rst1("rawcurrentint") <> 0 then %>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<%if not hasShownMulti_and_Num then%>
							<td align="center" colspan="1"><%=rst1("Meternum")%></td>
							<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<%end if%>
						<td align="right" <%if hasShownMulti_and_Num then%>Colspan=3<%end if%>>Mid Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawpreviousint"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrentint"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("usedint"),2)%></td>
						<td align="right" colspan=3><%if rst1("demand_int")<>"" then%><%=Formatnumber(rst1("demand_int"),2)%><%end if%></td>
					</tr>
					<%hasShownMulti_and_Num=true
						loopgroup = loopgroup +1
					end if
				else%>
				<tr bordercolor="#FFFFFF">
					<td></td>
					<td align="center" <%if not extusage then%>colspan="2"<%end if%>><%=rst1("Meternum")%></td>
					<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
					<%if extusage then%><td></td><%end if%>
					<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
					<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
					<%if not(extusage) then%>
						<td align="right"><%=Formatnumber(rst1("onpeak"),2)%></td>
						<%if calcintpeak then%>
							<td align="right"><%=Formatnumber(rst1("intpeak"),2)%></td>
						<%end if%>
						<td align="right"><%=Formatnumber(rst1("offpeak"),2)%></td>
					<%end if%>
					<td align="right" <%if extusage then%>colspan="3"<%end if%>><%=Formatnumber(rst1(usagelabel),2)%></td>
					<td align="right"><%=Formatnumber(rst1("demand_P"),2)%></td>
				</tr>
				<%
				end if
			end if
		end select
		
		
		'Calculate Overall Totals
	if calcintpeak then
		tot_intpeak= tot_intpeak+ formatnumber(cdbl(rst1("intpeak")),2)
		abs_intpeak = abs_intpeak + tot_intpeak
		
		tot_demandoff_p= tot_demandoff_p + formatnumber(cdbl(rst1("demand_off")),2)
		abs_demandoff_p = abs_demandoff_p + tot_demandoff_p				
		
		tot_demandint_p= tot_demandint_p + formatnumber(cdbl(rst1("demand_int")),2)
		abs_demandint_p= abs_demandint_p + tot_demandint_p
		
	end if 
	tot_offpeak= tot_offpeak + formatnumber(cdbl(rst1("offpeak")),2)
	abs_offpeak= abs_offpeak + tot_offpeak
	
	tot_onpeak = tot_onpeak + cdbl(rst1("onpeak"))
	abs_onpeak= abs_onpeak + tot_onpeak
	
	if utilityid = 3 or utilityid = 10 then
		tot_kwhused= tot_kwhused + (cdbl(rst1(usagelabel))/usagedivisor)
		abs_kwhused= abs_kwhused + tot_kwhused
	else
		if extusage then 
			if rst1("mextusg") then 
				tot_kwhusedon 	= tot_kwhusedon + (cdbl(rst1("used"))/usagedivisor)
				abs_kwhusedon= abs_kwhusedon + tot_kwhusedon
				
				tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(rst1("usedoff"))/usagedivisor)
				abs_kwhusedoff= abs_kwhusedoff + tot_kwhusedoff
				
				tot_kwhusedint 	= tot_kwhusedint + (cdbl(rst1("usedint"))/usagedivisor)
				abs_kwhusedint= abs_kwhusedint + tot_kwhusedint
				
				tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
				abs_kwhused= abs_kwhused + tot_kwhused
				
			else
				tot_kwhusedon 	= tot_kwhusedon + (cdbl(rst1("onpeak"))/usagedivisor)
				abs_kwhusedon= abs_kwhusedon + tot_kwhusedon
				
				tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(rst1("offpeak"))/usagedivisor)
				abs_kwhusedoff= abs_kwhusedoff + tot_kwhusedoff
				
				tot_kwhusedint 	= tot_kwhusedint + (cdbl(rst1("intpeak"))/usagedivisor)
				abs_kwhusedint= abs_kwhusedint + tot_kwhusedint
				
				tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
				abs_kwhused= abs_kwhused + tot_kwhused
				
			end if
		else
			tot_kwhused = tot_kwhused + (cdbl(rst1(usagelabel))/usagedivisor)
			abs_kwhused= abs_kwhused + tot_kwhused
		end if
	end if
	if isnumeric(rst1("demand_C")) then 
	tot_demand_c=cdbl(rst1("demand_C"))
		abs_demand_c=abs_demand_c + tot_demand_c
	end if
		tot_demand_p= tot_demand_p + cdbl(rst1("demand_P"))
		abs_demand_p= abs_demand_p + tot_demand_p
	'End of Totals
	
	
		
		
	
			
			
			
		
	
	
		rst1.movenext
			
			
			
			'response.write metercount &"<br>"
			'response.Write(maxmeters)
		
	if metercount > 5  and not(billrs.eof) then
		loopgroup = 0
		metercount = 0
		'firstLoop= 0
		Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
		response.Write  "mainTable.height = 20%" & vbCrlf
		Response.Write "</SCR" & "IPT>" & vbCrlf
		%>
 		<tr>
			<td align="right" colspan="10">Continues On The Next Page...</td>
		</tr> 
		</table></td></tr>
		</table>
		
		<WxPrinter PageBreak>
		<table width="102%" cellpadding="2" cellspacing="0" border="0" height="100%" align="center" >
		<tr><td align="center"><br><font size="10"><strong><%=utilityname%> Bill Summary</strong></font><br><br></td></tr>
		<tr><td valign="top" align="center">
		<%mainheader rst2("tenantname"), rst2("tenantnum"), rst2("billingaddress"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("billingaddress"), rst2("btcity"), rst2("btstate"), rst2("billzip"), rst2("onlinebill"), textheader,demo, paymentterm%>
		</td></tr>
		<tr><td>

		<table width="640" cellpadding="5" cellspacing="2" border="0" align="center">
		<tr><td width="384" rowspan="2" colspan=<% if showusagedetails = false then %>"6"<%else%>"4"<%end if %>><font size="4"><u>Billing Summary</u></font></td>
		<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Number</td>
		<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Date</td>
		</tr>
		<tr>
		<td colspan="2" width="256" bgcolor="#eeeeee" align="center">PGE/<%=rst2("billperiod") &"/" & Right(rst2("billyear"),2) &"/"& getPropNumber(building) %></td>
		<td colspan="2" width="256" bgcolor="#eeeeee" align="center"><%if isdate(rst2("dateend")) then response.write formatdatetime(rst2("dateend"),2)%></td>
		</tr>
		<%select case utilityid
		case 3, 10,6,1,4
		  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility")
		case else
  		if rst2("calcintpeak") then calcintpeak = true else calcintpeak = false
	  	if extUsage then 
		  metertableheadsExtUsage rst2("billyear"), rst2("billperiod"), rst2("datestart"), rst2("dateend")
	  	else
		  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart"), rst2("dateend"), calcintpeak
	  	end if
		end select
	end if
	
wend
rst1.close

end if 'if not rs2.eof
rst2.close

billrs.movenext

loop

%>
<tr bordercolor="#FFFFFF"><td colspan=<%if showusagedetails = false then %>"6"<%else%>"4"<%end if %>></td><td colspan="7"><hr noshade size="1"></td></tr>
<%select case utilityid
case 3, 10,6,1,4 'Water bill%>
<tr bordercolor="#FFFFFF">
	<td colspan="5"></td>
	<td align="center">Totals</td>
	<td align="right" colspan="3"><b><%=Formatnumber(abs_kwhused,2)%></b></td>
	<%if utilityid <> 1 and utilityid <> 4 then%>
	<td align="right"><b><%if coincidentflag then response.write formatcurrency(FormatNumber(abs_demand_C,2)) else response.write formatcurrency(FormatNumber(abs_demand_P,2))%></b></td>
   <%end if%>
</tr>
<%case else
	if extusage then
	%>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (KWH)</td>
		<td align="right">On</td>
		<td align="right">Off</td>
		<td align="right">Mid</td>
		<td align="right">Total</td>
	</tr>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right">&nbsp;</td>
		<td align="right"><%=Formatnumber(abs_kwhusedon,2)%></td>
		<td align="right"><%=Formatnumber(abs_kwhusedoff,2)%></td>
		<td align="right"><%=Formatnumber(abs_kwhusedint,2)%></td>
		<td align="right"><%=Formatnumber(abs_kwhused,2)%></td>
	</tr>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (KW)</td>
		<td align="right"><%=Formatnumber(abs_demand_p,2)%></td>
		<td align="right"><%=Formatnumber(abs_demandoff_p,2)%></td>
		<td align="right"><%=Formatnumber(abs_demandint_p,2)%></td>
		<td align="right"><%=Formatnumber(abs_demand_p+abs_demandint_p+abs_demandoff_p,2)%></td>
	</tr>
	<%else%>
	<tr bordercolor="#FFFFFF">
		<% if showusagedetails = false and showDemanddetails = false then %>
		<td colspan="5"></td>
		<%else%>
		<td colspan="3"></td>
		<%end if %>
		<td align="center">Totals</td>
	   <% if showusagedetails = false then %>
		<td align="right"><%if meterbreakdown then%><%=Formatnumber(abs_onpeak,2)%><%end if%></td>
	  <%if calcintpeak then%><td align="right"><%if meterbreakdown then%><%=Formatnumber(abs_intpeak,2)%><%end if%></td><%end if%>
		<td align="right"><%if meterbreakdown then%><%=Formatnumber(abs_offpeak,2)%><%end if%></td>
	  <% end if %>
		<td align="right"><%=Formatnumber(abs_kwhused,2)%></td>
		<% if showdemanddetails = false then %>
		<td align="right"><%=FormatNumber(abs_demand_P,2)%><%if coincidentflag then %>*<%end if%></b></td>
			<%if calcintpeak then%>
				<td align="right"><%if coincidentflag then response.write FormatNumber(abs_demandint_p,2) else response.write FormatNumber(abs_demandint_p,2)%></td>
				<td align="right"><%if coincidentflag then response.write FormatNumber(abs_demandoff_p,2) else response.write FormatNumber(abs_demandoff_p,2)%></td>
			<%end if%>
		<%end if%>
	</tr>
	<%if coincidentflag then %><tr><td colspan="6">*Totalized Demand:&nbsp;<%=FormatNumber(abs_demand_c,2)%> KW</td></tr><%end if%>
	<%
	end if
end select%>
</table>
<%

billrs.close


		'Mapping 
			
			'Monthly Charge = serviceFee
			'Industry Settlment = fuelAdj
			'Vat = salesTax
			'Day Units = penalty
			'Night Units = adminfeedollar
			'Maximum Demand = demand
			'CCL total units = demand_cost
			'Total This Invoice = totalamt
			
sql = "select sum(serviceFee) as monthly_charge,sum(fuelAdj) as industry_settlment,sum(tax) as vat,sum(penalty) as day_units,sum(adminfeedollar) as night_units,sum(demand) as max_demand,sum(demand_cost) as ccl_tot_units,sum(totalamt) as tot_invoices from tblbillbyperiod where bldgnum = '" & building & "' and billperiod = " & bperiod & " and billyear = " & byear & " and reject =0"

billrs.open sql,cnn1
if (not billrs.eof) then


             monthly_charge = monthly_charge + cdbl(replaceNull(billrs("monthly_charge")))
			industry_settlment = industry_settlment + cdbl(replaceNull(billrs("industry_settlment")))
			 vat = vat + cdbl(replaceNull(billrs("vat")))
			 day_units = day_units + cdbl(replaceNull(billrs("day_units")))
			 night_units = night_units + cdbl(replaceNull(billrs("night_units")))
			 max_demand = max_demand + cdbl(replaceNull(billrs("max_demand")))
			 ccl_tot_units = ccl_tot_units + cdbl(replaceNull(billrs("ccl_tot_units")))
			 tot_invoices = tot_invoices + cdbl(replaceNull(billrs("tot_invoices")))
			Dim ava_cap_charge
			ava_cap_charge = tot_invoices - (monthly_charge+industry_settlment+vat+day_units+night_units+max_demand+ccl_tot_units)

%>
<br>
<table width="640" align="center" border="0">
<tr><td><!--optional detail header can go here for consumption --></td>
	<td><!--optional detail header can go here for demand --></td></tr>
<tr><td valign="top">
<table border="0" width = 50%>
  <tr>
    <td><strong>Fixed Charges</strong></td>
    <td align="right"><strong>Amount</strong></td>
  </tr>
  <tr>
  <td>Available Capacity</td>
  <td align="right"><%=formatnumber(ava_cap_charge,2)%></td>
  </tr>
  <tr>
    <td>Monthly Charges</td>
    <td align="right"><%=formatnumber(billrs("monthly_charge"),2)%></td>
  </tr>
  <tr>
    <td>Industry settlement</td>
    <td align="right"><%=formatnumber(billrs("industry_settlment"),2)%></td>
  </tr>
  <tr>
    <td><strong>Variable Charges</strong></td>
  </tr>
  <tr>
    <td>Day Units</td>
    <td align="right"><%=formatnumber(billrs("day_units"),2)%></td>
  </tr>
  <tr>
    <td>Night Units</td>
    <td align="right"><%=formatnumber(billrs("night_units"),2)%></td>
  </tr>
  <tr>
    <td>Maximum Demand</td>
    <td align="right"><%=formatnumber(billrs("max_demand"),2)%></td>
  </tr>
  <tr>
    <td>CCL (total units)</td>
    <td align="right"><%=formatnumber(billrs("ccl_tot_units"),2)%></td>
  </tr>
  <tr>
    <td>VAT</td>
    <td align="right"><%=formatnumber(billrs("vat"),2)%></td>
  </tr>
    <tr>
    <td><strong>Total</strong></td>
    <td align="right"><strong>£<%=formatnumber(billrs("tot_invoices"),2)%></strong></td>
  </tr>
</table>


</td>
    <td valign="top"></td></tr>
</table>
<%end if%>
</td>
</tr>
<tr><td valign="bottom" height="175">
<%
dim hidedemand
if trim(ucase(serviceclass))="AVG COST 1" or utilityid=3 or utilityid = 10 or utilityid = 1 or utilityid = 4 then hidedemand="true" else hidedemand = ""

if showusagedetails = false then %>
<table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">

<tr><td width="10%" align="center"><img src="http://<%=request.servervariables("SERVER_NAME")%>/genergy2/invoices/MakeChartyrly.asp?summary=true&genergy2=<%=trim(request("genergy2"))%>&lid=<%=leaseid%>&by=<%=byear%>&bp=<%=bperiod%>&billid=<%=billid%>&hidedemand=<%=hidedemand%>&building=<%=replace(building," ","%20")%>&unittype=<%=unittype%><%if extusage then %>&includepeaks=false&extusg=true<%else%>&includepeaks=<%=meterbreakdown%><%end if%>&calcintpeak=<%=calcintpeak%>" width="600" height="175"></td></tr>
</table>
<% end if%>
</td>
</tr>
<tr><td><%'newfooter%>
</td></tr>
</table>

<WxPrinter PageBreak>
<%
end function


sub showtenantbill(leaseid, billid, utilityid, extusage)

dim loopgroup, metercount, tot_onpeak, tot_offpeak, tot_kwhused,tot_kwhusedoff,tot_kwhusedon,tot_kwhusedint, tot_demand_p, tot_demand_c, coincidentflag, usagedivisor, unittype, tot_intpeak, tot_demandoff_p, tot_demandint_p,currentdemandP,currentUsage, serviceclass,serviceclassid, utilityname, paymentterm


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
case 3, 10,6,1,4
	select case utilityid
	case 6,1,4
		usagedivisor = 1
	case else
		usagedivisor = 100
  end select 
	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto, r.addonfee as myaddonfee, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, (select case count(distinct isnull(addonfee,0)) when 0 then 0 else 1 end as aoncnt from tblmetersbyperiod where bill_id=r.id group by leaseutilityid,ypid,bill_id) as showaddonfee, r.invoice_note as invoiceNote, isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.*, l.onlinebill FROM tblbillbyperiod r, tblleases l, buildings b, "&DBmainIP&"ratetypes rt WHERE r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&billid
case else
	if SJPproperties then
  	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_Bill_summary_NoBill r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	else
  	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_bill_summary_NoBill r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	end if
end select
rst2.open sql, cnn1, 2
if not rst2.eof then
pid = rst2("portfolioid")
unittype = rst2("unittype")
serviceclass = rst2("rt")
serviceclassid = rst2("rtid")
select case serviceclassid 
	case 36
		maxmeters = 3
	case else	
		maxmeters = 5
end select

'get payment text
rst3.open "SELECT isnull(paymentterm,'') as paymentterm FROM portfolio WHERE id="&pid, getConnect(pid,building,"billing")
if not rst3.eof then
	paymentterm = rst3("paymentterm")
	if trim(paymentterm)="" then paymentterm = "*Payment due upon receipt"
end if
rst3.close
%>

<!-- header -->
<table width="102%" cellpadding="2" cellspacing="0" border="0" height="100%">
<tr><td align="center"><br><font size="10"><strong><%=utilityname%> Bill</strong></font><br><br></td></tr>
<tr><td valign="top" align="center" width ="50%">
<%mainheader rst2("tenantname"), rst2("tenantnum"), rst2("billingaddress"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("billingaddress"), rst2("btcity"), rst2("btstate"), rst2("billzip"), rst2("onlinebill"), textheader,demo, paymentterm%>
</td></tr>
<tr><td>
<!-- meterlisting -->
<table width="640" cellpadding="5" cellspacing="2" border="0" align="center">
<tr><td width="384" rowspan="2" colspan=<% if showusagedetails = false then %>"6"<%else%>"4"<%end if %>><font size="4"><u>Billing Details</u></font></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Number</td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Date</td>
</tr>
<tr>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">PGE/<%=rst2("billperiod") &"/" & Right(rst2("billyear"),2) &"/"& getPropNumber(building) %></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center"><%if isdate(rst2("dateend")) then response.write formatdatetime(rst2("dateend"),2)%></td>
</tr>
<%
select case utilityid
case 3, 10,6,1,4
  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility")
case else
  if rst2("calcintpeak") then calcintpeak = true else calcintpeak = false
	  if extUsage then 
		  metertableheadsExtUsage rst2("billyear"), rst2("billperiod"), rst2("datestart"), rst2("dateend")
	  else
		  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart"), rst2("dateend"), calcintpeak
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
					<tr bordercolor="#FFFFFF">
						<td></td>
						<td align="center" colspan="2"><%=rst1("Meternum")%></td>
						<td align="right"><%=Formatnumber(rst1("manualmultiplier"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<%currentUsage = Formatnumber(((cdbl(rst1(usagelabel))/usagedivisor)),2)%>
						<td align="right" colspan="3"><%=currentUsage%></td>
						<%currentDemandP = cdbl(formatnumber(currentUsage*formatnumber(rst2("demanddetail"),6),2))+formatnumber(currentUsage*formatnumber(rst2("energydetail"),6),2)%>
						<td align="right"><%=currentDemandP%></td>
						<%tot_demand_p = tot_demand_p + currentDemandP%>
					</tr>
					<%
				end if
			case 1, 4
				if meterbreakdown then
					%>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<td align="center" colspan="2"><%=rst1("Meternum")%></td>
						<td align="right"><%=Formatnumber(rst1("manualmultiplier"),3)%></td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<%currentUsage = Formatnumber(((cdbl(rst1(usagelabel))/usagedivisor)),2)%>
						<td align="right" colspan="3"><%=currentUsage%></td>
					</tr>
					<%
				end if
			case else 'electricity
				if meterbreakdown then
					dim hasShownMulti_and_Num
					hasShownMulti_and_Num = false
					if (extusage and not(isnull(rst1("used")) or isnull(rst1("usedoff")) or isnull(rst1("usedint")))) and rst1("mextusg") then 
					if rst1("rawcurrent") <> 0 or (rst1("rawcurrentoff") = 0 and rst1("rawcurrentint") = 0)then 
						'in other words if on peak is zero don't show it unless off peak and int peak are 0 too (for these crazy instances of zeroed out usages that still get a charge.)%>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<%if not hasShownMulti_and_Num then%>
							<td align="center" colspan="1"><%=rst1("Meternum")%></td>
							<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<%end if%>
						<td align="right">On Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("used"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("demand_P"),2)%></td>
					</tr>
					<%hasShownMulti_and_Num=true
					end if 
					if rst1("rawcurrentoff") <> 0 then %>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<%if not hasShownMulti_and_Num then%>
							<td align="center" colspan="1"><%=rst1("Meternum")%></td>
							<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<%end if%>
						<td align="right" <%if hasShownMulti_and_Num then%>Colspan=3<%end if%>>Off Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawpreviousoff"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrentoff"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("usedoff"),2)%></td>
						<td align="right" colspan=3><%if rst1("demand_off")<>"" then%><%=Formatnumber(rst1("demand_off"),2)%><%end if%></td>
					</tr>
					<%hasShownMulti_and_Num=true
						loopgroup = loopgroup +1
					end if 
					if rst1("rawcurrentint") <> 0 then %>
					<tr bordercolor="#FFFFFF">
						<td></td>
						<%if not hasShownMulti_and_Num then%>
							<td align="center" colspan="1"><%=rst1("Meternum")%></td>
							<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
						<%end if%>
						<td align="right" <%if hasShownMulti_and_Num then%>Colspan=3<%end if%>>Mid Peak</td>
						<td align="right"><%=Formatnumber(rst1("rawpreviousint"),2)%></td>
						<td align="right"><%=Formatnumber(rst1("rawcurrentint"),2)%></td>
						<td align="right" colspan=3><%=Formatnumber(rst1("usedint"),2)%></td>
						<td align="right" colspan=3><%if rst1("demand_int")<>"" then%><%=Formatnumber(rst1("demand_int"),2)%><%end if%></td>
					</tr>
					<%hasShownMulti_and_Num=true
						loopgroup = loopgroup +1
					end if
				else%>
				<tr bordercolor="#FFFFFF">
					<td></td>
					<td align="center" <%if not extusage then%>colspan="2"<%end if%>><%=rst1("Meternum")%></td>
					<td align="center"><%=Formatnumber(rst1("manualmultiplier"),1)%></td>
					<%if extusage then%><td></td><%end if%>
					<td align="right"><%=Formatnumber(rst1("rawprevious"),2)%></td>
					<td align="right"><%=Formatnumber(rst1("rawcurrent"),2)%></td>
					<%if not(extusage) then%>
						<td align="right"><%=Formatnumber(rst1("onpeak"),2)%></td>
						<%if calcintpeak then%>
							<td align="right"><%=Formatnumber(rst1("intpeak"),2)%></td>
						<%end if%>
						<td align="right"><%=Formatnumber(rst1("offpeak"),2)%></td>
					<%end if%>
					<td align="right" <%if extusage then%>colspan="3"<%end if%>><%=Formatnumber(rst1(usagelabel),2)%></td>
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
			if rst1("mextusg") then 
				tot_kwhusedon 	= tot_kwhusedon + (cdbl(rst1("used"))/usagedivisor)
				tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(rst1("usedoff"))/usagedivisor)
				tot_kwhusedint 	= tot_kwhusedint + (cdbl(rst1("usedint"))/usagedivisor)
				tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
			else
				tot_kwhusedon 	= tot_kwhusedon + (cdbl(rst1("onpeak"))/usagedivisor)
				tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(rst1("offpeak"))/usagedivisor)
				tot_kwhusedint 	= tot_kwhusedint + (cdbl(rst1("intpeak"))/usagedivisor)
				tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
			end if
		else
			tot_kwhused = tot_kwhused + (cdbl(rst1(usagelabel))/usagedivisor)
		end if
	end if
	if isnumeric(rst1("demand_C")) then tot_demand_c=cdbl(rst1("demand_C"))
		tot_demand_p= tot_demand_p + cdbl(rst1("demand_P"))
		
		rst1.movenext
	if loopgroup>maxmeters and not(rst1.eof) then
		loopgroup = 0%>
<!-- 		<tr>
			<td align="right" colspan="10">Continues On The Next Page...</td>
		</tr> -->
		</table></td></tr>
		</table>
		
		<WxPrinter PageBreak>
		
		<table width="100%" cellpadding="2" cellspacing="0" border="0" height="100%">
		<tr><td <%if (lcase(trim(serviceclass))="lpls2" and pid <> "15") or serviceclassid= 36 then%><%else%>height="380"<%end if%> valign="top">
		<table width="640" border="0" align="center" cellpadding="5" cellspacing="2">
		<%select case utilityid
			case 3, 10, 6,1
			  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility")
			case else
			  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), calcintpeak
		end select
	end if
	loopgroup = loopgroup + 1
wend
%>
<tr bordercolor="#FFFFFF"><td colspan=<%if showusagedetails = false then %>"6"<%else%>"4"<%end if %>></td><td colspan="7"><hr noshade size="1"></td></tr>
<%select case utilityid
case 3, 10,6,1,4 'Water bill%>
<tr bordercolor="#FFFFFF">
	<td colspan="5"></td>
	<td align="center">Totals</td>
	<td align="right" colspan="3"><b><%=Formatnumber(tot_kwhused,2)%></b></td>
	<%if utilityid <> 1 and utilityid <> 4 then%>
	<td align="right"><b><%if coincidentflag then response.write formatcurrency(FormatNumber(tot_demand_C,2)) else response.write formatcurrency(FormatNumber(tot_demand_P,2))%></b></td>
   <%end if%>
</tr>
<%case else
	if extusage then
	%>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (KWH)</td>
		<td align="right">On</td>
		<td align="right">Off</td>
		<td align="right">Mid</td>
		<td align="right">Total</td>
	</tr>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right">&nbsp;</td>
		<td align="right"><%=Formatnumber(tot_kwhusedon,2)%></td>
		<td align="right"><%=Formatnumber(tot_kwhusedoff,2)%></td>
		<td align="right"><%=Formatnumber(tot_kwhusedint,2)%></td>
		<td align="right"><%=Formatnumber(tot_kwhused,2)%></td>
	</tr>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (KW)</td>
		<td align="right"><%=Formatnumber(tot_demand_p,2)%></td>
		<td align="right"><%=Formatnumber(tot_demandoff_p,2)%></td>
		<td align="right"><%=Formatnumber(tot_demandint_p,2)%></td>
		<td align="right"><%=Formatnumber(tot_demand_p+tot_demandint_p+tot_demandoff_p,2)%></td>
	</tr>
	<%else%>
	<tr bordercolor="#FFFFFF">
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
		<td align="right"><%=FormatNumber(tot_demand_P,2)%><%if coincidentflag then %>*<%end if%></b></td>
			<%if calcintpeak then%>
				<td align="right"><%if coincidentflag then response.write FormatNumber(tot_demandint_p,2) else response.write FormatNumber(tot_demandint_p,2)%></td>
				<td align="right"><%if coincidentflag then response.write FormatNumber(tot_demandoff_p,2) else response.write FormatNumber(tot_demandoff_p,2)%></td>
			<%end if%>
		<%end if%>
	</tr>
	<%if coincidentflag then %><tr><td colspan="6">*Totalized Demand:&nbsp;<%=FormatNumber(tot_demand_c,2)%> KW</td></tr><%end if%>
	<%
	end if
end select%>
</table>
<%if detailed="true" and utilityid<>3 then%>
<table width="640" align="center" border="0">
<tr><td><!--optional detail header can go here for consumption --></td>
	<td><!--optional detail header can go here for demand --></td></tr>
<tr><td valign="top"><%if rst2("energydetail")<>"" then%><%=replace(replace(replace(rst2("energydetail"),"|","<br>")," ","&nbsp;"),"_"," ")%><%end if%></td>
    <td valign="top"><%if rst2("demanddetail")<>"" then%><%=replace(replace(replace(rst2("demanddetail"),"|","<br>")," ","&nbsp;"),"_"," ")%><%end if%></td></tr>
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
<!-- end Totals section -->
</td>
</tr>
<tr><td valign="bottom" height="175">
<%
dim hidedemand
if trim(ucase(serviceclass))="AVG COST 1" or utilityid=3 or utilityid = 10 or utilityid = 1 or utilityid = 4 then hidedemand="true" else hidedemand = ""

if showusagedetails = false then %>
<table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">
<tr><td width="10%" align="center"><img src="http://<%=request.servervariables("SERVER_NAME")%>/genergy2/invoices/MakeChartyrly.asp?genergy2=<%=trim(request("genergy2"))%>&lid=<%=leaseid%>&by=<%=rst2("billyear")%>&bp=<%=rst2("billperiod")%>&billid=<%=billid%>&hidedemand=<%=hidedemand%>&building=<%=replace(building," ","%20")%>&unittype=<%=unittype%><%if extusage then %>&includepeaks=false&extusg=true<%else%>&includepeaks=<%=meterbreakdown%><%end if%>&calcintpeak=<%=calcintpeak%>" width="600" height="175"></td></tr>
</table>
<%'dim e
'e="http://"&request.servervariables("SERVER_NAME")&"/genergy2/invoices/MakeChartyrly.asp?genergy2=" &trim(request("genergy2"))&"&lid="&leaseid&"&by="&rst2("billyear")&"&bp="&rst2("billperiod")&"&billid="&billid&"&hidedemand="&hidedemand&"&building="&building&"&unittype="&unittype
'response.write e
'response.end%>
<%end if%>
</td>
</tr>
<tr><td><%footer rst2("tenantname"), rst2("tenantnum"), rst2("billingaddress"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("billingaddress"), rst2("btcity"), rst2("btstate"), rst2("billzip"), rst2("onlinebill"), textheader,demo, paymentterm%>
</td></tr>
</table>
<%if ucase(trim(serviceclass)) = "LPLS2" and detailed="true" then makeLPLS2totals billid%>

<WxPrinter PageBreak>
<%end if%>

<!-- </div> -->
<%

rst2.close
set rst2 = nothing
end sub

'###########################################################################################
sub mainheader(tenantname, tenantnum, tstrt, tcity, tstate, tzip, btbldgname, billingaddress, btcity, btstate, btzip, isonlinebill,textheader,demo, paymentterm)%>

	<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td width="40%" valign="top" ><font size="4"><u>Invoice Address:</u></font><br>
		<font size="4"><b><%=btbldgname%><br>
    <%
    %>
		<%=replace(rst2("billto"),vbNewLine,"<br>")%><br>
		<%=rst2("btcity")%>, <%=rst2("btstate")%>&nbsp;<%=rst2("billzip")%></b></font>
	  </td>
	  <td width="50%" valign="top" ><font size="4"><u>Supply Address:</u></FONT><br>
	  <font size="4"><b><%=tenantname%> (<%=tenantnum%>)<br>
	  <%=replace(tstrt,vbNewLine,"<br>")%><br>
	  <%=tcity%>, <%=tstate%>&nbsp;<%=tzip%></b></font>
	  </td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	</table>

<%end sub
sub footer(tenantname, tenantnum, tstrt, tcity, tstate, tzip, btbldgname, billingaddress, btcity, btstate, btzip, isonlinebill,textheader,demo, paymentterm)%>
	<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
  <td colspan="3">&nbsp;<br><center>PAYMENT DUE UPON RECEIPT<br><u>For any enquiry, please phone 020 7377 8944 or fax to 020 7247 2885</u><br>Registered in England 4476663<br>Registered  Address: 15 Devonshire Square London EC2M 4YW<br> VAT Registration Number: 798504675
<!--      <%'if trim(isonlinebill)="True" then%>To view online bill, login to www.genergyonline.com with access code <b><%'=tenantnum%>.<%'=building%></b>.<%'end if%>--></center></td></tr>
	</table>
<%end sub

sub newfooter()%>
	<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
  <td colspan="3">&nbsp;<br><center>PAYMENT DUE UPON RECEIPT<br><u>For any enquiry, please phone 020 7377 8944 or fax to 020 7247 2885</u><br>Registered in England 4476663<br>Registered  Address: 15 Devonshire Square London EC2M 4YW<br> VAT Registration Number: 798504675
<!--      <%'if trim(isonlinebill)="True" then%>To view online bill, login to www.genergyonline.com with access code <b><%'=tenantnum%>.<%'=building%></b>.<%'end if%>--></center></td></tr>
	</table>
<%end sub

sub metertableheads(billyear, billperiod, datestart, dateend, showintpeaks)%>
    <tr bgcolor="#eeeeee">
    	<td align="center">Period</td>
    	<td align="center">From</td>
    	<td align="center">To</td>
    	<td align="center">No.&nbsp;Days</td>
		<%if showUsagedetails = false then %><td colspan="2" width="20%" align="center">READINGS</td><%end if%>
    	<td <%if showusagedetails = false then %>colspan=<%if showintpeaks then%>"4"<%else%>"3"<%end if%> <%end if%> width="30%" align="center">CONSUMPTION</td>
    	<%if showDemanddetails = false then %> <td <%if showintpeaks then%>colspan="3"<%end if%> align="center">DEMAND</td><%end if%>
    </tr>
    <tr>
    	<td width="64" align="center"><%=billyear%>/<%=billperiod%></td>
    	<td width="64" align="center"><%=datestart%></td>
    	<td width="64" align="center"><%=dateend%></td>
    	<td width="64" align="center"><%=(dateend-datestart)+1%></td>
		<%if showUsagedetails = false then %>
    	<td bgcolor="#eeeeee" width="64" align="center">Previous</td>
    	<td bgcolor="#eeeeee" width="64" align="center">Current</td>
    	<td bgcolor="#eeeeee" width="64" align="center">On&nbsp;Peak</td>
      <%if showintpeaks then%>
      	<td bgcolor="#eeeeee" width="64" align="center">Int&nbsp;Peak</td>
      <%end if%>
    	<td bgcolor="#eeeeee" width="64" align="center">Off&nbsp;Peak</td>
		<%end if %>
    	<td bgcolor="#eeeeee" width="64" align="center">Total&nbsp;Usage</td>
	 <%if showDemanddetails = false then %>
    	<td bgcolor="#eeeeee" width="64" align="center"><%if showintpeaks then%>On<%else%>KW<%end if%></td>
      <%if showintpeaks then%>
        <td bgcolor="#eeeeee" width="64" align="center">Int</td>
        <td bgcolor="#eeeeee" width="64" align="center">Off</td>
      <%end if%>
	 <%end if %>
    </tr>
    <tr bordercolor="#FFFFFF">
    	<td align="center"></td>
    	<td colspan="2" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="center" bgcolor="#eeeeee">Multi.</td>
    	<td colspan="6" align="center"></td>
    </tr>
<%
end sub

sub metertableheadsExtUsage(billyear, billperiod, datestart, dateend)%>
    <tr bgcolor="#eeeeee">
    	<td align="center">Period</td>
    	<td align="center">From</td>
    	<td align="center">To</td>
    	<td align="center">No.&nbsp;Days</td>
		<td colspan="2" width="20%" align="center">READINGS</td>
		<td colspan="3" width="30%" align="center">CONSUMPTION</td>
    	<td colspan="3" align="center">DEMAND</td></tr>
    <tr>
    	<td width="64" align="center"><%=billyear%>/<%=billperiod%></td>
    	<td width="64" align="center"><%=datestart%></td>
    	<td width="64" align="center"><%=dateend%></td>
    	<td width="64" align="center"><%=dateend-datestart%></td>
    	<td bgcolor="#eeeeee" width="64" align="center">Previous</td>
    	<td bgcolor="#eeeeee" width="64" align="center">Current</td>
    	<td bgcolor="#eeeeee" width="64" align="center" colspan=3>&nbsp;KWH</td>
    	<td bgcolor="#eeeeee" width="64" align="right" colspan=3>KW</td>
    </tr>
    <tr bordercolor="#FFFFFF">
    	<td align="center"></td>
    	<td colspan="1" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="left" bgcolor="#eeeeee">Multi.</td>
    	<td colspan="8" align="center"></td>
    </tr>
<%end sub

sub miscCredits(numcells, tmpleaseid, billid)
	if cint(rst2("creditnonull")) <> 0 or cint(rst2("adjustmentnonull")) <> 0 then
		dim rstMiscCred, credSql
		credSql = "select isnull(description,'Misc Credit') as [desc], credit, convert(integer,adj) as adj FROM tblcreditbyperiod where bill_id="&billid&" and credit<>0 ORDER BY adj"
		set rstMiscCred = server.createobject("adodb.recordset")
		rstMiscCred.open credSql, getLocalConnect(building)
		'response.write credSql
		if not rstMiscCred.eof then
			do while not rstMiscCred.eof
				dim desc
				desc = rstMiscCred("desc")	%>
					<tr>
						<td align="right" width="15%"><%if rstMiscCred("adj")=1 then%>Adjustment:<%else%>Credit:<%end if%></td>
						<td align="right"><%=desc%></td>
						<%if numcells = 3 then%><td>&nbsp;</td><%end if%>
						<td <%if numcells <> 3 then%>width="15%"<%end if%> align="right"><%=formatcurrency(abs(cdbl(rstMiscCred("credit"))),2)%></td>
					</tr><%
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
				case 4
					 	usageHeader = "CF Used"
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
    	<td colspan="2" width="20%" align="center">READINGS</td>
		
    	<td colspan="<%if utilityid <> 1 and utilityid <> 4 then%>3<%else%>4<%end if%>"  width="30%" align="center">CONSUMPTION</td>
    	<%if utilityid <> 1 and utilityid <> 4 then %><td align="center" rowspan="2">Sub&nbsp;Total</td><%end if%>
    </tr>
    <tr>
    	<td width="64" align="center"><%=billyear%>/<%=billperiod%></td>
    	<td width="64" align="center"><%=datestart%></td>
    	<td width="64" align="center"><%=dateend%></td>
    	<td width="64" align="center"><%=dateend-datestart%></td>
    	<td bgcolor="#eeeeee" width="64" align="center">Previous</td>
    	<td bgcolor="#eeeeee" width="64" align="center">Current</td>
    	<td bgcolor="#eeeeee" align="right" colspan="<%if utilityid <> 1 and utilityid <> 4 then%>3<%else%>4<%end if%>"><%=usageHeader%></td>
    </tr>
    <tr bordercolor="#FFFFFF">
    	<td align="center"></td>
    	<td colspan="2" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="center" bgcolor="#eeeeee">Multi.</td>
    	<td colspan="6" align="center"></td>
    </tr>

<%end sub
	function getPropNumber (building)
	dim tempRec
	dim sql,ret
	ret = "000"
	sql = "select propertyID from custom_mch where bldgnum = '" & building & "'"
	set tempRec = server.createobject("adodb.recordset")
		tempRec.open sql, getLocalConnect(building)
	if not tempRec.eof then
		ret = trim(tempRec("propertyID"))
	end if
	getPropNumber = ret
	end function
	function replaceNull(val)
	Dim temp
	temp = val
	if isnull(val) or val ="" then temp = 0
	replaceNull = temp
	end function
%>