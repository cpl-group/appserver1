<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

'N.Ambo 5/21/2009 accomodated for new utility 'Condenser Water', utilityid 21
'Michelle T 5/27/2009 modify subtotal calculation on water bill and to eliminte detail bill information on a meter level
Server.ScriptTimeout = 60*15
dim leaseid, ypid, building, pid, byear, bperiod, utilityid, detailed, meterbreakdown, calcintpeak, SJPproperties, ShowUsageDetails, ShowDemandDetails,onlinereview,totdeliverychrgs, maxmeters, textheader, masterTotal,demo, reject, energydetail
dim pdfsession,startCount,endCount
leaseid = trim(Request("l"))
ypid = trim(request("y"))
utilityid = trim(request("utilityid"))
building = trim(request("building"))
'building = "260 Mad"
pid = trim(request("pid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
logo = trim(request("logo"))
detailed = trim(request("detailed"))
textheader = trim(request("textheader"))
demo = request("demo")
pdfsession = request("fdp")
startCount = request("s")
endCount = request("e")


if startCount = "" then startCount = 0
if (endCount = "" or endCount = "-1") then endCount = 1000


if demo = "" then demo = false end if 
if trim(request("reject"))="" then reject = 0 else reject = 1

if trim(request("SJPproperties"))="true" then SJPproperties = true else SJPproperties = false
if trim(request("meterbreakdown"))="no" then meterbreakdown = false else meterbreakdown = true
if trim(request("summaryusage"))="true" then showusagedetails= true else showusagedetails = false
if trim(request("summarydemand"))="true" then showdemanddetails = true else showdemanddetails = false
'response.Write(isempty(session("xmlUserObj")))
'response.End()

'if pdfsession = "pdffdp" then

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
	'response.write bldgrs("billid")
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
			'response.Write("here")
			bldgrs.movenext
		loop
	elseif byear<>"" and bperiod<>"" then
    sql = "SELECT b.id as billid, b.leaseutilityid, ypid, b.utility, extusg FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
    if isnumeric(utilityid) then sql = sql & " and utility="&utilityid
    sql = sql & "  ORDER BY TenantName"
		
		'response.Write(sql & " <p> ")
		bldgrs.open sql, cnn1
		dim count
		dim tempCount
		tempCount = 0
		count = startCount
			if count > 0 then 'skip to the proper bill
			bldgrs.movefirst
			for tempCount = 0 to count-1 Step 1
			bldgrs.movenext
			next
			end if
		do until bldgrs.eof
		'on Error resume next
			templid = trim(bldgrs("leaseutilityid"))
			tempypid = trim(bldgrs("ypid"))
			temputility = trim(bldgrs("utility"))
			Extusage = trim(bldgrs("extusg"))
			showtenantbill templid, bldgrs("billid"), temputility, extusage
			count = count + 1
				if count >= endCount+1 then 
				response.write "</body></html>"
				response.End()
				end if
			'response.write("here2")
			bldgrs.movenext
		loop
	end if
end if
set cnn1 = nothing
response.write "</body></html>"

'response.write Totalout




'### begin of showtenantbill, is rest of file ###

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
'response.write billid
select case utilityid
case 3, 10,6,1,4,21,22
	select case utilityid
	case 6,1,4,21
		usagedivisor = 1
	case else
		usagedivisor = 100
  end select 
	'sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto, r.addonfee as myaddonfee, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, (select case count(distinct isnull(addonfee,0)) when 0 then 0 else 1 end as aoncnt from tblmetersbyperiod where bill_id=r.id group by leaseutilityid,ypid,bill_id) as showaddonfee, r.invoice_note as invoiceNote, isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.*, l.onlinebill FROM tblbillbyperiod r, tblleases l, buildings b, "&DBmainIP&"ratetypes rt WHERE r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&billid
	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto, r.addonfee as myaddonfee, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, (select case count(distinct isnull(addonfee,0)) when 0 then 0 else 1 end as aoncnt from tblmetersbyperiod where bill_id=r.id group by leaseutilityid,ypid,bill_id) as showaddonfee, r.invoice_note as invoiceNote, isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.*, l.onlinebill FROM tblbillbyperiod r, tblleases l, buildings b, ratetypes rt WHERE r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&billid

case else
	if SJPproperties then
  	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_Bill_summary_NoBill r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	
	else
  	sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_bill_summary r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	
	end if
end select
'response.write sql
'response.end
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
		maxmeters = 9
end select

'11/15/2007 Added by N. Ambo
' For utilities 3,6, 10 the Sub Toal is being calculated by using energydetail which is actually a text value. This if statement was added
' to deal with 0 values
'Also change to line 551 and 280 due to same issue - modified to read the variable 'energydetail' rather than the direct field from the table

if IsNull(rst2("energydetail")) then 
   energydetail = 0
   'Response.Write "ENERGY DETAIL:   " & energydetail
   else 
   	energydetail = rst2("energydetail")
 end if						
				
'if cold water bill and build id = 270MA
'change Bill name to Cold Water/Sewer Bill - mod by Kamto Cheng 7/31/2008
if (trim(utilityname) = "Cold Water") And (trim(building) = "910") Then
    utilityname = "Cold Water/Sewer"
End If

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
<%if textheader="advRealty.asp" then%>
<tr><td align="center"><!--#INCLUDE FILE="advRealty.asp" --></td></tr>
<%end if%>
<tr><td 
<%if (lcase(trim(serviceclass))="lpls2" and pid <> "15") or serviceclassid= 36 then%><%else%>height="380"<%end if%> valign="top">
<!-- meterlisting -->
<table width="640" cellpadding="5" cellspacing="2" border="0" align="center">
<%if utilityname="Cold Water" and building ="240E38" then%>
<tr><td width="384" rowspan="2" colspan=<% if showusagedetails = false then %>"6"<%else%>"4"<%end if %>><font size="+1"><strong>Domestic Water Bill</strong></font></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Number</td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Date</td>
</tr>
<%else%>
<tr><td width="384" rowspan="2" colspan=<% if showusagedetails = false then %>"6"<%else%>"4"<%end if %>><font size="+1"><strong><%=utilityname%> Bill</strong></font></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Number</td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">Invoice Date</td>
</tr>
<%end if%>

<tr>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center">EL.<%=rst2("billperiod") & Right(rst2("billyear"),2) &"."& rst2("tenantnum") %></td>
	<td colspan="2" width="256" bgcolor="#eeeeee" align="center"><%if isdate(rst2("postdate")) then response.write formatdatetime(rst2("postdate"),2)%></td>
</tr>

<%
select case utilityid
case 3,10,6,1,4,21,22
  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility"), leaseid
case else
  if rst2("calcintpeak") then calcintpeak = true else calcintpeak = false
	  if extUsage then 
		  metertableheadsExtUsage rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend")
	  else
		  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), calcintpeak
	  end if
end select
Set rst1 = Server.CreateObject("ADODB.recordset")
rst1.open "select tblmetersbyperiod.*,dbo.Meters.Location from tblmetersbyperiod INNER JOIN dbo.Meters ON tblmetersbyperiod.MeterId = dbo.Meters.MeterId where bill_id="&billid, cnn1
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
		'response.write "Flag is: " &"<BR>"
		'response.write coincidentflag
		'response.end
		select case utilityid
			case 3, 10, 6,21,22
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
						<%'RESPONSE.Write "1:" & rst2("demanddetail") &"<br>"
							'RESPONSE.Write "2" &rst2("energydetail") 
						   'RESPONSE.END
						%>

						<%currentDemandP = cdbl(formatnumber(currentUsage*formatnumber(rst2("demanddetail"),6),2))+formatnumber(currentUsage*formatnumber(energydetail,6),2)%> 
						<!--<td align="right"><%=currentDemandP%></td>-->
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
							<td align="center"><%=trim(rst1("location"))%></td>
					    </tr>
					    <tr>
						<%end if%>
						<td align="right"></td>
						<td align="right"></td>
						<td align="right"></td>
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
					<td align="center"><%=Formatnumber(rst1("location"),1)%></td>
					</tr>
					<tr>
					
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
			case 3, 10, 6, 1,22
			  metertableheadsWater rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), rst2("utility"), leaseid
			case else
			  metertableheads rst2("billyear"), rst2("billperiod"), rst2("datestart")-1, rst2("dateend"), calcintpeak
		end select
	end if
	loopgroup = loopgroup + 1
wend
%>
<tr bordercolor="#FFFFFF"><td colspan=<%if showusagedetails = false then %>"6"<%else%>"4"<%end if %>></td><td colspan="7"><hr noshade size="1"></td></tr>
<%select case utilityid
case 3, 10,6,1,4,21,22 'Water bill%>
<tr bordercolor="#FFFFFF">
	<td colspan="5"></td>
	<!--Michelle T. 5/29/2009.   commented out to eliminate subtotal detail on a meter level-->
	<td align="center">Totals</td>
	<td align="right" colspan="3"><b><%=Formatnumber(tot_kwhused,2)%></b></td>       
	<%'if utilityid <> 1 and utilityid <> 4 then%> 
	<!--commented line below to accomodate for subtotal not calculating correctly, as per kimberly formula should be ((cw+sc)*cff) from water bill. Michelle T. 5/27/2009 -->
   <td align="right"><b><%'if coincidentflag then response.write formatcurrency(FormatNumber(tot_demand_C,2)) else response.write formatcurrency(FormatNumber(tot_demand_P,2))%></b></td>
   <td align="right"><b><%'if coincidentflag then response.write formatcurrency(FormatNumber(tot_demand_C,2)) else response.write formatcurrency(((cdbl(rst2("energydetail"))+cdbl(rst2("demanddetail"))) * tot_kwhused),2)%></b></td>
 
   
   
   
   <%'end if%>
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
	<%
	        dim rst7, cDemand
	        cDemand = 0
	        set rst7 = server.CreateObject("ADODB.recordset")
	       
	        rst7.open "Select demand from coincidentdemand where leaseutilityid="&leaseid&" and billyear="&byear&" and billperiod="&bperiod, cnn1
            
            if (NOT rst7.EOF) then
                cDemand = rst7("demand")
            else
                cDemand = 0    
            end if 
            rst7.close%>
	<tr bordercolor="#FFFFFF">
		<td colspan="5"></td>
		<td align="right" nowrap>Totals (Coin. Demand)</td>
		<td align="right"><%=Formatnumber(0,2)%></td>
		<td align="right"><%=Formatnumber(0,2)%></td>
		<td align="right"><%=Formatnumber(0,2)%></td>
		<td align="right"><%=Formatnumber(cDemand,2)%></td>
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
<!--added utiltiy <> 10 to accodmodate for hot water as well as cold Michelle T. 5/27/2009-->
<%if detailed="true" and utilityid<>3 and utilityid<>10 and utilityid<>6 and utilityid<>21 and utilityid<>22 then%>
<table width="640" align="center" border="0">
<tr><td><%if trim(rst2("energydetail"))<>"" then%><%if ucase(trim(serviceclass)) <> "SC4 RA1" and ucase(trim(serviceclass)) <> "SC4 RA2" then%>Consumption<%else%>Invoice<%end if%> Details:<%end if%></td>
	<td><%if trim(rst2("demanddetail"))<>"" then%>Demand Details:<%end if%></td></tr>
<tr><td valign="top"><%if rst2("energydetail")<>"" then%><%=replace(replace(rst2("energydetail"),"|","<br>")," ","&nbsp;")%><%end if%></td>
    <td valign="top"><%if rst2("demanddetail")<>"" then%><%=replace(replace(rst2("demanddetail"),"|","<br>")," ","&nbsp;")%><%end if%></td></tr>
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
<!-- </td></tr> --><tr><td height="105">
<!-- Totaling section -->
<%if 1=1 then'ucase(trim(serviceclass)) <> "LPLS2" or (pid = "15" and detailed<>"true") then%>
<table border=0 cellpadding="0" cellspacing="0" width="640" align="center">
<tr>
          <td width="50%" valign="top"> 
            <!-- grey box numbers -->
            <table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#eeeeee" valign="top">
    <%
    dim rst, cUnit, dUnit
    
    cUnit = ""
    dUnit = ""
    
    if (utilityid = 6 OR utilityid = 21) then
		    set rst = server.CreateObject("ADODB.recordset")
            rst.open "Select * from tblleasespecificmeasure where LeaseutilityId="&leaseid, cnn1
            
            if (NOT rst.EOF) then
                if ( rst("ConsumptionMeasure") <> "default") then
                    cUnit = rst("ConsumptionMeasure")
                end if 
                if ( rst("DemandMeasure") <> "default") then
                    dUnit = rst("DemandMeasure")
                end if
            end if
             
		end if 		
		
	if (cUnit = "") then cUnit = "Ton/Hr"
	if (dUnit = "") then dUnit = "Ton"
  
	select case utilityid
    case 3, 10, 6,21,22		'cold or hot water
	%>
      <table border=0 cellpadding="2" cellspacing="0" bgcolor="#eeeeee">
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td><%=serviceclass%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td>Service Fee</td>
        <td width="30">&nbsp;</td>
        <td><%if rst2("SHOWADDONFEE")=1 then %><%=FormatCurrency(rst2("myaddonfee"),2)%>/per meter<%else%><%=FormatCurrency(rst2("servicefeenull"),2)%><%end if%></td>
      </tr>
	  <tr bgcolor="#eeeeee">
        <td>Admin Fee</td>
        <td width="30">&nbsp;</td>
        <td><%=Formatnumber(rst2("adminfee")*100,2)%>%</td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td><%Select Case Utilityid
				case 3,22
					 %>CW per CCF<%
				case 6,21
					 %>Usage Cost Per <%=cUnit%>
				<%	 
				case 10
					%>HW per CCF<%
				end Select 
		%></td>
        <td width="30">&nbsp;</td>
        <td><%=FormatCurrency(energydetail,6)%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td><%Select Case Utilityid
				case 3,10,22
					 %>SC per CCF<%
				case 6,21
					 %>Demand Cost Per <%=dUnit%>
		        <%
				end Select 
		%></td>
        <td width="30">&nbsp;</td>
        <td><%=FormatCurrency(rst2("demanddetail"),6)%></td>
      </tr>
      </table>
    <%case 1, 4%>
	<table border=0 cellpadding="2" cellspacing="0" bgcolor="#eeeeee">
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td><%=serviceclass%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td>Service Fee</td>
        <td width="30">&nbsp;</td>
        <td><%if rst2("SHOWADDONFEE")=1 then %><%=FormatCurrency(rst2("myaddonfee"),2)%>/per meter<%else%><%=FormatCurrency(rst2("servicefeenull"),2)%><%end if%></td>
      </tr>
	  <tr bgcolor="#eeeeee">
        <td>Admin Fee</td>
        <td width="30">&nbsp;</td>
        <td><%=Formatnumber(rst2("adminfee")*100,2)%>%</td>
      </tr>
      </table>
	<%case else%>
      <table border=0 cellpadding="2" cellspacing="0" bgcolor="#eeeeee">
      
      <% if rst2("datestart") >= "4/1/2010" and ucase(trim(serviceclass)) = "SC4R1" then %>	
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td>SC9R1<%end if%></td>
      </tr>
       <% if rst2("datestart") >= "4/1/2010" and ucase(trim(serviceclass)) = "SC4RA1" then %>	
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td>SC9RA1<%end if%></td>
      </tr>
       <% if rst2("datestart") >= "4/1/2010" and ucase(trim(serviceclass)) = "SC4R2" then %>	
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td>SC9R2<%end if%></td>
      </tr>
       <% if rst2("datestart") >= "4/1/2010" and ucase(trim(serviceclass)) = "SC4RA2" then %>	
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td>SC9RA2<%end if%></td>
      </tr>
      <% if rst2("datestart") >= "4/1/2010" and (ucase(trim(serviceclass)) <> "SC4R1" and ucase(trim(serviceclass)) <> "SC4RA1" and ucase(trim(serviceclass)) <> "SC4R2" and ucase(trim(serviceclass)) <> "SC4RA2") then%>
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td><%=serviceclass%><%end if%></td>
      </tr>
      <% if rst2("datestart") < "4/1/2010" then%>
      <tr bgcolor="#eeeeee">
        <td>Service Class</td>
        <td width="30">&nbsp;</td>
        <td><%=serviceclass%><%end if%></td>
      </tr>
	  <% if ucase(trim(serviceclass)) <> "LPLS2" then %>	  
      <tr bgcolor="#eeeeee">
        <td>Service Fee</td>
        <td width="30">&nbsp;</td>
        <td><%if rst2("SHOWADDONFEE")=1 then %><%=FormatCurrency(rst2("myaddonfee"),2)%>/per meter<%else%><%=FormatCurrency(rst2("servicefee"),2)%><%end if%></td>
      </tr>
	  <% else %>
      <tr bgcolor="#eeeeee">
        <td>Tenant Service Fee</td>
        <td width="30">&nbsp;</td>
        <td><%if rst2("SHOWADDONFEE")=1 then %><%=FormatCurrency(rst2("myaddonfee"),2)%>/per meter<%else%><%=FormatCurrency(rst2("servicefee"),2)%><%end if%></td>
      </tr>
	  <% if cdbl(rst2("rate_servicefee")) > 0 then %>
      <tr bgcolor="#eeeeee">
        <td>Utility Service Fee</td>
        <td width="30">&nbsp;</td>
        <td><%=FormatCurrency(rst2("rate_servicefee"),2)%></td>
      </tr>  
	  <% 
	  	end if 
	  end if %>
      <tr bgcolor="#eeeeee">
                      <td>Admin Fee</td>
        <td width="30">&nbsp;</td>
                      <td><%=formatpercent(rst2("adminFee"))%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td>EL Adjust Factor</td>
        <td width="30">&nbsp;</td>
        <td><%if isnumeric(rst2("fAdj")) and not(isnull(rst2("fAdj"))) and trim(rst2("fAdj"))<>"0" then response.write formatnumber(rst2("fAdj"), 6) else response.write "NA"%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td>Average Cost KWH</td>
        <td width="30">&nbsp;</td>
        <td><%if isnumeric(rst2("akwhdisplay")) and cdbl(rst2("akwhdisplay"))<>0 then response.write FormatCurrency(rst2("akwhdisplay"),6) else response.write "NA"%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td>Average Cost KW</td>
        <td width="30">&nbsp;</td>
        <td><%if isnumeric(rst2("akwdisplay")) and cdbl(rst2("akwdisplay"))<>0 then response.write FormatCurrency(rst2("akwdisplay"),6) else response.write "NA"%></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td><%if isnumeric(rst2("unit_credit")) and trim(rst2("unit_credit"))<>"0" then response.write "LMEP "%>Rate</td>
        <td width="30">&nbsp;</td>
        <td><%if isnumeric(rst2("unit_credit")) and trim(rst2("unit_credit"))<>"0" then response.write rst2("unit_credit") else response.write "NA"%></td>
      </tr>
      </table>
    <%end select%>
  </td>
</tr>
</table>
<!-- end grey box numbers -->
  </td>
  <td width="50%" align="right" valign="top">
<!-- start Totals section -->
<%
dim subTotal
select case utilityid
case 3, 10, 6,21,22		'water utilities%>
  <table border="0" cellpadding="3" cellspacing="1">
  <tr>
  <%
 'subTotal = tot_demand_P
 'as per kimberly, subtotal in water bill detail does not caluculate accoding to formula, modify subtotal commented out in line above. Michelle T. 5/27/2009  
 
 subTotal=((cdbl(rst2("energydetail"))+cdbl(rst2("demanddetail"))) * formatnumber(tot_kwhused,2))


 
 %>
  	<td align="right" colspan=2>Sub Total</td>
  	<td width="20">&nbsp;</td>
    <td align="right"><%=FormatCurrency(subTotal,2)%></td>
  	 	<!--<td align="right"><%=FormatCurrency(rst2("subtotal"),2)%></td>-->
  </tr>
  <%if cdbl(rst2("rateservicefee_dollar"))>0 then%>
  <tr>
  	<td align="right" colspan=2>Rate&nbsp;Service&nbsp;Fee</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=formatcurrency(rst2("rateservicefee_dollar"),2)%></td>
  </tr>
  <%end if%>
  <%miscCredits 3,leaseid,billid%>
  <tr>
  
  	<td align="right" colspan=2>Admin Fee</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=formatcurrency((formatnumber(rst2("adminfee"),2)*subTotal),2)%></td>
  </tr>
  <%subTotal = ((subTotal + subTotal) * Formatnumber(rst2("adminfee"),2))%>
  <tr>
  	<td align="right" colspan=2>Service&nbsp;Fee</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=formatcurrency(cdbl(rst2("serviceFeenonull")))%>
	</td>
  </tr>

  <%subTotal = subTotal + cdbl(rst2("serviceFeenonull"))%>

  <tr>
  	<td align="right" colspan=2>Sub Total</td>
  	<td width="20">&nbsp;</td>
  	<!--<td align="right"><%=FormatCurrency(rst2("subtotal"),2)%></td>-->
  	<td align="right"><%=FormatCurrency(rst2("subTotal"),2)%>
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
<%case 1 'Steam
	Dim SteamBreakdown, i,details,tempval
	steamBreakdown = split(cstr(rst2("energydetail")),"|")
	%>
            <table border="0" cellpadding="3" cellspacing="1">
              <%
	  for each i in steamBreakdown
		details = split(i,",") 
	  %>
              <tr> 
                <td align="right" nowrap><%=details(1)%></td>
                <td align="right" nowrap><%=(split(details(0),"="))(0)%> =</td>
                <td align="right" nowrap><%=(split(details(0),"="))(1)%></td>
              </tr>
              <%
	  next 	
	  %>
              <tr> 
                <td align="right" nowrap>Consumption Sub Total</td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right" nowrap><%=formatcurrency(cdbl(rst2("energy")))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Rate Adjustment</td>
                <td align="right" nowrap>Consumption Sub Total @ <%=rst2("ratemodify")%> 
                  =</td>
                <%tempval = formatcurrency(cdbl(rst2("energy"))*cdbl(rst2("ratemodify")))%>
                <td align="right" nowrap><%=tempval%></td>
              </tr>
              <%miscCredits 2,leaseid, billid%>
              <tr>
                <td align="right" nowrap>Admin Fee</td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right" nowrap><%=formatnumber(rst2("adminfeedollar"),2)%></td>
              </tr>
              <tr>
                <td align="right" nowrap>Service Fee</td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right" nowrap><%=formatcurrency(cdbl(rst2("servicefeenonull")))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Sub Total </td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right" nowrap><%=formatcurrency(rst2("subtotal"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Sales Tax</td>
                <td align="right" nowrap>@ <%=formatnumber(rst2("salestax"),5)%> 
                  =</td>
                <td align="right" nowrap><%=formatcurrency(rst2("tax"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap><strong>Total&nbsp;Due</strong></td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right"  nowrap><strong><%=formatcurrency(rst2("totalamt"))%></strong></td>
              </tr>
            </table>
		  <%masterTotal = rst2("Totalamt")%>
<%case else
	Select Case serviceclassid
	case 36 
		sql = "select * from custom_JCPLbill where bill_id="&billid
		rst3.open sql, cnn1
		if rst3.EOF then 
			response.write "<div align=center>RATE DETAILS ARE UNAVAILABLE</div>"
		else
		totdeliverychrgs =0
		totdeliverychrgs = cdbl(rst3("deliv_b1_kwh_chrg"))+cdbl(rst3("deliv_b1_kw_chrg"))+cdbl(rst3("deliv_b2_kwh_chrg"))+cdbl(rst3("deliv_b2_kw_chrg"))
		%>
			<table border="0" cellpadding="3" cellspacing="1">
              <tr> 
                <td align="right" nowrap>Customer Charge</td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right" nowrap><%=formatcurrency(cdbl(rst3("customer_chrg")))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>BGS Energy Charges</td>
                <td align="left" nowrap><%=Formatnumber(tot_kwhused,2)%> KWH x 
                  <%=rst3("energy_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(cdbl(rst3("energy_chrg")))%></td>
              </tr>
              <%miscCredits 3, leaseid,billid%>
              <tr> 
                <td align="right" nowrap>BGS Transmission Charges</td>
                <td align="left" ><%=Formatnumber(tot_kwhused,2)%> KWH x <%=rst3("trans_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("trans_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap><strong>Sub Total</strong></td>
                <td align="left" nowrap><hr noshade size="1"></td>
                <td align="right" nowrap><strong><%=formatcurrency(cdbl(rst3("customer_chrg"))+cdbl(rst3("energy_chrg"))+cdbl(rst3("trans_chrg")))%></strong></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Delivery Charges</td>
                <td align="left" nowrap><%=rst3("deliv_b1_kwh")%> KWH x <%=rst3("deliv_b1_kwh_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("deliv_b1_kwh_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>&nbsp;</td>
                <td align="left" nowrap><%=rst3("deliv_b2_kwh")%> KWH x <%=rst3("deliv_b2_kwh_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("deliv_b2_kwh_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>&nbsp;</td>
                <td align="left" nowrap><%=rst3("deliv_b1_kw")%> KW x <%=rst3("deliv_b1_kw_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("deliv_b1_kw_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>&nbsp;</td>
                <td align="left" nowrap><%=rst3("deliv_b2_kw")%> KW x <%=rst3("deliv_b2_kw_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("deliv_b2_kw_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap><strong>Total Delivery Charges</strong></td>
                <td align="right" ><hr noshade size="1"></td>
                <td align="right"  nowrap><strong><%=formatcurrency(totdeliverychrgs)%></strong></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Market Transition Charges</td>
                <td align="right" nowrap><%=Formatnumber(tot_kwhused,2)%> KWH 
                  x <%=rst3("market_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("market_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Societal Benefits Charges</td>
                <td align="right" nowrap><%=Formatnumber(tot_kwhused,2)%> KWH 
                  x <%=rst3("societal_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("societal_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap>Transitional Assessment Charges</td>
                <td align="right" nowrap><%=Formatnumber(tot_kwhused,2)%> KWH 
                  x <%=rst3("tefa_rate")%></td>
                <td align="right" nowrap><%=formatcurrency(rst3("tefa_chrg"))%></td>
              </tr>
              <tr> 
                <td align="right" nowrap><strong>Total Due</strong></td>
                <td align="right" nowrap>&nbsp;</td>
                <td align="right" nowrap><strong><%=formatcurrency(rst2("Totalamtnonull"))%></strong></td>
              </tr>
            </table>		
		  <%masterTotal = rst2("Totalamtnonull")%>
		<%
		end if
		rst3.close
		
	case else
  %>
  <table border="0" cellpadding="3" cellspacing="1">
  <tr>
  <%
  subTotal = (cDbl(rst2("energynonull"))+cDbl(rst2("demandnonull")))
  %>
  	<td align="right" colspan=2>Sub Total</td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><%=FormatCurrency(subTotal,2)%></td>
  </tr>
	  <%if SJPproperties then%>
		<tr>
			<td align="right" colspan=2><%if ucase(trim(serviceclass)) = "LPLS2" then%>Tenant <%end if%>Service&nbsp;Fee</td>
			<td width="20">&nbsp;</td>
			<td align="right"><%=FormatCurrency((rst2("servicefeenonull")),2)%></td>
		</tr>
	  <%end if%>
  <%
    IF UTILITYID <> 4 THEN
		if not isnumeric(rst2("unit_credit")) or trim(rst2("unit_credit"))="0" then
	  miscCredits 3, leaseid,billid
	end if
  END IF

	
  %>
  <tr>
  	<td align="right" colspan=2>Admin Fee</td>
  	<td width="20">&nbsp;</td>
  <%dim adminfeecalc
  if trim(rst2("Adminfeedollar"))<>"" then  adminfeecalc = cdbl(rst2("Adminfeedollar")) else adminfeecalc = 0
  
  %> 
  	<td align="right"><%=FormatCurrency(adminfeecalc,2)%></td>
  </tr>
  <%if not(SJPproperties) then%>
    <tr>
    	<td align="right" colspan=2>Service&nbsp;Fee</td>
    	<td width="20">&nbsp;</td>
    	<td align="right"><%=FormatCurrency((rst2("servicefeenonull")),2)%></td>
    </tr>
  <%end if%>
  <%subTotal = subTotal + adminfeecalc + cDbl(rst2("servicefeenonull")) - cDbl(rst2("credit"))%>
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
<%
  IF UTILITYID <> 4 THEN 
   if isnumeric(rst2("unit_credit")) and trim(rst2("unit_credit"))<>"0" then
	  miscCredits 3, leaseid,billid
	end if
  END IF
%>
  <tr>
  	<td align="right" colspan=2><font size="+1"><b>Total&nbsp;Due</b></font></td>
  	<td width="20">&nbsp;</td>
  	<td align="right"><font size="+1"><b><%=FormatCurrency(rst2("Totalamtnonull"),2)%></b></font></td>
  </tr>
  </table>
  <%masterTotal = rst2("Totalamtnonull")%>
<%
	end select
end select%>
<%if ucase(trim(serviceclass)) = "LPLS2" and detailed="true" then response.write "Continues On The Next Page..."%>
  </td>
</tr>
</table>
<%else
'	response.write "MAK2" & serviceclass
end if%>
<!-- end Totals section -->
</td>
</tr>
<tr><td valign="bottom" height="175">
<%
dim hidedemand
if trim(ucase(serviceclass))="AVG COST 1" or utilityid=3 or utilityid = 10 or utilityid = 1 or utilityid = 4 then hidedemand="true" else hidedemand = ""
if showusagedetails = false then %>
<table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">
<tr><td width="10%" align="center"><img src="http://<%=request.servervariables("SERVER_NAME")%>/genergy2/invoices/MakeChartyrly.asp?genergy2=<%=trim(request("genergy2"))%>&lid=<%=leaseid%>&by=<%=rst2("billyear")%>&bp=<%=rst2("billperiod")%>&billid=<%=billid%>&hidedemand=<%=hidedemand%>&building=<%=building%>&unittype=<%=unittype%><%if extusage then %>&includepeaks=false&extusg=true<%else%>&includepeaks=<%=meterbreakdown%><%end if%>&calcintpeak=<%=calcintpeak%>" width="600" height="175"></td></tr>
</table>
<%end if%>
</td>
</tr>
<tr><td valign="top">
<%footer rst2("tenantname"), rst2("tenantnum"), rst2("billingaddress"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("billingaddress"), rst2("btcity"), rst2("btstate"), rst2("billzip"), rst2("onlinebill"), textheader,demo, paymentterm%>
</td></tr>
</table>
<%
'if ucase(trim(serviceclass)) = "LPLS2" and detailed="true" then makeLPLS2totals billid 
if ucase(trim(serviceclass)) = "LPLS2-111 RIVER" and detailed="true" then makeLPLS2totals billid%>

<WxPrinter PageBreak>
<%end if%>

<!-- </div> -->
<%

rst2.close
set rst2 = nothing
end sub

'###########################################################################################
sub footer(tenantname, tenantnum, tstrt, tcity, tstate, tzip, btbldgname, billingaddress, btcity, btstate, btzip, isonlinebill,textheader,demo, paymentterm)%>
	<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr><td colspan="3" align="center"><font size="2"><% if textheader="advRealty.asp" then %>KEEP ABOVE PART FOR YOUR RECORDS<%end if%>&nbsp;</font></td></tr>
  <tr><td colspan="3" bgcolor="black" height="1"><img src="/images/spacer.gif" width="1" height="1"></td></tr>
  <tr><td colspan="3" align="center"><font size="2"><% if textheader="advRealty.asp" then %>DETACH HERE AND RETURN WITH PAYMENT<%end if%></font>&nbsp;<br></td></tr>
	<tr>
<% if demo then %>
		<td width="50%" valign="top"><font size="1">Tenant Name and Address:</FONT><br>
	  <font size="3"><b>Demo Tenant (<%=tenantnum%>)<br>
	  123 Main Street<br>
	  <%=tcity%>, <%=tstate%>&nbsp;<%=tzip%></b></font>
  </td>
		<td width="40%" valign="top"><font size="1">Make Check Payable To:</font><br>
		<font size="3"><b>Demo Management Company or Owner<br>
    <%
    %>
		456 Main Street<br>
		<%=rst2("btcity")%>, <%=rst2("btstate")%>&nbsp;<%=rst2("billzip")%><br>
	  <%=paymentterm%></b><br></font>
	  </td>
<%
	else
	%>
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
		<%=rst2("btcity")%>, <%=rst2("btstate")%>&nbsp;<%=rst2("billzip")%><br>
	  <%=paymentterm%></b><br></font>
	  </td>
		<% if textheader="advRealty.asp" then %><td width="10%" valign="top" align="right" nowrap><br><strong><font size="3"><u>TOTAL DUE: <%=FormatCurrency(masterTotal,2)%></u><font></strong><br></td><% end if %>
	<% end if%>
	</tr>
  <tr>
  <% 
    Dim txtcontanct,txtcontanct2
	txtcontanct="If you have any questions about your bill, please call CPL Energy Management Services @ 212-664-7600 ext 2 or email us at rb@cplems.com"
	txtcontanct2=""

	if textheader<>"advRealty.asp" then %><td colspan="3">&nbsp;<br><center><%=txtcontanct2%><br>&nbsp;<br><center><%=txtcontanct%><br>
      <%if trim(isonlinebill)="True" then%>To view online bill, login to www.genergyonline.com with access code <b><%=tenantnum%>.<%=building%></b>.<%end if%></center></td><%end if%></tr>
	</table>
<%end sub

sub metertableheads(billyear, billperiod, datestart, dateend, showintpeaks)%>
    <%if trim(rst2("tenantname")) = "Davis Polk & Wardwell" then %>
	 <tr bgcolor="#eeeeee">
        <td colspan="3" align="center"><%=trim(rst2("tenantname"))%></td>
    </tr> <%
	end if %>
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
    	<td width="64" align="center"><%=dateend-datestart%></td>
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
    	<td colspan="1" bgcolor="#eeeeee" align="center">Meter No.</td>
    	<td align="center" bgcolor="#eeeeee">Multi.</td>
    	<td align="left" bgcolor="#eeeeee">Location</td>
    	<td colspan="7" align="center"></td>
    </tr>
<%
end sub



sub metertableheadsExtUsage(billyear, billperiod, datestart, dateend)%>
    <%if trim(rst2("tenantname")) = "Davis Polk & Wardwell" then %>
	 <tr bgcolor="#eeeeee">
        <td colspan="3" align="center"><%=trim(rst2("tenantname"))%></td>
    </tr> <%
	end if %>
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
    	<td align="left" bgcolor="#eeeeee">Location</td>
    	<td colspan="7" align="center"></td>
    </tr>
<%end sub

sub miscCredits(numcells, tmpleaseid, billid)
	'response.write "HereCreditnonull" &rst2("creditnonull") &"<BR>"
	'response.write "Hereadjustmentnonull" &rst2("adjustmentnonull")
'response.end
     '  if cint(rst2("creditnonull")) <> 0 or cint(rst2("adjustmentnonull")) <> 0 then
      if clng(rst2("creditnonull")) <> 0 or cint(rst2("adjustmentnonull")) <> 0 then
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

sub metertableheadsWater(billyear, billperiod, datestart, dateend, utilityid, leaseid)
		dim usageHeader
		Select Case Utilityid
				case 3,10,22
					 	usageHeader = "CCF Used"
				case 4
					 	usageHeader = "CF Used"
				case 6,21
						usageHeader = "Ton/Hrs"
				case 1 
						usageHeader = "MLBS Used"
				end Select 
				
		if (Utilityid = 6 OR Utilityid = 21) then
		    dim rst
		    set rst = server.CreateObject("ADODB.recordset")
            rst.open "Select * from tblleasespecificmeasure where LeaseutilityId="&leaseid, cnn1
            
            if (NOT rst.EOF) then
                if ( rst("ConsumptionMeasure") <> "default") then
                    usageHeader = rst("ConsumptionMeasure")
                end if
            end if
             
		end if 		
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

sub meterlist()
				
end sub

sub makeLPLS2totals(billid)
	Set rst3 = Server.CreateObject("ADODB.recordset")
	
	sql = "select * from custom_psegbill where bill_id="&billid
	rst3.open sql, cnn1
	if rst3.EOF then 
		response.write "<div align=center>RATE DETAILS ARE UNAVAILABLE</div>"
	else
		Dim subTotalDelivery,subTotalSupply,rst3
		subTotalDelivery = 0
		subTotalSupply = 0
		%>
<WxPrinter PageBreak>

		<%while not rst3.EOF%>
		<div align="center">
		<table width="640" border="0" cellspacing="0" cellpadding="0">
          <tr valign="bottom"> 
            <td height="50" colspan = 5 >&nbsp;</td>
          </tr>
          <tr> 
            <td colspan = 5 bgcolor="#eeeeee" ><strong>Rate Details</strong></td>
          </tr>
          <tr> 
            <td colspan = 5 >&nbsp;</td>
          </tr>
          <tr> 
            <td colspan = 5 bgcolor="eeeeee" >Delivery</td>
          </tr>
          <tr> 
            <td  colspan=4>Service Charge</td>
            <td width="6%"  align="right"><%=formatcurrency(rst2("servicefeenonull"),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + cdbl(rst2("servicefeenonull"))%>
          <tr> 
            <td colspan=5 >Distribution Charges</td>
          </tr>
          <tr> 
            <td width="40%">&nbsp;&nbsp;Annual Demand</td>
            <td width="17%" align="right"><%=cdbl(rst3("measuredemand"))%>&nbsp;</td>
            <td width="5%">KW</td>
            <td width="32%" >@ <%=cdbl(rst3("annualrate"))%></td>
            <td  align="right"><%=formatcurrency(cdbl(rst3("measuredemand")) * cdbl(rst3("annualrate")),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + (cdbl(rst3("measuredemand")) * cdbl(rst3("annualrate")))%>
          <% if formatnumber(rst3("summerrate")) > 0 then %>
          <tr> 
            <td >&nbsp;&nbsp;Summer Demand</td>
            <td align="right" ><%=cdbl(rst3("summerdemand"))%>&nbsp;</td>
            <td >KW </td>
            <td >@ <%=cdbl(rst3("summerrate"))%></td>
            <td  align="right"><%=formatcurrency(cdbl(rst3("summerdemand")) * cdbl(rst3("summerrate")),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + (cdbl(rst3("summerdemand")) * cdbl(rst3("summerrate")))%>
          <%end if%>
          <tr> 
            <td >&nbsp;&nbsp;KWH - On-Peak</td>
            <td align="right" ><%=cdbl(rst3("onpeak"))%>&nbsp;</td>
            <td >KWH </td>
            <td >@ <%=cdbl(rst3("delivery_on"))%></td>
            <td  align="right"><%=formatcurrency(cdbl(rst3("onpeak"))*cdbl(rst3("delivery_on")),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + (cdbl(rst3("onpeak"))*cdbl(rst3("delivery_on")))%>
          <tr> 
            <td >&nbsp;&nbsp;KWH - Off-Peak</td>
            <td align="right" ><%=cdbl(rst3("offpeak"))%>&nbsp;</td>
            <td >KWH</td>
            <td >@ <%=cdbl(rst3("delivery_off"))%></td>
            <td  align="right"><%=formatcurrency(cdbl(rst3("offpeak"))*cdbl(rst3("delivery_off")),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + (cdbl(rst3("offpeak"))*cdbl(rst3("delivery_off")))%>
          <tr> 
            <td >Societal benefits</td>
            <td align="right" ><%=cdbl(rst3("offpeak")) + cdbl(rst3("onpeak"))%>&nbsp;</td>
            <td >KWH</td>
            <td >@ <%=cdbl(rst3("soc_benefit"))%></td>
            <td  align="right"><%=formatcurrency((cdbl(rst3("offpeak")) + cdbl(rst3("onpeak")))*cdbl(rst3("soc_benefit")),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + ((cdbl(rst3("offpeak")) + cdbl(rst3("onpeak")))*cdbl(rst3("soc_benefit")))%>
          <tr> 
            <td >Securitization transition</td>
            <td align="right" ><%=cdbl(rst3("offpeak")) + cdbl(rst3("onpeak"))%>&nbsp;</td>
            <td >KWH</td>
            <td >@ <%=cdbl(rst3("sec_trans"))%></td>
            <td  align="right"><%=formatcurrency((cdbl(rst3("offpeak")) + cdbl(rst3("onpeak")))*cdbl(rst3("sec_trans")),2)%></td>
          </tr>
          <% subTotalDelivery = subTotalDelivery + ((cdbl(rst3("offpeak")) + cdbl(rst3("onpeak")))*cdbl(rst3("sec_trans")))%>
          <tr> 
            <td colspan=5>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan=3>&nbsp;</td>
            <td>Sub-Total Delivery</td>
            <td align="right"><%=formatcurrency(subTotalDelivery,2)%></td>
          </tr>
          <tr> 
            <td colspan=5>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan=5 bgcolor="eeeeee">Supply</td>
          </tr>
          <tr> 
            <td colspan=5>BGS Capacity</td>
          </tr>
          <tr> 
            <td>&nbsp;&nbsp;Generation</td>
            <td align="right">&nbsp;<%=cdbl(rst3("generationkw"))%>&nbsp;</td>
            <td>KW</td>
            <td>@ <%=cdbl(rst3("generationcost"))%></td>
            <td align="right">&nbsp;<%=formatcurrency(cdbl(rst3("generationkw")) * cdbl(rst3("generationcost")),2)%></td>
          </tr>
          <% subTotalSupply = subTotalSupply + (cdbl(rst3("generationkw")) * cdbl(rst3("generationcost")))%>
          <tr> 
            <td>&nbsp;&nbsp;Transmission</td>
            <td align="right">&nbsp;<%=cdbl(rst3("transmissionkw"))%>&nbsp;</td>
            <td>KW</td>
            <td>@ <%=cdbl(rst3("transmissioncost"))%></td>
            <td align="right">&nbsp;<%=formatcurrency(cdbl(rst3("transmissionkw")) * cdbl(rst3("transmissioncost")),2)%></td>
          </tr>
          <% subTotalSupply = subTotalSupply + (cdbl(rst3("transmissionkw")) * cdbl(rst3("transmissioncost")))%>
          <tr> 
            <td colspan=5>BGS Energy</td>
          </tr>
          <tr> 
            <td>&nbsp;&nbsp;On-Peak</td>
            <td align="right">&nbsp;<%=cdbl(rst3("onpeak"))%>&nbsp;</td>
            <td>KWH</td>
            <td>@ <%=cdbl(rst3("supply_on"))%></td>
            <td align="right">&nbsp;<%=formatcurrency(cdbl(rst3("onpeak"))*cdbl(rst3("supply_on")),2)%></td>
          </tr>
          <% subTotalSupply = subTotalSupply + (cdbl(rst3("onpeak"))*cdbl(rst3("supply_on")))%>
          <tr> 
            <td>&nbsp;&nbsp;Off-Peak</td>
            <td align="right">&nbsp;<%=cdbl(rst3("offpeak"))%>&nbsp;</td>
            <td>KWH</td>
            <td>@ <%=cdbl(rst3("supply_off"))%></td>
            <td align="right">&nbsp;<%=formatcurrency(cdbl(rst3("offpeak"))*cdbl(rst3("supply_off")),2)%></td>
          </tr>
          <tr> 
            <td colspan=5>&nbsp;</td>
          </tr>
          <tr> 
            <% subTotalSupply = subTotalSupply + (cdbl(rst3("offpeak"))*cdbl(rst3("supply_off")))%>
            <td colspan=3>&nbsp;</td>
            <td>Sub-Total Supply</td>
            <td align="right"><%=formatcurrency(subTotalSupply,2)%></td>
          </tr>
          <tr> 
            <td colspan=3>&nbsp;</td>
            <td>&nbsp;</td>
            <td align="right">&nbsp;</td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td colspan=3>&nbsp;</td>
            <td>Sub-Total Electric Charges</td>
            <td align="right">&nbsp;<%=formatcurrency(formatnumber(rst2("subtotal")),2)%></td>
          </tr>
          <tr> 
            <td colspan=5 height="235">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan=5></td>
          </tr>
          <%if abs(formatnumber(rst2("creditnonull"))) > 0 then%>
          <%end if%>
          <%if formatnumber(rst2("adminfee")) > 0 then%>
          <%end if%>
          <%if abs(formatnumber(rst2("creditnonull"))) > 0 or formatnumber(rst2("adminfee")) then%>
          <%end if%>
        </table>
		</div>
		<%
		rst3.movenext		
		wend
		rst3.close
	end if
end sub
%>