<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="./genergy2/secure.inc"-->
<%
dim pid, building, byear, bperiod, utilityid, mscroll, yscroll, yscroll2, reccount,rowColor, ischecked, tenantname, hilight, kwvartemp, kwhvartemp,kwhvarbold, dollvarbold,displaymode,checkopentickets, link, extusg, historic, showscroll, showscroll2

pid = request.querystring("pid")
building = trim(request("building"))
utilityid = request.querystring("utilityid")
if instr(request("bperiod"),"/")>0 then
	byear = split(request("bperiod"),"/")(1)
	bperiod = split(request("bperiod"),"/")(0)
else
	byear = request("byear")
	bperiod = request("bperiod")
end if
mscroll = request("mscroll")
yscroll = request("yscroll")
yscroll2 = request("yscroll2")
showscroll = lcase(trim(request("showscroll")))
showscroll2 = lcase(trim(request("showscroll2")))
if showscroll<>"inline" and showscroll<>"block" and showscroll<>"none" then showscroll = "block"
if showscroll2<>"inline" and showscroll2<>"block" and showscroll2<>"none" then showscroll2 = "none"

if utilityid = "" then utilityid = 0
if byear = "" then byear = 0
if bperiod = "" then bperiod = 0
if trim(mscroll)="" then mscroll = 0
if trim(yscroll)="" then yscroll = 0
if trim(yscroll2)="" then yscroll2 = 0
if lcase(request("historic"))="true" then historic=true else historic=false

dim rst1, cnn1, strsql, cnn2, cnn3
set rst1 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")

set cnn3 = server.createobject("ADODB.connection")
cnn1.open getConnect(pid,building,"billing")

if building<>"" then cnn3.open getLocalConnect(building)
dim DBlocalIP
if trim(building)<>"" then DBlocalIP = ""
%>
<html>
<head>
<title>Bill Validation</title>
<script>
var checkboxf = 0
function updatemeter(meterid, byear, bperiod, tnumber,tid, tname, posted)
{	if(checkboxf==0)
	{	var newwin = open('update_billentryG1.asp?tid='+tid+'&meterid='+meterid+'&byear='+byear+'&bperiod='+bperiod+'&tname='+tname+'&tnumber='+tnumber+'&building=<%=building%>&pid=<%=pid%>&utilityid=<%=utilityid%>&posted='+posted, 'update_billentryG1','left=8,top=8,scrollbars=yes,width=1024, height=380, status=no');
		newwin.focus();
	}
}

function scrollpoint(i){ 
	try{
		return(document.all['meterlist'][i].scrollTop);
	}catch(exception){
		try{
			return(document.all['meterlist'].scrollTop);
		}catch(exception){
			return(0);
		}
	}
}

function mainscrollpoint(){ 
	try{
		return(document.body.scrollTop);
	}catch(exception){
		return(0);
	}
}

function displaypoint(i){ 
	try{
		return(document.all[i].style.display);
	}catch(exception){
		return('none');
	}
}

function movemeterlist(i,y){ 
	try{
	    document.all['meterlist'][i].scrollTop = y;
	}catch(exception){
		try{
		    document.all['meterlist'].scrollTop = y;
		}catch(exception){}
	}
}

function nullfunction()
{
}

function clearSelects(sel)
{ var frm = document.forms['form1'];
  if((frm.building!=null)&&(sel=='pid'))frm.building.value="";
  if((frm.utilityid!=null)&&((sel=='pid')||(sel=='building')))frm.utilityid.value="";
  if((frm.byear!=null)&&((sel=='pid')||(sel=='building')||(sel=='utilityid')))frm.byear.value="";
  if((frm.bperiod!=null)&&((sel=='pid')||(sel=='building')||(sel=='utilityid')||(sel=='byear'))) frm.bperiod.value="";
}
function JumpTo(url){
	var frm = document.forms['form1'];
	var url = url + "?pid=<%=pid%>&bldg=<%=building%>&building=<%=building%>&utilityid=<%=utilityid%>&byear=<%=byear%>&bperiod=<%=bperiod%>";
	window.document.location=url;
}
function display(id){
	var tag = document.getElementById(id) 
	tag.style.display = (tag.style.display == "block" ? "none" : "block");
	var func = eval('document.all.d'+id)
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
}

//function loadperiod()
//{	var frm = document.forms['form1'];
//	if((frm.building.value!='')&&(frm.byear.value!='')&&(frm.bperiod.value!=''))
//	{	var newhref = "bill_validation.asp?pid="+frm.pid.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&utilityid="+frm.utilityid.value+"&bperiod="+frm.bperiod.value;
//		document.frames['mainval'].location=newhref;
//	}
//}

</script>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #dddddd; }
</style>
</head>
<body bgcolor="#FFFFFF" LINK="#000099" vlink="#000099" alink="#000099" onLoad="movemeterlist(0,<%=yscroll%>);movemeterlist(1,<%=yscroll2%>);window.scrollTo(0,<%=mscroll%>)">
<form name="form1" method="get" action="">
	
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr> 
      <td bgcolor="#6699cc"><span class="standardheader">Review/Edit+</span></td>
      <td align="right" bgcolor="#6699cc">
	<label style="border:1px solid #6699cc; color:white; font-weight: bold; border-bottom-style: solid;cursor:hand" onClick="
	document.location='validation_select.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&bperiod=<%=request("bperiod")%>&historic=<%if historic then%>false<%else%>true<%end if%>'
	" onMouseOver="this.style.borderColor='white';" onMouseOut="this.style.borderColor='#6699cc';" type="" src="" value="New Job">&nbsp;<%if historic then%>Hide<%else%>Show<%end if%>&nbsp;Historical&nbsp;Periods&nbsp;</label>
	<%if Not Cint(pid) = 7 Then %>
	  <% if building <> "" then %><select name="select" onChange="JumpTo(this.value)">
        <option value="#" selected>Jump to...</option>
        <option value="/genergy2/billing/processor_select.asp">Bill Processor</option>
        <option value="/genergy2/setup/buildingedit.asp">Building Setup</option>
        <option value="/genergy2/manualentry/entry_select.asp">Manual Entry</option>
        <option value="/genergy2/billentry/entry.asp">Utility Bill Entry</option>
        <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem 
        Report</option>
		<option value="/genergy2/validation/re_index.asp">Review/Edit+ v.2 Test</option>
      </select></td>
      <% end If%>
      </select><% end if%></td>
      
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td> 
              <%if allowgroups("Genergy Users") then%>
              <select name="pid" onChange="clearSelects(this.name);submit()">
                <option value="">Select Portfolio</option>
                <%
					'rst1.open "SELECT distinct id, name FROM portfolio p ORDER BY name", getConnect(0,0,"dbCore")
					 rst1.open "SELECT distinct id, name FROM portfolio p WHERE id="&pid,  getConnect(0,0,"dbCore")
									do until rst1.eof%>
                <option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then response.write " SELECTED"%>><%=rst1("name")%></option>
                <%rst1.movenext
									loop
									rst1.close%>
              </select> 
              <%elseif isnumeric(pid) then
								rst1.open "SELECT distinct name FROM portfolio p WHERE id="&pid, cnn1
								if not rst1.eof then response.write rst1("name")
								rst1.close%>
              <input type="hidden" name="pid" value="<%=pid%>"> 
              <%end if%>
            </td>
            <%if trim(pid)<>"" then%>
            <td> <select name="building" onChange="clearSelects(this.name);submit()">
                <option value="">Select Building</option>
                <%
									rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
									do until rst1.eof%>
                <option <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>> 
                <%=rst1("strt")%>, <%=trim(rst1("Bldgnum"))%> </option>
                <%rst1.movenext
									loop
									rst1.close
									%>
              </select> </td>
            <%end if
						if trim(building)<>"" then%>
            <td> <select name="utilityid" onChange="clearSelects(this.name);submit()">
                <option value="">Select Utility</option>
                <%
								dim tempSqlUtil
								tempSqlUtil = "SELECT DISTINCT byp.Utility as utilityid, u.Utilitydisplay FROM BillYrPeriod byp inner join dbo.tblutility u ON byp.Utility = u.utilityid WHERE (BldgNum = '" & trim(building) &"')"
								'tempSqlUtil = "SELECT distinct lup.Utility as utilityid, u.utilitydisplay FROM tblLeases l left outer join tblLeasesUtilityPrices lup ON l.BillingId = lup.BillingId inner join ["& application("superIP") &"].mainmodule.dbo.tblutility u ON lup.Utility = u.utilityid WHERE (l.BldgNum = '" & trim(building) &"') ORDER BY u.utilitydisplay"
								response.write tempSqlUtil
								rst1.open tempSqlUtil, cnn3
								do until rst1.eof%>
                <option value="<%=trim(rst1("utilityid"))%>"<%if trim(rst1("utilityid"))=trim(utilityid) then%>SELECTED<%end if%>> 
                <%=rst1("utilitydisplay")%> </option>
                <%rst1.movenext
								loop
								rst1.close
								%>
              </select> </td>
            <%end if
			if trim(utilityid)<>0 then%>
            <td>
				<select name="bperiod">
                <%
				strsql = "SELECT distinct cast(billperiod as varchar)+'/'+billyear as periodyear, billyear, billperiod FROM billyrperiod WHERE "
				if not(historic) then strsql = strsql & "billyear>=year(getdate())-1 and "
				strsql = strsql & "bldgnum='"&building&"' and utility="&utilityid&" order by billyear, billperiod"
				rst1.open strsql, cnn1
				if rst1.eof then
					%><option value="">No Billing Periods</option><%
				end if
				do until rst1.eof
					%><option value="<%=rst1("periodyear")%>"<%if trim(rst1("periodyear"))=trim(bperiod&"/"&byear) or (bperiod="0" and month(dateadd("m",-1,now))&"/"&year(dateadd("m",-1,now))=rst1("periodyear")) then response.write " SELECTED"%>><%=rst1("periodyear")%></option><%
					rst1.movenext
				loop
				rst1.close%>
				</select>
			</td>
            <td><input type="button" name="action" value="View" onClick="submit()"></td>
            <%end if%>
          </tr>
        </table></td>
    </tr>
  </table>
  	<input type="hidden" name="historic" value="<%=historic%>">
</form>
<!-- beginning of display -->
<%if trim(bperiod)<>0 then

	dim super, btabcolor, stabcolor, flagColor, flagColorHilight, procPage
	
	
	if allowgroups("GY_Supervisors_ES,IT Services") then
		super=true
	else
		super=false
	end if
	
	dim usage, demand, UBtable
	select case cint(utilityid)
		case 1
			usage = "Mlbs/hr"
			demand = "Mlbs"
			UBtable = "utilitybill_steam"
		case 2
			usage = "KWH"
			demand = "KW"
			UBtable = "utilitybill"
		case 3
			usage = "CCF"
			demand = "-"
			UBtable = "utilitybill_coldwater"
		case 4
			usage = "CF"
			demand = "-"
			UBtable = "utilitybill_gas"
		case 6
			usage = "Ton/hr"
			demand = "Tons"
			UBtable = "utilitybill_chilledwater"
		case else
			usage = "?"
			demand = "?"
	end select
	if super then
		'Supervisor
		stabcolor="#0099FF"
		btabcolor="#CCCCCC"
		flagColor = "#009900"
		flagColorHilight = "#009900"
		procPage = "super_process.asp"
		strsql = _
			"SELECT (case when (avgKWH=0 or avgKW=0 or kwhvarience > variance*100+8 or kwvarience > variance*100 or ((kwhOFFvarience > variance*100+8 or kwhINTvarience > variance*100+8 or kwOFFvarience > variance*100 or kwINTvarience > variance*100) and extusg=1) or AMTvarience > variance*100) then '0' else '1' end) as belowVarience, * FROM ("&vbcrlf&_
			"SELECT distinct m.meterid, m.meternum, m.extusg, m.variance, v.revdate, c.validate, c.svalidate, bbp.posted, m.bldgnum,c.[current], isNull(c.used,0) as kwhused, isNull(c.usedoff,0) as kwhoff, isNull(c.usedint,0) as kwhint, isNull(pd.demand,0) as demand, isNull(pd.demand_off,0) as demand_off, isNull(pd.demand_int,0) as demand_int, l.tenantnum, l.billingname, isnull(l.billingid,'') as billingid, isNull(bbp.totalamt,0) as totalamt, bbp.adminfee, bbp.sqft, v.biller, v.org_kwh, v.org_kw, case when bbp.sqft=0 then 0 else(bbp.demand/bbp.sqft)end as wsqft, lup.coincident,lup.coincident_peak, lup.leaseutilityid, isnull(cd.demand,0) as coindemand, isnull(avgKWH,0) as avgKWH, isnull(avgKWHoff,0) as avgKWHoff, isnull(avgKWHint,0) as avgKWHint, isNuLL(avgKW,0) as avgKW, isNuLL(avgKWoff,0) as avgKWoff, isNuLL(avgKWint,0) as avgKWint, isNuLL(avgAmt,0) as avgAmt, "&vbcrlf&_
			"isNull(case when isNull(avgKWH,0)=0 then '0' else abs((c.used - (isNull(avgKWH,0)))/isNull(avgKWH,0)*100) end, 0) as kwhvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWHoff,0)=0 then '0' else abs((c.usedoff - (isNull(avgKWHoff,0)))/isNull(avgKWHoff,0)*100) end, 0) as kwhOFFvarience, "&_
			"isNull(case when isNull(avgKWHint,0)=0 then '0' else abs((c.usedint - (isNull(avgKWHint,0)))/isNull(avgKWHint,0)*100) end, 0) as kwhINTvarience, "&_
			"isNull(case when isNull(avgKW,0)=0 then '0' else abs((pd.demand - (isNull(avgKW,0)))/isNull(avgKW,0)*100) end, 0) as kwvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWoff,0)=0 then '0' else abs((pd.demand_off - (isNull(avgKWoff,0)))/isNull(avgKWoff,0)*100) end, 0) as kwOFFvarience, "&_
			"isNull(case when isNull(avgKWint,0)=0 then '0' else abs((pd.demand_int - (isNull(avgKWint,0)))/isNull(avgKWint,0)*100) end, 0) as kwINTvarience, "&_
			"isNull(case when isNull(avgAmt,0)=0 then '0' else abs((bbp.totalamt - (isNull(avgAmt,0)))/isNull(avgAmt,0)*100) end, 0) as Amtvarience "
		if Cint(pid) = 7 and cint(utilityid)= 6 then
			strsql = strsql & ", isNull(cb.mintons,0) as mintons "
		end if
		strSql = strsql &vbcrlf&_
			"FROM consumption c "&_
			"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
			"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
			"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
			"INNER JOIN tblleases l on lup.billingid=l.billingid "
			
		If Cint(pid) = 7 and cint(utilityid)= 6 then
			strSql = strsql & "LEFT JOIN Custom_OucBill CB ON c.billyear=CB.billyear " & _
				" AND c.billperiod=CB.billperiod AND m.LeaseutilityId = CB.LeaseUtilityId  " 
		End If	
		strsql = strsql & _
			"LEFT JOIN ("&_
			"SELECT c2.meterid, isNull(avg(used),0) as avgKWH, avg(usedoff) as avgKWHoff, avg(usedint) as avgKWHint, isNull(avg(demand),0) as avgKW, avg(demand_off) as avgKWoff, avg(demand_int) as avgKWint FROM consumption c2 INNER JOIN peakdemand d2 ON d2.meterid=c2.meterid and c2.billyear=d2.billyear and c2.billperiod=d2.billperiod WHERE ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)) GROUP BY c2.meterid "&_
			") CAvg ON CAvg.meterid=m.meterid "&_
			"LEFT JOIN ("&_
			"SELECT leaseutilityid, isNull(avg(totalamt),0) as avgAmt FROM tblbillbyperiod WHERE ((billyear="&byear&"-1 and billperiod>="&bperiod&"+9)or(billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)) and reject=0 GROUP BY leaseutilityid "&_
			") BBAvg ON BBAvg.leaseutilityid=lup.leaseutilityid "&_
			"LEFT JOIN coincidentdemand cd on cd.leaseutilityid = lup.leaseutilityid and cd.billyear = c.billyear and cd.billperiod = c.billperiod "&_
			"LEFT JOIN tblbillbyperiod bbp on bbp.reject=0 and m.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
			"LEFT JOIN validation v on m.Meterid=v.Meterid and c.billyear=v.billyear and c.billperiod=v.billperiod "&_
			"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and 	m.bldgnum='"&building&"' and lup.utility="&utilityid&" and Online='1' and leaseexpired=0 "&_
			") final ORDER BY belowVarience, billingname, meternum, revdate desc"
	else 'Biller
		stabcolor="#CCCCCC"
		btabcolor="#6699cc"
		flagColor = "#cc0000"
		flagColorHilight = "#ff0000"
		procPage = "biller_processG1.asp"
		strsql = _
			"SELECT (case when (avgKWH=0 or avgKW=0 or kwhvarience > variance*100+8 or kwvarience > variance*100 or ((kwhOFFvarience > variance*100+8 or kwhINTvarience > variance*100+8 or kwOFFvarience > variance*100 or kwINTvarience > variance*100) and extusg=1) or AMTvarience > variance*100) then '0' else '1' end) as belowVarience, * FROM ("&vbcrlf&_
			"SELECT Distinct m.meterid, m.meternum, m.extusg, m.variance, c.validate, isNull(bbp.totalamt,0) as totalamt, bbp.posted, c.svalidate, m.bldgnum,c.[current], isNull(c.used,0) as kwhused, isNull(c.usedoff,0) as kwhoff, isNull(c.usedint,0) as kwhint, isNull(pd.demand,0) as demand, isNull(pd.demand_off,0) as demand_off, isNull(pd.demand_int,0) as demand_int, l.tenantnum, isnull(l.billingname,'') as billingname, isnull(l.billingid,'') as billingid,lup.coincident,lup.coincident_peak, isnull(cd.demand,0) as coindemand,lup.leaseutilityid, isnull(avgKWH,0) as avgKWH, isnull(avgKWHoff,0) as avgKWHoff, isnull(avgKWHint,0) as avgKWHint, isNuLL(avgKW,0) as avgKW, isNuLL(avgKWoff,0) as avgKWoff, isNuLL(avgKWint,0) as avgKWint, isNuLL(avgAmt,0) as avgAmt, "&vbcrlf&_
			"isNull(case when isNull(avgKWH,0)=0 then '0' else abs((c.used - (isNull(avgKWH,0)))/isNull(avgKWH,0)*100) end, 0) as kwhvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWHoff,0)=0 then '0' else abs((c.usedoff - (isNull(avgKWHoff,0)))/isNull(avgKWHoff,0)*100) end, 0) as kwhOFFvarience, "&_
			"isNull(case when isNull(avgKWHint,0)=0 then '0' else abs((c.usedint - (isNull(avgKWHint,0)))/isNull(avgKWHint,0)*100) end, 0) as kwhINTvarience, "&_
			"isNull(case when isNull(avgKW,0)=0 then '0' else abs((pd.demand - (isNull(avgKW,0)))/isNull(avgKW,0)*100) end, 0) as kwvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWoff,0)=0 then '0' else abs((pd.demand_off - (isNull(avgKWoff,0)))/isNull(avgKWoff,0)*100) end, 0) as kwOFFvarience, "&_
			"isNull(case when isNull(avgKWint,0)=0 then '0' else abs((pd.demand_int - (isNull(avgKWint,0)))/isNull(avgKWint,0)*100) end, 0) as kwINTvarience, "&_
			"isNull(case when isNull(avgAmt,0)=0 then '0' else abs((bbp.totalamt - (isNull(avgAmt,0)))/isNull(avgAmt,0)*100) end, 0) as Amtvarience "
			
			if Cint(pid) = 7 and cint(utilityid)= 6 then
				strsql = strsql & ", isNull(cb.mintons,0) as mintons "
			end if
			strSql = strsql &vbcrlf&_		
			"FROM consumption c "&_
			"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
			"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
			"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
			"INNER JOIN tblleases l on lup.billingid=l.billingid "
			If Cint(pid) = 7 and cint(utilityid)= 6 then
				strSql = strsql & "LEFT JOIN Custom_OucBill CB ON c.billyear=CB.billyear " & _
				" AND c.billperiod=CB.billperiod AND m.LeaseutilityId = CB.LeaseUtilityId  " 
			End If		
			strsql = strsql & _	
			"LEFT JOIN ("&_
			"SELECT c2.meterid, isNull(avg(used),0) as avgKWH, avg(usedoff) as avgKWHoff, avg(usedint) as avgKWHint, isNull(avg(demand),0) as avgKW, avg(demand_off) as avgKWoff, avg(demand_int) as avgKWint FROM consumption c2 INNER JOIN peakdemand d2 ON d2.meterid=c2.meterid and c2.billyear=d2.billyear and c2.billperiod=d2.billperiod WHERE ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)) GROUP BY c2.meterid "&_
			") CAvg ON CAvg.meterid=m.meterid "&_
			"LEFT JOIN ("&_
			"SELECT leaseutilityid, isNull(avg(totalamt),0) as avgAmt FROM tblbillbyperiod WHERE ((billyear="&byear&"-1 and billperiod>="&bperiod&"+9)or(billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)) and reject=0 GROUP BY leaseutilityid "&_
			") BBAvg ON BBAvg.leaseutilityid=lup.leaseutilityid "&_
			"LEFT JOIN coincidentdemand cd on cd.leaseutilityid = lup.leaseutilityid and cd.billyear = c.billyear and cd.billperiod = c.billperiod "&_
			"LEFT JOIN tblbillbyperiod bbp on bbp.reject=0 and m.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
			"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and 	m.bldgnum='"&building&"' and lup.utility="&utilityid&" and Online='1' and leaseexpired=0 "&_
			") final ORDER BY belowVarience, billingname, meternum"
	end if
	'response.write strsql
	'response.end
	
	dim displaydate, perioddates ' get start and end dates of current period for header display
	set perioddates = server.createobject("ADODB.recordset")
	perioddates.open "SELECT distinct DateStart, DateEnd FROM tblbillbyperiod WHERE reject=0 and bldgnum='"&building&"' and billperiod="&bperiod&" and utility="&utilityid&" and billyear="&byear, cnn3
	if not perioddates.EOF then
		displaydate = " ("&month(perioddates("DateStart"))&"/"&day(perioddates("DateStart"))&" - "&month(perioddates("DateEnd"))&"/"&day(perioddates("DateEnd"))&")"
	end if
	perioddates.close
	
	dim previousMeterid, isposted, needAcceptButton
	needAcceptButton = false
	
	
	dim numoftenants, numofmeters, numoftenantsPrev, numofmetersPrev 'need to fill these variables next section (assume are zero)
	numoftenants =0
	numofmeters =0
	numoftenantsPrev =0
	numofmetersPrev =0
	dim rst2, strsql2, prevbyear, prevbperiod
	set rst2 = server.createobject("ADODB.Recordset")
	prevbyear = byear
	prevbperiod = bperiod-1
	if prevbperiod<1 then 
		prevbperiod = 12
		prevbyear = prevbyear-1
	end if
	strsql2 = "SELECT (SELECT count(Distinct m.LeaseUtilityID) FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l WHERE m.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=m.leaseutilityid and m.bldgnum='"&building&"' and m.billperiod="&bperiod&" and m.billyear="&byear&" and l.utility="&utilityid&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c, tblbillbyperiod bbp, tblleasesutilityprices l WHERE c.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=c.leaseutilityid and c.bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and l.utility="&utilityid&") as meters"
	
	rst2.open strsql2, cnn3
	if not rst2.EOF then
		numofmeters = rst2("meters")
		numoftenants = rst2("tenants")
	end if
	rst2.close
	
	strsql2 = "SELECT (SELECT count(Distinct m.LeaseUtilityID) FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l WHERE m.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=m.leaseutilityid and m.bldgnum='"&building&"' and m.billperiod="&prevbperiod&" and m.billyear="&prevbyear&" and l.utility="&utilityid&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c, tblbillbyperiod bbp, tblleasesutilityprices l WHERE c.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=c.leaseutilityid and c.bldgnum='"&building&"' and c.BillYear="&prevbyear&" and c.BillPeriod="&prevbperiod&" and l.utility="&utilityid&") as meters"
	rst2.open strsql2, cnn3
	if not rst2.EOF then
		numofmetersPrev = rst2("meters")
		numoftenantsPrev = rst2("tenants")
	end if
	rst2.close
	
	dim prevbuildingAvgKW, prevbuildingAvgKWH, prevbuildingAvgBillAmt, avgBuildingCostKW, avgBuildingCostKWH, avgBuildingFuelAdj'now get these header fields as well
	if utilityid=2 then
	  strsql2 = "SELECT isNull(avg(ub.fuelAdj),0) as avgfueladj, isNull(avg(TotalKW),0) as avgTotalKW, isNull(avg(TotalKWH),0) as avgTotalKWH, isNull(avg(TotalBillAmt),0) as avgTotalBillAmt, isNull(avg(CostKW),0) as avgCostKW, isNull(avg(CostKWH),0) as avgCostKWH FROM tblBillByPeriod bbp INNER JOIN "&UBtable&" ub ON ub.ypid=bbp.ypid Where bbp.reject=0 and ((billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)or(billyear="&byear-1&" and billperiod>="&bperiod&"+9)) and bbp.utility="&utilityid&" and bbp.bldgnum='"& building &"'"
	  rst2.open strsql2, cnn3
	  'response.write strsql2
	  'response.end
	  if not rst2.EOF then
		prevbuildingAvgKW = rst2("avgTotalKW")
		prevbuildingAvgKWH = rst2("avgTotalKWH")
		prevbuildingAvgBillAmt = rst2("avgTotalBillAmt")
		avgBuildingCostKW = rst2("avgCostKW")
		avgBuildingCostKWH = rst2("avgCostKWH")
		avgBuildingFuelAdj = rst2("avgfueladj")
	  end if
	  rst2.close
	end if
	
	dim currentkw, currentkwh, currentcostkw, currentcostkwh, currentfueladj, buildingname, currentBillAmt 'totals (and building name for top)
	'dim aveKWH'averages
	if utilityid=2 then
	  strsql2 = "select (select distinct strt from buildings where bldgnum='"&building&"') as building, FuelAdj, sum(TotalKWH) as TotalKWH, case when sum(totalkwh)=0 then 0 else sum(CostKWH)/sum(TotalKWH) end as AvgKWH, sum(TotalKW) as TotalKW, sum(CostKWH) as CostKWH, sum(CostKW) as CostKW, sum(TotalBillAmt) as TotalBillAmt  FROM "&UBtable&" where ypid in (select ypid FROM billyrperiod where bldgnum='"&building&"' and Billyear="&byear&" and BillPeriod="&bperiod&") GROUP BY FuelAdj"
	  rst2.open strsql2, cnn3
	  if not rst2.EOF then
		buildingname = rst2("building")
		currentkw = rst2("TotalKW")
		currentkwh = rst2("TotalKWH")
		currentcostkw = rst2("CostKW")
		currentcostkwh = rst2("CostKWH")
		currentBillAmt = rst2("TotalBillAmt")
		currentfueladj = rst2("FuelAdj")
	  end if
	  rst2.close
	end if
	
	Dim opencriticalTickets, tcount
	strsql2 = "select count(*) as tcount from ["&Application("CoreIP")&"].dbCore.dbo.tickets where ticketfortype = 'bldgnum' and ticketfor = '" & building & "' and billyear = "&byear&" and billperiod = "&bperiod&" and closed = 0"
	
	rst2.open strsql2, cnn3
	if not rst2.EOF then 
		tcount = rst2("tcount")
		if tcount <> 0 then 
		opencriticalTickets = true
		else
		opencriticalTickets = false
		end if 
	end if
	rst2.close	
	
	dim bKWflag, bKWHflag, bCostKWflag, bCostKWHflag, tenantsFlag, metersFlag'the four building variance flags get set if building variance below is to high
	if prevbuildingAvgKW<>0 then if abs(currentkw-prevbuildingAvgKW)/prevbuildingAvgKW>.2 then bKWflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
	if prevbuildingAvgKWH<>0 then if abs(currentkwh-prevbuildingAvgKWH)/prevbuildingAvgKWH>.2 then bKWHflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
	if avgBuildingCostKW<>0 then if abs(currentcostkw-avgBuildingCostKW)/avgBuildingCostKW>.2 then bCostKWflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
	if avgBuildingCostKWH<>0 then if abs(currentcostkwh-avgBuildingCostKWH)/avgBuildingCostKWH>.2 then bCostKWHflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
	if numoftenants<>numoftenantsPrev then tenantsFlag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
	if numofmeters<>numofmetersPrev then metersFlag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
	
	%>
	<table border=0 cellspacing="0" cellpadding="3" width="97%" align="center">
		<tr valign="top"><td>
			<%="<b>"&request("processnote")&"</b>"%>
		</td></tr>
		<tr valign="top">
			<td>
				<table border=0 cellpadding="3" cellspacing="0" width="100%">
					<tr>
					<% if opencriticalTickets = false then %>
						<td bgcolor="#6699cc">
							<span class="standardheader"><%=buildingname%> &nbsp;&nbsp;Period&nbsp;<%=bperiod%><%=displaydate%>,&nbsp;<%=byear%></span>
						</td>
					<% else %>
						
          <td bgcolor="#FF0000"> <span class="standardheader"><%=buildingname%> 
            &nbsp;&nbsp;Period&nbsp;<%=bperiod%><%=displaydate%>,&nbsp;<%=byear%></span>  <b>[<a href="#" onClick="window.open('/genergy2_intranet/itservices/ttracker/troublesearch.asp?searchstring=<%=building%>&action=Search&searchbox=false&buildings=True','SearchNotes','width=800,height=400, scrollbars=no')">there are <%=tcount%> criticl tickets open for this building</a>]</b>
          </td>
					<% end if %>
					</tr>
				</table>
	
				<table border=0 cellpadding="3" cellspacing="1" width="100%">
					<tr bgcolor="#dddddd">
						<td width="8%">&nbsp;</td>
						<td width="10%" align="right"><%=demand%></td>
						<td width="10%" align="right"><%=usage%></td>
						<td width="10%" align="right">Cost <%=demand%></td>
						<td width="10%" align="right">Cost <%=usage%></td>
						<td width="10%" align="right">Fuel Adjustment</td>
						<td width="12%" align="right">Total Bill Amount</td>
						<td width="10%">&nbsp;</td>
						<td width="10%" align="right">Tenants&nbsp;Billed</td>
						<td width="10%" align="right">Meters&nbsp;Billed</td>
					</tr>
					<tr>
						<td align="right" class="tblunderline">Average:</td>
						<td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(prevbuildingAvgKW)%></td>
						<td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(prevbuildingAvgKWH,0)%></td>
						<td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(avgBuildingCostKW)%></td>
						<td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(avgBuildingCostKWH)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(avgBuildingFuelAdj,6)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(prevbuildingAvgBillAmt)%></td>
						<td align="right" bgcolor="#eeeeee">This&nbsp;Period:</td>
						<td bgcolor="#eeeeee" align="right"<%=tenantsFlag%>><a href="javascript:nullfunction()" onClick="window.open('tenantmeterlist.asp?building=<%=building%>&bperiod=<%=bperiod%>&byear=<%=byear%>&utilityid=<%=utilityid%>&checking=Tenant', '', 'toolbar=no,width=250,height=200, resizable=no,scrollbars=yes')"><%=numoftenants%></a></td>
						<td bgcolor="#eeeeee" align="right"<%=metersFlag%>><a href="javascript:nullfunction()" onClick="window.open('tenantmeterlist.asp?building=<%=building%>&bperiod=<%=bperiod%>&byear=<%=byear%>&utilityid=<%=utilityid%>&checking=Meter', '', 'toolbar=no,width=250,height=200, resizable=no,scrollbars=yes')"><%=numofmeters%></a></td>
					</tr>
					<tr>
						<td align="right" class="tblunderline">Current:</td>
						<td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(currentkw)%></td>
						<td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(currentkwh,0)%></td>
						<td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(currentcostkw)%></td>
						<td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(currentcostkwh)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(currentfueladj,6)%></td>
						<td class="tblunderline"align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(currentBillAmt)%></td>
						<td align="right" bgcolor="#eeeeee">Last&nbsp;Period:</td>
						<td bgcolor="#eeeeee" align="right"<%=tenantsFlag%>><%=numoftenantsPrev%></td>
						<td bgcolor="#eeeeee" align="right"<%=metersFlag%>><%=numofmetersPrev%></td>
					</tr>
				</table>
				
				<table border=0 cellpadding="3" cellspacing="0">
					<tr>
						<td>
						<%if not super then%>
							<b>Biller&nbsp;Validation</b>
						<%else%>
							<b>Supervisor&nbsp;Validation</b>
						<%end if%>
						</td>
					</tr>
				</table>
	
	<%
	dim cellw
	'if not isposted then cellw = "90" else 
	cellw = "70"
	%>
		<form name="form2" method="post" action="<%=procPage%>">
	<% 
	strsql2 = "SELECT distinct tname as Tenantname,tix.id as tnumber, l.billingid   FROM dbo.Consumption c, dbo.meters m,  tblleasesutilityprices l,tblleases tbL, ["&Application("CoreIP")&"].dbcore.dbo.tickets tix  WHERE c.meterid = m.meterid and l.leaseutilityid=m.leaseutilityid and tbl.billingid = l.billingid and m.bldgnum='"&building&"' and c.billperiod="&bperiod&" and c.billyear="&byear&" and l.utility="&utilityid&"  and ticketfortype = 'tid' and ticketfor = '"&split(getBuildingIP(building),"\")(1)&"-' + ltrim(convert(varchar(10),l.billingid)) and closed=0 and tix.billperiod = c.billperiod and tix.billyear = c.billyear"

	rst1.open strsql2, cnn3
	if not rst1.eof then 
	checkopentickets = true
	%>
		<table width="100%"><tr><td><span id="dticketlisting" name = "empty">[-]</span>&nbsp;&nbsp;<a href="#" onClick="display('ticketlisting')">Tenants with Open Trouble Tickets</a> [<span id="mTicketlist" name = "empty"></span>]</td></tr></table>
		<div id="ticketlisting" style="display:block;">
		<table width="100%" border=0 cellspacing="1" cellpadding="3">
			<tr bgcolor="#dddddd" valign="bottom">
				<td nowrap width="10%"><span class="standardheader"><font color="black">TT Number</font></span></td>
				<td width="40%"><span class="standardheader"><font color="black">Tenant Name</font></span></td>
				<td colspan=13>&nbsp;</td>
			</tr>
		</table>
	<div style="overflow:auto;height:300;border: 1px solid #cccccc;margin:3px;width:100%;">
		<table width="100%" border=0 cellspacing="1" cellpadding="3">
		<% 
		reccount = 0				
		do until rst1.eof 
		reccount = reccount + 1
		%>
		<tr valign="top" >
				<td nowrap width="10%"><span class="standardheader"><font color="black"><%=rst1("tnumber")%></font></span></td>
				<td width="40%"><span class="standardheader"><font color="black"><%=rst1("tenantname")%></font></span></td>
<%
			Link = "window.open('/genergy2/setup/tenantedit.asp?pid="&pid&"&bldg="&building&"&tid="&rst1("billingid")&"','TenantSetup','width=900,height=525,resizable=yes,toolbar=no,scrollbars=yes')"
%>				
				<td>&nbsp;<a href="#" onClick="<%=link%>">Account Setup</a>&nbsp;</td>
<%
			Link = "window.open('/genergy2_intranet/itservices/ttracker/ticket.asp?mode=update&tid="&rst1("tnumber")&"&child=1','TenantSetup','width=660,height=330,resizable=yes,toolbar=no,scrollbars=yes')"
			%>				
				<td colspan=12>&nbsp;<a href="#" onClick="<%=link%>">Open Trouble 
                  Ticket</a>&nbsp;</td>
			</tr>
		<% 
		rst1.movenext
		loop
		%>
		</table>
		</div>
		</div>
		<script>
			var func = eval('document.all.mTicketlist')
			func.innerHTML = '<%=reccount%>';
		</script>
	<% 
	else
		checkopentickets = false
	end if 
	rst1.close
	%>
		
		<table width="100%"><tr><td><span id="dabovevar" name = "empty"><% if checkopentickets then %>[+]<% else %>[-]<% end if%></span>&nbsp;&nbsp;<a href="#" onClick="try{display('abovevar')}catch(exception){}">Meters Above Allowable Set Variance</a> [<span id="mabovevar" name = "empty">0</span>]</td></tr></table>
		<%
		rst1.open strsql, cnn3
		if not rst1.eof then
			reccount= 0 
			if checkopentickets = true then displaymode = "none" else displaymode = "block" end if 	
			listmeters "abovevar", showscroll, 0
		end if 
%>

		<table width="100%"><tr><td><span id="dbelowvar" name = "empty">[+]</span>&nbsp;&nbsp;<a href="#" onClick="try{display('belowvar')}catch(exception){}">Meters Below Allowable Set Variance</a> [<span id="mbelowvar" name = "empty">0</span>]</td></tr></table>
		<%
		if not rst1.eof then
			reccount= 0	
			listmeters "belowvar",showscroll2, 1
		end if 
		rst1.close
%>
		<%if needAcceptButton and not(isBuildingOff(building)) then%>
			<table border=0 cellpadding="3" cellspacing="0" width="100%">
				<tr>
					<td style="padding-top:4px;" bgcolor="#eeeeee">
						<input type="submit" value="Accept" onClick="
						var f = document.forms[1];
						f.mscroll.value=mainscrollpoint();
						f.yscroll.value=scrollpoint(0);
						f.yscroll2.value=scrollpoint(1);
						f.showscroll.value=displaypoint('abovevar');
						f.showscroll2.value=displaypoint('belowvar');
						" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
					</td>
				</tr>
			</table>
		<%end if%>
	  <table cellpadding="3">
        <tr style="font-family: Arial, Helvetica, sans-serif;font-size:13"> 
          <td colspan="2"> <b>Bold</b> face indicates coincidental peak values </td>
        </tr>
      </table>
		<input type="hidden" name="byear" value="<%=byear%>">
		<input type="hidden" name="utilityid" value="<%=utilityid%>">
		<input type="hidden" name="bperiod" value="<%=bperiod%>">
		<input type="hidden" name="building" value="<%=building%>">
		<input type="hidden" name="pid" value="<%=pid%>">
		<input type="hidden" name="buildingname" value="<%=buildingname%>">
		<input type="hidden" name="mscroll" value="0">
		<input type="hidden" name="yscroll" value="0">
		<input type="hidden" name="yscroll2" value="0">
		<input type="hidden" name="showscroll" value="0">
		<input type="hidden" name="showscroll2" value="0">
		</form>
	
	</td>
	</tr>
	</table>

<%end if%>
</body>
</html>
<%
function listmeters(tag, displaymode, rmode)
	if not isnumeric(rmode) then rmode=0%>
	<div id="<%=tag%>" style="display:<%=displaymode%>;">
		<table width="100%" border=0 cellspacing="1" cellpadding="3">
			<tr bgcolor="#dddddd" valign="bottom">
			<td width="40"><font color="black">Accept</font></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Tenant Number</font></span></td>
			<td width="<%=cellw+15%>"><span class="standardheader"><font color="black">Tenant&nbsp;Name</font></span></td>
			<td width="<%=cellw+30%>"><span class="standardheader"><font color="black">Meter</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Average <%=usage%></font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Curr. <%=usage%> Usage</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Variance <%=usage%></font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Avg. <%=demand%></font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Curr. <%=demand%></font></span></td>
			<%if Cint(pid) = 7 and  cint(utilityid)= 6 then %>
				<td width="<%=cellw%>"><span class="standardheader"><font color="black">Cont. <%=demand%></font></span></td>
				<td width="<%=cellw%>"><span class="standardheader"><font color="black">Billed <%=demand%></font></span></td>
			<%End If %>			
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Variance <%=demand%></font></span></td>
			
			<td width="<%=cellw-15%>"><span class="standardheader"><font color="black">Bill Amount</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Average Amount</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Variance Amount</font></span></td>
			<%if super then%>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Updated <%=usage%>/<%=demand%></font></span></td>
			<%end if%>
			<td width="15">&nbsp;</td>
			</tr>
		</table>
		<div id="meterlist" style="overflow:auto;height:300;border: 1px solid #cccccc;margin:3px;width:100%;">
			<table width="100%" border="0" cellspacing="1" cellpadding="3">
				<%
				previousMeterid = ""
				do until rst1.eof 
					if cint(rmode)<>cint(rst1("belowVarience")) then exit do
					if lcase(rst1("extusg"))="true" then extusg = true else extusg = false
					reccount= reccount+ 1
					isposted = false
					if rst1("posted")="True" then isposted = true else needAcceptButton = true
					if previousMeterid<>trim(rst1("meterid")) then
						kwvartemp = rst1("kwvarience")
						if not isnumeric(trim(kwvartemp)) then kwvartemp = 0
						kwhvartemp = rst1("kwhvarience")
						if not isnumeric(trim(kwhvartemp)) then kwhvartemp = 0
						if cint(rst1("belowVarience"))<>1 then
							rowColor = flagColor
							kwhvarbold = "<b>"
							hilight = flagColorHilight
							ischecked = ""
						else
							kwhvarbold = ""
							rowColor = "#666666"
							hilight = "#DDDDDD"
							ischecked = " CHECKED"
						end if
						if isPosted and (cdbl(rst1("Amtvarience")) >= 25) then
							rowColor = flagColor
							hilight = flagColorHilight
							dollvarbold = "<b>"
						else
						dollvarbold = ""
					end if
					if rst1("validate")="True" and Not(super) then ischecked = " CHECKED"
					if rst1("svalidate")="True" and super then ischecked = " CHECKED"
					tenantname = rst1("billingname")
					if len(tenantname)>10 then
						tenantname = left(tenantname,10)&"..."
					end if
					%>
					<tr style="color:<%=rowColor%>;cursor:hand" valign="top" onClick="updatemeter('<%=rst1("meterid")%>', <%=byear%>, <%=bperiod%>, '<%=rst1("tenantnum")%>','<%=rst1("billingid")%>', '[<%=server.urlencode(rst1("billingname"))%>]', '<%=isposted%>')" onMouseOver="this.style.backgroundColor='<%=hilight%>';this.style.color='#ffffff';" onMouseOut="this.style.backgroundColor='#ffffff';this.style.color='<%=rowcolor%>';">
					<td width="40">
					<%if not isposted then%>
						<input type="checkbox" value="<%=rst1("meterid")%>" name="meters" onMouseOver="checkboxf=1" onMouseOut="checkboxf=0" style="cursor:auto" <%=ischecked%>>
					<%else%>
						S. Acc
					<%end if%>
					</td>
					<td width="<%=cellw%>"><nobr><%=rst1("tenantnum")%></nobr></td>
					<td width="<%=cellw%>"><nobr><%=tenantname%></nobr></td>
					<td width="<%=cellw%>"><%=rst1("meternum")%></td>
					<td width="<%=cellw%>" ><%if extusg then response.write "On:"%><%=formatnumber(rst1("avgKWH"),0)%>
					<%if extusg then%><br><nobr>Off:<%=formatnumber(rst1("AvgKWHoff"),0)%></nobr><br><nobr>Int:<%=formatnumber(rst1("AvgKWHint"),0)%></nobr><%end if%>
					</td>
					<td width="<%=cellw%>"><%if extusg then response.write "On:"%><%=formatnumber(rst1("kwhused"),0)%>
					<%if extusg then%><br><nobr>Off:<%=formatnumber(rst1("KWHoff"),0)%></nobr><br><nobr>Int:<%=formatnumber(rst1("KWHint"),0)%></nobr><%end if%>
					</td>
					<td width="<%=cellw%>"><%if extusg then response.write "On:"%><%=formatnumber(kwhvartemp,0)%>%
					<%if extusg then%><br><nobr>Off:<%=formatnumber(rst1("KWHoffvarience"),0)%>%</nobr><br><nobr>Int:<%=formatnumber(rst1("KWHintvarience"),0)%>%</nobr><%end if%>
					</td>
					
					<td width="<%=cellw%>"><%if extusg then response.write "On:"%>
					<%=formatnumber(rst1("AvgKW"),2)%>
					<%if lcase(trim(rst1("coincident"))) = "true" or lcase(trim(rst1("coincident_peak"))) = "true" then
						dim tempRST
						set tempRST = server.createobject("adodb.recordset")
						dim tempBPeriod1, tempBPeriod2, tempBYear1, tempBYear2
						tempBPeriod1 = bperiod - 1
						tempBYear1 = bYear
						if tempBPeriod1 <= 0 then
							tempBPeriod1 = tempBPeriod2 + 12
							tempBYear1 = tempBYear1 - 1
						end if
							tempBPeriod2 = tempBPeriod1 -1
							tempBYear2 = tempBYear1
						if tempBPeriod2 <= 0 then
							tempBPeriod2 = tempBPeriod2 + 12
							tempBYear2 = tempBYear2 - 1
						end if
						dim tempSQL
						tempSQL = "select isnull(avg(demand),0) as avgcoindemand from coincidentdemand where ((billyear= " & byear & " and billperiod = " & _
						bperiod & ") or  (billyear = " & tempbyear1 & " and billperiod = " & tempbperiod1 & ") or  (billyear = " & tempbyear2 & _
						" and billperiod = " & tempbperiod2 & ")) AND leaseutilityid = " & rst1("leaseutilityid")
						tempRST.open tempSQL, cnn3
						%><!--<%'=tempSQL'%>--><%
						%>; <b><%=formatnumber(temprst("avgcoindemand"),2)%></b><%
					end if%>
					<%if extusg then%><br><nobr>Off:<%=formatnumber(rst1("AvgKWoff"),2)%></nobr><br><nobr>Int:<%=formatnumber(rst1("AvgKWint"),2)%></nobr><%end if%>
					</td>
					<td width="<%=cellw%>">
					<%if extusg then response.write "On:"%><%=formatnumber(rst1("demand"),2)%>
					<%if lcase(trim(rst1("coincident"))) = "true" or lcase(trim(rst1("coincident_peak"))) = "true"then%>
						; <b><%=formatnumber(rst1("coindemand"),2)%></b>
					<%end if%>
					<%if extusg then%><br><nobr>Off:<%=formatnumber(rst1("demand_off"),2)%></nobr><br><nobr>Int:<%=formatnumber(rst1("demand_int"),2)%></nobr><%end if%>
					</td>
					<%if Cint(pid) = 7 and  cint(utilityid)= 6 then %>
						<td width="<%=cellw%>"><%=formatnumber(rst1("mintons"),2)%></td>
						<%
						If CDbl(rst1("mintons")) > Cdbl(rst1("demand")) Then %>
							<td width="<%=cellw%>"><%=formatnumber(rst1("mintons"),2)%></td>
						<%Else%>
							<td width="<%=cellw%>"><%=formatnumber(rst1("demand"),2)%></td>
						<%End If
						%>
					<%End If %>						
					<td width="<%=cellw%>"><%=kwhvarbold%><%if extusg then response.write "On:"%><%=formatnumber(kwvartemp,0)%>%
					<%if extusg then%><br><nobr>Off:<%=formatnumber(rst1("KWoffvarience"),0)%>%</nobr><br><nobr>Int:<%=formatnumber(rst1("KWintvarience"),0)%>%</nobr><%end if%>
					</td>
					<td width="<%=cellw%>"><%=formatcurrency(rst1("totalamt"),2)%></td>
					<td width="<%=cellw%>"><%=formatcurrency(rst1("avgAmt"),2)%></td>
					<td width="<%=cellw%>"><%=dollvarbold%><%=formatnumber(rst1("Amtvarience"),0)%>%</td>
					<%if super then%>
						<td width="<%=cellw%>">
						<%if rst1("validate")="True" then%>
							Accepted<br>
						<%end if%>
						<%if trim(rst1("biller"))<>"" then %>
							<%=rst1("biller")%><br><%=rst1("org_kwh")%><%=usage%>/ <%=rst1("org_kw")%><%=demand%>
						<%end if%>
						</td>
					<%end if%>
					<!--<td><%'=formatnumber(rst1("variance")*100,2)%>%</td>-->
					</tr>
					<%previousMeterid=trim(rst1("meterid"))
					end if
					rst1.movenext
				loop
				'response.write strsql
				%>
			</table>
			<script>
			var func = eval('document.all.m<%=tag%>')
			func.innerHTML = '<%=reccount%>';
			</script>
		</div>
	</div>
	<%
end function
%>