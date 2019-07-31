<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="./genergy2/secure.inc"-->
<%
dim pid, building, byear, bperiod, utilityid, mscroll, yscroll, yscroll2, reccount,rowColor, ischecked, tenantname, hilight, kwvartemp, kwhvartemp,kwhvarbold, dollvarbold,displaymode,checkopentickets, link, extusg, historic, showscroll, showscroll2
dim gloTenant,gloTenantNum
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
Dim t
t = request("t")
if t = "" then t="t"
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
<title>Review/Edit + V.2.1</title>
<script>
var checkboxf = 0
function updatemeter(meterid, byear, bperiod, tnumber,tid, tname, posted)
{	if(checkboxf==0)
	{	var newwin = open('update_billentry.asp?t=<%=t%>&tid='+tid+'&meterid='+meterid+'&byear='+byear+'&bperiod='+bperiod+'&tname='+tname+'&tnumber='+tnumber+'&building=<%=building%>&pid=<%=pid%>&utilityid=<%=utilityid%>&posted='+posted, 'update_billentry','left=8,top=8,scrollbars=yes,width=1024, height=380, status=no');
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
	var func = document.getElementById('d_'+id)
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
}

function switchView(tag)
{
	
	if ("<%=t%>" == tag.value) return;
	var url ="re_index.asp?pid=<%=pid%>&bldg=<%=building%>&building=<%=building%>&utilityid=<%=utilityid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&t=" + tag.value;
	window.document.location=url;
}


function showall(treename){
	var func = document.getElementById('toggle_'+treename) 
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
	var displaytype = (func.innerHTML != '[-]' ? 'none':'block');
	var tag = document.all//('note162');
	for (i = 0; i < tag.length; i++){
		if (tag[i].name == 'tenantlist_'+treename) tag[i].style.display = displaytype;
		if (tag[i].name == 'tenantlist_titles_'+treename) tag[i].innerHTML = func.innerHTML;
	} 
}

var ol_fgcolor = "#ffffcc";
var ol_bgcolor = "#ffffcc";
var ol_textsize = "11px";

</script>
<script type="text/javascript" src="\includes\overlib.js"><!-- overLIB (c) Erik Bosrup --></script>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #dddddd; }
</style>
</head>
<body bgcolor="#FFFFFF" LINK="#000099" vlink="#000099" alink="#000099" onLoad="movemeterlist(0,<%=yscroll%>);movemeterlist(1,<%=yscroll2%>);window.scrollTo(0,<%=mscroll%>)">
<div id="overDiv" style="font:10px/12px Arial,Helvetica,sans-serif; border:solid 1px #666666; width:270px; padding:1px; position:absolute; visibility:hidden; color:#333333; top:80px; left:90px; background-color:#ffffcc; z-index:1000;"></div>
<!--Selection Bar Start -->
<form name="form1" method="get" action="">	
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr> 
      <td bgcolor="#6699cc"><span class="standardheader">Review/Edit+ v.2.1</span></td>
      <td align="right" bgcolor="#6699cc">
	  <% if building <> "" then %><select name="select" onChange="JumpTo(this.value)">
        <option value="#" selected>Jump to...</option>
        <option value="/genergy2/billing/processor_select.asp">Bill Processor</option>
        <option value="/genergy2/setup/buildingedit.asp">Building Setup</option>
        <option value="/genergy2/manualentry/entry_select.asp">Manual Entry</option>
        <option value="/genergy2/billentry/entry.asp">Utility Bill Entry</option>
        <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem 
        Report</option>
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
                <%rst1.open "SELECT distinct id, name FROM portfolio p ORDER BY name", getConnect(0,0,"dbCore")
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
			<td>	<label style="border:1px solid #6699cc; color:black; font-weight: bold; border-bottom-style: solid;cursor:hand" onClick="
	document.location='validation_select.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&bperiod=<%=request("bperiod")%>&historic=<%if historic then%>false<%else%>true<%end if%>'
	" onMouseOver="this.style.borderColor='green';" onMouseOut="this.style.borderColor='#6699cc';" type="" src="" value="New Job">&nbsp;<%if historic then%>Hide<%else%>Show<%end if%>&nbsp;Historical&nbsp;Periods&nbsp;</label></td>
            <%end if%>
          </tr>
        </table></td>
    </tr>
  </table>
  	<input type="hidden" name="historic" value="<%=historic%>">
</form>
<!--Selection Bar End -->
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
	dim pagelink
	dim mainSql
	if super then
		'Supervisor
		flagColor = "#009900"
		flagColorHilight = "#009900"
		procPage = "super_process.asp"
		pagelink = "s"
		mainSql = _
			"SELECT (case when (avgKWH=0 or avgKW=0 or kwhvarience > variance*100+8 or kwvarience > variance*100 or ((kwhOFFvarience > variance*100+8 or kwhINTvarience > variance*100+8 or kwOFFvarience > variance*100 or kwINTvarience > variance*100) and extusg=1) or AMTvarience > variance*100) then '0' else '1' end) as belowVarience, * FROM ("&vbcrlf&_
			"SELECT distinct m.meterid, m.meternum, m.extusg, m.variance, v.revdate, c.validate, c.svalidate, bbp.posted, m.bldgnum,c.[current], isNull(c.used,0) as kwhused, isNull(c.usedoff,0) as kwhoff, isNull(c.usedint,0) as kwhint, isNull(pd.demand,0) as demand, isNull(pd.demand_off,0) as demand_off, isNull(pd.demand_int,0) as demand_int, l.tenantnum, l.billingname, isnull(l.billingid,'') as billingid, isNull(bbp.totalamt,0) as totalamt, bbp.adminfee, bbp.sqft, v.biller, v.org_kwh, v.org_kw, case when bbp.sqft=0 then 0 else(bbp.demand/bbp.sqft)end as wsqft, lup.coincident,lup.coincident_peak, lup.leaseutilityid, isnull(cd.demand,0) as coindemand, isnull(avgKWH,0) as avgKWH, isnull(avgKWHoff,0) as avgKWHoff, isnull(avgKWHint,0) as avgKWHint, isNuLL(avgKW,0) as avgKW, isNuLL(avgKWoff,0) as avgKWoff, isNuLL(avgKWint,0) as avgKWint, isNuLL(avgAmt,0) as avgAmt, "&vbcrlf&_
			"isNull(case when isNull(avgKWH,0)=0 then '0' else abs((c.used - (isNull(avgKWH,0)))/isNull(avgKWH,0)*100) end, 0) as kwhvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWHoff,0)=0 then '0' else abs((c.usedoff - (isNull(avgKWHoff,0)))/isNull(avgKWHoff,0)*100) end, 0) as kwhOFFvarience, "&_
			"isNull(case when isNull(avgKWHint,0)=0 then '0' else abs((c.usedint - (isNull(avgKWHint,0)))/isNull(avgKWHint,0)*100) end, 0) as kwhINTvarience, "&_
			"isNull(case when isNull(avgKW,0)=0 then '0' else abs((pd.demand - (isNull(avgKW,0)))/isNull(avgKW,0)*100) end, 0) as kwvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWoff,0)=0 then '0' else abs((pd.demand_off - (isNull(avgKWoff,0)))/isNull(avgKWoff,0)*100) end, 0) as kwOFFvarience, "&_
			"isNull(case when isNull(avgKWint,0)=0 then '0' else abs((pd.demand_int - (isNull(avgKWint,0)))/isNull(avgKWint,0)*100) end, 0) as kwINTvarience, "&_
			"isNull(case when isNull(avgAmt,0)=0 then '0' else abs((bbp.totalamt - (isNull(avgAmt,0)))/isNull(avgAmt,0)*100) end, 0) as Amtvarience "&vbcrlf&_
			"FROM consumption c "&_
			"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
			"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
			"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
			"INNER JOIN tblleases l on lup.billingid=l.billingid "&_
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
		pagelink = "b"
		flagColor = "#cc0000"
		flagColorHilight = "#ff0000"
		procPage = "biller_process.asp"
	mainSql	= _
			"SELECT (case when (avgKWH=0 or avgKW=0 or kwhvarience > variance*100+8 or kwvarience > variance*100 or ((kwhOFFvarience > variance*100+8 or kwhINTvarience > variance*100+8 or kwOFFvarience > variance*100 or kwINTvarience > variance*100) and extusg=1) or AMTvarience > variance*100) then '0' else '1' end) as belowVarience, * FROM ("&vbcrlf&_
			"SELECT Distinct m.meterid, m.meternum, m.extusg, m.variance, c.validate, isNull(bbp.totalamt,0) as totalamt, bbp.posted, c.svalidate, m.bldgnum,c.[current], isNull(c.used,0) as kwhused, isNull(c.usedoff,0) as kwhoff, isNull(c.usedint,0) as kwhint, isNull(pd.demand,0) as demand, isNull(pd.demand_off,0) as demand_off, isNull(pd.demand_int,0) as demand_int, l.tenantnum, isnull(l.billingname,'') as billingname, isnull(l.billingid,'') as billingid,lup.coincident,lup.coincident_peak, isnull(cd.demand,0) as coindemand,lup.leaseutilityid, isnull(avgKWH,0) as avgKWH, isnull(avgKWHoff,0) as avgKWHoff, isnull(avgKWHint,0) as avgKWHint, isNuLL(avgKW,0) as avgKW, isNuLL(avgKWoff,0) as avgKWoff, isNuLL(avgKWint,0) as avgKWint, isNuLL(avgAmt,0) as avgAmt, "&vbcrlf&_
			"isNull(case when isNull(avgKWH,0)=0 then '0' else abs((c.used - (isNull(avgKWH,0)))/isNull(avgKWH,0)*100) end, 0) as kwhvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWHoff,0)=0 then '0' else abs((c.usedoff - (isNull(avgKWHoff,0)))/isNull(avgKWHoff,0)*100) end, 0) as kwhOFFvarience, "&_
			"isNull(case when isNull(avgKWHint,0)=0 then '0' else abs((c.usedint - (isNull(avgKWHint,0)))/isNull(avgKWHint,0)*100) end, 0) as kwhINTvarience, "&_
			"isNull(case when isNull(avgKW,0)=0 then '0' else abs((pd.demand - (isNull(avgKW,0)))/isNull(avgKW,0)*100) end, 0) as kwvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWoff,0)=0 then '0' else abs((pd.demand_off - (isNull(avgKWoff,0)))/isNull(avgKWoff,0)*100) end, 0) as kwOFFvarience, "&_
			"isNull(case when isNull(avgKWint,0)=0 then '0' else abs((pd.demand_int - (isNull(avgKWint,0)))/isNull(avgKWint,0)*100) end, 0) as kwINTvarience, "&_
			"isNull(case when isNull(avgAmt,0)=0 then '0' else abs((bbp.totalamt - (isNull(avgAmt,0)))/isNull(avgAmt,0)*100) end, 0) as Amtvarience "&vbcrlf&_
			"FROM consumption c "&_
			"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
			"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
			"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
			"INNER JOIN tblleases l on lup.billingid=l.billingid "&_
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

	dim displaydate, perioddates ' get start and end dates of current period for header display
	set perioddates = server.createobject("ADODB.recordset")
	perioddates.open "SELECT distinct DateStart, DateEnd FROM tblbillbyperiod WHERE reject=0 and bldgnum='"&building&"' and billperiod="&bperiod&" and utility="&utilityid&" and billyear="&byear, cnn3
	if not perioddates.EOF then
		displaydate = " ("&month(perioddates("DateStart"))&"/"&day(perioddates("DateStart"))&" - "&month(perioddates("DateEnd"))&"/"&day(perioddates("DateEnd"))&")"
	end if
	perioddates.close
	
	dim previousMeterid, isposted, needAcceptButton
	needAcceptButton = false
	
	
	dim numoftenants, numofmeters, numoftenantsPrev, numofmetersPrev,totalbillamt,totalbillamtprev 'need to fill these variables next section (assume are zero)
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
'	strsql2 = "SELECT 'Cur' as rectype,(SELECT count(Distinct m.LeaseUtilityID) FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l WHERE m.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=m.leaseutilityid and m.bldgnum='"&building&"' and m.billperiod="&bperiod&" and m.billyear="&byear&" and l.utility="&utilityid&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c, tblbillbyperiod bbp, tblleasesutilityprices l WHERE c.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=c.leaseutilityid and c.bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and l.utility="&utilityid&") as meters, isnull((SELECT sum(totalamt) FROM tblbillbyperiod bbp, tblleasesutilityprices l WHERE bbp.reject=0 and l.leaseutilityid=bbp.leaseutilityid and bbp.bldgnum='"&building&"' and bbp.BillYear="&byear&" and bbp.BillPeriod="&bperiod&" and l.utility="&utilityid&"),0) as TotalAmt union SELECT 'Prev' as rectype,(SELECT count(Distinct m.LeaseUtilityID) FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l WHERE m.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=m.leaseutilityid and m.bldgnum='"&building&"' and m.billperiod="&prevbperiod&" and m.billyear="&prevbyear&" and l.utility="&utilityid&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c, tblbillbyperiod bbp, tblleasesutilityprices l WHERE c.bill_id=bbp.id and bbp.reject=0 and l.leaseutilityid=c.leaseutilityid and c.bldgnum='"&building&"' and c.BillYear="&prevbyear&" and c.BillPeriod="&prevbperiod&" and l.utility="&utilityid&") as meters,isnull((SELECT sum(totalamt) FROM tblbillbyperiod bbp, tblleasesutilityprices l WHERE bbp.reject=0 and l.leaseutilityid=bbp.leaseutilityid and bbp.bldgnum='"&building&"' and bbp.BillYear="&prevbyear&" and bbp.BillPeriod="&prevbperiod&" and l.utility="&utilityid&"),0) as TotalAmt"
'	
'	rst2.open strsql2, cnn3
'	if not rst2.EOF then
'	While not rst2.eof	
'		select case lcase(trim(rst2("rectype")))
'		case "cur" 
'			numofmeters = rst2("meters")
'			numoftenants = rst2("tenants")
'			totalbillamt = rst2("totalamt")
'		case "prev"
'			numofmetersPrev = rst2("meters")
'			numoftenantsPrev = rst2("tenants")
'			totalbillamtprev = rst2("totalamt")
'		end select 
'	rst2.movenext
'	wend
'	end if
'	rst2.close
	
	dim prevbuildingAvgKW, prevbuildingAvgKWH, prevbuildingAvgBillAmt, avgBuildingCostKW, avgBuildingCostKWH, avgBuildingFuelAdj, prevkw, prevkwh,prevcostkw, prevcostkwh, prevBillAmt,prevfueladj
	dim currentkw, currentkwh, currentcostkw, currentcostkwh, currentfueladj, buildingname, currentBillAmt 'totals (and building name for top)
	'now get these header fields as well
	if utilityid=2 then
	  strsql2 = "select 'Cur' as rectype,(select distinct strt from buildings where bldgnum='"&building&"') as building, isnull(FuelAdj,0) as fueladj,  isnull(sum(TotalKW),0) as TotalKW, isnull(sum(TotalKWH),0) as TotalKWH, isnull(sum(TotalBillAmt),0) as TotalBillAmt, isnull(sum(CostKW),0) as CostKW, isnull(sum(CostKWH),0) as CostKWH FROM utilitybill where ypid in (select ypid FROM billyrperiod where bldgnum='"& building &"' and Billyear="&byear&" and BillPeriod="&bperiod&") GROUP BY FuelAdj union select 'Prev' as rectype,(select distinct strt from buildings where bldgnum='"&building&"') as building, FuelAdj,  sum(TotalKW) as TotalKW,sum(TotalKWH) as TotalKWH, sum(TotalBillAmt) as TotalBillAmt, sum(CostKW) as CostKW, sum(CostKWH) as CostKWH FROM utilitybill where ypid in (select ypid FROM billyrperiod where bldgnum='"&building&"' and Billyear="&prevbyear&" and BillPeriod="&prevbperiod&") GROUP BY FuelAdj union SELECT 'AVG' as rectype,'NA' as building, isNull(avg(ub.fuelAdj),0) as fueladj, isNull(avg(TotalKW),0) as TotalKW, isNull(avg(TotalKWH),0) as TotalKWH, isNull(avg(TotalBillAmt),0) as TotalBillAmt, isNull(avg(CostKW),0) as CostKW, isNull(avg(CostKWH),0) as CostKWH FROM tblBillByPeriod bbp INNER JOIN utilitybill ub ON ub.ypid=bbp.ypid Where bbp.reject=0 and ((billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&prevbperiod&"-2)or(billyear="&prevbyear&" and billperiod>="&prevbperiod&"+9)) and bbp.utility="&utilityid&" and bbp.bldgnum='"&building&"'" 

	  rst2.open strsql2, cnn3
	  if not rst2.EOF then
	  while not rst2.eof	  
		  select case lcase(trim(rst2("rectype")))
		  case "avg"
			prevbuildingAvgKW 		= rst2("TotalKW")
			prevbuildingAvgKWH 		= rst2("TotalKWH")
			prevbuildingAvgBillAmt 	= rst2("TotalBillAmt")
			avgBuildingCostKW 		= rst2("CostKW")
			avgBuildingCostKWH 		= rst2("CostKWH")
			avgBuildingFuelAdj 		= rst2("fueladj")
		  case "cur"
			currentkw = rst2("TotalKW")
			currentkwh = rst2("TotalKWH")
			currentcostkw = rst2("CostKW")
			currentcostkwh = rst2("CostKWH")
			currentBillAmt = rst2("TotalBillAmt")
			currentfueladj = rst2("FuelAdj")
		  case "prev"
			prevkw = rst2("TotalKW")
			prevkwh = rst2("TotalKWH")
			prevcostkw = rst2("CostKW")
			prevcostkwh = rst2("CostKWH")
			prevBillAmt = rst2("TotalBillAmt")
			prevfueladj = rst2("FuelAdj")
		  end select
		  rst2.movenext
	  wend
	  end if
	  rst2.close
	end if
	
	'dim aveKWH'averages
	
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
						<td width="10%" align="center">Fuel Adjustment</td>
						<td width="12%" align="center">Total Amount Billed</td>
					</tr>
					<tr>
						<td align="right" class="tblunderline">Current</td>
						<td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(currentkw)%></td>
						<td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(currentkwh,0)%></td>
						<td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(currentcostkw)%></td>
						<td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(currentcostkwh)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(currentfueladj,6)%></td>
						<td class="tblunderline"align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(currentBillAmt)%></td>
				  </tr>
					<tr>
						<td align="right" class="tblunderline">Previous</td>
						<td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(prevkw)%></td>
						<td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(prevkwh,0)%></td>
						<td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(prevcostkw)%></td>
						<td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(prevcostkwh)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(prevfueladj,6)%></td>
						<td class="tblunderline"align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(prevBillAmt)%></td>
					</tr>
					<tr>
						<td align="right" class="tblunderline"><a onMouseOver="overlib('Average: Average taken from the previous 3 months');" onMouseOut="nd();" style="cursor:hand">Average</a></td>
						<td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(prevbuildingAvgKW)%></td>
						<td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(prevbuildingAvgKWH,0)%></td>
						<td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(avgBuildingCostKW)%></td>
						<td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(avgBuildingCostKWH)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(avgBuildingFuelAdj,6)%></td>
						<td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(prevbuildingAvgBillAmt)%></td>
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
					
						<br>View <input type="radio" name="t" value="t" onClick="switchView(this);" <%if t="t" then response.Write "checked"%>> Tenant <input type="radio" name="t" value="m" <%if t="m" then response.Write "checked"%> onClick="switchView(this);"> Meter 
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
		<table width="100%"><tr><td><span id="d_ticketlisting" name = "empty">[-]</span>&nbsp;&nbsp;<a href="#" onClick="display('ticketlisting')">Tenants with Open Trouble Tickets</a> [<span id="mTicketlist" name = "empty"></span>]</td></tr></table>
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
		
		<table width="100%"><tr><td><span id="d_abovevar" name = "empty"><% if checkopentickets then %>[+]<% else %>[-]<% end if%></span>&nbsp;&nbsp;<a href="#" onClick="try{display('abovevar')}catch(exception){}">Meters Above Allowable Set Variance</a> [<span id="mabovevar" name = "empty">0</span>] <span id= "sabovevar" style = "display:none;"><a href = "csvdownloadval.asp?link=<%=pagelink%>&bldg=<%=building%>&bp=<%=bperiod%>&by=<%=byear%>&u=<%=utilityid%>&r=0"><img border = 0 src="../../images/pmsave.gif"></span></a></td></tr></table>
<%
			if checkopentickets = true then displaymode = "none" else displaymode = "block" end if 
			if t = "m" then
			rst1.open mainSql,cnn3	
			listOriginalmeters "abovevar", showscroll, 0
			else
			listmeters "abovevar", showscroll, 0
			end if
%>

		<table width="100%"><tr><td><span id="d_belowvar" name = "empty">[+]</span>&nbsp;&nbsp;<a href="#" onClick="try{display('belowvar')}catch(exception){}">Meters Below Allowable Set Variance</a> [<span id="mbelowvar" name = "empty">0</span>] <span id= "sbelowvar" style = "display:none;"><a href = "csvdownloadval.asp?link=<%=pagelink%>&bldg=<%=building%>&bp=<%=bperiod%>&by=<%=byear%>&u=<%=utilityid%>&r=1"><img border = 0 src="../../images/pmsave.gif"></a></span></td></tr></table>
<%
			if t = "m" then
			listOriginalmeters "belowvar",showscroll2, 1
			rst1.close
			else
			listmeters "belowvar", showscroll2, 1
			end if
%>
		<%if not(isBuildingOff(building)) then%>
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


<%end if%>
</body>
</html>
<%

function listmeters(tag, displaymode, rmode)

	Dim cTenant,endLoop,tagcolor_u,tagcolor_d,ischecked 
	if not isnumeric(rmode) then rmode=0
	
	select case tag
	case "abovevar"
		strsql = "select isnull(avgused_cy,0) as avgusedcy,isnull(avgused_py,0) as avgusedpy,isnull(avgdemand_cy,0) as avgdemandcy,isnull(avgdemand_py,0) as avgdemandpy,isnull(varused_cy,0) as varusedcy,isnull(varused_py,0) as varusedpy,isnull(vardemand_cy,0) as vardemandcy,isnull(vardemand_py,0) as vardemandpy, * from (select tl.billingid, tl.bldgnum, billingname, tenantnum,m.meterid,m.meternum,m.variance, case when avgused_py <> 0 then (((used-avgused_py)*100)/avgused_py)/100 else avgused_py end as varused_py, case when avgused_cy <> 0 then (((used-avgused_cy)*100)/avgused_cy)/100 else avgused_cy end as varused_cy, case when avgdemand_py <> 0 then abs(((demand-avgdemand_py)*100)/avgdemand_py)/100 else avgdemand_py end as vardemand_py, case when avgdemand_cy <> 0 then abs(((demand-avgdemand_cy)*100)/avgdemand_cy)/100 else avgdemand_cy end as vardemand_cy,used, demand, avgused_py, avgdemand_py,avgused_cy, avgdemand_cy,validate, svalidate,coincident, isnull(posted,0) as posted from tblleases tl inner join tblleasesutilityprices lpt on lpt.billingid = tl.billingid inner join meters m on m.leaseutilityid = lpt.leaseutilityid inner join consumption c on c.meterid = m.meterid inner join peakdemand p on p.meterid = m.meterid and c.billyear = p.billyear and c.billperiod=p.billperiod left join (select c.meterid, avg(used) as avgused_py, avg(demand) as avgdemand_py from consumption c left join peakdemand p on p.meterid = c.meterid and c.billyear = p.billyear and p.billperiod=c.billperiod where c.billyear = "&byear-1&" and c.billperiod = "&bperiod&" group by c.meterid) avgpy on avgpy.meterid = m.meterid left join (select c.meterid, avg(used) as avgused_cy, avg(demand) as avgdemand_cy from consumption c left join peakdemand p on p.meterid = c.meterid and c.billyear = p.billyear and p.billperiod=c.billperiod where c.billyear = "&byear&" and c.billperiod = "&prevbperiod&" group by c.meterid) avgcy on avgcy.meterid = m.meterid left join (select leaseutilityid, billperiod, billyear, posted from tblbillbyperiod) bbp on bbp.leaseutilityid=lpt.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod and posted=1 where c.billyear = "&byear&" and c.billperiod = "&bperiod&" and tl.bldgnum='"&building&"' and lpt.utility="&utilityid&" and online=1 and nobill=0) rs where varused_cy > variance or vardemand_cy > variance order by billingname, meternum" 
	case "belowvar"
		strsql = "select isnull(avgused_cy,0) as avgusedcy,isnull(avgused_py,0) as avgusedpy,isnull(avgdemand_cy,0) as avgdemandcy,isnull(avgdemand_py,0) as avgdemandpy,isnull(varused_cy,0) as varusedcy,isnull(varused_py,0) as varusedpy,isnull(vardemand_cy,0) as vardemandcy,isnull(vardemand_py,0) as vardemandpy, * from (select tl.billingid, tl.bldgnum, billingname, tenantnum,m.meterid,m.meternum,m.variance, case when avgused_py <> 0 then (((used-avgused_py)*100)/avgused_py)/100 else avgused_py end as varused_py, case when avgused_cy <> 0 then (((used-avgused_cy)*100)/avgused_cy)/100 else avgused_cy end as varused_cy, case when avgdemand_py <> 0 then abs(((demand-avgdemand_py)*100)/avgdemand_py)/100 else avgdemand_py end as vardemand_py, case when avgdemand_cy <> 0 then abs(((demand-avgdemand_cy)*100)/avgdemand_cy)/100 else avgdemand_cy end as vardemand_cy,used, demand, avgused_py, avgdemand_py,avgused_cy, avgdemand_cy,validate, svalidate,coincident, isnull(posted,0) as posted from tblleases tl inner join tblleasesutilityprices lpt on lpt.billingid = tl.billingid inner join meters m on m.leaseutilityid = lpt.leaseutilityid inner join consumption c on c.meterid = m.meterid inner join peakdemand p on p.meterid = m.meterid and c.billyear = p.billyear and c.billperiod=p.billperiod left join (select c.meterid, avg(used) as avgused_py, avg(demand) as avgdemand_py from consumption c left join peakdemand p on p.meterid = c.meterid and c.billyear = p.billyear and p.billperiod=c.billperiod where c.billyear = "&byear-1&" and c.billperiod = "&bperiod&" group by c.meterid) avgpy on avgpy.meterid = m.meterid left join (select c.meterid, avg(used) as avgused_cy, avg(demand) as avgdemand_cy from consumption c left join peakdemand p on p.meterid = c.meterid and c.billyear = p.billyear and p.billperiod=c.billperiod where c.billyear = "&byear&" and c.billperiod = "&prevbperiod&" group by c.meterid) avgcy on avgcy.meterid = m.meterid left join (select leaseutilityid, billperiod, billyear, posted from tblbillbyperiod) bbp on bbp.leaseutilityid=lpt.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod and posted=1 where c.billyear = "&byear&" and c.billperiod = "&bperiod&" and tl.bldgnum='"&building&"'  and lpt.utility="&utilityid&" and online=1 and nobill=0) rs where varused_cy <= variance and vardemand_cy <= variance order by billingname,meternum"
	end select
	
	rst1.open strsql, cnn1,1,1

	if not rst1.eof then	
	%>
	<div id="<%=tag%>" style="display:<%=displaymode%>;">
		<table><tr><td>&nbsp;<span id="toggle_<%=tag%>" style="display:none">[+]</span>&nbsp;</td><td bgcolor="#dddddd" width=150 align="center" onMouseOver="this.style.background='lightgreen'" onMouseOut="this.style.background='#dddddd'"><a onClick="showall('<%=tag%>')" style="cursor:hand"><font color="#000000"><b>Expand/Collapse</b></font></a></td></tr></table>
	<%
	ctenant=rst1("billingid")
	do until rst1.eof
	isposted = false
	
	if rst1("posted") then isposted = true end if 
	%>
		<table border=0 cellspacing="1" cellpadding="3">
		<tr><td>&nbsp;</td><td align="left"><span id="d_<%=tag%>_<%=rst1("billingid")%>_row" name="tenantlist_titles_<%=tag%>">[+]</span>&nbsp;<a href="javascript:display('<%=tag%>_<%=rst1("billingid")%>_row')"><%=rst1("billingname")%></a> [TN: <%=rst1("tenantnum")%> BID: <%=rst1("billingid")%>]</td></tr>
		<tr id="<%=tag%>_<%=rst1("billingid")%>_row" style="display:none" name="tenantlist_<%=tag%>"><td>&nbsp;</td><td>	
			<table width="700" border=0 cellspacing="1" cellpadding="3">
			<% if rst1("coincident") then %>
			<tr bgcolor="#dddddd" valign="bottom">
			<td colspan=8 align="center" style="border:2px solid black">THIS TENANT IS ON COINCIDENTAL DEMAND</td>
			</tr>
			<% end if %>
				<tr bgcolor="#dddddd" valign="bottom">
				<td width="40">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td colspan=2 align="center"><span class="standardheader"><font color="black"><%=usage%> Comparison<br>(<font color="#FF0000">Over</font>/<font color="#0066FF">Under</font> Variance)</font></span></td>
				<td>&nbsp;</td>

				<td colspan=2 align="center"><span class="standardheader"><font color="black"><%=demand%> Comparison<br>(<font color="#FF0000">Over</font>/<font color="#0066FF">Under</font> Variance)</font></span></td>
				</tr>
				<tr bgcolor="#dddddd" valign="bottom">
				<td width="40" align="center"><span class="standardheader"><font color="black">Accept</font></span></td>
				<td align="center"><span class="standardheader"><font color="black">Meter</font></span></td>
				<td align="center"><span class="standardheader"><font color="black">Current <%=usage%></font></span></td>
				<td align="center"><span class="standardheader"><font color="black"><a onMouseOver="overlib('Usage from Last Period');" onMouseOut="nd();" style="cursor:hand">Previous<br>Period</a></font></span></td>
				<td align="center"><span class="standardheader"><font color="black"><a onMouseOver="overlib('Usage for same period Last Year');" onMouseOut="nd();" style="cursor:hand">This Period,<br>Last Year</a></font></span></td>
				<td align="center"><span class="standardheader"><font color="black">Current <%=demand%></font></span></td>
				<td align="center"><span class="standardheader"><font color="black"><a onMouseOver="overlib('Demand from Last Period');" onMouseOut="nd();" style="cursor:hand">Previous<br>Period</a></font></span></td>
				<td align="center"><span class="standardheader"><font color="black"><a onMouseOver="overlib('Demand for same Period Last Year');" onMouseOut="nd();" style="cursor:hand">This Period,<br>Last Year</a></font></span></td>
				</tr>
				<% 
				endLoop=false
				Dim tagcolor_up, tagcolor_dp
				do until rst1.EOF or endLoop
				tagcolor_u="#0066FF"
				tagcolor_d="#0066FF"
				tagcolor_up="#0066FF"
				tagcolor_dp="#0066FF"
					
				if cint(rst1("validate"))=true or Not(super) then ischecked = " CHECKED"
				if cint(rst1("svalidate"))=true and super then ischecked = " CHECKED"

				if cdbl(rst1("varusedcy")) < cdbl(rst1("variance")) then 
					tagcolor_u="#FF0000"
				end if 
				if cdbl(rst1("vardemandcy")) < cdbl(rst1("variance")) then
					tagcolor_d="#FF0000"
				end if  
				if cdbl(rst1("varusedpy")) < cdbl(rst1("variance")) then 
					tagcolor_up="#FF0000"
				end if 
				if cdbl(rst1("vardemandpy")) < cdbl(rst1("variance")) then
					tagcolor_dp="#FF0000"
				end if  
				
				%>
				<tr bgcolor="#dddddd" valign="bottom" style="cursor:hand" onClick="updatemeter('<%=rst1("meterid")%>', <%=byear%>, <%=bperiod%>, '<%=rst1("tenantnum")%>','<%=rst1("billingid")%>', '[<%=server.urlencode(rst1("billingname"))%>]', '<%=isposted%>')" onMouseOver="this.style.backgroundColor='lightgreen';" onMouseOut="this.style.backgroundColor='#dddddd';">
				<td width="40" align="center"><input type="checkbox" value="<%=rst1("meterid")%>" name="meters" onMouseOver="checkboxf=1" onMouseOut="checkboxf=0" style="cursor:auto" <%=ischecked%>></td>
				<td><%=rst1("meternum")%></td>
				<td align="right"><font color="#336699"><b><%=rst1("used")%></b></font></td>
				<td align="right"><%=formatnumber(rst1("avgusedcy"),2)%><br>(<font color="<%=tagcolor_u%>"><%=formatpercent(abs(cdbl(rst1("varusedcy"))))%></font>)</td>
				<td align="right"><%=formatnumber(rst1("avgusedpy"),2)%><br>(<font color="<%=tagcolor_up%>"><%=formatpercent(abs(cdbl(rst1("varusedpy"))))%></font>)</td>
				<td align="right"><font color="#336699"><b><%=rst1("demand")%></b></font></td>
				<td align="right"><%=formatnumber(rst1("avgdemandcy"),2)%><br>(<font color="<%=tagcolor_d%>"><%=formatpercent(abs(cdbl(rst1("vardemandcy"))))%></font>)</td>
				<td align="right"><%=formatnumber(rst1("avgdemandpy"),2)%><br>(<font color="<%=tagcolor_dp%>"><%=formatpercent(abs(cdbl(rst1("vardemandpy"))))%></font>)</td>
				</tr>
				<%				
				rst1.movenext
				if not rst1.eof then
					if rst1("billingid") <> ctenant then 
						endLoop = true
						ctenant=rst1("billingid")
					end if
				end if
				loop
				%>
			</table>
			</td></tr>
		</table>
	<%
	loop
	%>
	<script>
	var func = eval('document.all.m<%=tag%>')
	func.innerHTML = '<%=rst1.recordcount%>';
				if (<%=rst1.recordcount%> != 0 )
			{
			var func = eval('document.all.s<%=tag%>');
			func.style.display = "inline";
			}
	</script>
	</div>
	<%
	end if
	rst1.close
end function

function listOriginalmeters(tag, displaymode, rmode)
	if not isnumeric(rmode) then rmode=0%>
	<div id="<%=tag%>" style="display:<%=displaymode%>;">
		<table width="100%" border=0 cellspacing="1" cellpadding="3">
			<tr bgcolor="#dddddd" valign="bottom">
			<td width="40"><font color="black">Accept</font></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Tenant Number</font></span></td>
			<td width="<%=cellw+15%>"><span class="standardheader"><font color="black">Tenant&nbsp;Name</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Meter</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Average <%=usage%></font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Current <%=usage%> Usage</font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Variance <%=usage%></font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Average <%=demand%></font></span></td>
			<td width="<%=cellw%>"><span class="standardheader"><font color="black">Current <%=demand%></font></span></td>
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
				reccount=0
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
					<td width="<%=cellw%>" nowrap><%if extusg then response.write "On:"%><%=formatnumber(rst1("avgKWH"),0)%>
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
				if (<%=reccount%> != 0 )
			{
			var func = eval('document.all.s<%=tag%>');
			func.style.display = "inline";
			}
			</script>
		</div>
	</div>
	<%
end function
%>



		
	
