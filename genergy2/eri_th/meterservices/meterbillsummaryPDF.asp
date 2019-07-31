<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, ypid, luid, b, m, building, bperiod, byear, utility, usagedivisor, bldg,demo
utility = cint(Request("utility"))
leaseid = Request("l")
ypid = request("y")
luid = request("luid")
b = request("b")
bldg = request("bldg")
m = request("m")
building = request("building")
demo = request("demo")

if trim(building)="" then building = b
if trim(building)="" then building = bldg
if trim(leaseid)="" then leaseid = luid

dim cnn1, cmd1, rst1, rst2, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
'dim DBmainIP
'DBmainIP = "["&application("superIP")&"].mainmodule.dbo."

dim usagelabel, genergy2, title, tot_onpeak, tot_offpeak, tot_intpeak, tot_kwhused, tot_demand_p, strt, units
genergy2 = "true"
usagelabel = "used"
cnn1.Open getLocalConnect(building)

rst1.open "SELECT strt FROM buildings b WHERE bldgnum='"&building&"'", cnn1
if not rst1.eof then strt = rst1("strt")
rst1.close

units = "BTU"
rst1.open "SELECT measure FROM tblutility u WHERE u.utilityid="&utility, cnn1
if not rst1.eof then 
  units=rst1("measure")
end if
rst1.close

if trim(ypid)<>"" then
  rst1.open "SELECT distinct billyear, billperiod FROM billyrperiod b WHERE ypid="&ypid, cnn1
  if not rst1.eof then 
    byear = rst1("billyear")
    bperiod = rst1("billperiod")
  end if
  rst1.close
end if

dim billlink, pid,billurl,logo,logoh,logow, portfolio
if trim(building)<>"" then
	rst1.open "SELECT location, b.bldgnum, b.portfolioid,billurl,logo, logoh, logow, p.portfolio FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&building&"'	", cnn1
	if not rst1.eof then 
		billlink = rst1("location")
		pid = rst1("portfolioid")
		billurl = rst1("billurl")
		logo = rst1("logo")
		logoh = rst1("logoh")
		logow = rst1("logow")		
		portfolio=rst1("portfolio")		
	end if
	rst1.close
	Dim billCount
	rst1.open "SELECT count(distinct lup.leaseutilityid) as billCount FROM tblleasesutilityprices lup, tblleases l, meters m WHERE l.billingid=lup.billingid and lup.leaseutilityid=m.leaseutilityid and m.nobill=0 and online=1 and l.leaseexpired=0 and l.bldgnum='"&building&"' and lup.utility="&utility , cnn1
	if not rst1.eof then 
	billCount = rst1("billCount")
	else
	billCount = ""
	end if
	rst1.close
end if


'response.Write(billlink)
'response.End()
'cmd1.ActiveConnection = cnn1
if leaseid<>0 then
	title = "Tenant"
	sql = "SELECT id as billid, billyear, billperiod, isnull(credit,0) as creditsum, isnull(Adminfee,0) as Adminfee, isnull(Addonfee,0) as Addonfee, isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, isnull(energy,0) as energy, isnull(demand,0) as demand from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid &" and posted=1"
else
	title = "Building"
	sql = "select billyear, billperiod, sum(isnull(credit,0)) as creditsum, sum(isnull(energy,0)) as energy, sum(isnull(demand,0)) as demand, avg(isnull(adminfee,0)) as adminfee, sum(isnull(addonfee,0)) as addonfee, sum(isnull(tax,0)) as tax, sum(isnull(totalamt,0)) as totalamt from tblbillbyperiod where reject=0 and ypid=" & ypid &" and posted=1 and utility="&utility&" group by billyear, billperiod"
end if
rst2.open sql, cnn1,1
'if not rst2.eof then
	'if trim(rst2("billperiod"))<>"" then bperiod = rst2("billperiod")
	'if trim(rst2("billyear"))<>"" then byear = rst2("billyear")
'end if



if not isnumeric(utility) then utility=0
dim charge1, charge2, charge1FLD, charge2FLD
select case utility
case 3
  usagedivisor = 100
  charge1 = "Water Charge"
  charge2 = "Sewer Charge"
  charge1FLD = "energy"
  charge2FLD = "demand"
case else
  usagedivisor = 1
  charge1 = "Energy Charge"
  charge2 = "Demand Charge"
  charge1FLD = "energy"
  charge2FLD = "demand"
end select
%>
<html>

<head>
<title>Meter Details</title>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<script>
function closeme(){
	window.close()
}
function viewbill(ypid,lid, uid, detailed) {
			var temp= "http://pdfmaker.genergyonline.com<%=billlink%>genergy2=<%=genergy2%>&billurl=<%=billurl%>&logo=<%=logo%>&logoh=<%=logoh%>&logow=<%=logow%>&pid=<%=pid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&y=" + ypid + "&ypid=" + ypid + "&l=" + lid+"&utilityid="+ uid + "&demo=<%=demo%>&devIP=<%=request.servervariables("SERVER_NAME")%>&building=<%=building%>&billCount=<%=billCount%>&detailed="+detailed;
			window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
	}
function viewpdfLinks(uid) {
			var temp= "PdfLinks.asp?pid=<%=pid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&utilityid="+ uid +"&building=<%=building%>";
			window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
	}
function viewDetailedPDF(uid)
{
    var tempURL = '/genergy2/billing/loading.asp?url=<%=server.urlencode("/genergy2/billing/PA_pdfLinks.asp?genergy2=true&devIP="&request.ServerVariables("SERVER_NAME")&"&billurl="&billurl&"&pid="&pid&"&logo="&logo&"&logoh="&logoh&"&logow="&logow&"&byear="&byear&"&bperiod="&bperiod&"&y=&building="&building&"&bldg="&building&"&b="&building&"&detailed=true&billCount="&billCount&"&utilityid="&utility)%>';
    window.open(tempURL,'','width=600,height=500,resizable=yes,scrollbars=yes');
}
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
  <tr> 
    <td bgcolor="#0099FF"><font face="Arial" size="3" color="#FFFFFF"><b>Bill 
      Details</b></font></td>
  </tr>
  <tr> 
    <td width="46%" bgcolor="#0099FF" align="left" style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;" nowrap> 
      <%if leaseid<>0 then%>
      <b><a href="javascript:viewbill('<%=ypid%>','<%=leaseid%>','<%=utility%>', false)" style="text-decoration:none;color:white" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><font size="2">View 
      Bill For This Period</font></a> | <a href="javascript:viewbill('<%=ypid%>','<%=leaseid%>','<%=utility%>',true)" style="text-decoration:none;color:white" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><font size="2">View 
      Detailed Bill For This Period</font></a></b> 
      <%else
		%>
			<a style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;" href="javascript:viewpdfLinks('<%=utility%>')" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><b>Download&nbsp;all&nbsp;bills&nbsp;for&nbsp;this&nbsp;building</b></a> 
			| <a style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;" href="javascript:viewDetailedPDF('<%=utility%>')" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><b>Download&nbsp;all&nbsp;detailed&nbsp;bills&nbsp;for&nbsp;this&nbsp;building</b></a> 
			| <a style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;" target="_blank" href="http://pdfmaker.genergyonline.com/pdfMaker/pdfBillSummary.asp?genergy2=<%=genergy2%>&pid=<%=getpid(trim(building))%>&building=<%=server.urlencode(building)%>&utilityid=<%=utility%>&byear=<%=byear%>&bperiod=<%=bperiod%>&demo=<%=demo%>&devIP=<%=request.servervariables("SERVER_NAME")%>&strt=<%=strt%>" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><b>Download&nbsp;Bill&nbsp;Summary&nbsp;Report</b></a> 
      <% end if%>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td> 
<%if not rst2.eof then%>
      <div align="left"></div>
      <table border="1" width="100%" height="1">
        <tr bgcolor="#0099FF"> 
          <td width="7%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Period</font></b></td>
            <td width="14%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=charge1%></font></b></td>
            <td width="12%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=charge2%></font></b></td>
          <%if utility<>3 then%>
            <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Admin Fee</font></b></td>
          <%end if%>
          <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Service Fee</font></b></td>
          <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Sales Tax</font></b></td>
          <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Total Amt</font></b></td>
        </tr>
        <tr> 
          <td width="7%" height="1%" align="center"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></b></td>
            <td width="14%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("energy"),2)%></font></b></td>
            <td width="12%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("demand"),2)%></font></b></td>
          <%if utility<>3 then%>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatpercent(rst2("Adminfee"),2)%></font></b></td>
          <%end if%>
          <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("Addonfee"),2)%></font></b></td>
          <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("tax"),2)%></font></b></td>
          <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("totalamt").value,2)%></font></b></td>
        </tr>
      </table>
      <font face="Arial"> <br>
      </font>
	<%if leaseid<>0 then%>
      <table border="1" width="100%" height="1" cellpadding="0" cellspacing="0" align="center">
        <tr bgcolor="#0099FF"> 
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Meter</font></b></td>
          <%if utility<>3 then%>
            <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">On Peak <%=units%></font></b></td>
            <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Off Peak <%=units%></font></b></td>
            <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Int Peak <%=units%></font></b></td>
          <%end if%>
          <td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Total <%=units%></font></b></td>
          <%if utility<>3 then%><td align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Total Demand</font></b></td><%end if%>
        </tr>
        <%
sql = "select meternum, onpeak, offpeak, intpeak, isnull(used,0)+isnull(usedoff,0)+isnull(usedint,0) as usedtot, usedint, isnull(demand_p,0)+isnull(demand_off,0)+isnull(demand_int,0) as demand from tblmetersbyperiod where bill_id="&rst2("billid")

rst1.open sql, cnn1

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
%>
        <tr> 
          <td align="center"><p align="left"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("Meternum")%></font></b></td>
          <%if utility<>3 then%>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("onpeak"),0)%></font></b></td>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("offpeak"),0)%></font></b></td>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("intpeak"),0)%></font></b></td>
          <%end if%>
          <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(cdbl(rst1("usedtot"))/cint(usagedivisor),2)%></font></b></td>
          <%if utility<>3 then%><td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("demand"),2)%></font></b></td><%end if%>
        </tr>
        <%
tot_onpeak = tot_onpeak + cDbl(rst1("onpeak"))
tot_offpeak= tot_offpeak+ cDbl(rst1("offpeak"))
tot_intpeak= tot_intpeak+ cDbl(rst1("intpeak"))
tot_kwhused= tot_kwhused + cDbl(cdbl(rst1("usedtot"))/usagedivisor)
tot_demand_p= tot_demand_p + cDbl(rst1("demand"))

rst1.movenext
wend
end if 'this one is for masking the meter table when building view
end if 'this one is for if has records
if leaseid<>0 then
%>
        <tr bgcolor="#CCCCCC"> 
          <td align="center"><div align="center"></div><p align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Totals</font></b></td>
          <%if utility<>3 then%>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_onpeak,0)%></font></b></td>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_offpeak,0)%></font></b></td>
            <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_intpeak,0)%></font></b></td>
          <%end if%>
          <td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_kwhused,2)%></font></b></td>
          <%if utility<>3 then%><td align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_demand_P,2)%></font></b></td><%end if%>
        </tr>
      </table>
      <%
end if
set cnn1 = nothing
%>
    </td>
  </tr>
</table>
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><i><font size="1"> 
  </font></i></b></font></p>
  <table bgcolor="#0099FF" cellpadding="0" cellspacing="0" width="100%"><tr><td>&nbsp;</td></tr></table>
</body>
</html>




