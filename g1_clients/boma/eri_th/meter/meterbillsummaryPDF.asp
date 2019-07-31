<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, ypid, luid, b, m, building, bperiod, byear, utility, usagedivisor, bldg
utility = cint(Request("utility"))
leaseid = Request("l")
ypid = request("y")
luid = request("luid")
b = request("b")
bldg = request("bldg")
m = request("m")
building = request("building")
if trim(building)="" then building = b
if trim(building)="" then building = bldg
if trim(leaseid)="" then leaseid = luid

dim cnn1, cmd1, rst1, rst2, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
dim DBmainIP
DBmainIP = "["&application("superIP")&"].mainmodule.dbo."

dim usagelabel, genergy2, title, tot_onpeak, tot_offpeak, tot_kwhused, tot_demand_p, strt, units
genergy2 = "true"
usagelabel = "used"
cnn1.Open getLocalConnect(building)

rst1.open "SELECT strt FROM buildings b WHERE bldgnum='"&building&"'", cnn1
if not rst1.eof then strt = rst1("strt")
rst1.close

units = "BTU"
rst1.open "SELECT measure FROM tblutility u WHERE u.utilityid="&utility, application("cnnstr_supermod")
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

dim billlink, pid
if trim(building)<>"" then
	rst1.open "SELECT location, b.bldgnum, b.portfolioid FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&building&"'", application("cnnstr_supermod")
	if not rst1.eof then 
    billlink = rst1("location")
    pid = rst1("portfolioid")
  end if
	rst1.close
end if

'cmd1.ActiveConnection = cnn1
if leaseid<>0 then
	title = "Tenant"
	sql = "SELECT billyear, billperiod, isnull(credit,0) as creditsum, isnull(Adminfee,0) as Adminfee, isnull(Addonfee,0) as Addonfee, isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, isnull(energy,0) as energy, isnull(demand,0) as demand from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid &" and posted=1"
else
	title = "Building"
	sql = "select billyear, billperiod, sum(isnull(credit,0)) as creditsum, sum(isnull(energy,0)) as energy, sum(isnull(demand,0)) as demand, avg(isnull(adminfee,0)) as adminfee, sum(isnull(addonfee,0)) as addonfee, sum(isnull(tax,0)) as tax, sum(isnull(totalamt,0)) as totalamt from tblbillbyperiod where ypid=" & ypid &" and posted=1 and utility="&utility&" group by billyear, billperiod"
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
%><html>

<head>
<title>Meter Details</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">
</head>

<script>
function closeme(){
	window.close()
}

function viewbill(ypid,lid) {
	var temp= "invoice.asp?building=<%=building%>&y=" + ypid + "&l=" + lid
	window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
}
function billSummary(building, billPeriod, billYear){
	var temp = "bill_summary.asp?building=" + building + "&bperiod=" + billPeriod + "&byear=" + billYear
	window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
}
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#6699CC">
  <tr>
    <td width="46%" bgcolor="#6699CC"><span class="standardheader">Bill Details</span></td>
    <td width="46%" bgcolor="#6699CC" align="right">
		<%if leaseid<>0 then%>
			<a href="javascript:viewbill('<%=ypid%>','<%=leaseid%>')" style="text-decoration:none;color:white" 
				onMouseOver="this.style.color= 'lightblue';" onMouseOut="this.style.color = 'white'">
				View Bill For This Period
			</a>
		<%else%>
			<font color="white" onMouseOver="this.style.color='lightblue';window.status='Link does not function in demo version'; return true" onMouseOut="this.style.color = 'white'; window.status=''; return true">
				Download&nbsp;all&nbsp;bills&nbsp;for&nbsp;this&nbsp;building
			</font>
			&nbsp;
			<%if utility<>3 then
				dim link, rstByBp
				Set rstByBp = Server.CreateObject("ADODB.recordset")
				dim sqlByBp
				sqlByBp = "select billyear, billperiod from tblMetersByPeriod where ypid='" & ypid & "'"
				%><script>//alert("<%=sqlByBp%>")</script><%
				rstByBp.open sqlByBp, cnn1
				%>
				|&nbsp;
				<a style="text-decoration:none;color:white" href="javascript:billSummary('<%=building%>','<%=rstByBp("billperiod")%>','<%=rstByBp("billyear")%>');" 
					onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'">Download&nbsp;Bill&nbsp;Summary&nbsp;Report
				</a>
			<%end if%>
		<%end if%>
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
      <table border="0" width="100%" height="1">
        <tr bgcolor="#6699CC"> 
          <td width="7%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Period</font></b></td>
            <td width="14%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=charge1%></font></b></td>
            <td width="12%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=charge2%></font></b></td>
          <%if utility<>3 then%>
            <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Admin Fee</font></b></td>
          <%end if%>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Service Fee</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Sales Tax</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Total Amt</font></b></td>
        </tr>
        <tr> 
          <td width="7%" height="1%" align="center"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></b></td>
            <td width="14%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("energy"),2)%></font></b></td>
            <td width="12%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("demand"),2)%></font></b></td>
          <%if utility<>3 then%>
            <td width="10%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatpercent(rst2("Adminfee"),2)%></font></b></td>
          <%end if%>
          <td width="10%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("Addonfee"),2)%></font></b></td>
          <td width="10%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("tax"),2)%></font></b></td>
          <td width="10%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("totalamt").value,2)%></font></b></td>
        </tr>
      </table>
      <font face="Arial"> <br>
      </font>
	<%if leaseid<>0 then%>
      <table border="0" width="100%" height="1" cellpadding="0" cellspacing="1" align="center">
        <tr bgcolor="#6699CC"> 
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Meter</font></b></td>
          <%if utility<>3 then%>
            <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">On Peak <%=units%></font></b></td>
            <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Off Peak <%=units%></font></b></td>
          <%end if%>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=units%></font></b></td>
          <%if utility<>3 then%><td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Demand</font></b></td><%end if%>
        </tr>
        <%
sql = "select * from tblmetersbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid

rst1.open sql, cnn1

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
%>
        <tr> 
          <td width="20%" height="1%" align="left"><%=rst1("Meternum")%></td>
          <%if utility<>3 then%>
            <td width="20%" height="1%" align="right"><%=formatnumber(rst1("onpeak"),0)%></td>
            <td width="20%" height="1%" align="right"><%=formatnumber(rst1("offpeak"),0)%></td>
          <%end if%>
          <td width="20%" height="1%" align="right"><%=formatnumber(cdbl(rst1(usagelabel))/cint(usagedivisor),2)%></td>
          <%if utility<>3 then%><td width="20%" height="1%" align="right"><%=formatnumber(rst1("demand_P"),2)%></td><%end if%>
        </tr>
        <%
tot_onpeak = tot_onpeak + cDbl(rst1("onpeak"))
tot_offpeak= tot_offpeak+ cDbl(rst1("offpeak"))
tot_kwhused= tot_kwhused + cDbl(cdbl(rst1(usagelabel))/usagedivisor)
tot_demand_p= tot_demand_p + cDbl(rst1("demand_P"))

rst1.movenext
wend
end if 'this one is for masking the meter table when building view
end if 'this one is for if has records
if leaseid<>0 then
%>
        <tr bgcolor="#CCCCCC"> 
          <td width="20%" height="1%" align="center"><div align="center"></div><p align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Totals</font></b></td>
          <%if utility<>3 then%>
            <td width="20%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_onpeak,0)%></font></b></td>
            <td width="20%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_offpeak,0)%></font></b></td>
          <%end if%>
          <td width="20%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_kwhused,2)%></font></b></td>
          <%if utility<>3 then%><td width="20%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_demand_P,2)%></font></b></td><%end if%>
        </tr>
      </table>
      <%
end if
set cnn1 = nothing
%>
    </td>
  </tr>
</table>

  <table bgcolor="#6699CC" cellpadding="0" cellspacing="0" width="100%"><tr><td>&nbsp;</td></tr></table>
</body>
</html>




