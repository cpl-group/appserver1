<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim building, byear, bperiod, pid
pid = Request("pid")
building = Request("building")
byear= Request("byear")
bperiod= Request("bperiod")

Dim cnn1, rst1, sql
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(pid,building,"billing")

%>
<html>
<head>
<title>Meter Problems Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
<script>
function clearSelects(n)
{ var frm = document.forms['form1']
  if((frm.building!=null)&&(n=='pid')) frm.building.value='';
  if((frm.byear!=null)&&((n=='pid')||(n=='building'))) frm.byear.value='';
  if((frm.bperiod!=null)&&((n=='pid')||(n=='building')||(n=='byear'))) frm.bperiod.value='';
}
function JumpTo(url){
	var frm = document.forms['form1'];
	var url = url + "?pid=<%=pid%>&bldg=<%=building%>&building=<%=building%>&utilityid=2&byear=<%=byear%>&bperiod=<%=bperiod%>";
	window.document.location=url;
}
</script>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td bgcolor="#6699CC" class="standardheader">Meter Problems Report</td>
    <td align="right" bgcolor="#6699CC" class="standardheader"><%if trim(building) <> "" then %>
	<select name="select" onChange="JumpTo(this.value)">
        <option value="#" selected>Jump to...</option>
        <option value="/genergy2/billing/processor_select.asp">Bill Processor</option>
        <option value="../validation/re_index.asp">Review Edit</option>
        <option value="/genergy2/setup/buildingedit.asp">Building Setup</option>
        <option value="/genergy2/manualentry/entry_select.asp">Manual Entry</option>
        <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem 
        Report</option>
      </select><% end if %></td>
  </tr>
  <tr>
    <td colspan="2"><form name="form1" action="meterProblemReport.asp"> 
  <tr bgcolor="#eeeeee"> 
    <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
      <table border=0 cellpadding="3" cellspacing="0">
        <tr>
          <td> <select name="pid" onChange="clearSelects(this.name);submit()">
              <option value="">Select Portfolio</option>
              <%rst1.open "SELECT distinct id, name FROM portfolio p ORDER BY name", getConnect(0,0,"dbCore")
      do until rst1.eof%>
              <option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then response.write " SELECTED"%>><%=rst1("name")%></option>
              <%	rst1.movenext
      loop
      rst1.close%>
            </select> </td>
          <%if trim(pid)<>"" then%>
          <td> <select name="building" onChange="clearSelects(this.name);submit()">
              <option value="">Select Building</option>
              <%
      rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
      do until rst1.eof%>
              <option <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>><%=rst1("strt")%>, 
              <%=trim(rst1("Bldgnum"))%></option>
              <%	rst1.movenext
      loop
      rst1.close
      %>
            </select> </td>
          <%end if
    if trim(building)<>"" then%>
          <td> <select name="byear" onChange="clearSelects(this.name);submit()">
              <%
      rst1.open "SELECT Distinct BillYear FROM BillYrPeriod WHERE BldgNum='"&building&"'", getLocalConnect(building)
      if rst1.eof then
        response.write "<option value="""">No Billing Years</option>"
      else
        response.write "<option value="""">Select Bill Year</option>"
      end if
      do until rst1.eof
        %>
              <option value="<%=rst1("Billyear")%>"<%if trim(rst1("billyear"))=trim(byear) then response.write " SELECTED"%>><%=rst1("Billyear")%></option>
              <%
        rst1.movenext
      loop
      rst1.close
      %>
            </select> </td>
          <%end if
    if trim(byear)<>"" and trim(building)<>"" then%>
          <td> <select name="bperiod" onChange="submit()">
              <option value="">Select Bill Period</option>
              <%rst1.open "SELECT distinct billperiod FROM billyrperiod WHERE bldgnum='"&building&"' and billyear="&byear&" order by billperiod", getLocalConnect(building)
      do until rst1.eof
        %>
              <option value="<%=rst1("billperiod")%>"<%if trim(rst1("billperiod"))=trim(bperiod) then response.write " SELECTED"%>><%=rst1("BillPeriod")%></option>
              <%
        rst1.movenext
      loop
      rst1.close
      %>
            </select> </td>
          <td> <input type="button" value="Print" onClick="window.print()"> </td>
          <%end if%>
        </tr>
      </table></td>
  </tr>
</table>
<%if trim(bperiod)<>"" then
sql = "SELECT DISTINCT meters.meternum,tblleases.billingName,tblleases.flr, peakdemand.billyear,peakdemand.billperiod, peakdemand.usernote as pnote,consumption.usernote as cnote, tenantnum FROM peakdemand JOIN meters ON peakdemand.meterid = meters.meterid JOIN consumption ON consumption.meterid = meters.meterid join tblleasesutilityprices tp on meters.leaseutilityid=tp.leaseutilityid join tblleases on tp.billingid=tblleases.billingid  WHERE peakdemand.billyear = '" & byear & "' AND peakdemand.billperiod = '" & bperiod & "' AND   consumption.billyear = '" & byear & "' AND consumption.billperiod = '" & bperiod & "' AND (peakdemand.usernote IS NOT NULL OR consumption.usernote IS NOT NULL) AND meters.bldgnum = '" & building & "'"
rst1.Open sql, getLocalConnect(building)
if rst1.EOF then%>
<table width="100%" border="0">
  <tr>
    <td>NO PROBLEMS FOUND</td>
  </tr>
</table>
<%else%>
<table width="100%" border="0" cellpadding="2" cellspacing="0">
<tr><td bgcolor="#6699CC" class="standardheader" align="center" colspan="7"><b>Genergy Meter Problem Report for Building Number <%if isBuildingOff(building) then%><i><%end if%><%=building%> <%if isBuildingOff(building) then%> (offline)</i><%end if%></b></td></tr>
<tr><td width="8%" bgcolor="#CCCCCC">Meter #</td>
	<td width="15%" bgcolor="#CCCCCC">Tenant Name</td>
	<td width="15%" bgcolor="#CCCCCC" align="center">Tenant Number</td>
	<td width="7%" bgcolor="#CCCCCC">Floor</td>
	<td width="10%" bgcolor="#CCCCCC" align="center">Period</td>
	<td width="20%" bgcolor="#CCCCCC">User Note (Consumption)</td>
	<td width="25%" bgcolor="#CCCCCC">User Note (Peakdemand)</td>
</tr>
<% While not rst1.EOF %>
<tr valign="top"> 
<td width="8%"><%=rst1("meternum")%></td>
<td width="15%"><%=rst1("billingname")%></td>
<td width="15%" align="center"><%=rst1("tenantnum")%></td>
<td width="7%"><%=rst1("flr")%></td>
<td width="10%" align="center"><%=rst1("billperiod")%>/<%=rst1("billyear")%></td>
<td width="20%"><%=rst1("cnote")%></td>
<td width="25%"><%=rst1("pnote")%></td>
</tr>
<tr valign="top"> 
<td colspan="7"><hr size="1" noshade></td>
</tr>
<%
rst1.movenext
Wend
%>
</table>
<p>&nbsp;</p>
<%
end if

end if
%>

</body>
</html>