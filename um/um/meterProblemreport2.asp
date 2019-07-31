<%option explicit%>
<%
dim building, byear, bperiod, pid
pid = Request("pid")
building = Request("building")
byear= Request("byear")
bperiod= Request("bperiod")

Dim cnn1, rst1, sql
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy2")
%>
<html>
<head>
<title>Meter Problems Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<script>
function clearSelects(n)
{ var frm = document.forms['form1']
  if((frm.building!=null)&&(n=='pid')) frm.building.value='';
  if((frm.byear!=null)&&((n=='pid')||(n=='building'))) frm.byear.value='';
  if((frm.bperiod!=null)&&((n=='pid')||(n=='building')||(n=='byear'))) frm.bperiod.value='';
}
//document.forms[0].building.value='';document.forms[0].byear.value='';document.forms[0].bperiod.value='';
</script>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#6699CC" class="standardheader">Bill Processing</span></td></tr>
<form name="form1" action="meterProblemReport2.asp">
<tr bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr><td>
      <select name="pid" onchange="clearSelects(this.name);submit()">
      <option value="">Select Portfolio</option>
      <%rst1.open "SELECT distinct id, name FROM portfolio p ORDER BY name", cnn1
      do until rst1.eof%>
        <option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then response.write " SELECTED"%>><%=rst1("name")%></option>
      <%	rst1.movenext
      loop
      rst1.close%>
      </select>
    </td>
    <%if trim(pid)<>"" then%>
    <td>
      <select name="building" onchange="clearSelects(this.name);submit()">
      <option value="">Select Building</option>
      <%
      'rst1.open "SELECT BldgNum, strt FROM (SELECT * FROM ["&join(getAllBuildingIP(),"].genergy2.dbo.buildings UNION SELECT * FROM [")&"].genergy2.dbo.buildings) b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
      rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
      do until rst1.eof%>
        <option value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>><%=rst1("strt")%>, <%=trim(rst1("Bldgnum"))%></option>
      <%	rst1.movenext
      loop
      rst1.close
      %>
      </select>
    </td>
    <%end if
    if trim(building)<>"" then%>
    <td>
      <select name="byear" onchange="clearSelects(this.name);submit()">
      <%
      'rst1.open "SELECT Distinct BillYear FROM "&DBlocalIP&"BillYrPeriod WHERE BldgNum='"&building&"' and utility="&utilityid, cnn1
      rst1.open "SELECT Distinct BillYear FROM BillYrPeriod WHERE BldgNum='"&building&"'", cnn1
      if rst1.eof then
        response.write "<option value="""">No Billing Years</option>"
      else
        response.write "<option value="""">Select Bill Year</option>"
      end if
      do until rst1.eof
        %><option value="<%=rst1("Billyear")%>"<%if trim(rst1("billyear"))=trim(byear) then response.write " SELECTED"%>><%=rst1("Billyear")%></option><%
        rst1.movenext
      loop
      rst1.close
      %>
      </select>
    </td>
    <%end if
    if trim(byear)<>"" and trim(building)<>"" then%>
    <td>
      <select name="bperiod" onchange="submit()">
      <option value="">Select Bill Period</option>
      <%rst1.open "SELECT distinct billperiod FROM billyrperiod WHERE bldgnum='"&building&"' and billyear="&byear&" order by billperiod", cnn1
      do until rst1.eof
        %><option value="<%=rst1("billperiod")%>"<%if trim(rst1("billperiod"))=trim(bperiod) then response.write " SELECTED"%>><%=rst1("BillPeriod")%></option><%
        rst1.movenext
      loop
      rst1.close
      %>
      </select>
    </td>
    <td>
      <input type="button" value="Print" onclick="window.print()">
    </td>
    <%end if%>
  </tr>
  </table>
  </td>
</tr>
</table>
<%if trim(bperiod)<>"" then
sql = "SELECT DISTINCT meters.meternum,tblleases.billingName,tblleases.flr, peakdemand.billyear,peakdemand.billperiod, peakdemand.usernote as pnote,consumption.usernote as cnote FROM peakdemand JOIN meters ON peakdemand.meterid = meters.meterid JOIN consumption ON consumption.meterid = meters.meterid join tblleasesutilityprices tp on meters.leaseutilityid=tp.leaseutilityid join tblleases on tp.billingid=tblleases.billingid  WHERE peakdemand.billyear = '" & byear & "' AND peakdemand.billperiod = '" & bperiod & "' AND   consumption.billyear = '" & byear & "' AND consumption.billperiod = '" & bperiod & "' AND (peakdemand.usernote IS NOT NULL OR consumption.usernote IS NOT NULL) AND meters.bldgnum = '" & building & "'"
rst1.Open sql, cnn1
if rst1.EOF then%>
<table width="100%" border="0">
  <tr>
    <td>NO PROBLEMS FOUND</td>
  </tr>
</table>
<%else%>
<table width="100%" border="0">
<tr><td bgcolor="#6699CC" class="standardheader" align="center"><b>Genergy Meter Problem Report for Building Number <%=building%></b></td></tr>
<tr><td>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr><td width="8%" bgcolor="#CCCCCC">Meter #</td>
		    <td width="15%" bgcolor="#CCCCCC">Tenant Name</td>
        <td width="7%" bgcolor="#CCCCCC">Floor</td>
        <td width="10%" bgcolor="#CCCCCC" align="center">Period</td>
        <td width="20%" bgcolor="#CCCCCC">User Note (Consumption)</td>
        <td width="40%" bgcolor="#CCCCCC">User Note (Peakdemand)</td>
    </tr>
    </table>
    </td>
</tr>
<tr><td> 
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <% While not rst1.EOF %>
    <tr valign="top"> 
        <td width="8%"><%=rst1("meternum")%></td>
        <td width="15%"><%=rst1("billingname")%></td>
        <td width="7%"><%=rst1("flr")%></td>
        <td width="10%" align="center"><%=rst1("billperiod")%>/<%=rst1("billyear")%></td>
        <td width="20%"><%=rst1("cnote")%></td>
        <td width="40%"><%=rst1("pnote")%></td>
    </tr>
    <tr valign="top"> 
        <td colspan="6"><hr size="1" noshade></td>
    </tr>
    <%
    rst1.movenext
    Wend
    %>
    </table>
    </td>
</tr>
</table>
<p>&nbsp;</p>
<%
end if

end if
%>

</body>
</html>