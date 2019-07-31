<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#include file="checksession.asp"-->
<%
dim byear, bperiod, meterid, tnum, tname, building, super, isposted, pest, cest, yscroll
meterid = request.querystring("meterid")
building = request.querystring("building")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
tname = request.querystring("tname")
tnum = request.querystring("tnumber")
isposted = request.querystring("posted")
yscroll = request("yscroll")
if session("isSuper")="True" then
	super=true
else
	super=false
end if


dim rst1, cnn1, strsql, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy1")

dim startdate, enddate, datasource
rst1.open "SELECT b.DateStart, b.DateEnd, datasource FROM billyrperiod b, meters m WHERE b.bldgnum=m.bldgnum and billyear="&byear&" and billperiod="&bperiod&" and m.meterid="&meterid, cnn1
if not rst1.eof then
startdate = rst1("DateStart")
enddate = rst1("DateEnd")
datasource = rst1("datasource")
end if
rst1.close
if request.form("action")="Save" then
  'stored proc scripting
  cnn1.CursorLocation = adUseClient
  'specify stored procedure to run
  cmd.CommandText = "sp_validation"
  cmd.CommandType = adCmdStoredProc
  'input params
  Set prm = cmd.CreateParameter("meterid", adInteger, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("BY", adInteger, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("BP", adSmallInt, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("kwh", adDouble, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("kw", adDouble, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("on", adDouble, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("off", adDouble, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("diff", adDouble, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("note", adVarChar, adParamInput, 250)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("PrevKWH", adDouble, adParamInput)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("pdnote", adVarChar, adParamInput, 250)
  cmd.Parameters.Append prm
  'set connection
  Set cmd.ActiveConnection = cnn1
  'set input params
  cmd.Parameters("meterid")	= meterid
  cmd.Parameters("BY")		= byear
  cmd.Parameters("BP")		= bperiod
  cmd.Parameters("kwh")		= request.form("CurrentKWH")
  cmd.Parameters("kw")		= request.form("Demand")
  cmd.Parameters("user")		= session("login")
  cmd.Parameters("on")		= request.form("OnPeak")
  cmd.Parameters("off")		= request.form("OffPeak")
  cmd.Parameters("diff")		= request.form("KWHUsed")
  cmd.Parameters("note")		= request.form("note")
  cmd.Parameters("PrevKWH")	= request.form("PrevKWH")
  cmd.Parameters("pdnote")		= request("pdnote")
  cmd.execute
  %>
  <script>
  var yscroll = opener.scrollpoint()
  opener.document.location.href="bill_validation.asp?building=<%=building%>&byear=<%=byear%>&bperiod=<%=bperiod%>&yscroll="+yscroll;
  window.close();
  </script>
  <%	response.end
end if

'strsql = "SELECT * FROM consumption INNER JOIN meters ON meters.meterid=consumption.meterid WHERE consumption.MeterId="& MeterId &" and ((BillYear="& byear &" and BillPeriod<="& bperiod &") or (BillYear="& byear-1 &" and BillPeriod>="& bperiod &")) ORDER BY BillYear desc, BillPeriod desc"
strsql = "SELECT c.estimated as cest, pd.estimated as pest, m.MeterNum, c.BillYear, c.BillPeriod, OnPeak+isnull(intpeak,0) as onpeak, OffPeak, PrevKWH, CurrentKWH, KWHUsed, demand, c.usernote as note, datelastread, pd.usernote as pdnote FROM consumption c INNER JOIN meters m ON m.meterid=c.meterid INNER JOIN PeakDemand pd ON pd.MeterId=c.MeterId and c.BillYear=pd.BillYear and c.BillPeriod=pd.BillPeriod LEFT JOIN validation v ON v.MeterId=c.MeterId AND v.BillPeriod=c.BillPeriod AND v.BillYear=c.BillYear WHERE c.MeterId="& MeterId &" and ((c.BillYear="& byear &" and c.BillPeriod<="& bperiod &") or (c.BillYear="& byear-1 &" and c.BillPeriod>="& bperiod &")) ORDER BY c.BillYear desc, c.BillPeriod desc"
'response.write strsql
'response.end
rst1.open strsql, cnn1
%>
<html>
<head><title>Bill Validation</title>
<script>
function makesubmit()
{	var frm = document.forms['form1']
	if(frm.PrevKWH.value!=frm.PrevKWHOrig.value)
	{	if(frm.note.value=='')
		{	alert('Must enter a note when changing the previous KWH value');
			return(0);
		}
	}
	frm.submit();
}
</script>
</head>
<body leftmargin="0" topmargin="0">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr style="font-family: Arial, Helvetica, sans-serif;font-size:12;">
    <td>Tenant #: <%=tnum%>&nbsp;&nbsp;Tenant Name: <%=tname%>&nbsp;&nbsp;Meter: <%=rst1("meternum")%></td>
    <!-- <td align="right">&nbsp;&nbsp;<%if trim(datasource)<>"" then%><a class="standardheader" href="/genergy2/UMreports/meterPulseReport.asp?meterid=<%=meterid%>&startdate=<%=startdate%>&enddate=<%=enddate%>&genergy1=yes" target="_blank">View Pulse Data</a><%end if%></td> -->
</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="3">
<tr style="background-color: #0099FF; font-family: Arial, Helvetica, sans-serif;font-size:13">
<td>Year</td>
<td>Period</td>
<td>Date&nbsp;Last&nbsp;Read</td>
<td>On&nbsp;Peak</td>
<td>Off&nbsp;Peak</td>
<td>Prev&nbsp;KWH Reading</td>
<td>Current&nbsp;KWH Reading</td>
<td>Current&nbsp;KWH Usage</td>
<td>KW</td>
<td>Consumption Notes</td>
<td>Peak&nbsp;Demand Notes</td>
<td></td>
</tr>
<%
dim currentbp
if not rst1.EOF then
	currentbp = trim(rst1("billperiod"))
%>
 
<form name="form1" method="post">
<tr style="background-color:#CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12;" valign="top">
<td><%=rst1("BillYear")%><input type="hidden" name="meterid" value="<%=meterid%>"><input type="hidden" name="byear" value="<%=byear%>"><input type="hidden" name="bperiod" value="<%=bperiod%>"></td>
<td><%=rst1("BillPeriod")%></td>
<td><%=rst1("datelastread")%></td>
<%if isposted<>"True" then%>
<td><input type="text" name="OnPeak" value="<%=rst1("OnPeak")%>" size="5"></td>
<td><input type="text" name="OffPeak" value="<%=rst1("OffPeak")%>" size="5"></td>
<td><input type="hidden" name="PrevKWHOrig" value="<%=rst1("PrevKWH")%>"><input type="text" name="PrevKWH" value="<%=rst1("PrevKWH")%>" size="7"></td>
<td><input type="text" name="CurrentKWH" value="<%=rst1("CurrentKWH")%>" size="7" onKeyUp="KWHUsed.value=this.value-PrevKWH.value"></td>
<td><input type="text" name="KWHUsed" readonly value="<%=rst1("KWHUsed")%>" size="7"></td>
<td><input type="text" name="demand" value="<%=rst1("Demand")%>" size="5"></td>
<td><textarea cols="10" rows="" name="note"><%=rst1("note")%></textarea></td>
<td><textarea cols="10" rows="" name="pdnote"><%=rst1("pdnote")%></textarea></td>
      <td> 
        <input type="button" onClick="makesubmit()" name="Action2" value="Save" size="15">
        <input type="hidden" name="Action" value="Save"></td>
<%else%>
<td><%=rst1("OnPeak")%></td>
<td><%=rst1("OffPeak")%></td>
<td><%=rst1("PrevKWH")%></td>
<td><%=rst1("CurrentKWH")%></td>
<td><%=rst1("KWHUsed")%></td>
<td><%=rst1("Demand")%></td>
<td><%=rst1("note")%></td>
<td><%=rst1("pdnote")%></td>
<td></td>
<%end if%>

<input type="hidden" name="building" value="<%=building%>">
</form>
</tr>
<%rst1.movenext
end if%>
<%
do until rst1.EOF
  if rst1("cest")="True" then cest = "*" else cest = ""
  if rst1("pest")="True" then pest = "*" else pest = ""
	if trim(rst1("billperiod"))<>currentbp then
		currentbp = trim(rst1("billperiod"))
		response.write "<tr style=""background-color:#CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12;"" valign=""top"">"
		response.write "<td>"& rst1("BillYear") &"</td>"
		response.write "<td>"& rst1("BillPeriod") &"</td>"
		response.write "<td>"& rst1("datelastread") &"</td>"
		response.write "<td>"& rst1("OnPeak") &"</td>"
		response.write "<td>"& rst1("OffPeak") &"</td>"
		response.write "<td>"& rst1("PrevKWH") &"</td>"
		response.write "<td>"& rst1("CurrentKWH") &"</td>"
		response.write "<td>"& rst1("KWHUsed") & cest &"</td>"
		response.write "<td>"& rst1("Demand") & pest &"</td>"
		response.write "<td>"& rst1("note") &"</td>"
		response.write "<td>"& rst1("pdnote") &"</td>"
		response.write "<td></td>"
		response.write "</tr>"
	end if
	rst1.movenext
loop
%>
<tr><td colspan="12" align="right" style="background-color:#CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12;">* reading is estimated</td></tr>
</table>
</body>
</html>