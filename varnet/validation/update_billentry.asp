<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#include file="checksession.asp"-->
<%
dim byear, bperiod, meterid, tnum, tname, building, super, isposted
meterid = request.querystring("meterid")
building = request.querystring("building")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
tname = request.querystring("tname")
tnum = request.querystring("tnumber")
isposted = request.querystring("posted")
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
'execution
cmd.execute
'if super="True" then 
'	rst1.open "SELECT bval FROM consumption WHERE meterid in (SELECT meterid FROM Meters WHERE bldgnum='"&building&"' and pp<>1) and billperiod="&bperiod&" and billyear="&byear, cnn1
'	if not rst1.eof then
'		response.write "Send mail to "&rst1("bval")&"."
'	else
'		response.write "No biller associated with biller validations!"
'	end if
'	response.end
'end if
%>
<script>
opener.document.location.href="bill_validation.asp?building=<%=building%>&byear=<%=byear%>&bperiod=<%=bperiod%>";
window.close();
</script>
<%	response.end
end if

'strsql = "SELECT * FROM consumption INNER JOIN meters ON meters.meterid=consumption.meterid WHERE consumption.MeterId="& MeterId &" and ((BillYear="& byear &" and BillPeriod<="& bperiod &") or (BillYear="& byear-1 &" and BillPeriod>="& bperiod &")) ORDER BY BillYear desc, BillPeriod desc"
strsql = "SELECT m.MeterNum, c.BillYear, c.BillPeriod, OnPeak, OffPeak, PrevKWH, CurrentKWH, KWHUsed, demand, c.usernote as comments, datelastread FROM consumption c INNER JOIN meters m ON m.meterid=c.meterid INNER JOIN PeakDemand pd ON pd.MeterId=c.MeterId and c.BillYear=pd.BillYear and c.BillPeriod=pd.BillPeriod LEFT JOIN validation v ON v.MeterId=c.MeterId AND v.BillPeriod=c.BillPeriod AND v.BillYear=c.BillYear WHERE c.MeterId="& MeterId &" and ((c.BillYear="& byear &" and c.BillPeriod<="& bperiod &") or (c.BillYear="& byear-1 &" and c.BillPeriod>="& bperiod &")) ORDER BY c.BillYear desc, c.BillPeriod desc"
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

<span style="font-family: Arial, Helvetica, sans-serif;font-size:12;">Tenant #: <%=tnum%>&nbsp;&nbsp;Tenant Name: <%=tname%>&nbsp;&nbsp;Meter: <%=rst1("meternum")%></span>
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
<td>Notes</td>
<td></td>
</tr>
<%if not rst1.EOF then%>
<form name="form1" method="post">
<tr style="background-color:#CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12;" valign="top">
<td><%=rst1("BillYear")%><input type="hidden" name="meterid" value="<%=meterid%>"><input type="hidden" name="byear" value="<%=byear%>"><input type="hidden" name="bperiod" value="<%=bperiod%>"></td>
<td><%=rst1("BillPeriod")%></td>
<td><%=rst1("datelastread")%></td>
<%if isposted<>"True" then%>
<td><input type="text" name="OnPeak" value="<%=rst1("OnPeak")%>" size="5"></td>
<td><input type="text" name="OffPeak" value="<%=rst1("OffPeak")%>" size="5"></td>
<td><input type="hidden" name="PrevKWHOrig" value="<%=rst1("PrevKWH")%>" size="15"><input type="text" name="PrevKWH" value="<%=rst1("PrevKWH")%>" size="10"></td>
<td><input type="text" name="CurrentKWH" value="<%=rst1("CurrentKWH")%>" size="15" onKeyUp="KWHUsed.value=this.value-PrevKWH.value"></td>
<td><input type="text" name="KWHUsed" readonly value="<%=rst1("KWHUsed")%>" size="15"></td>
<td><input type="text" name="demand" value="<%=rst1("Demand")%>" size="5"></td>
<td><textarea cols="15" rows="" name="note"><%=rst1("comments")%></textarea></td>
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
<td><%=rst1("comments")%></td>
<td></td>
<%end if%>

<input type="hidden" name="building" value="<%=building%>">
</form>
</tr>
<%rst1.movenext
end if%>
<%
do until rst1.EOF
	response.write "<tr style=""background-color:#CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12;"" valign=""top"">"
	response.write "<td>"& rst1("BillYear") &"</td>"
	response.write "<td>"& rst1("BillPeriod") &"</td>"
	response.write "<td>"& rst1("datelastread") &"</td>"
	response.write "<td>"& rst1("OnPeak") &"</td>"
	response.write "<td>"& rst1("OffPeak") &"</td>"
	response.write "<td>"& rst1("PrevKWH") &"</td>"
	response.write "<td>"& rst1("CurrentKWH") &"</td>"
	response.write "<td>"& rst1("KWHUsed") &"</td>"
	response.write "<td>"& rst1("Demand") &"</td>"
	response.write "<td>"& rst1("comments") &"</td>"
	response.write "<td></td>"
	response.write "</tr>"
	rst1.movenext
loop
%>
</table>
</body>
</html>