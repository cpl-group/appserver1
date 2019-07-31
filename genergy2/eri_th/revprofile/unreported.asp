<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim building, pid, date1, action, i, description, amt, period, entrytype, id, utility, utilitydisplay
building = Request("building")
if trim(building) = "" then Request("building")
pid = Request("pid")
date1 = Request("date1")
if trim(date1)="" then date1 = year(date())
action = Request("action")
description = Request("description")
amt = Request("amt")
period = Request("period")
entrytype = Request("type")
id = request("id")
utility = request("utility")
dim rst1, cnn1, sql
Set rst1 = Server.CreateObject("ADODB.recordset")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(building)


if trim(action)="Save" then
  sql = "Insert into tblRPentries (pid, description, amt, period,bldgnum, year, type, utility) values ('"&pid&"', '"&description&"', "&amt&", '"&period&"', '"&building&"', '"&date1&"', '"&entrytype&"', '"&utility&"')"
elseif trim(action)="Update" then
  sql = "Update tblRPentries set description='"&description&"', amt="&amt&", period='"&period&"', type='"&entrytype&"', utility='"&utility&"' where id="&id
elseif trim(action)="Delete" then
  sql = "DELETE FROM tblRPentries WHERE id="&id
end if
'response.write sql
'response.end

if trim(sql)<>"" then cnn1.execute sql

rst1.open "SELECT * FROM tblutility WHERE utilityid="&utility, getConnect(pid,building,"billing")
if not rst1.eof then utilitydisplay = rst1("utilitydisplay")
rst1.close
%>
<html>
<head>
<title>Adjustments</title>
<script>
function loadentry(id, type, amt, period, description)
{ var frm = document.forms[0];
  frm.type[(type == 'True' ? 1 : 0)].checked = true
  frm.description.value = description;
  frm.amt.value = amt;
  frm.period.value = period;
  frm.id.value = id;
  document.all['update'].style.display='inline';
  document.all['save'].style.display='none';
}
</script>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" cellpadding="3" cellspacing="0" border="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#3399cc" class="standardheader">Adjustments : Expenses &amp; Revenue for building <%=building%>, <%=date1%> (<%=utilitydisplay%>)</td></tr>
<tr><td bgcolor="#eeeeee" class="standard"><b>Current Entries</b></td></tr>
</table>

<%
sql = "select *, convert(varchar,entrydate,101) as date from tblRPentries where pid='"&pid&"' and bldgnum='"&building&"' and year ='"&date1&"' and utility="&utility
'response.write sql
'response.end
rst1.Open sql, cnn1
%>
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #cccccc;" width="100%">
<tr bgcolor="#cccccc" class="standard">
    <td width="15%"><b>date</b></td>
    <td width="5%"><b>Exp</b></td>
    <td width="5%"><b>Rev</b></td>
    <td width="35%"><b>Description</b></td>
    <td width="15%"><b>Period</b></td>
    <td width="25%"><b>Total Amount</b></td>
</tr>
</table>
<div align="center" style="overflow: auto; height: 150px;"> 
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<%do until rst1.EOF 
%>
<tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="loadentry(<%=rst1("id")%>, '<%=rst1("type")%>', <%=rst1("amt")%>, <%=rst1("period")%>, '<%=rst1("description")%>')">
    <td width="15%"><%=rst1("date")%></td>
    <td width="5%" align="center"><% if not rst1("type") then %><img src="images/greencheck.gif" width="13" height="15"><%end if%></td>
    <td width="5%" align="center"><% if rst1("type") then %><img src="images/greencheck.gif" width="13" height="15"><%end if%></td>
    <td width="35%"><%=rst1("description")%></font></td>
    <td width="15%"><%=rst1("period")%></td>
    <td width="25%" align="right"><%=FormatCurrency(rst1("amt"))%></td>
</tr>
<%
rst1.movenext
loop
%>
</table>
</div>
<%rst1.close%>

<form name="form1" method="post" action="unreported.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td bgcolor="#CCCCCC">Entry</td></tr>
<tr><td>
    <table border="0" cellspacing="0" cellpadding="0">
    <tr><td><input type="radio" name="type" value="0" checked> Expense <input type="radio" name="type" value="1"> Revenue</td>
    </tr>
    <tr><td>Description (max 150 characters)</td></tr>
    <tr><td><textarea name="description" cols="50" rows="5" wrap="PHYSICAL"></textarea></td></tr>
    <tr><td>Amount $ <input type="text" name="amt" size="15" maxlength="15" value=""> Period 
        <select name="period">
          <option value="0">date1</option>
          <%for i = 1 to 12%>
            <option value="<%=i%>">Period <%=i%></option>
          <%next%>
        </select></td>
    </tr>
    </table>
</td></tr>
</table>&nbsp;<br>
<input type="hidden" name="id" value="">
<input type="hidden" name="date1" value="<%=date1%>">
<input type="hidden" name="building" value="<%=building%>">
<input type="hidden" name="utility" value="<%=utility%>">
<input type="hidden" name="pid" value="<%=pid%>">
<div id="save" style="display: inline;">
<input type="submit" name="action" value="Save">
</div>
<div id="update" style="display: none;">
<input type="submit" name="action" value="Update">
<input type="submit" name="action" value="Delete">
<input type="button" value="Cancel" onclick="loadentry('', '', '', '', '');document.all['update'].style.display='none';document.all['save'].style.display='inline';">
</div><br>
<!-- <input type="button" value="Close" onClick="window.close()"> -->
</form>
</body>
</html>
