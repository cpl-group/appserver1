<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if isempty(getKeyValue("name")) then
%>
<script>
top.location="http://www.genergyonline.com"
</script>
<%
			'	Response.Redirect "http://www.genergyonline.com"
end if	
if trim(request("name")) = "" then 
	user="ghnet\"&trim(getXMLUserName())
else
	user="ghnet\"&trim(request("name")) 
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

sql="select username, startweek, endweek from user_cost where username='"&user&"'"
rst1.Open sql, cnn1, 0, 1, 1

if not rst1.eof then
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
rst1.close
strsql = "SELECT *, matricola AS Expr1 FROM Times WHERE (matricola = '"& user &"'  and [date] between '" & Startweek - 18  & "' and '" & endweek  &"') order by date desc"
rst1.Open strsql, cnn1, 0, 1, 1
%>
<html>
<head>
<title>Timesheet</title>
<script>
function openpopup(){
//configure "Open Logout Window
    parent.document.location.href="../index.asp";
}
function loadpopup(){
    openpopup()
}
function updateEntry(id){
	parent.frames.tsbottom.location="timedetail.asp?id="+id+"&name=<%=trim(request("name"))%>"
}
function displaytotal(hrs, ot, expn, lastdate){

	var temp = "Totals as of " + lastdate + " : Hours = " + hrs + ", Overtime Hours =  " + ot + ", Expenses = " + expn
//	alert(temp) 
  document.all.totalspan.innerHTML = temp;

} 
function delete1(key,u){
	if(confirm("Are you sure you want to delete this entry?")){
	document.location="deletetime.asp?key="+key+"&u="+u
	}
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #cccccc; }
</style>
<head>
<body bgcolor="FFFFFF" class="innerbody">
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
<!--
<tr bgcolor="#ffffff">
  <td colspan="10"><input type="button" name="Submit3" value="Totals" onClick="displaytotal(form2.hrstotal.value,form2.bhrstotal.value,form2.expensetotal.value, form2.lastdate.value)" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">&nbsp;<span id="totalspan" class="notetext"></span></td>
</tr>
-->
<tr bgcolor="#dddddd" valign="bottom">
  <td align="center" valign="middle" class="tblunderline"><input type="button" name="Submit3" value="Totals" onClick="displaytotal(form2.hrstotal.value,form2.bhrstotal.value,form2.expensetotal.value, form2.lastdate.value)" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
  <td colspan="4" class="tblunderline" valign="middle"><span id="totalspan" class="notetext"></span>&nbsp;</td>
  <td colspan="3" class="tblunderline" width="21%" align="center"><b>Hours</b></td>
  <td colspan="2" class="tblunderline" width="14%" align="center"><b>Expenses</b></td>
</tr>
<%
dim showheader, daynumber
showheader = 0
daynumber = -1
  
if not rst1.eof then
timesheettotal_hrs = 0
timesheettotal_ot = 0
expensetotal=0
lastdate=rst1("date")
	Do until rst1.EOF 
%>
<form name=form1 method="post" action="">
<input type="hidden" name="key" value="<%=rst1("id")%>">
<input type="hidden" name="u" value="<%=rst1("expr1")%>">
<% 


if (weekday(rst1("date")) = 6) then showheader = 1 else showheader = 0
daynumber = weekday(rst1("date"))
if (showheader) then
 %>
<tr bgcolor="#eeeeee" style="font-weight:bold;">
  <td class="tblunderline" width="10%">&nbsp;</td>
  <td colspan="2" class="tblunderline" width="10%">Day &amp; Date</td>
  <td class="tblunderline" width="8%">Job#</td>
  <td class="tblunderline">Description</td>
  <td class="tblunderline" width="7%" bgcolor="#ffffee" style="border-left:1px solid #e3e3d3;">Regular</td>
  <td class="tblunderline" width="7%" bgcolor="#f0f0e0">Billable</td>
  <td class="tblunderline" width="7%" bgcolor="#e3e3d3">Overtime</td>
  <td class="tblunderline" width="7%">Descr.</td>
  <td class="tblunderline" width="7%">Amount</td>
</tr>
<% end if %>
<tr bgcolor="#ffffff" valign="middle"> 
  <td align="center" class="tblunderline">
  <input type="button" name="edit" value="Edit" onClick="updateEntry(key.value)" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
  <a href="javascript:delete1('<%=rst1("id")%>','<%=rst1("expr1")%>')" style="background-color:#eeeeee;"><img src="/um/opslog/delete.gif" align="absmiddle" border="0"></a>
  </td>
  <td align="center" class="tblunderline"><%=shortDay(weekday(rst1("date")))%>&nbsp;</td>
  <td class="tblunderline"><%=rst1("date")%>&nbsp;</td>
  <td class="tblunderline"><%=rst1("jobno")%>&nbsp;</td>
  <td class="tblunderline"><%=rst1("description")%>&nbsp;</td>
  <td bgcolor="#ffffee" align="right" class="tblunderline" style="border-left:1px solid #e3e3d3;"> 
  <%=rst1("hours")%> 
  <%  if rst1("date") >= startweek then
  timesheettotal_hrs=timesheettotal_hrs + Formatnumber(rst1("hours"))
  end if%>
  &nbsp;</td>
  <td bgcolor="#f0f0e0" align="right" class="tblunderline"><%=rst1("hours_bill")%>&nbsp;</td>
  <td bgcolor="#e3e3d3" align="right" class="tblunderline"> 
  <%=rst1("overt")%> 
  <% if rst1("date") >= startweek then
  timesheettotal_ot=timesheettotal_ot + Formatnumber(rst1("overt")) 
  end if%>
  &nbsp;</td>
  <td bgcolor="#ffffff" class="tblunderline"><%=rst1("expense")%>&nbsp;</td>
  <td bgcolor="#ffffff" class="tblunderline"><%=FormatCurrency(rst1("value"))%> 
  <% if rst1("date") >= startweek then 
  expensetotal=expensetotal + Formatnumber(rst1("value"))
  end if %>
  &nbsp;</td>
</tr>
</form>
  <%  
    rst1.movenext
    loop
end if
%>
</table>
<form name="form2" method="post" action="">

<input type="hidden" name="hrstotal" value="<%=timesheettotal_hrs%>">
<input type="hidden" name="bhrstotal" value="<%=timesheettotal_ot%>">
<input type="hidden" name="expensetotal" value="<%=Formatcurrency(expensetotal)%>">
<input type="hidden" name="lastdate" value="<%=lastdate%>">
</form>

<%
function shortDay(someday)
  select case someday
    case 1
      shortDay = "Su"
    case 2
      shortDay = "M"
    case 3
      shortDay = "Tu"
    case 4
      shortDay = "W"
    case 5
      shortDay = "Th"
    case 6
      shortDay = "F"
    case 7
      shortDay = "Sa"
  end select
end function 
%>
</body>
</html>
