<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if isempty(getKeyValue("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'	Response.Redirect "http://www.genergyonline.com"
else
	if getKeyValue("ts") < 4 then 
		getKeyValue("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
		Response.Redirect "../main.asp"
	end if	
end if	
dim user, user1, cnn1, rst1, sql
user="ghnet\"& secureRequest("user")
user1=secureRequest("user")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")
sql="select startweek, endweek from time_submission where username='payroll'"


rst1.Open sql, cnn1, 0, 1, 1

if not rst1.eof then
	dim startweek, endweek
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
rst1.close
dim strsql
strsql = "SELECT t.*, t.matricola AS Expr1, e.[first name]+ ' '+e.[last name] as name1, substring(e.username,7,20) as user1 FROM Times t join employees e on e.username=t.matricola WHERE (t.matricola = '"& user &"'  and t.[date] between '" & Startweek &"' and '" & endweek &"' ) order by t.JobNo, t.[date] desc"

rst1.Open strsql, cnn1, 0, 1, 1

%>
<html>
<head>
<script language="JavaScript" type="text/javascript">
function openpopup(){
//configure "Open Logout Window
    parent.document.location.href="../index.asp";
}
function loadpopup(){
    openpopup()
}
function printtime(uname){
    //var temp="timeprint.asp"
	if (uname == 'Print All Timesheets'){
		var temp="timetemplateall.asp"
	} else {
		var temp="timetemplate.asp?user=" + uname 
		}
	window.open(temp,"", "scrollbars=yes,width=800, height=600, status=no" );
}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}
function openwin(url,mwidth,mheight){
cwin = window.open(url,"childwin","status=no, menubar=no,HEIGHT="+mheight+", WIDTH="+mwidth)
cwin.focus()
}
function openjob(jobno,jid)
{
var urlLink  = "https://appserver1.genergy.com/um/war/jc/jc1.asp?c=GY&jg=" + jobno+"&jid="+jid
window.open(urlLink,"window","scrollbars=no,width=900,height=600,resizeable")
}
function updateentry(id,name){
	var urllink="timedetail.asp?id="+id+"&name="+name+"&source=review"
	window.open(urllink,"window","scrollbars=no,width=600,height=150")
}
function sortByDate(pUser){
	var urllink="timesheetreview.asp?user="+pUser
	window.navigate(urllink)
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #cccccc; }
</style>
</head>
<body bgcolor="FFFFFF" style="border-top:1px solid #ffffff;">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<form name="form2" method="post" action="">
<tr bgcolor="#eeeeee">
  
  <td><b>Time Sheet for <% if not(rst1.eof) then Response.write(rst1("name1")) end if%> for the week between <%=startweek%> and <%=endweek%></b></td>
  <td align="right">
  
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
   	<td><button name="SortDate" value="Sort By Date" onClick="sortByDate('<%=user1%>')" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" class="standard">&nbsp;Sort By Date</button></td>
    <td><button name="Submit" value="Print Time Sheet" onClick="printtime(user.value)" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" class="standard"><img src="/um/opslog/images/printer.gif" align="absmiddle" hspace="3" border="0">&nbsp;Print...</button></td>
    <td>&nbsp;</td>
    <td><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
  </tr>
  </table>
  </td>
</tr>
</form>
</table>
<%
if not rst1.eof then
	dim currentJobNum, workingJobNum
	currentJobNum = rst1("JobNo")		' the job number outside the inner loop
	Do until rst1.EOF 
	workingJobNum = rst1("JobNo")		' the job number inside the inner loop, is
										' compared to current to make sure we're still
										' in the same group.
%>
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
<tr bgcolor="#eeeeee">
  <td colspan="2" class="tblunderline"><b>Job Number: <a href="javascript:openjob('<%=trim(rst1("jobno"))%>','<%=trim(rst1("jobno"))%>')"><%=trim(rst1("jobno"))%></a></b></td>
  <td colspan="3" class="tblunderline" width="21%" align="center"><b>Hours</b></td>
  <td colspan="2" class="tblunderline" width="14%" align="center"><b>Expenses</b></td>
</tr>
<tr bgcolor="#eeeeee" style="font-weight:bold;">
  <td class="tblunderline" width="20%">Date</td>
  <td class="tblunderline" width="45%">Description of Time</td>
  <td class="tblunderline" width="7%" bgcolor="#ffffee" style="border-left:1px solid #e3e3d3;">Regular</td>
  <td class="tblunderline" width="7%" bgcolor="#f0f0e0">Billable</td>
  <td class="tblunderline" width="7%" bgcolor="#e3e3d3">Overtime</td>
  <td class="tblunderline" width="10%">Descr.</td>
  <td class="tblunderline" width="5%">Amount</td>
</tr>
<%
	dim totalRegHours, totalBillhours, totalOTHours, totalExpenses
	totalRegHours = 0
	totalBillHours = 0
	totalOTHours = 0
	totalExpenses = 0
	Do until ((currentJobNum <> workingJobNum) OR (rst1.eof))
		
		totalRegHours  = cDbl(totalRegHours) +  cDbl(rst1("hours"))
		totalBillHours = cDbl(totalBillHours) + cDbl(rst1("hours_bill"))
		totalOTHours   = cDbl(totalOTHours) +   cdbl(rst1("overt"))
		totalExpenses  = cdbl(totalExpenses) +  cdbl(rst1("value"))
		
%>
<tr bgcolor="#ffffff" valign="top"> 
<form name=form1 method="post" action="">
<input type="hidden" name="key" value="<%=rst1("id")%>"> 
  <td class="tblunderline"><a href="javascript:updateentry('<%=rst1("id")%>','<%=user1%>')"><%=rst1("date")%></a></td>
 
  <td class="tblunderline"><%=rst1("description")%>&nbsp;</td>
  <td bgcolor="#ffffee" align="right" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("hours")%>&nbsp;</td>
  <td bgcolor="#f0f0e0" align="right" class="tblunderline"><%=rst1("hours_bill")%>&nbsp;</td>
  <td bgcolor="#e3e3d3" align="right" class="tblunderline"><%=rst1("overt")%>&nbsp;</td>
  <td class="tblunderline"><%=rst1("expense")%>&nbsp;</td>
  <td class="tblunderline">$<%=rst1("value")%>&nbsp;</td>
</form>
</tr>
<%  
	rst1.movenext
	if (not(rst1.eof)) then
		workingJobNum = rst1("JobNo")
	end if
	

	loop  'end of inner loop, which steps through all of the entries for a particular job num

	if (not(rst1.eof)) then
		currentJobNum = rst1("JobNo")
	end if
%>
<tr bgcolor="#ffffff" valign="top"> 
<form name=form1 method="post" action="">
  <td class="tblunderline">&nbsp;</td>
  <td class="tblunderline" align = "right"><b>Totals:</b></td>
  <td bgcolor="#ffffee" align="right" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=totalRegHours%>&nbsp;</td>
  <td bgcolor="#f0f0e0" align="right" class="tblunderline"><%=totalBillHours%>&nbsp;</td>
  <td bgcolor="#e3e3d3" align="right" class="tblunderline"><%=totalOTHours%>&nbsp;</td>
  <td class="tblunderline">&nbsp;</td>
  <td class="tblunderline">$<%=totalExpenses%>&nbsp;</td>
</form>
</tr>
</table>
<br>
<%
loop 'end of outer loop, which steps through all the jobnums

end if
rst1.close
%>

 
</body>
</html>
