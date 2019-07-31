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
end if	

dim user, user1, cnn1, rst1, sql

user="ghnet\"& Request.querystring("user")
user1=Request.querystring("user")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")
sql="select startweek, endweek from time_submission where username='payroll'"


rst1.Open sql, cnn1, 0, 1, 1
dim startweek, endweek
if not rst1.eof then
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
rst1.close

dim strsql
strsql = "SELECT t.*, t.matricola AS Expr1, e.[first name]+ ' '+e.[last name] as name1, substring(e.username,7,20) as user1 FROM Times t join employees e on e.username=t.matricola WHERE (t.matricola = '"& user &"'  and t.[date] between '" & Startweek &"' and '" & endweek &"' ) order by t.date desc"

dim rsname, rsuser

rst1.Open strsql, cnn1, 0, 1, 1
if not rst1.eof then
	rsname = rst1("name1")
	rsuser = rst1("user1")
end if
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
function sortByJobNum(pUser){
	var urllink="timesheetreviewbyjob.asp?user="+pUser
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
  <td><b>Time Sheet for <%=rsname%> for the week between <%=startweek%> and <%=endweek%></b></td>
  <td align="right">
  <input type="hidden" name="user" value="<%=rsuser%>">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
  	<td><button name="SortJNum" value="Sort By Job Number" onClick="sortByJobNum('<%=user1%>')" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" class="standard">&nbsp;Sort By Job Number</button></td>
    <td><button name="Submit" value="Print Time Sheet" onClick="printtime(user.value)" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" class="standard"><img src="/um/opslog/images/printer.gif" align="absmiddle" hspace="3" border="0">&nbsp;Print...</button></td>
    <td>&nbsp;</td>
    <td><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
  </tr>
  </table>
  </td>
</tr>
</form>
</table>
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
<tr bgcolor="#eeeeee">
    <td colspan="3" class="tblunderline">&nbsp;</td>
  <td colspan="3" class="tblunderline" width="21%" align="center"><b>Hours</b></td>
  <td colspan="2" class="tblunderline" width="14%" align="center"><b>Expenses</b></td>
</tr>
<tr bgcolor="#eeeeee" style="font-weight:bold;">
  <td class="tblunderline">Date</td>
  <td class="tblunderline">Job #</td>
  <td class="tblunderline">Description of Time</td>
  <td class="tblunderline" width="7%" bgcolor="#ffffee" style="border-left:1px solid #e3e3d3;">Regular</td>
  <td class="tblunderline" width="7%" bgcolor="#f0f0e0">Billable</td>
  <td class="tblunderline" width="7%" bgcolor="#e3e3d3">Overtime</td>
  <td class="tblunderline" width="7%">Descr.</td>
  <td class="tblunderline" width="7%">Amount</td>
</tr>
<%
if not rst1.eof then
	Do until rst1.EOF 
%>
<tr bgcolor="#<%if cdbl(rst1("value"))>0 then%>ffcccc<%else%>ffffff<%end if%>" valign="top"> 
<form name=form1 method="post" action="">
<input type="hidden" name="key" value="<%=rst1("id")%>">
  <td class="tblunderline"><a href="javascript:updateentry('<%=rst1("id")%>','<%=user1%>')"><%=rst1("date")%></a></td>
  <td class="tblunderline"><a href="javascript:openjob('<%=trim(rst1("jobno"))%>','<%=trim(rst1("jobno"))%>')"><%=trim(rst1("jobno"))%></a></td>
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
loop
end if
'rst1.close
%>
</table>
 
</body>
</html>
