<%@Language="VBScript"%>
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim jid,longjob,cStatus,permissionflag,company,jtype,tcolor
jid=secureRequest("jid")

'if session("corp")=5 then
  permissionflag=""
'else
 ' permissionflag="1"
'end if

Dim cnn1,rst1,sqlstr,orderdate
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

sqlstr="select change_order.id,jobno,company,status,type,change_order.description,amount,accepted,TimeStamp from change_order join master_job on jobno=job where master_job.id=" & jid &" order by timestamp"
rst1.Open sqlstr
if rst1.eof then
  rst1.close
  rst1.open "select job,status,company,type from master_job where id='" & jid & "'"
  longjob=rst1("job")
  cStatus=rst1("status")
  company=rst1("company")
  jtype=rst1("type")
  rst1.movenext
else
  longjob=rst1("jobno")
  cStatus=rst1("status")
  company=rst1("company")
  jtype=rst1("type")
end if
Select Case cStatus
case "In progress"
  tcolor = "#66ff66"
case "Unstarted"
  tcolor = "#ffcc00"
case "Closed"
  tcolor = "#cc0033"
end select 

showChangeOrderButton = true
%>
<html>
<head>
<title>View Change Order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
//<!--
var theURL;
function highlight(tRow){
  tRow.style.backgroundColor = "lightgreen";
}

function unlight(tRow){
  tRow.style.backgroundColor = "white";
  
}
function openwin(url,mwidth,mheight){
  newwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
function newchange(){
  theURL="updatechange.asp?jid=<%=longjob%>&mode=new"
  openwin(theURL,375,165)
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
	}
	
function editchange(changeID){
  theURL="updatechange.asp?jid=<%=longjob%>&change_id="+changeID+"&mode=edit"
  openwin(theURL,375,165)
}
var loaded = 0;
function preloadImages(){
  edit_jobOn = new Image(); edit_jobOn.src = "images/btn-edit_job-1.gif";
  edit_jobOff = new Image(); edit_jobOff.src = "images/btn-edit_job.gif";
  new_poOn = new Image(); new_poOn.src = "images/btn-new_po-1.gif";
  new_poOff = new Image(); new_poOff.src = "images/btn-new_po.gif";
  invoiceOn = new Image(); invoiceOn.src = "images/btn-invoice-1.gif";
  invoiceOff = new Image(); invoiceOff.src = "images/btn-invoice.gif";
  
  loaded = 1;
}

function msover(img){
  if ((loaded) && (document.all)) {
    img.src = eval (img.id + "On.src");
  }
}

function msout(img){
  if ((loaded) && (document.all)) {
    img.src = eval (img.id + "Off.src");
  }
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
function edit_job(jid) {
  theURL="/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid=" + jid
  //window.document.all['genjobtable'].Border ="1"
  window.document.all['genjobtable'].bgColor ="#dddddd"
  openwin(theURL,750,400)
}
//-->
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">
</head>
<body bgcolor="#eeeeee" onload="preloadImages();">
<form>
<table border=0 cellpadding="3" cellspacing="0" width="100%" id="genjobtable" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;">
<tr> 
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
  <table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
            <td> <a href="<%="/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" & jid %>">General 
              Info</a> |&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=longjob%>&c=<%=company%>&mode=view">Job 
              Estimates</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobtime.asp?jid=" &jid%>">Job 
              Time</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job 
              Folder</a> &nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requisition 
              Forms</a>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=longjob%>&jid=<%=jid%>&caller=joblog">Job 
              Cost</a>&nbsp;|&nbsp;<b>Change Orders</b>&nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a> </td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;" bgcolor="#eeeeee"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr> 
  <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
    <td width="130">Job:</td>
    <td><%=company%>: &nbsp;<%=longjob%> &nbsp;(<%=jtype%>)</td>
  </tr>
  <tr>
    <td>Status:</td>
    <td>
    <table border=0 cellpadding="0" cellspacing="0">
    <tr>
      <td><div style="position:inline;width:18px;height:12px;background:<%=tcolor%>;border:1px solid #999999;">&nbsp;</div></td>
      <td width="6">&nbsp;</td>
      <td><%=cStatus%></td>
    </tr>
    </table>
    </td>
  </tr><%
if rst1.EOF then%>
<tr>
  <td colspan="2" bgcolor="#ffffff" style="border-top:2px outset #ffffff;padding:6px;">No change orders were found under the requested job id<br>&nbsp;</td>
</tr>
</table>
<%  else   %>
</table>
<table border=0 cellpadding="3" cellspacing="1" width="100%">
<tr bgcolor="#dddddd" style="font-weight:bold;"> 
  <td width="60%">Description</td>
  <td width="20%">Amount</td>   	
  <td width="9%">Accepted</td>
  <td width="10%">Date</td>
</tr>
</table>
<div style="width:100%; overflow:auto; height:200px;">  
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#dddddd">
<% While not rst1.EOF %>
<tr bgcolor="#ffffff" onMouseOver="highlight(this);" onMouseOut="unlight(this);" onClick="editchange(<%=rst1("id")%>)" style="cursor:hand"> 
  <td width="60%"><%=rst1("description")%></td>
  <td width="20%"><%=FormatCurrency(rst1("amount"))%></td>
  <td width="9%"><%=replace(replace(rst1("accepted"),"False","No"),"True","Yes")%></td>
 <% if isnull(rst1("TimeStamp")) then orderdate="N/A" else orderdate =rst1("TimeStamp") end if%> 
 <td width="10%"><%=orderdate%></td>
</tr>
<%
x=x+1
rst1.movenext
Wend
%>
<tr><td colspan="5" bgcolor="#eeeeee"><i><%=x%> change orders found</i></td></tr>
</table>
</div>

<%
  rst1.close
  end if
%>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;">
<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
  <td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
</form>
</body>
</html>