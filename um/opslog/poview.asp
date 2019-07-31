<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if


dim po, poid, jid,pdf,onload,callerstatus
PO= Request("po")
POID = Request("poid")
jid = Request("jid")
callerstatus= trim(request("caller"))


dim ticket
set ticket = New tickets
ticket.Label="PO"
ticket.Note = "PO Master Ticket for PoID "&po
ticket.ccuid  = ""
ticket.client = 1
ticket.requester = "POADMIN"
ticket.department = "OPERATIONS"
ticket.userid = "POADMIN"
if po<>"0" and po<>"" and isnumeric(po) then ticket.findtickets "PoID", po

Dim cnn1,rst1,rst2,permissionflag, sqlstr, printview, printwidth,printcolor
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
if not pdf then 
	if allowGroups("Genergy_Corp") then
	  permissionflag=""
	else
	  permissionflag="1"
	end if
end if 

onload = "preloadImages();"
if lcase(trim(request("printview")))="yes" then
	pdf = true 
	printview = true
else
	printview = false
	pdf = false
end if

printwidth = "100%"

if printview then printcolor = "#000000" else printcolor="#cccccc" end if
dim num1,num2,num3
if  POID=""  then 
if len(split(PO,".")(0)) = 5 then num1=7 else num1=6 
if len(split(PO,".")(0)) = 5 then num2=5 else num2=4 
if num1=7 then num3=4 else num3=3
end if

if  POID=""  then 
	'10000 series update
	sqlstr = "select company, po.*,po.id as poid, ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber from po INNER JOIN MASTER_JOB mj ON mj.id=po.jobnum where po.jobnum=substring('" & PO & "',1,"& num2 &") and po.ponum=substring('" & PO & "',"& num1 &","& num3 &")"
else
	sqlstr = "select company, po.*,po.id as poid, ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber from po INNER JOIN MASTER_JOB mj ON mj.id=po.jobnum where po.id=" & POID
end if
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then 
	if rst1("company") = "gy" then 
		if rst1("accepted") then 
		titletext = "Purchase Order"
		else
		titletext = "Requisition"
		end if 
	else
		titletext = "Purchase Order"
	end if 
	if poid = "" then poid = rst1("poid") end if 
%>

<html>
<head>
<title>View <%=titletext%> Form</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<% if not pdf then %>
<script language="JavaScript" type="text/javascript">
//<!--
function processpo(poid,action, shipping) {
  var poaction
  if (action=="Submit For Review") {
    poaction="submit"
  } else {
    poaction="delete"
  } 
  
  var temp = "processpo.asp?poid=" + poid + "&poaction=" + poaction + "&jid=<%=jid%>&caller=<% if request("caller")="joblog" then %>joblog<% end if %>" + "&ponum=<%=server.urlencode(rst1("ponum"))%>&jobnum=<%=server.urlencode(rst1("jobnum"))%>&jobaddr=<%=server.urlencode(rst1("Jobaddr"))%>&shipaddr=<%=server.urlencode(rst1("shipaddr"))%>&poamt=<%=server.urlencode(rst1("po_total"))%>&podesc=<%=server.urlencode(replace(rst1("description"),vbCrlf," "))%>"
  document.location=temp
}

function closepo(id1,acctponum){
	if ((acctponum == null) || (acctponum == "" )){
		alert("please enter a value for PO Number before closing this RF.");
	}else {
		document.location="accpofilter.asp?id1="+id1+"&acctponum="+acctponum
	}
}

function openwin(url,mwidth,mheight){
  window.name="opener";
  popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  popwin.focus();
}

function checkshipping(shipping) {

  if (shipping <= 0) {
  
    alert("Shipping is currently $0, please double check.")
  }

}
function printpo(poid){
  var temp = "poreport.asp?id1="+poid
  window.open(temp,"", "scrollbars=yes,width=800, height=600, status=no" );

}
function withdrawlpo(id1,podate,ponum) {

  var temp = "wdpo.asp?id1=" + id1 + "&podate="+ podate+"&ponum=" + ponum  + "&jid=<%=jid%>&caller=<% if request("caller")="joblog" then %>joblog<% end if %>"
  document.location = temp;
//  window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );
}
function comment(poid,ponum,pocomment){
  var temp = "pocomment.asp?id1="+poid+"&ponum=" + ponum+ "&pocomment=" +pocomment
  window.open(temp,"", "scrollbars=yes,width=550, height=300, status=no" );

}

function editjob(jid) {
  theURL="/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid=" + jid
  //window.document.all['genjobtable'].Border ="1"
  window.document.all['genjobtable'].bgColor ="#990000"
  openwin(theURL,750,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

disable = 0;

function updateEntry(id,poid,q,u,i,p,d,togglestate){
  if (togglestate) {
    textdisplay = "none";
    fielddisplay = "inline";
    document.all[id + "row"].style.backgroundColor = "#eeeeee";
    disable = 1; //disable row highlighting
  } else {
    textdisplay = "inline";
    fielddisplay = "none";  
    document.all[id + "row"].style.backgroundColor = "#ffffff";
    disable = 0;
  }
  document.all[id + "qty"].style.display = textdisplay;
 // document.all[id + "unit"].style.display = textdisplay;
  document.all[id + "invnum"].style.display = textdisplay;
  document.all[id + "unitprice"].style.display = textdisplay;
  document.all[id + "description"].style.display = textdisplay;
  document.all[id + "buttons"].style.display = textdisplay;

  document.all[id + "qty" + "field"].style.display = fielddisplay;
  //document.all[id + "unit" + "field"].style.display = fielddisplay;
  document.all[id + "invnum" + "field"].style.display = fielddisplay;
  document.all[id + "unitprice" + "field"].style.display = fielddisplay;
  document.all[id + "description" + "field"].style.display = fielddisplay;
  document.all[id + "buttons" + "field"].style.display = fielddisplay;
}


function highlight(tRow){
  if (!disable) {
    tRow.style.backgroundColor = "lightgreen";
  }
}

function unlight(tRow){
  if (!disable) {
    tRow.style.backgroundColor = "white";
  }
}

function newPO(){
  <% if request("caller")="joblog" then %>
  document.location="newpo.asp?jid=<%=jid%>&caller=joblog"
  <% else %>
  document.location="newpo.asp?jid=<%=jid%>"
  <% end if %>
}

function viewinvoice(jid) {
  theURL="timesheetmain.asp?flag=0&job=" + jid
  //window.document.all['genjobtable'].Border ="1"
  window.document.all['genjobtable'].bgColor ="#999999"
  openwin(theURL,750,550)
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

//-->
</script>
<% end if %>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body <%if printview then%>bgcolor="#ffffff"<%else%>bgcolor="#eeeeee" <%end if%>text="#000000">

<%
if callerstatus = "joblog" then 'joblog callerstatus

		Dim tcolor,Desc,company,job,jtype,cstatus
		rst2.Open "SELECT description, type, company, job, status FROM MASTER_JOB WHERE id='"&jid&"'", cnn1
		if not rst2.EOF then
			Desc = rst2("description")
			company = rst2("company") 
			job = rst2("job")
			jtype=left(rst2("type"),6)
			cStatus = lcase(rst2("status"))
			Select Case cStatus
				case "in progress"
					tcolor = "#66ff66"
				case "unstarted"
					tcolor = "#ffcc00"
				case "closed"
					tcolor = "#cc0033"
			end select 
		end if
		rst2.close
		
		%>
		<table border=0 cellpadding="3" cellspacing="0" width="<%=printwidth%>" id="genjobtable" style="border-top:1px solid #cccccc;">
		<tr> 
		  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
		  <table border="0" cellspacing="0" cellpadding="3" width="<%=printwidth%>">
		  <tr>
			<td>
			<a href="<%="/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" & jid %>">General Info</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobtime.asp?jid=" &jid%>">Job Time</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job Folder</a> &nbsp;|&nbsp; <b><a href="posearch.asp?caller=<%=request("caller")%>&select=jobnum&findvar=<%=jid%>">Requisition Forms</a></b>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>&caller=<%=request("caller")%>">Job Cost</a> &nbsp;|&nbsp; <a href="/genergy2_intranet/opsmanager/joblog/viewchange.asp?jid=<%=jid%>">Change Orders</a>
			</td>
		  </tr>
		  </table>
		  </td>
		  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;" bgcolor="#eeeeee"><img src="images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
		</tr>
		<tr> 
		  <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
		  <table border=0 cellspacing="0" cellpadding="3" width="100%">
		  <tr>
			<td width="130">Job:</td>
			<td><%=Desc%>&nbsp; (<%=company%>: &nbsp;<%=job%>) &nbsp;<%=jtype%></td>
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
		  </tr>
		  <tr>
			<td>Total requisition forms:</td>
			<td>
			<b><!--[[span id="totalPOs"]][[/span]]--> 
			<%
			 dim rst3,str
			 Set rst3 = Server.CreateObject("ADODB.recordset")
			str="select sum(po_total) as sum1 from po where jobnum=" & jid & " and accepted = 1"
			rst3.Open str, cnn1, 0, 1, 1
			if rst3("sum1") > 0 then %>
			  <%=Formatcurrency(rst3("sum1"))%>
			  <% else%>
			  $0 
			<% end if 
			rst3.close
			%>
			</b>
			</td>
		  </tr>
		  <tr>
				<td height="25">Total hours invoiced:</td>
			<td>
			<b>
			<%
			str="select sum(hours) as sum1 from invoice_submission where jobno=" & jid & " and submitted=1"
			rst3.Open str, cnn1, 0, 1, 1
			if trim(rst3("sum1")) > 0 then %>
			  <%=rst3("sum1")%>
			  <% else %>
			  0    
			<% end if
			rst3.close 
			%>
			</b>
			</td>
		  </tr>
		  </table>
		  </td>
		</tr></table>
		<%  
end if 'end Joblog callerstatus


dim editable, titletext
editable = (not rst1("submitted")) and not (rst1("accepted"))

if printview and trim(lcase(rst1("company"))) = "ge" then editable = false end if 
%>
<form name="form1" method="post" action="poupdate.asp">
  <table border="0" cellpadding="3" cellspacing="0" width="<%=printwidth%>" id="genjobtable">
<%if printview then%><tr><td colspan=2><img src="/logos/<%=rst1("company")%>_letterhead.gif"></td></tr><%end if%>
    <tr bgcolor="#dddddd"> 
	<%
	


	
	
	Dim headertext
	if printview then 
		headertext = titletext &" Form: # "&rst1("jobnum")&"."& rst1("ponum") 
	else
		headertext = titletext &" Form: # <a href=""javascript:openwin('/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid="& server.urlencode(rst1("jobnum"))&"',550,400)"">"&server.urlencode(rst1("jobnum"))&"</a>."& rst1("ponum") 
	    
	end if 
	%>
     <td><b><%=headertext%></b></td><td><%ticket.Display 0,true, true, false%></td>
       <%if request("printview") <> "yes" then%>
      <td align="right">&nbsp;
        <% if rst1("submitted") and not rst1("closed") then %><input type="button" name="Button3" value="Withdraw Submitted RF" onclick="withdrawlpo(id1.value,podate.value,ponum.value)"><%end if%><input type="button" name="print" value="Print <%=titletext%>" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?devIP=<%=request.servervariables("server_name")%>&sn=<%=request.servervariables("script_name")%>&qs=<%=server.urlencode("poid=" & poid & "&printview=yes")%>','','')">
      <%end if%></td>
    </tr>
    <tr>       
    <td colspan="2">
	<% 
		if printview then 
			printblock
		else
			editblock
		end if 
	%>
	
	</td>
	</tr>
	</table>
</form>
<div id="poitem" style="overflow:auto;width:<%=printwidth%>;background-color:#ffffff;">
  <table border=0 cellpadding="3" cellspacing="1" bgcolor="<%=printcolor%>" width="<%=printwidth%>">
    <form name="newitemform" method="post" action="savepoitem.asp">
      <tr bgcolor="#dddddd" style="font-weight:bold;">
 			 <td style="border-top:1px solid #ffffff;" colspan=5 bgcolor="#eeeeee"><b><%=titletext%> Items</b></td>
      </tr>
      <tr bgcolor="#dddddd" style="font-weight:bold;"> 
        <td width="5%">Qty</td>
        <!--<td width="15%">Unit</td>-->
        <td width="15%">Item #</td>
        <td width="40%">Description</td>
	  <td width="10%" align="center">
          <%if not(rst1("accepted")) and not printview then%>
          Estimated<br>
          <%end if%>
          Unit Cost</td>
        <td align="center" width="10%">
          <%if printview then%>
          Total Cost
          <% end if%>
        </td>
	</tr>
      <% if (not rst1("submitted") and not rst1("accepted")) and editable then  %>
      <tr bgcolor="#eeeeee"> 
        <td><input type="text" name="qty" size="4"></td>
        <!--<td><input type="text" name="unit"></td>-->
        <td><input type="text" name="invnum"></td>
        <td><input type="text" name="description" size="70"></td>
        <td>$&nbsp;
          <input type="text" name="unitprice" size="16"></td>
        <td> <input type="submit" name="choice2"  value="Save"> <input type="hidden" name="poid" value="<%=rst1("id")%>"> <input type="hidden" name="caller" value="<%=callerstatus%>">  
        </td>
      </tr>
      <% end if %>
    </form>
    <!--[[/table]]-->
    <%
sqlstr = "select * from po_item where poid='" & rst1("id") & "' order by id desc"

rst2.Open sqlstr, cnn1, 0, 1, 1


dim i, posubtotal,emptyrows
posubtotal 	= 	0
emptyrows 	=	1
While not rst2.EOF 
%>
    <form name="form<%=rst2("id")%>" method="post" action="poitemupdate.asp">
      <% if not rst1("submitted") and not rst1("accepted") and editable then  %>
      <tr id="<%=rst2("id")%>row" bgcolor="#ffffff"> 
        <td> <span id="<%=rst2("id")%>qty" style="display:inline;"><%=rst2("qty")%></span> 
          <span id="<%=rst2("id")%>qtyfield" style="display:none;">
          <input type="text" name="qty" size="4" value="<%=rst2("qty")%>">
          </span> </td>
        <!--<td>
  <span id="<%'=rst2("id")%>unit" style="display:inline;"><%'=rst2("unit")%></span>
  <span id="<%'=rst2("id")%>unitfield" style="display:none;"><input type="text" name="unit" value="<%'=rst2("unit")%>"></span>
  </td>-->
        <td> <span id="<%=rst2("id")%>invnum" style="display:inline;"><%=server.htmlencode(rst2("invnum"))%></span>
          <span id="<%=rst2("id")%>invnumfield" style="display:none;">
          <input type="text" name="invnum" value="<%=server.htmlencode(rst2("invnum"))%>">
          </span> </td>
        <td> <span id="<%=rst2("id")%>description" style="display:inline;"> <%=server.htmlencode(rst2("description"))%>
          </span> <span id="<%=rst2("id")%>descriptionfield" style="display:none;">
          <input type="text" name="description" size="70" value="<%=server.htmlencode(rst2("description"))%>">
          </span> </td>
                <td> <span id="<%=rst2("id")%>unitprice" style="display:inline;"><%=server.htmlencode(FormatCurrency(rst2("unitprice")))%></span> 
          <span id="<%=rst2("id")%>unitpricefield" style="display:none;">$&nbsp;
          <input type="text" name="unitprice" size="16" value="<%=server.htmlencode(rst2("unitprice"))%>">
          </span> </td>
		<td width="180" height="30"> <span id="<%=rst2("id")%>buttons" style="display:inline;"><img src="images/btn-edit.gif" value="Edit" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this, '#ffffff');" onclick="updateEntry('<%=replace(server.htmlencode(rst2("id")),"'","\'")%>',poid.value,'<%=replace(server.htmlencode(rst2("qty")),"'","\'")%>','<%=replace(server.htmlencode(rst2("unit")),"'","\'")%>','<%=replace(server.htmlencode(rst2("invnum")),"'","\'")%>','<%=replace(server.htmlencode(FormatCurrency(rst2("unitprice"))),"'","\'")%>','<%=replace(server.htmlencode(rst2("description")),"'","\'")%>',1);" style="border:1px solid #ffffff;"></span> 
          <span id="<%=rst2("id")%>buttonsfield" style="display:none;"> 
          <input type="hidden"	name="boolDelete"	value="true">
          <input type="button"	name="update"	value="Update" onclick="document.form<%=rst2("id")%>.boolDelete.value = 'false'; form<%=rst2("id")%>.submit()">
          <input type="button"	name="delete"	value="Delete" onclick="document.form<%=rst2("id")%>.boolDelete.value = 'true' ; form<%=rst2("id")%>.submit()">
          <input type="button"	name="cancel"	value="Cancel" onclick="updateEntry('<%=rst2("id")%>',poid.value,'<%=replace(server.htmlencode(rst2("qty")),"'","\'")%>','<%=replace(server.htmlencode(rst2("unit")),"'","\'")%>','<%=replace(server.htmlencode(rst2("invnum")),"'","\'")%>','<%=replace(server.htmlencode(FormatCurrency(rst2("unitprice"))),"'","\'")%>','<%=replace(server.htmlencode(rst2("description")),"'","\'")%>',0);">
          </span> </td>
      </tr>
      <input type="hidden" name="key" value="<%=rst2("id")%>">
      <input type="hidden" name="poid" value="<%=rst1("id")%>"> <input type="hidden" name="caller" value="<%=callerstatus%>"> 
      <input type="hidden" name="jid" value="<%=jid%>">
      <% if request("caller")="joblog" then %>
      <input type="hidden" name="caller" value="joblog">
      <% end if %>
      <% else %>
      <tr bgcolor="#ffffff"> 
        <td height="30"><%=rst2("qty")%></td>
        <!--<td><%'=rst2("unit")%></td>-->
        <td height="30"><%=rst2("invnum")%></td>
        <td height="30"><%=rst2("description")%></td>
        <td height="30" align="right"><%=FormatCurrency(rst2("unitprice"))%></td>
        <td width="180" height="30" align="right">
          <%if printview then%>
          <%=formatcurrency(rst2("unitprice")*rst2("qty"))%>
          <% end if%>
        </td>
      </tr>
      <input type="hidden" name="key" value="<%=rst2("id")%>">
      <input type="hidden" name="poid" value="<%=rst1("id")%>"> <input type="hidden" name="caller" value="<%=callerstatus%>"> 
      <input type="hidden" name="jid" value="<%=jid%>">
      <% end if %>
    </form>
<%
posubtotal = posubtotal + (rst2("unitprice")*rst2("qty"))
emptyrows 	= 	emptyrows + 1
rst2.movenext
Wend

rst2.close
%>
<%
if printview then
	if emptyrows <= 13 then 
		for i  = emptyrows to 13
		%>
			<tr  bgcolor="#ffffff"> 
			  <td height=30></td>
			  <td height=30></td>
			  <td height=30></td>
			  <td height=30></td>
			  <td height=30></td>
			</tr>
	<%
		next
	end if 
%>
<tr bgcolor="#FFFFFF">
	<td height=30>&nbsp;</td>
	<td height=30>&nbsp;</td>
	<td height=30>&nbsp;</td>
	<td align="right" height=30><b>Subtotal</b></td>
	<td align="right" height=30>&nbsp;<% if posubtotal<>0 then%><b><%=formatcurrency(posubtotal)%></b><% end if%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td height=30>&nbsp;</td>
<td height=30>&nbsp;</td>
<td height=30>&nbsp;</td>
<td align="right" height=30><b>Tax</b></td>
        <td align="right">&nbsp;<% if (posubtotal*rst1("tax"))<>0 then%><%=formatcurrency(posubtotal*rst1("tax"))%><% end if%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td height=30>&nbsp;</td>
<td height=30>&nbsp;</td>
<td height=30>&nbsp;</td>
<td align="right" height=30><b>Total</b></td>
        <td align="right">&nbsp;<% if (rst1("po_total"))<>0 then%><%=formatcurrency(rst1("po_total"))%><% end if%></td>
</tr>
	<%
	if ucase(rst1("company")) = "GE" then 
	%>
		<tr  bgcolor="#ffffff"> 
		  <td height=30 colspan=5 align="center">FOREMAN:_________________________ 
			CALL IN ORDER:________________________</td>
		</tr>
		<%
	end if 


end if 
%>
  </table>
<% 
if printview then 
	disclaimer(rst1("company"))
end if
%>
</div>
<%if not rst1("submitted") and not rst1("accepted") and editable then  %>
	<div id="podetail" style="overflow:auto;width:100%;border:1px solid #cccccc;display:none;">
		<table border=0 cellpadding="3" cellspacing="1" width="100%">
			<form name="podetailform" method="post" action="poitemupdate.asp">
				<tr bgcolor="#dddddd" style="font-weight:bold;">
					<td width="8%">Qty</td>
					<!--<td width="15%">Unit</td>-->
					<td width="15%">Item #</td>
					<td>Description</td>
					<td width="15%">Estimated Price</td>
				</tr>
				
				<tr> 
					<td><input type="text" name="qty" size="4"></td>
					<!-- <td><input type="text" name="unit"></td>-->
					<td><input type="text" name="invnum"></td>
					<td>
						<input type="text" name="description" size="70">
						<input type="submit" name="choice2"  value="Update">
						<input type="hidden" name="key">
						<input type="hidden" name="poid" value="<%=rst1("id")%>"> <input type="hidden" name="caller" value="<%=callerstatus%>"> 
					</td>
					<td>$&nbsp;<input type="text" name="unitprice" size="16"></td>
				</tr>  
			</form>
		</table>
	</div>
	
	<div id="newitem" style="overflow:auto;width:100%;border:1px solid #cccccc;display:none;">
		<table border=0 cellpadding="3" cellspacing="1" width="100%">
			<form name="newitemform" method="post" action="savepoitem.asp">
				<tr bgcolor="#dddddd" style="font-weight:bold;">
					<td width="8%">Qty</td>
					<!--<td width="15%">Unit</td>-->
					<td width="15%">Item #</td>
					<td>Description</td>
					<td width="15%"><%if not(rst1("accepted")) then%>Estimated<%end if%> Price</td>
				</tr>
				
				<tr> 
					<td><input type="text" name="qty" size="4"></td>
					<!--<td><input type="text" name="unit"></td>-->
					<td><input type="text" name="invnum"></td>
					<td>
						<input type="text" name="description" size="70">
						<input type="submit" name="choice2"  value="Save">
						<input type="hidden" name="poid" value="<%=rst1("id")%>"> <input type="hidden" name="caller" value="<%=callerstatus%>"> 
					</td>
					<td>$&nbsp;<input type="text" name="unitprice" size="16"></td>
				</tr>  
			</form>
		</table>
</div>
<% if request("caller")="joblog" and not pdf then %>
<br>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;">
<tr>
  <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr><%
    if cStatus<>"closed" or session("temp_permis_closejob")="1" then %>
    <td><img src="/um/opslog/images/btn-edit_job.gif" value="Edit Job" align="middle" onclick="editjob('<%=jid%>');return false;" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td><%end if%>
    <td><img src="/um/opslog/images/btn-invoice.gif" value="Invoice Job" id="invoice" align="middle" onclick="viewinvoice('<%=jid%>');" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
    <td><img src="/um/opslog/images/btn-new_po.gif" value="New Requisition Form" id="new_po" align="middle" onclick="newPO();" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
  </tr>
  </table> 
  </td>
  <td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>

<%
end if 'caller=joblog
end if
%>
</body>
</html>
<% else %>
PO <%=po%> NOT FOUND
<%
end if 
rst1.close

function editblock()
%>
<table border=0 align="left" cellpadding="3" cellspacing="0">
  <tr valign="top"> 
            <td> 
              <!-- begin first column -->
              <table border=0 cellpadding="3" cellspacing="0">
                <tr> 
                  <td align="right">Date</td>
                  <td><input type="hidden" name="podate" value="<%=rst1("podate")%>"> 
                    <%=rst1("podate")%> </td>
                </tr>
                <tr valign="middle"> 
                  <td align="right">Status</td>
                  <td> <table border=0 cellpadding="0" cellspacing="0">
                      <tr valign="middle">
                        <%
			dim color, text
			
			if (rst1("submitted")) and not (rst1("accepted")) then
				color = "red"
				text = "Submitted"
			elseif rst1("accepted") and (not rst1("approved")) then
				color = "yellow"
				text = "Accepted"
			elseif rst1("approved") and (not rst1("closed")) then
				color = "green"
				text = "Approved"
			elseif rst1("closed") then
				color = "green"
				text = "Closed"
			else
				color = "red"
				text = "Not yet submitted"
			end if		%>
                        <td><div style="position:inline;width:18px;height:12px;background-color:<%=color%>;border:1px solid #999999;padding:2px;font-size:7pt;">RF</div></td>
                        <td width="6">&nbsp;</td>
                        <td><%=text%></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td align="right">Vendor</td>
                  <td> 
                    <% if (editable) then %>
                    <%if trim(rst1("vendor"))<>"" then%>
                    <input type="text" name="vendor" value="<%=rst1("vendor")%>" size="20" maxlength="40"> 
                    <%end if%>
                    <select name="vid">
                      <%
			rst2.Open "SELECT distinct vendor, name FROM " & rst1("company") & "_MASTER_APM_VENDOR order by name", getConnect(0,0,"intranet")
			do until rst2.eof
				%>
                      <option value="<%=rst2("vendor")%>" <%if rst1("vid")=rst2("vendor") then%>SELECTED<%end if%>><%=rst2("name")%></option>
                      <%
				rst2.movenext
			loop
			rst2.close
			%>
                    </select> 
                    <%
					else
						rst2.open "SELECT name FROM " & rst1("company") & "_MASTER_APM_VENDOR WHERE vendor='"&rst1("vid")&"'", getConnect(0,0,"intranet")
						if rst2.eof then
							response.write rst1("vendor")
						else
							response.write rst2("name")
						end if
						rst2.close
                    end if %>
                  </td>
                </tr>
                <tr valign="top"> 
                  <td align="right">Submitted by</td>
                  <td>
                    <%=rst1("submittedby")%> 
                  </td>
                </tr>
                <tr valign="top"> 
                  <td align="right">Requisitioner</td>
                  <td>
                    <%
			if (editable) then %>
                    <select name="req">
                      <%
					sqlstr = "select [first name]+' '+ [last name] as name, substring(username,7,20) as user1 from employees order by [last name]"
					rst2.Open sqlstr, cnn1, 0, 1, 1
					if not rst2.eof then
						do until rst2.eof				%>
                      <option value="<%=rst2("user1")%>" <%If lcase(trim(rst1("requistioner")))= lcase(trim(rst2("user1"))) then%> selected<%end if%>><%=rst2("name")%></option>
                      <%
							rst2.movenext
						loop
					end if
					rst2.close
					%>
                    </select>
                    <%
                     
			else 
					%>
                    <%=rst1("requistioner")%> 
                    <%
			end if %>
                  </td>
                </tr>
                <%
	if (rst1("accepted")) then 
		dim acceptedOutput
		acceptedOutput = rst1("accepted_user")
		if trim(acceptedOutput) = "" then
			acceptedOutput= "N/A"
		end if%>
                <tr> 
                  <td align="right">Accepted By</td>
                  <td><%=acceptedOutput%></td>
                </tr>
                <%
	end if
	if (rst1("approved")) then 
		dim approveOutput
		approveOutput = rst1("approved_user")
		if trim(approveOutput) = "" then
			approveOutput = "N/A"
		end if%>
                <tr> 
                  <td align="right">Approved By</td>
                  <td><%=approveOutput%></td>
                </tr>
                <%
	end if%>
                <tr> 
                  <td align="right">Job Address</td>
                  <td>
                    <% if (editable) then %>
                    <input type="text" name="jobaddr" value="<%=server.htmlencode(rst1("jobaddr"))%>" size="40" maxlength="40">
                    <% else %>
                    <%=rst1("jobaddr")%>
                    <% end if %>
                  </td>
                </tr>
                <tr valign="top"> 
                  <td align="right">Shipping Address</td>
                  <td>
                    <% if (editable) then %>
                    <input type="text" name="shipaddr" value="<%=server.htmlencode(rst1("shipaddr"))%>" >
                    <% else %>
                    <%=rst1("shipaddr")%>
                    <% end if %>
                  </td>
                </tr>
                <tr valign="top"> 
                  <td>Administrative Comment</td>
                  <td> <%=rst1("admin_comment")%> <input type="hidden" name="ponum" value="<%=rst1("ponumber")%>"> 
                    <input type="hidden" name="id1" value="<%=rst1("id")%>"> </td>
                </tr>
                <%
		if ((not rst1("closed")) and (rst1("approved"))  and  allowGroups("Genergy Accounting") ) then%>
                <tr valign="top"> 
                  <td>PO Number</td>
                  <td> <input type="text" name="acctponum" value=""> </td>
                </tr>
                <tr> 
                  <td></td>
                  <td nowrap> <input type="button" value="Close" onclick="closepo(id1.value,acctponum.value);"> 
                  </td>
                </tr>
                <%
		end if	
		if rst1("closed") and rst1("approved") then		%>
                <tr valign="top"> 
                  <td>PO Number</td>
                  <td> <%=rst1("acct_ponum")%> </td>
                </tr>
                <%
		end if		%>
                <% if not rst1("submitted") and not rst1("accepted") and editable then %>
                <tr> 
                  <td>&nbsp;</td>
                  <td> 
                    <% if request("caller")="joblog" then %>
                    <input type="hidden" name="caller" value="joblog"> <input type="hidden" name="jid" value="<%=jid%>"> 
                    <% end if %>
                    <input type="submit" name="Button5" value="Update"> <input type="button" name="Button3" value="Submit For Review" onClick="processpo(id1.value,this.value,0)"> 
                    <input type="button" name="Button6" value="Delete" onClick="processpo(id1.value,this.value,0)"> 
                <% end if %>
				   &nbsp;
				   </td>
				   </tr>
              </table>
              <!-- end first column -->
            </td>
            <td width="30">&nbsp;</td>
            <td> 
              <!-- begin second column -->
              <p style="padding:3px;"> <u><%=titletext%> Form Description</u><br>
                <% if (editable) then %>
                <textarea name="description" cols="25" rows="3" wrap="PHYSICAL" ><%=rst1("description")%></textarea>
                <% else %>
                <%=rst1("description")%>
                <% end if %>
              </p>
			  <% if not(rst1("accepted")) then %>
              <table border=0 cellpadding="3" cellspacing="0" style="border:1px solid #999999;padding:3px;">
                <tr valign="top"> 
                  <td align="right">Tax Rate</td>
                  <td>%</td>
                  <td>
                    <% if (editable) then %>
                    <input type="text" name="tax1" value="<%=rst1("tax")%>" >
                    <% else %>
                    <%=rst1("tax")%>
                    <% end if %>
                  </td>
                </tr>
                <tr> 
                  <td align="right"><b>
                    <%if not(rst1("accepted")) then%>
                    Estimated
                    <%end if%>
                    Total</b></td>
                  <td>$</td>
                  <td><b><%=FormatCurrency(rst1("po_total"))%> 
                    <input type="hidden" name="total" value="<%=FormatCurrency(rst1("po_total"))%>">
                    </b></td>
                </tr>
              </table>
			  <%end if%>
			  </td>
    </tr>
  </table>
<%
end function 
function printblock()
%>
<table border=0 align="left" cellpadding="3" cellspacing="0" width="<%=printwidth%>">
  <tr valign="top">
    <td colspan=3><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="2"><b>Vendor:</b></td>
          <td colspan="2"><b>Job Address:</b></td>
        </tr>
        <tr> 
          <td colspan="2" valign="top">
		  	<%
		  		rst2.open "SELECT * FROM " & rst1("company") & "_MASTER_APM_VENDOR WHERE vendor='"&rst1("vid")&"'", getConnect(0,0,"intranet")
						if rst2.eof then
							response.write rst1("vendor") 
						else
							if rst2("name") <> "" then 
								response.write rst2("name")& "<br>" 
							end if
							if rst2("Contact_1_Name")  <> "" then 
								response.write rst2("Contact_1_Name") & "<br>" 
							end if
							if rst2("telephone") <> "" then 
								response.write "Tel: " & rst2("telephone") & "<br>" 
							end if
							if rst2("Fax_Number") <> "" then 
								response.write "Fax: " & rst2("Fax_Number")  & "<br>"
							end if
							if rst2("Address_1") <> "" then 
								response.write rst2("Address_1")  & "<br>" 
							end if
							if rst2("Address_2") <> "" then 
								response.write  rst2("Address_2")  & "<br>"
							end if
							if rst2("city") <> "" then 
								response.write rst2("city")  & "," 
							end if
							if rst2("state") <> "" then 
								response.write rst2("state")  & " "
							end if
							if rst2("zip") <> "" then 
								response.write rst2("zip")
							end if
						end if
						rst2.close
			%>
		  </td>
          <td colspan="2" valign="top"><%=rst1("jobaddr")%>&nbsp;</td>
        </tr>
		<tr><td colspan=4><hr></td></tr>
        <tr>
          <td width="24%" align="left" colspan=2><b>Requisitioner</b></td>
          <td width="26%" align="left"><b>P.O. Date</b></td>
          <td width="25%" align="left"><b>P.O. Number</b></td>
        </tr>
        <tr> 
          <td colspan=2><font size="3">
		  <%					
		  	sqlstr = "select [first name]+' '+ [last name] as name, substring(username,7,20) as user1 from employees where substring(username,7,20) = '"&trim(rst1("requistioner"))&"'"
			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
				response.write rst2("name")
			else
				response.write rst1("requistioner")
			end if
			rst2.close
		  %>
		</font></td>
          <td><font size="3"><%=rst1("podate")%></font></td>
          <td><font size="3"><%=rst1("jobnum")&"."& rst1("ponum")%></font></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
end function 
function disclaimer(company)

	select case lcase(company)
	
	case "ge"
		%>
		<br>1. Please Notify us <u>immediately</u> if you are unable to ship as specified
		<br>2. Please if there are any <u>back orders</u> call (212) 974 - 5199
		<br>3. Please notify us <u>immediately</u> if there are any freight charges above noted on PO
	<%
	case "gy"
		%>
		<br>1. Please Notify us <u>immediately</u> if you are unable to ship as specified
		<br>2. Please if there are any <u>back orders</u> call _________________ 
		<br>3. Please notify us <u>immediately</u> if there are any freight charges above noted on PO
	<%
	end select
end function
%>
