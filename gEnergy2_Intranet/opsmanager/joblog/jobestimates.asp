<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim cnn, rst, strsql, cstatus
dim job,c,eid,fmode,jid, esttype,datestamp
dim est_LCPH, est_LH, est_LC, est_BC, est_OC, est_MC,est_CO, overhead_percent, burden_percent, profit, total,total_job_cost
cstatus = true
job = request("job")
jid = request("jid")
c 	= request("c")
eid = request("eid")
esttype = request("esttype")

select case lcase(request("mode"))
  case "new","edit"
   %>
    <html>
    <head>
    <title>Job Estimate</title>
    <script>
    function checkform()
    {
	update_amt();
      if (document.form1.est_LCPH.value == "" || document.form1.est_LH.value == "" || document.form1.est_LC.value == "" || document.form1.est_OC.value == "" || document.form1.est_BC.value == "" || document.form1.est_MC.value == ""){
		  alert("All Fields Are Required")
      }
	  else{
          document.form1.submit()
      }
    }
function closepage(action)
{
  switch (action)
  {
  case "cancel":
    if (confirm("Cancel Job Estimate?")){
      history.back()
    }

    break;
  default:
    break;
  }
}
function openwin(url,mwidth,mheight){
  newwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
  
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
	}

function edit_job(jid) {
  theURL="/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid=" + jid
  //window.document.all['genjobtable'].Border ="1"
  //window.document.all['genjobtable'].bgColor ="#dddddd"
  openwin(theURL,750,400)
}

function deleteestimate(eid,co)
{
	alert("Deleting a Primary Job Estimate, will delete all change orders for that estimate as well!")
    if (confirm("Delete Job Estimate"+co+"?")){
      document.location="jobestimates.asp?job=<%=job%>&jid=<%=jid%>&c=<%=c%>&mode=delete&eid="+eid+"&esttype="+co
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

function update_amt(){

 	document.form1.est_LC.value = document.form1.est_LCPH.value * document.form1.est_LH.value;
 	document.form1.est_OC.value = document.form1.est_LC.value * document.form1.overhead_per.value;
 	document.form1.est_BC.value = document.form1.est_LC.value * document.form1.burden_per.value;
  	document.form1.total.value = eval(document.form1.est_LC.value) + eval(document.form1.est_OC.value) + eval(document.form1.est_BC.value) + eval(document.form1.est_MC.value) + eval(document.form1.profit.value)
}
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">
</head>
<body bgcolor="#dddddd" onunload="closepage('close')">
<form name="form1" method="get" action="jobestimates.asp">
<%
if trim(esttype) = "" then 
	est_CO = false
else
	est_CO = true
end if 

if eid <> "" then 
	
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")
	
strsql = "select * from "&c&"_job_estimates where id = " & eid 

rst.open strsql, cnn

if not rst.eof then 

	est_LCPH		=	rst("est_Labor_Unit_Cost") 
	est_LH 			=   rst("est_Labor_units") 
	est_LC			= 	rst("est_labor_cost")
	est_BC			=  	rst("est_Burden_Cost") 
	est_OC 			= 	rst("est_overhead_cost") 
	est_MC			= 	rst("est_Material_cost") 
	est_CO			=	rst("change_order") 
	datestamp 		= 	rst("datestamp")
	overhead_percent= rst("ohc_percent")
	burden_percent 	= rst("bc_percent")
	profit 			= rst("profit")
	total 			= est_LC + est_BC + est_OC + est_MC + profit
	fmode 			=	"update"
end if
rst.close
set cnn = nothing
set rst = nothing
else
	fmode 		= 	"save"
	overhead_percent= .5
	burden_percent 	= .91
	est_LCPH		=	0
	est_LH 			=   0
	est_LC			= 	0
	est_BC			=  	0
	est_OC 			= 	0
	est_MC			= 	0
	profit 			=   0
end if 


%>    
    
  <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" ><table border="0" cellspacing="0" cellpadding="3" width="100%">
          <tr> 
            <td nowrap><a href="<%="/gEnergy2_Intranet/opsmanager/joblog/viewjob.asp?jid="&jid%>">General 
              Info</a>&nbsp;| <b>Job Estimates</b> |&nbsp; <a href="<%="/gEnergy2_Intranet/opsmanager/joblog/jobtime.asp?jid="&jid&"&sbu=0"%>">Job 
              Time</a> &nbsp;|&nbsp; <a href="<%="/gEnergy2_Intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job 
              Folder</a>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requisition 
              Forms</a>&nbsp;|&nbsp; <a href="/um/war/jc/jc.asp?c=<%=c%>&j<%=lcase(left(c,1))%>=<%=job%>&jid=<%=jid%>&job=<%=job%>&caller=joblog">Job 
              Cost</a>&nbsp;|&nbsp; <a href="viewchange.asp?jid=<%=jid%>">Change 
              Orders </a></td>
          </tr>
        </table></td>
      <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee"><div id="backbutton2"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></div></td>
    </tr>
    <tr> 
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="3" bgcolor="#6699cc"><span class="standardheader"> 
        <% if est_CO then %>
        Change Order 
        <%else%>
        Primary Job Estimate 
        <% end if %>
        for <%=job%> 
        <%if  datestamp<>"" then %>
        (last changed <%=datestamp%>) 
        <%end if%>
        <input name="est_CO" type="hidden" value="<%=est_CO%>">
        <input name="c" type="hidden" value="<%=c%>">
        <input name="mode" type="hidden" value="<%=fmode%>">
        <input name="job" type="hidden" value="<%=job%>">
        <input name="jid" type="hidden" value="<%=jid%>">
        <input name="eid" type="hidden" value="<%=eid%>">
        </span></td>
    </tr>
    <tr> 
      <td colspan="3" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">All 
        fields are required except if noted<br></td>
    </tr>
    <tr> 
      <td colspan="3" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp; 
      </td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Average 
        Labor Cost Per Hour</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="est_LCPH" value="<%=est_LCPH%>" onchange="update_amt();"></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Labor Hours</td>
      <td width="476" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp;&nbsp; 
        <input type="text" name="est_LH" value="<%=est_LH%>" onchange="update_amt();"></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Labor Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="est_LC" value="<%=est_LC%>"></td>
    </tr>
    <tr> 
      <td width="373" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Overhead Cost</td>
      <td width="382" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Apply 
        Percent 
        <input type="text" name="overhead_per" value="<%=Overhead_Percent%>" onchange="update_amt();"></td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="est_OC" value="<%=est_OC%>" ></td>
    </tr>
    <tr> 
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Burden Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Apply 
        Percent 
        <input type="text" name="burden_per" value="<%=Burden_Percent%>" onchange="update_amt();"></td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="est_BC" value="<%=est_BC%>"></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Materials Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="est_MC" value="<%=est_MC%>" onchange="update_amt()"> </td>
    </tr>
    <tr>
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Profit 
        Amount </td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="profit" value="<%=Profit%>" onchange="update_amt()"></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimate Total</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">$ 
        <input type="text" name="total" value="<%=total%>" disabled></td>
    </tr>
    <tr> 
      <td colspan="3" bgcolor="#dddddd"> <div style="margin-left:5px;"
> 
          <input type="button" value="Save" onclick="checkform();">
          &nbsp; 
          <input type="button" value="Cancel" onclick="closepage('cancel');">
          <% if eid <> "" then %>
          &nbsp; 
          <input type="button" value="Delete" onclick="deleteestimate(<%=eid%>,'<%if est_CO then %> Change Order <%end if%>');">
          <% end if %>
        </div></td>
    </tr>
  </table>
  </form>
    </body>
    </html>
  <%
  
  case "save","update"'Save Job Estimate
  
	est_LCPH	=	request("est_LCPH") 
	est_LH 		=   request("est_LH") 
	est_LC		= 	request("est_LC")
	est_BC		=  	request("est_BC") 
	est_OC 		= 	request("est_OC") 
	est_MC		= 	request("est_MC") 
	est_CO		=	request("est_CO") 
	profit		= 	request("profit")
	overhead_percent = request("overhead_per")
	burden_percent = request("burden_per")
	
	if lcase(est_CO) = "true" then 
	 	est_CO 	= 1
	else
		est_CO	= 0
	end if 
	
    set cnn = server.createobject("ADODB.connection")
	set rst = server.createobject("ADODB.recordset")

    cnn.open getConnect(0,0,"intranet")
	if trim(request("mode")) = "update" then 
		eid 	= 	request("eid")
    	strsql = "update "&c&"_Job_Estimates set est_Labor_Unit_Cost='"&est_LCPH&"',est_Labor_Units='"&est_LH&"', est_Labor_Cost='"&est_LC&"', est_Burden_Cost='"&est_BC&"', est_Overhead_cost='"&est_OC&"',est_Material_cost='"&est_MC&"',datestamp='"&date()&"', ohc_percent="&overhead_percent&", bc_percent= "&burden_percent&", profit="&profit&" where id = " & eid
	else
    	strsql = "insert into "&c&"_Job_Estimates (job,est_Labor_Unit_Cost,est_Labor_Units, est_Labor_Cost, est_Burden_Cost, est_Overhead_cost, est_Material_cost,change_order, ohc_percent, bc_percent,profit) values ('"&job&"','"&est_LCPH&"','"&est_LH&"','"&est_LC&"','"&est_BC&"','"&est_OC&"','"&est_MC&"', '"&est_CO&"',"&overhead_percent&", "&burden_percent&","&profit&")"
	end if
	cnn.Execute strsql
	
	strsql = "select sum(est_overhead_cost)+sum(est_burden_cost)+sum(est_material_cost) + sum(est_labor_cost) + sum(profit) as TotalCost from "&c&"_Job_Estimates where job = '"&job&"'"
	
	rst.open strsql, cnn
	
	if not rst.eof then 
		total_job_cost = rst("totalcost")
	end if
	rst.close
	
	strsql = "update master_job set amt_1 = " & total_job_cost & " where job='"&job&"'"
	
	cnn.Execute strsql
	
	set cnn = nothing

	response.redirect "jobestimates.asp?job="&job&"&jid="&jid&"&c="&c&"&mode=view"
	
  case "delete"'deletes job estimate, if Primary estimate - all change orders are deleted via table trigger
    set cnn = server.createobject("ADODB.connection")
    cnn.open getConnect(0,0,"intranet")
	
	if esttype <> "" then 
    	strsql = "delete from "&c&"_Job_Estimates where id="&eid
	else
    	strsql = "delete from "&c&"_Job_Estimates where job='"&job&"'"
	end if
	cnn.Execute strsql

	set cnn = nothing
	response.redirect "jobestimates.asp?job="&job&"&jid="&jid&"&c="&c&"&mode=view"
  case "view"'List all Job Estimates & Change Orders
   %>
    <html>
    <head>
    <title>Job Estimate</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<script>
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
  newwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function edit_job(jid) {
  theURL="/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid=" + jid
  //window.document.all['genjobtable'].Border ="1"
  //window.document.all['genjobtable'].bgColor ="#dddddd"
  openwin(theURL,750,400)
}


</script>
<body bgcolor="#dddddd">
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee" id="genjobtable">
<tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" ><table border="0" cellspacing="0" cellpadding="3" width="100%">
          <tr> 
            <td nowrap><a href="<%="/gEnergy2_Intranet/opsmanager/joblog/viewjob.asp?jid="&jid%>">General Info</a>&nbsp;| <b>Job Estimates</b> |&nbsp; <a href="<%="/gEnergy2_Intranet/opsmanager/joblog/jobtime.asp?jid="&jid&"&sbu=0"%>">Job 
              Time</a> &nbsp;|&nbsp; <a href="<%="/gEnergy2_Intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job 
              Folder</a>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requisition 
              Forms</a>&nbsp;|&nbsp; <a href="/um/war/jc/jc.asp?c=<%=c%>&j<%=lcase(left(c,1))%>=<%=job%>&jid=<%=jid%>&job=<%=job%>&caller=joblog">Job 
              Cost</a>&nbsp;|&nbsp; <a href="viewchange.asp?jid=<%=jid%>">Change 
              Orders </a>&nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a></td>
          </tr>
        </table></td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee"><div id="backbutton2"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></div></td>
    </tr>
<%
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")
	
strsql = "select * from "&c&"_job_estimates where job = '" & job &"' order by change_order"

rst.open strsql, cnn
if rst.eof then 
	%>
	<tr>
		<td colspan=2>&nbsp;&nbsp;<input type="button" value="Add Estimate" onclick="document.location='jobestimates.asp?job=<%=job%>&jid=<%=jid%>&c=<%=c%>&mode=new'"></td>
	</tr>
	<%
else
	%>
	<tr> 
		<td colspan="2"><input type="button" value="Change Order" onclick="document.location='jobestimates.asp?job=<%=job%>&jid=<%=jid%>&c=<%=c%>&esttype=co&mode=new'">
		</td>
	</tr>
	<%
end if
while not rst.eof 

	est_LCPH	=	formatcurrency(rst("est_Labor_Unit_Cost"))
	est_LH 		=   formatnumber(rst("est_Labor_units"))
	est_LC		= 	formatcurrency(rst("est_labor_cost"))
	est_BC		=  	formatcurrency(rst("est_Burden_Cost")) 
	est_OC 		= 	formatcurrency(rst("est_overhead_cost")) 
	est_MC		= 	formatcurrency(rst("est_Material_cost"))
	est_CO		=	rst("change_order")
	profit		= 	formatcurrency(rst("profit")) 
	datestamp	=	rst("datestamp")	
	eid 		= 	rst("id")
	total 		= 	formatcurrency(rst("est_labor_cost") + rst("est_Burden_Cost") + rst("est_overhead_cost") + rst("est_Material_cost") + rst("profit"))
%>    
    <tr bgcolor="#6699cc"> 
      <td colspan="3"><span class="standardheader"><%if est_CO then %>Change Order<% else %>Primary Job Estimate<%end if%> <%if datestamp <> "" then %> (last changed <%=datestamp%>) <%end if%></span>&nbsp;
</td>
    </tr>
    <tr> 
      <td colspan="3" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp; 
      </td>
    </tr>
    <tr>
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Average 
        Labor Cost Per Hour</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=est_LCPH%></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Labor Hours</td>
      <td width="412" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=est_LH%></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Labor Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=est_LC%></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Overhead Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=est_OC%></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Burden Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=est_BC%></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Estimated 
        Materials Cost</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=est_MC%></td>
    </tr>
    <tr> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Profit</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=profit%></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><b>Estimate Total</b></td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><b><%=total%></b></td>
    </tr>
    <tr bgcolor="#dddddd"> 
      <td colspan="3"><div style="margin-left:5px;"
> 
          <input type="button" value="Edit" onclick="document.location='jobestimates.asp?job=<%=job%>&jid=<%=jid%>&c=<%=c%>&esttype=<%=est_CO%>&mode=edit&eid=<%=eid%>'">
        </div></td>
    </tr>
    <tr bgcolor="#dddddd"> 
      <td colspan="3"><hr></td>
    </tr>
<%
rst.movenext
wend
	rst.close
	set cnn = nothing
	set rst = nothing
 
%>
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;">
<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
  <td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
    </body>
    </html>
  <%
  case else
end select
%>