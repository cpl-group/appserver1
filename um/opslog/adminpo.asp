<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<%
		if isempty(getKeyValue("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function fillup(typebox){
	document.location="oplogindex.asp?typebox=" + typebox
}
function searchjob(typebox, searchitem, comp, var2) {
//alert(typebox)
//alert(searchitem)
	var temp
	if(searchitem != ""){
		if (document.forms[0].comp.checked) {
		    if(var2 !=""){
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=1&var="+var2
			}else{
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=1"
			}
		} else {
		    if(var2 !=""){
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=0&var="+var2
			}else{
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=0"
			}
		}
		document.frames.oplog.location=temp
    }else{
		alert("At least type something...")
	}
}
function report(spec, job) {
document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
//alert(typebox)
//alert(document.forms[0].findvar.value)
	var temp
	if(typebox =="[entry id]"){
	    if(job == ""){
		    alert("Please enter job number")
		}else{
            if(isNaN(job)){
				alert("Not a valid number")
            }else{
                temp="timesheetmain.asp?job="+job
//				temp="timesheetsearch.asp?job="+job
				document.frames.oplog.location=temp
			}
		}
	}else{
		temp="null.htm"
		document.frames.oplog.location=temp
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


</script>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>
<%
ReDim Ay(5)
ReDim By(5)
Ay(0) = "[entry id]"
Ay(1) = "customer"
Ay(2) = "manager"
Ay(3) = "[current status]"
Ay(4) = "[description]"
By(0) = "Job Number"
By(1) = "Customer ID"
By(2) = "Manager ID"
By(3) = "Status"
By(4) = "Description"
msg = Request.querystring("msg")
typebox = Request("typebox")
			if isempty(msg) then
				msg="Please enter search and click the FIND button to begin"
			end if
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

		
%>
<body bgcolor="#eeeeee" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#666699"> 
    <td><span class="standardheader">Requisition Form Administration</span></td>
    <td align="right"><input type="button" name="Submit" value="Print Current View" onClick='javascript:document.frames.poadmin.focus();document.frames.poadmin.print()'></td>
  </tr>
</table>
<!--
[[table border=0 cellpadding="6" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"]]
  [[tr]]
    [[td]][[a href="corppoview.asp" target="poadmin" style="color:#333366;"]]Approve/Reject Submitted PO's[[/a]] &nbsp;|&nbsp; [[a href="acctpoview.asp" target="poadmin" style="color:#333366;"]]View Approved PO's[[/a]]
[[!~~
      [[img src="images/btn-approve_pos.gif" alt="Approve/Reject PO's" onClick="Javascript:frames.poadmin.location='corppoview.asp'" border="0" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this, '#ffffff');" style="border:1px solid #ffffff;"]]
      [[img src="images/btn-view_approved_pos.gif" alt="Approved PO's" onClick="Javascript:frames.poadmin.location='acctpoview.asp'" border="0" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this, '#ffffff');" style="border:1px solid #ffffff;"]]
~~]]
    [[/td]]
  [[/tr]]
[[/table]]
-->
<IFRAME name="poadmin" width="100%" height="90%" src="corppoview.asp" scrolling="auto" marginwidth="0" marginheight="0" border="0" frameborder="0"></IFRAME>
</body>
</html>