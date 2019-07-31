<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
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

</script>
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

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

		
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Invoice 
        Administration</font></b></font></div>
    </td>
  </tr>
</table>
  
<table width="100%" border="0">
  <tr> 
    <td width="90%"> 
      <input type="button" name="Button" value="Review Submitted Invoices" onClick="Javascript:frames.admin.location='corpview.asp'">
      <input type="button" name="Button2" value="Review Approved Invoices" onClick="Javascript:frames.admin.location='accview.asp'">
    </td>
  </tr>
</table>
<IFRAME name="admin" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>
</body>
</html>