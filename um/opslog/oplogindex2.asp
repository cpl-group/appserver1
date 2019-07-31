<html>
<head>
<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function searchjob(typebox, searchitem) {
if (document.forms[0].comp.checked) {
var temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=1"
} else {
var temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=0"
}
document.frames.oplog.location=temp

}
function report(spec, job) {
document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
//alert(typebox)
//alert(job)
	var temp
	if(typebox =="[entry id]"){
		temp="timesheetsearch.asp?job="+job
	}else{
		temp="null.htm"
	}
	document.frames.oplog.location=temp
}

</script>
</head>
<%@Language="VBScript"%>

<%
msg = Request.querystring("msg")

			if isempty(msg) then
				msg="Please enter search and click the FIND button to begin"
			end if
			
%>
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post">
  <table width="100%" border="0" align="center">
    <tr> 
      <td align="left" height="36" valign="middle"> 
        <div align="left"> 
          <div align="center">
            <div align="left"></div>
          </div>
        </div>
        <font face="Arial, Helvetica, sans-serif">Search for Job by 
        <select name="typebox" size="1">
          <option value="[entry id]" selected><font face="Arial, Helvetica, sans-serif">Job 
          Number</font></option>
          <option value="customer"><font face="Arial, Helvetica, sans-serif">Customer 
          ID</font></option>
          <option value="manager"><font face="Arial, Helvetica, sans-serif">Project 
          Manager ID</font></option>
          <option value="[current status]"><font face="Arial, Helvetica, sans-serif">Status</font></option>
          <option value="[description]"><font face="Arial, Helvetica, sans-serif">Description</font></option>
        </select>
        : 
        <input type="text" name="findvar" size="50" maxlength="50">
        <input type="checkbox" name="comp" value="1">
        <font size="2"><i>show completed/cancelled </i></font> 
        <input type="button" name="Submit" value="Find" onClick="searchjob(typebox.value,findvar.value)">
        </font></td>
    </tr>
    <tr>
      <td align="center"><%=msg%></td>
    </tr>
    <tr> 
      <td align="center" height="12"> 
        <div align="left">
<table width="100%" border="0">
            <tr>
              <td width="54%">
                <input type="hidden" name="report1" value="opslogwip.asp">
                <input type="hidden" name="report2" value="opslogopenrfp.asp">
                <input type="button" name="button" value="Genergy WIP" onClick="report(report1.value)">
                <input type="button" name="button2" value="Genergy Open RFP" onClick="report(report2.value)">
                <input type="button" name="time" value="Invoice" onClick="timesheetjob(typebox.value, findvar.value)">
              </td>
              <td width="46%"> 
                <div align="right">
                  <input type="button" name="print" value="Print Current View" onClick="javascript:document.frames.oplog.focus();document.frames.oplog.print()">
                </div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  <font face="Arial, Helvetica, sans-serif"> </font> 
</form>
<IFRAME name="oplog" width="100%" height="70%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>