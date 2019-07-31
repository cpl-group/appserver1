<html>
<head>
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
	var temp
	if (typebox=="[entry id]" && searchitem != "") {
		temp="opslogview.asp?job=" + searchitem;
		document.frames.oplog.location=temp;
	} else {
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
			alert("Please provide a search criteria")
		}
	}
}
function report(spec, job) {
document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
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
Ay(5) = "[Entry Type]"
By(0) = "Job Number"
By(1) = "Customer ID"
By(2) = "Manager ID"
By(3) = "Status"
By(4) = "Description"
By(5) = "Job Type"

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
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Job 
        Log - Job Administration</font></b></font></div>
    </td>
  </tr>
</table>
<form name="form1" method="post">
  <table width="100%" border="0" align="center">
    <tr> 
      <td align="left" height="36"> 
        <div align="left"> 
          <div align="center"> 
            <div align="left"></div>
          </div>
        </div><font face="Arial, Helvetica, sans-serif">
        Search for Job by 
        <select name="typebox" size="1" onChange=fillup(this.value)>
          <%
		  for i=0 to 5
		    if Ay(i) = typebox then
		  %>
          <option value="<%=typebox%>" selected><font face="Arial, Helvetica, sans-serif"><%=By(i)%> 
          </font></option>
          <%
		    else
		  %>
          <option value="<%=Ay(i)%>" ><font face="Arial, Helvetica, sans-serif"><%=By(i)%></font> 
          <%
		    end if
		  Next
		  %>
        </select>
        : 
        <%
		if (typebox = "customer") then
		%>
        <select name="findvar">
          <%
			sqlstr = "select distinct customerid, companyname from customers order by companyname"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
				do until rst1.eof	
		%>
          <option value="<%=rst1("customerid")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("companyname")%></font></option>
          <%
					rst1.movenext
					loop
					end if
				%>
        </select>
		
        <input type="text" name="findvar2" size="25" maxlength="50">
    <tr> 
      <td align="center"> 
        <div align="left"> 
          <input type="button" name="Submit3" value="Find" onClick="searchjob(typebox.value,findvar.value, comp.value, findvar2.value)">
          <input type="checkbox" name="comp" value="1">
          <font size="2"><i>show completed/cancelled </i></font> </div>
      </td>
    </tr>
    <%
		else
			if (typebox = "manager") then
			%>
    <select name="findvar">
      <%
				sqlstr = "select distinct [job log]. manager , employees.[first name], employees.[last name] from employees join [job log] on [job log].manager = employees.[id] where employees.active=1 order by employees.[last name] "
				rst1.Open sqlstr, cnn1, 0, 1, 1
				if not rst1.eof then
					do until rst1.eof
		    %>
      <option value="<%=rst1("manager")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("first name")%>&nbsp<%=rst1("last name")%></font></option>
      <%
					rst1.movenext
					loop
				end if
				%>
    </select>
    <input type="button" name="Submit2" value="Find" onClick='searchjob(typebox.value,findvar.value, comp.value, "")'>
	<input type="checkbox" name="comp" value="1">
    <font size="2"><i>show completed/cancelled </i></font> 
    <%
	else
			if (typebox = "[current status]") then
			%>
    <select name="findvar">
      <%
				sqlstr = "select status from status "
				rst1.Open sqlstr, cnn1, 0, 1, 1
				if not rst1.eof then
					do until rst1.eof
		    %>
      <option value="<%=rst1("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("status")%></font></option>
      <%
					rst1.movenext
					loop
				end if
				%>
    </select>
    <input type="button" name="Submit2" value="Find" onClick='searchjob(typebox.value,findvar.value, comp.value, "")'>
	<input type="checkbox" name="comp" value="1">
    <font size="2"><i>show completed/cancelled </i></font> 
    <%
			else
		%>
    <input type="text" name="findvar" size="25" maxlength="50">
    <input type="button" name="Submit" value="Find" onClick='searchjob(typebox.value,findvar.value, comp.value,"")'>
    <input type="checkbox" name="comp" value="1">
    <font size="2"><i>show completed/cancelled </i></font> 
    <%
		    end if
		end if
		%></font></td></tr>
    <tr> 
      <td align="center"><b><font face="Arial, Helvetica, sans-serif"><%=msg%></font></b></td>
    </tr>
    <tr> 
      <td align="center">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td align="center"> 
        <div align="left"> 
          <input type="hidden" name="report1" value="opslogwip.asp">
          <input type="hidden" name="report2" value="opslogopenrfp.asp">
          <input type="button" name="button" value="Genergy WIP" onclick="report(report1.value)">
          <input type="hidden" name="button2" value="Genergy Open RFP" onclick="report(report2.value)">
          <input type="button" name="time" value="Invoice" onclick="timesheetjob(typebox.value, findvar.value)">
          <input type="hidden" name="nj" value="newjob.asp">
		  <input type="hidden" name="nc" value="newcustomer.asp">
          <%if Session("opslog") > 3 then %>
          <input type="button" name="job" value="New Job"  onclick="report(nj.value)">
		  <input type="button" name="customer" value="New Customer"  onclick="report(nc.value)">
          <%end if%>
		  <%end if%>
        </div>
      </td>
      <td> 
        <div align="right"> 
          <input type="button" name="print" value="Print Current View" onClick="javascript:document.frames.oplog.focus();document.frames.oplog.print()">
        </div>
      </td>
    </tr>
  </table>
  <font face="Arial, Helvetica, sans-serif"> </font> 
</form>
<IFRAME name="oplog" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>