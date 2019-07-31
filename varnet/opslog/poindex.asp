<html>
<head>
<title>PO</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function fillup(typebox){
	document.location="poindex.asp?typebox=" + typebox
}
function searchjob(typebox, searchitem) {
//alert(typebox)
//alert(searchitem)
	var temp
	if (typebox=="ponum"){
		temp="poview.asp?po=" + searchitem
	} 
	else {
		temp="posearch.asp?select=" + typebox + "&findvar="+searchitem
	}
	document.frames.oplog.location=temp
    
}
function report(spec, job) {
document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
//alert(typebox)
//alert(document.forms[0].findvar.value)
	var temp
	if(typebox =="Job Number"){
	    if(job == ""){
		    alert("Please Enter Job Number")
		}else{
			temp="timesheetsearch.asp?job="+job
			document.frames.oplog.location=temp
		}
	}else{
		temp="null.htm"
		document.frames.oplog.location=temp
	}
	
}

</script>
</head>
<%@Language="VBScript"%>
<%
ReDim Ay(5)
ReDim By(5)
Ay(0) = "jobnum"
Ay(1) = "vendor"
Ay(2) = "ponum"
Ay(3) = "description"
Ay(4) = "requistioner"
By(0) = "Job Number"
By(1) = "Vendor"
By(2) = "PO Number"
By(3) = "Description"
By(4) = "Requistioner"


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
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Purchase 
        Orders</font></b></font></div>
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
        </div>
        <font face="Arial, Helvetica, sans-serif">Search for Purchase Order by: 
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
       <%
		if (typebox = "requistioner") then
		%>
        <select name="findvar">
          <%
			sqlstr = "select [first name]+' '+ [last name] as name from employees where active=1"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
				do until rst1.eof	
		%>
          <option value="<%=rst1("name")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("name")%></font></option>
          <%
					rst1.movenext
					loop
					end if
					%>
		</select>
		<input type="button" name="Submit" value="Find" onClick="searchjob(typebox.value,findvar.value)">  
        
		<%
		else	
		if (typebox = "vendor") then
		%>
        <select name="findvar">
          <%
			sqlstr = "select distinct vendor from po"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
				do until rst1.eof	
		%>
          <option value="<%=rst1("vendor")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("vendor")%></font></option>
          <%
					rst1.movenext
					loop
					end if
					%>
		</select>
        
        
        <input type="button" name="Submit3" value="Find" onClick="searchjob(typebox.value,findvar.value)">
        <%else
		%>
		
        <input type="text" name="findvar" size="50" maxlength="50">
        <input type="button" name="Submit3" value="Find" onClick="searchjob(typebox.value,findvar.value)">
		<%
		    
		end if
		end if
		%>
        </font></td>
    </tr>
    <tr>
      <td align="center"><%=msg%></td>
    </tr>
    <tr> 
      <td align="left">
        
		  
          <input type="hidden" name="np" value="newpo.asp">
		  <input type="button" name="job" value="New PO"  onclick="report(np.value)">
        
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