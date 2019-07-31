<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<title>Facility Information</title>
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
function report() {
document.frames.finfo.location="newbldg.asp"
}


</script>
</head>

<%
ReDim Ay(4)
ReDim By(4)
Ay(0) = "ownerid"
Ay(1) = "bldgid"
Ay(2) = "floor"
Ay(3) = "occupant"

By(0) = "Owner"
By(1) = "Building"
By(2) = "Floor/Space"
By(3) = "Occupant"



msg = Request.querystring("msg")
typebox = Request("typebox")
			if isempty(msg) then
				msg="Please enter search and click the FIND button to begin"
			end if
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.110;uid=genergy1;pwd=g1appg1;database=main;"

		
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td height="20"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Facility Information</font></b></font></div>
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
        <font face="Arial, Helvetica, sans-serif">Search for Facility Infomation 
        by: 
        <select name="typebox" size="1" onChange=fillup(this.value)>
          <%
		  for i=0 to 4
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
		if (typebox = "owner") then
		%>
        <select name="findvar">
          <%
			sqlstr = "select  corp_name from owners "
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
				do until rst1.eof	
		%>
          <option value="<%=rst1("id")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("corp_name")%></font></option>
          <%
					rst1.movenext
					loop
					end if
					%>
		</select>
		<input type="button" name="Submit" value="Find" onClick="searchjob(typebox.value,findvar.value)">  
        
		<%
		else	
		if (typebox = "bldgid") then
		%>
        <select name="findvar">
          <%
			sqlstr = "select address from facilityinfo where ownerid"
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
        
		  
          
		  <input type="button" name="job" value="New Building"  onclick="report()">
        
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
<IFRAME name="finfo" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>