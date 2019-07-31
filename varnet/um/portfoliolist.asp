<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("UM") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewbills(ypid) {
	var temp
		temp="invoicebldg.asp?ypid=" + ypid
		document.frames.admin.location=temp
} 
function loadlist(pid) {
	var temp = "buildingtc.asp?pid=" + pid
	document.frames.admin.location=temp
}
</script>
</head>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

		
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="left"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Building 
        Trip Code Setup</font></b></font></div>
    </td>
  </tr>
</table><table width="100%" border="0">
  <tr>
    <td width="48%" height="2"> 
      <select name="pid" onchange="loadlist(this.value)">
	  <OPTGROUP label='Select Building'>

                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select Distinct portfolioid from buildings order by portfolioid"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
                  <option value="<%=rst2("portfolioid")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("portfolioid")%></font></option>
                  <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
                </select>
	  <input type="button" name="Button2" value="View Building List" onClick="loadlist(pid.value)">
    </td>
    <td width="52%" height="2"> 
      <div align="right">
        <input type="button" name="Submit" value="Print Trip Code List" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'>
      </div>
    </td>
  </tr>
</table>
<p><IFRAME name="admin" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></p>
</body>
</html>