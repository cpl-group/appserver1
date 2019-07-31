<html>
<head>
<%@Language="VBScript"%>
<%
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewrevprof(bldg, year) {
	var temp
		temp="revenueprofile.asp?bldgnum=" + bldg +"&year=" + year
		document.frames.admin.location=temp
} 
function loadypidlist(bldg,pid) {
	var temp = "revbldglist.asp?bldg=" + bldg + "&pid="+pid
	document.location = temp
}
function bldglist(pid){
document.location="revbldglist.asp?pid=" + pid
}
</script>
</head>
<%
pid = Request.QueryString("pid")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

		
%>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#0099FF" width="77%"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Revenue 
      Profile </font></b></font></td>
    <td bgcolor="#0099FF" width="23%"> 
      <div align="right"><font color="#FFFFFF"><b> <font face="Arial, Helvetica, sans-serif"><a href="<%="index.asp?pid="&Request.Querystring("pid") %>">Cost 
        &amp; Revenue Analysis</a></font></b></font></div>
    </td>
  </tr>
</table>
<form name="form1" method="post" action="">
<table width="100%" border="0">
  <tr>
    <td width="48%" height="2"> 
	<% if isempty(Request.Querystring("bldg")) then %>
      <select name="bldg" onchange="loadypidlist(this.value,pid.value)">
	  <OPTGROUP label='Select Building'>

    	<%  Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select bldgnum, strt from buildings where portfolioid='"& pid&"' order by strt"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			
			if not rst2.eof then
					do until rst2.eof
		%>
                  <option value="<%=rst2("bldgnum")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("strt")%></font></option>
                  <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
                </select>
		  <input type="hidden" name="pid" value="<%=pid%>">	
	  <input type="button" name="Button2" value="View Available Years" onClick="loadypidlist(bldg.value,pid.value)">
	  <% else %>
	        <select name="year" >
	  <OPTGROUP label='Select Year'>

        <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select distinct billyear from billyrperiod where bldgnum= '" & Request.Querystring("bldg") & "' ORDER BY billyear DESC"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
					if rst2("billyear") > 1999 then
						 currentyear= Year(now())-1
		%>
                  <option value="<%=rst2("billyear")%>"<%if Formatnumber(rst2("billyear")) = FormatNumber(currentyear) then response.write " selected" end if%>><font face="Arial, Helvetica, sans-serif">Revenue Year : <%=rst2("billyear")%></font></option>
                  <%end if
					rst2.movenext
					loop
					end if
					rst2.close
				%>
                </select>
	  <input type="hidden" name="pid" value="<%=pid%>">	
	  <input type="hidden" name="bldg" value="<%=Request.QueryString("bldg")%>">
      <input type="button" name="Button23" value="View Profile" onClick="viewrevprof(bldg.value, year.value)">
        <font face="Arial, Helvetica, sans-serif"><i> 
        <input type="button" name="Submit2" value="Building List" onclick="Javascript:bldglist(pid.value)">
        </i></font> 
        <% end if %>
      </td>
    <td width="52%" height="2"> 
      <div align="right">
        <input type="button" name="Submit" value="Print Current Profile" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'>
      </div>
    </td>
  </tr>
</table></form>
<p>
<IFRAME name="admin" width="100%" height="100%" src=<%="revenueprofile.asp?bldgNUM=" & Request.Querystring("bldg") & "&year=" & currentyear %>  scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></p>

</body>
</html>