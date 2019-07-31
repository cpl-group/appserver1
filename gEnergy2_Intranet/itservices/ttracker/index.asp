<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim cnn, rst, strsql,defaultpage

set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Job Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>
<body bgcolor="#eeeeee">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#6699cc"> 
    <td ><span class="standardheader">Trouble Tickets</span></td>
  </tr>
  <tr> 
	<td><a href="troublesearch.asp?searchbox=true" target="app" onClick="javascript:document.all.Function.innerHTML = 'Search Tickets';document.all.count.innerHTML='';">Search Tickets</a> | 
		<a href="troublesearch.asp?status=0&listlength=7" target="app" onclick="javascript:document.all.Function.innerHTML = 'All Open Tickets';">All Open Tickets</a> | 
		<a href="troublesearch.asp?status=1" target="app" onclick="javascript:document.all.Function.innerHTML = 'All Closed Tickets';">All Closed Tickets</a> | 
		<a href="ticket.asp?mode=enterid" target="app" onclick="javascript:document.all.Function.innerHTML = 'Open Ticket by Ticket ID';">Open Ticket by Ticket ID</a> | 
      <%if checkgroup("IT Services,Department Supervisors") then%>
      <a href="./userstats.asp" target="app" onclick="javascript:document.all.Function.innerHTML = 'User Statistics : All Tickets except tickets classified as running or ongoing tickets';javascript:document.all.count.innerHTML = '';">User 
      Statistics</a> | 
      <%end if%>
      <a href="ticket.asp?mode=new" target="app" onclick="javascript:document.all.Function.innerHTML = 'New Ticket';javascript:document.all.count.innerHTML = '';">New 
      Ticket </a></td>
  </tr>
  <tr> 
    <td width="29%" height="5" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><div align="center"><font size="2"><b><span id="Function"></span></b> </font></div></td>
  </tr>
  <tr>
    <td width="29%" ><div align="center"><font size="2"><b><span id="count"></span></b> </font></div></td>
  </tr>
</table>
<br> 
<%       
	if allowgroups("IT Services,Department Supervisors") then
		defaultpage = "troublesearch.asp?searchbox=true&action=Search&internalops=True&searchstring="&getXmlUserName()
	else
		defaultpage = "troublesearch.asp?status=0"
	end if
%>		
<iframe src="<%=defaultpage%>" name="app" width="100%" height="500" frameborder="0" scrolling="no"></iframe>
</body>
</html>
