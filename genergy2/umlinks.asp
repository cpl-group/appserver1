<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
		else
			if  Session("um") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	
			
dim umgate
umgate = "http://10.0.7.23/um_gate.jsp?username=" & Trim(Session("login")) & "&userlevel=" & Session("um") & Chr(34) & "," & Chr(34) & "GenergyOne" & Chr(34) & "," & Chr(34) & "resizeable=yes,statusbar=yes,toolbar=yes,height=768,width=1024"

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
        "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>UM Version 1</title>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>
<body bgcolor="#ffffff">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Utilility Manager, Version 1</span></td>
</tr>
<tr>
  <td><br>
  <blockquote>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="<%=umgate%>" target="_new">Utility Manager</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/um/bldglist.asp">Bill Printer/Viewer</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/billsummary/billsummary_select.asp">Bill Summary</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/um/meterlist.asp">Meter List View</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/um/meternotes.asp">Meter Problems</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/um/lmsetup.asp">Meter LM Setup</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/um/portfoliolist.asp">Building TC Setup</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/um/tenantbilllist.asp">Bill Processor</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/um/validation/validation_select.asp" target="_blank" >Review/Edit+</a><br>
  <img src="/genergy2/SETUP/images/aro-rt.gif" border="0">&nbsp;<a href="/client_entry/entry.asp">Bill Entry</a><br>
  </blockquote>
  </td>
</tr>
</table>



</body>
</html>
