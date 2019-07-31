<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Job Update Request</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/styles.css" type="text/css">    
</head>
<% 
Dim  process, reqlabel
process = request("process")

if process = "" then 
	process = false
else
	process = true
	sendmsg
end if 

if not process then 
	reqlabel="Sending Request..."
else
	reqlabel="Request Sent"
end if
%>
<body bgcolor="#FFFFCC">
<table width="300" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1%" align="center" bgcolor="#6699cc"><span class="standardheader">Job 
      Status Update Request </span></td>
  </tr>
  <tr>
    <td align="center" valign="middle"><font size="4"><%=reqlabel%></font></td>
  </tr>
  <tr>
    <td align="center"><% if not process then%><form name="form1" method="post" action="statusrequest.asp"><input name="pm" type="hidden" value="<%=request("pm")%>"><input name="desc" type="hidden" value="<%=request("desc")%>"><input name="jid" type="hidden" value="<%=request("jid")%>"><input name="process" type="hidden" value="true"></form><%else%>
      <input name="closewindow" type="button" id="closewindow" onclick="window.close()" value="close window">
      <%end if%></td>
  </tr>
</table>
</body>
</html>
<% if not process then %>
<script>
	document.form1.submit()
</script>
<%end if%>
<%
function sendmsg()
Dim pm, jid, sql, cnn, rs, email, subject, message, jdesc

pm = request("pm")
jid = request("jid")
jdesc = request("desc")

    set cnn = server.createobject("ADODB.connection")
    set rs 	 = server.createobject("ADODB.recordset")
    cnn.open getConnect(0,0,"dbCore")
	
		'sql = "select email from employees where [First Name]+' '+[Last Name] in ('" & pm & "','"&getKeyValue("fullname")&"')"
		sql = "select email from ADusers_GenergyUsers where fullname in ('" & pm & "','"&getKeyValue("fullname")&"')"
		
		rs.open sql, cnn 
		if not rs.eof then 
			while not rs.eof
			if email <> "" then 
			email = email & ";" & rs("email")
			else
			email = rs("email")			
			end if 
			
			rs.movenext
			wend
		end if
		rs.close
		
		subject = "A request was made for a status update for Job " & jid
		message = getKeyValue("fullname") & " has requested that you, as Project Manager of Job "& jid &" ("&jdesc&"), update the status of the job by going to the job log and entering a status note into the job" 
		sendmail email,"GSA",subject, message
end function
%>