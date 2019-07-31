<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(getKeyValue("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if getKeyValue("admin")<>"" then 
				if getKeyValue("admin") < 5 then 
					setKeyValue "fMessage", "Sorry, the module you attempted to access is unavailable to you."
	
					Response.Redirect "../main.asp"
				end if	
			end if
		end if
user2=Request.Querystring("user1")

%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
	Dim cnn1
	Set cnn1 = Server.CreateObject("ADODB.connection")
	Set rst1 = Server.CreateObject("ADODB.recordset")
	
cnn1.Open getConnect(0,0,"intranet")

		
	%>

      
       
  <%Set rst1 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select e.[first name] +' '+e.[last name] as name, substring(e.username,7,20) as user1 from employees e  where e.username='"&user2&"' "
			
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
			%>
			
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="processtimereject.asp">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td bgcolor="#666699"><span class="standardheader">Reject Time Sheet Of <%=rst1("name")%></span></td>
</tr>
</table>
<br>
<table width="90%" border="0" cellpadding="3" cellspacing="0" align="center">
<tr valign="top"> 
  <td width="36%">Send notice to:</td>
  <td width="64%"> 
  <%=rst1("name")%>
  <input type="hidden" name="user" value="<%=rst1("user1")%>"></font>
  <%
  end if
  'response.write sqlstr
  'response.end
    rst1.close
    
    set cnn1=nothing
  %>
  </td>
</tr>
<tr valign="top"> 
  <td width="36%">Reason for rejection:</td>
  <td width="64%"><textarea name="message" cols="20" rows="5"></textarea></td>
</tr>
<tr valign="top"> 
  <td width="36%">&nbsp;</td>
  <td width="64%">
  <input type="submit" name="Submit" value="Send">
  <input type="button" name="Submit2" value="Cancel" onclick="window.close();">
  </td>
</tr>
</table>
</form>
</body>
</html>
