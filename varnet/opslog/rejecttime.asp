<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("admin") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
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
	
cnn1.Open application("cnnstr_main")

		
	%>

      
       
  <%Set rst1 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select e.[first name] +' '+e.[last name] as name, substring(e.username,7,20) as user1 from employees e  where e.username='"&user2&"' "
			
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
			%>
			
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="processtimereject.asp">
  <table width="90%" border="0" align="center">
  <tr>
    <td bgcolor="#3399CC"><b><font face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">REJECT 
      TIMESHEET	 <%=rst1("name")%></font></font></b></td>
  </tr>
</table>
<table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr valign="top" bgcolor="#999999"> 
     
    <td width="64%">&nbsp;</td>
  </tr>
  <tr valign="top"> 
 
    <td width="36%"><font face="Arial, Helvetica, sans-serif">Send Notice to :</font></td>
    <td width="64%"> 

			<font face="Arial, Helvetica, sans-serif"><%=rst1("name")%>
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
    <td width="36%"><font face="Arial, Helvetica, sans-serif">Reason for Rejection:</font></td>
    <td width="64%"> 
      <div align="right"> 
        <textarea name="message" cols="20" rows="5"></textarea>
      </div>
    </td>
  </tr>
  <tr valign="top" bgcolor="#3399CC"> 
    <td width="36%"><font face="Arial, Helvetica, sans-serif"> 
      <input type="submit" name="Submit" value="Send" >
      <input type="button" name="Submit2" value="Cancel" onclick="javascript:window.close()">
      </font></td>
    <td width="64%">&nbsp;</td>
  </tr>
</table></form>
</body>
</html>
