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
		
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="processreject.asp">
  <table width="90%" border="0" align="center">
  <tr>
    <td bgcolor="#3399CC"><b><font face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">REJECT 
      INVOICE DATED <%=Request.Querystring("d")%> FOR JOB NUMBER <%=Request.Querystring("job")%></font></font></b></td>
  </tr>
</table>
<table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr valign="top" bgcolor="#999999"> 
    <td width="36%"><font face="Arial, Helvetica, sans-serif">
      <input type="hidden" name="job" value="<%=Request.Querystring("job")%>">
	  <input type="hidden" name="lastinvdate" value="<%=Request.Querystring("d")%>">
      </font></td>
    <td width="64%">&nbsp;</td>
  </tr>
  <tr valign="top"> 
    <td width="36%"><font face="Arial, Helvetica, sans-serif">Send Notice to :</font></td>
    <td width="64%"> 
	<%
	Dim cnn1
	Set cnn1 = Server.CreateObject("ADODB.connection")
	Set rst1 = Server.CreateObject("ADODB.recordset")
	
	cnn1.Open application("cnnstr_main")

		
	%>

      <div align="right"> 
              <select name="user">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select [last name]+', '+[first name]  as name, substring(username,7,20) as user1 from employees where active=1 order by [last name]"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
                  <option value="<%=rst2("user1")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("name")%></font></option>
                  <%
					rst2.movenext
					loop
					end if
					rst2.close
					set cnn1=nothing
				%>
                </select>

      </div>
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
      <input type="submit" name="Submit" value="Send">
      <input type="button" name="Submit2" value="Cancel" onclick="javascript:window.close()">
      </font></td>
    <td width="64%">&nbsp;</td>
  </tr>
</table></form>
</body>
</html>
