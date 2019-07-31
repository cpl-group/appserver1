<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
poid=Request.Querystring("id1")

%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="processporeject.asp">
  <table width="90%" border="0" align="center">
  <tr>
      <td bgcolor="#3399CC"><b><font face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">WITHDRAWL</font><font color="#FFFFFF"> 
        PO DATED <%=Request.Querystring("podate")%> FOR PO NUMBER <%=Request.Querystring("ponum")%></font></font></b></td>
  </tr>
</table>
<table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr valign="top" bgcolor="#999999"> 
    <td width="36%"><font face="Arial, Helvetica, sans-serif">
      <input type="hidden" name="ponum" value="<%=Request.Querystring("ponum")%>">
	  <input type="hidden" name="podate" value="<%=Request.Querystring("podate")%>">
	  <input type="hidden" name="poid" value="<%=Request.Querystring("id1")%>">
	  <input type="hidden" name="status" value="'Withdrawl'">
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
       
   <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select e.[first name] +' '+e.[last name]  as name, substring(e.username,7,20) as user1 from employees e join po p on substring(e.username,7,20)=p.requistioner  where substring(e.username,7,20)=p.requistioner and p.id="&poid&""
			
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			%>
			<font face="Arial, Helvetica, sans-serif"><%=rst2("name")
			%>
			<input type="hidden" name="user" value=<%=rst2("user1")%>>
			
		</font>
			
				<%	
					rst2.close
					
					set cnn1=nothing
				%>
                

      </div>
    </td>
  </tr>
  <tr valign="top"> 
      <td width="36%"><font face="Arial, Helvetica, sans-serif">Reason for Withdrawl:</font></td>
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
