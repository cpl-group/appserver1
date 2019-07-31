<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
poid=request("id1")

%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#eeeeee" text="#000000">
<form name="form1" method="post" action="processporeject.asp">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#dddddd">
  <td colspan="2"><b>Withdraw RF #<%=request("ponum")%> &nbsp;(<%=request("podate")%>)</b></td>
</tr>
<tr valign="top"> 
  <td width="30%">Send notice to:</td>
  <td> 
  <input type="hidden" name="ponum" value="<%=request("ponum")%>">
  <input type="hidden" name="podate" value="<%=request("podate")%>">
  <input type="hidden" name="poid" value="<%=request("id1")%>">
  <input type="hidden" name="status" value="'Withdrawl'">
  <input type="hidden" name="jid" value="<%=request("jid")%>">
  <input type="hidden" name="caller" value="<%=request("caller")%>">
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

  
%>

     
 <%Set rst2 = Server.CreateObject("ADODB.recordset")
    sqlstr = "select e.[first name] +' '+e.[last name]  as name, substring(e.username,7,20) as user1 from employees e join po p on substring(e.username,7,20)=p.requistioner  where substring(e.username,7,20)=p.requistioner and p.id="&poid&""
    
      rst2.Open sqlstr, cnn1, 0, 1, 1
    %>
    <%=rst2("name")%>
    <input type="hidden" name="user" value=<%=rst2("user1")%>>
    
      <%	
        rst2.close
        
        set cnn1=nothing
      %>

  </td>
</tr>
<tr valign="top"> 
  <td>Reason for withdrawal:</td>
  <td><textarea name="message" cols="20" rows="5"></textarea></td>
</tr>
<tr valign="top"> 
  <td>&nbsp;</td>
  <td>
  <input type="submit" name="Submit" value="Send">
  <input type="button" name="Submit2" value="Cancel" onclick="javascript:history.back()">
  </td>
</tr>
</table>
</form>
</body>
</html>
