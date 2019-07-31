<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/gEnergy2/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
'if isempty(Session("name")) then
'	Response.Redirect "http://www.genergyonline.com"
'else
'	if Session("admin") < 5 then 
'		Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
'		Response.Redirect "../main.asp"
'	end if	
'end if
dim poid, cnnMain, rs, sqlstr	
poid=Request("poid")

Set cnnMain = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.recordset")
	
cnnMain.Open getConnect(0,0,"intranet")

if Request("poaction") = "reject" then
	%>
	<form name="form1" method="post" action="processporeject.asp">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Reject RF #<%=Request("ponum")%> (dated <%=Request("podate")%>)</span></td>
		</tr>
	</table>
	<br>	
	<table width="90%" border="0" cellpadding="3" cellspacing="0" align="center">	
		<tr valign="top"> 
			<td width="36%">Send notice to:</td>
			<td width="64%"> 
				<input type="hidden" name="ponum" value="<%=Request("ponum")%>">
				<input type="hidden" name="podate" value="<%=Request("podate")%>">
				<input type="hidden" name="poid" value="<%=Request("poid")%>">
				<input type="hidden" name="status" value="Reject"><%
				sqlstr = "select e.[first name] +' '+e.[last name]  as name, substring(e.username,7,20) as user1 from employees e join po p on "_
					& "substring(e.username,7,20)=p.requistioner  where substring(e.username,7,20)=p.requistioner and p.id="&poid&""
				rs.Open sqlstr, cnnMain, 0, 1, 1
				%>
				<%=rs("name")%>
				<input type="hidden" name="user" value="<%=rs("user1")%>">					
			</td>
		</tr>		
		<tr valign="top"> 
			<td width="36%">Reasons for rejection:</td>
			<td width="64%"> 
				<textarea name="message" cols="20" rows="5"></textarea>
			</td>
		</tr>		
		<tr valign="top"> 
			<td width="36%">&nbsp;</td>
			<td width="64%"> 
				<input type="submit" name="Submit" value="Send">
				<input type="button" name="Submit2" value="Cancel" onclick="javascript:window.close()">
			</td>
		</tr>
		
	</table>
	</form>
<%
elseif Request("poaction") = "approve" then
	%>
	<form name="form1" method="post" action="processpoaccept.asp?action=approve">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Approve RF #<%=Request("ponum")%> (dated <%=Request("podate")%>)</span></td>
		</tr>
	</table>
	<br>
	
	<table width="90%" border="0" cellpadding="3" cellspacing="0" align="center">
		<tr valign="top"> 
			<td width="36%">Send notice to:</td>
			<td width="64%"> 
				<input type="hidden" name="ponum" value="<%=Request("ponum")%>">
				<input type="hidden" name="podate" value="<%=Request("podate")%>">
				<input type="hidden" name="poid" value="<%=Request("poid")%>">
				<input type="hidden" name="acctponum" value="<%=request("acctponum")%>">
				<input type="hidden" name="status" value="'Reject'">
				
				<%
				sqlstr = "select e.[first name] +' '+e.[last name]  as name, substring(e.username,7,20) as user1 from employees e join po p on "_
					&"substring(e.username,7,20)=p.requistioner  where substring(e.username,7,20)=p.requistioner and p.id="&poid&""
				rs.Open sqlstr, cnnMain, 0, 1, 1
				%>
				<%=rs("name")%>
				<input type="hidden" name="user" value="<%=rs("user1")%>">					
			</td>
		</tr>
		<tr valign="top"> 
			<td width="36%">Comments</td>
			<td width="64%"> 
				<textarea name="message" cols="20" rows="5"></textarea>
			</td>
		</tr>
		<tr valign="top"> 
			<td width="36%">&nbsp;</td>
			<td width="64%"> 
				<input type="submit" name="Submit" value="Send">
				<input type="button" name="Submit2" value="Cancel" onclick="javascript:window.close()">
			</td>
		</tr>
	</table>
	</form>	
<%
elseif Request("poaction") = "accept" then
	%>
	<form name="form1" method="post" action="processpoaccept.asp?action=accept">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Accept RF #<%=Request("ponum")%> (dated <%=Request("podate")%>)</span></td>
		</tr>
	</table>
	<br>
	
	<table width="90%" border="0" cellpadding="3" cellspacing="0" align="center">
		<tr valign="top"> 
			<td width="36%">Send notice to:</td>
			<td width="64%"> 
				<input type="hidden" name="ponum" value="<%=Request("ponum")%>">
				<input type="hidden" name="podate" value="<%=Request("podate")%>">
				<input type="hidden" name="poid" value="<%=Request("poid")%>">
				<input type="hidden" name="status" value="'Reject'">
				
				<%
				sqlstr = "select e.[first name] +' '+e.[last name]  as name, substring(e.username,7,20) as user1 from employees e join po p on "_
					&"substring(e.username,7,20)=p.requistioner  where substring(e.username,7,20)=p.requistioner and p.id="&poid&""
				rs.Open sqlstr, cnnMain, 0, 1, 1
				%>
				<%=rs("name")%>
				<input type="hidden" name="user" value="<%=rs("user1")%>">					
			</td>
		</tr>
		<tr valign="top"> 
			<td width="36%">Comments</td>
			<td width="64%"> 
				<textarea name="message" cols="20" rows="5"></textarea>
			</td>
		</tr>
		<tr valign="top"> 
			<td width="36%">&nbsp;</td>
			<td width="64%"> 
				<input type="submit" name="Submit" value="Send">
				<input type="button" name="Submit2" value="Cancel" onclick="javascript:window.close()">
			</td>
		</tr>
	</table>
	</form>
<%
elseif Request("poaction") = "question" then
	%>
	<form name="form1" method="post" action="processpoq.asp">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Question RF #<%=Request("ponum")%> (dated <%=Request("podate")%>)</span></td>
		</tr>
	</table>
	<br>
	
	<table width="90%" border="0" cellpadding="3" cellspacing="0" align="center">
		<tr valign="top"> 
			<td width="36%">Send notice to:</td>
			<td width="64%"> 
				<input type="hidden" name="ponum" value="<%=Request("ponum")%>">
				<input type="hidden" name="podate" value="<%=Request("podate")%>">
				<input type="hidden" name="poid" value="<%=Request("poid")%>">
				<input type="hidden" name="status" value="'question'">	
				<%
				sqlstr = "select e.[first name] +' '+e.[last name]  as name, substring(e.username,7,20) as user1 from employees e join po p on " &_
					"substring(e.username,7,20)=p.requistioner  where substring(e.username,7,20)=p.requistioner and p.id="&poid&""
				rs.Open sqlstr, cnnMain, 0, 1, 1
				%>
				<%=rs("name")%>
				<input type="hidden" name="user" value="<%=rs("user1")%>">
			</td>
		</tr>
		<tr valign="top"> 
			<td width="36%">Comments</td>
			<td width="64%"> 
				<textarea name="message" cols="20" rows="5"></textarea>
			</td>
		</tr>
		<tr valign="top"> 
			<td width="36%">&nbsp;</td>
			<td width="64%"> 
				<input type="submit" name="Submit" value="Send" >
				<input type="button" name="Submit2" value="Cancel" onclick="javascript:window.close()">
			</td>
		</tr>
	</table>
	</form>
	
	</body>
	</html>
	<%
end if

rs.close
set rs = nothing
set cnnMain=nothing
%>