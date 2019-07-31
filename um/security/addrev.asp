<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'Response.Redirect "http://www.genergyonline.com"
		
		end if		
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Revision Management</title>
</head>
<script>
function changeaction(){
	if (document.forms[0].del.checked == true){
		document.forms[0].action.value = 0
	} else {
		document.forms[0].action.value = 2
	}
	}
</script>
<body bgcolor="#FFFFFF">
<%

if request("action") = "2" then 

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
strsql = "select * from revtrack where id ='"& Request("id") & "'"
rst1.Open strsql, cnn1, 0, 1, 1

	if not rst1.EOF then 
	
%>
	<form name="form1" method="post" action="updaterev.asp">
	<div align="center">
	<p><font face="Arial, Helvetica, sans-serif"><b>ADD/UPDATE REVISION</b></font></p>
	<table width="45%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
	  <tr> 
		<td width="38%" bgcolor="#66CCFF" > 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">Rev 
			Date</font></div>
		</td>
		<td width="62%" > 
		  <div align="right"> 
			<input type="text" name="revdate" size="15" value="<%=rst1("revDate")%>">
		  </div>
		</td>
	  </tr>
	  <tr> 
		<td width="38%" bgcolor="#66CCFF"> 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">GUID</font></div>
		</td>
		<td width="62%"> 
		  <div align="right"> 
			<input type="text" name="guid" size="15" value="<%=rst1("guid")%>">
		  </div>
		</td>
	  </tr>
	  <tr> 
		<td width="38%" bgcolor="#66CCFF"> 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">PID 
			</font></div>
		</td>
		<td width="62%"> 
		  <div align="right"> 
			<input type="text" name="pid" size="15" maxlength="25" value="<%=rst1("pid")%>">
		  </div>
		</td>
	  </tr>
	  <tr valign="top"> 
		<td width="38%" bgcolor="#66CCFF"> 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">Revision 
			Description</font></div>
		</td>
		<td width="62%"> 
		  <div align="right"> 
			<textarea name="revdesc" cols="25" rows="5" wrap="PHYSICAL"><%=rst1("revdescriptor")%></textarea>
		  </div>
		</td>
	  </tr>
	  <tr> 
		<td width="38%"> 
		  <div align="left"> <font face="Arial, Helvetica, sans-serif"> 
            <input type="hidden" name="action" value="2">
            <input type="hidden" name="id" value="<%=rst1("id")%>">
            <input type="checkbox" name="del" value="0" onclick="changeaction()">
            <font size="2">Delete</font></font></div>
		</td>
		<td width="62%"> 
          <div align="right"><font face="Arial, Helvetica, sans-serif"> 
            <input type="submit" name="Submit2" value="Update">
            </font></div>
        </td>
	  </tr>
	</table>
	</div>
	</form>
<% 
	rst1.close
	Set cnn1 = nothing
	end if
else %>

	<form name="form1" method="post" action="updaterev.asp">
	<div align="center">
	<p><font face="Arial, Helvetica, sans-serif"><b>ADD/UPDATE REVISION</b></font></p>
	<table width="45%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
	  <tr> 
		<td width="38%" bgcolor="#66CCFF" > 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">Rev 
			Date</font></div>
		</td>
		<td width="62%" > 
		  <div align="right"> 
			<input type="text" name="revdate" size="15" value="<%=Date()%>">
		  </div>
		</td>
	  </tr>
	  <tr> 
		<td width="38%" bgcolor="#66CCFF"> 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">GUID</font></div>
		</td>
		<td width="62%"> 
		  <div align="right"> 
			<input type="text" name="guid" size="15" value="<%=Session("login")%>">
		  </div>
		</td>
	  </tr>
	  <tr> 
		<td width="38%" bgcolor="#66CCFF"> 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">PID 
			</font></div>
		</td>
		<td width="62%"> 
		  <div align="right"> 
			<input type="text" name="pid" size="15" maxlength="25" value="<%=Request("pid")%>">
		  </div>
		</td>
	  </tr>
	  <tr valign="top"> 
		<td width="38%" bgcolor="#66CCFF"> 
		  <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">Revision 
			Description</font></div>
		</td>
		<td width="62%"> 
		  <div align="right"> 
			<textarea name="revdesc" cols="25" rows="5" wrap="PHYSICAL">Enter Revision Description Here</textarea>
		  </div>
		</td>
	  </tr>
	  <tr> 
		<td width="38%"> 
		  <div align="left"> <font face="Arial, Helvetica, sans-serif"> 
			<input type="hidden" name="action" value="1">
            </font></div>
		</td>
		<td width="62%"> 
          <div align="right"><font face="Arial, Helvetica, sans-serif"> 
            <input type="submit" name="Submit" value="Submit">
            </font></div>
        </td>
	  </tr>
	</table>
	</div>
	</form>
<%end if %>