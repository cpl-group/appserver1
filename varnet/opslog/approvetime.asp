<%@Language="VBScript"%>
<%
		if isempty(Session("login")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if
name1=Request.Querystring("userapp")

		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
	
	sqlstr = "exec sp_corpaccepttime_email  '" & name1 & "'"

		msg = "TimeSheet has been approved - all parties are being notified via email"
	

cnn1.Execute sqlstr 
%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
      <div align="center"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=msg%></font></b></i></div>
    </td>
  </tr>
</table>
<div align="center"><i><b></b></i></div>

