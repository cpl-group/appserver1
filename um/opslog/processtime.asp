<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("login")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if
name=Request.Querystring("name")
start=Request.Querystring("start")	
end1=Request.Querystring("end1") 	
		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

	sqlstr = "exec sp_timesubmitted_email  '" & name & "','" &  start & "','" & end1 & "'"
		msg = "TimeSheet has been submitted - all parties are being notified via email"
	

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

