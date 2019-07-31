<%@Language="VBScript"%>
<%
		if isempty(Session("login")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if
		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

	
	if Request.Querystring("poaction") = "submit" then
	
		sqlstr = "exec sp_po_submitted " & Request.Querystring("poid") & ", [" & Session("login") & "]"
		msg = "PO has been Submitted - all parties are being notified via email"
	else
		sqlstr = "delete po where id=" & Request.Querystring("poid") & " delete po_item where poid=" & Request.Querystring("poid") & ""
		msg = "PO has been Deleted"
	end if
	

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

