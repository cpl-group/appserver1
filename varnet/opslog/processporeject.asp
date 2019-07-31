<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
poid=Request.form("poid")
user=Session("login")
message=request.form("message")
status=request.form("status")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

strsql="sp_po_reject " & poid & "," & status & ",'" & user & "','" & message & "'"

cnn1.execute(strsql)

set cnn1=nothing


%>
<body bgcolor="#FFFFFF">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
	  <%if status="Reject" then
	  %>
        <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>PO 
          WAS REJECTED SUCCESSFULLY</b></font></div>
		 <%
		 else
		 %>
		<div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>PO 
          WAS WITHDRAWN SUCCESSFULLY</b></font></div> 
		  <%
		  end if
		  %>
      </td>
    </tr>
  </table>
  <p> 
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.close()">
  </p>
</div>
