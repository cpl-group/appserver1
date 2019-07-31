<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
user1=Request.Form("user")
message=request.form("message")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

strsql="sp_reject_email '" & user1 & "','" & message & "'"

cnn1.execute strsql

set cnn1=nothing


%>
<body bgcolor="#FFFFFF">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
	  
        <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>TIMESHEET
          WAS REJECTED SUCCESSFULLY</b></font></div>
		
      </td>
    </tr>
  </table>
  <p> 
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.close()">
  </p>
</div>
