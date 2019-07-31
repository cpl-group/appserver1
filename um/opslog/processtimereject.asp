<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
user1=Request.Form("user")
message=request.form("message")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

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
