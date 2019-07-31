<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
poid=Request.form("poid")
user=Session("login")
message=request.form("message")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")

strsql="sp_po_question " & poid & ",'" & user & "','" & message & "'"

cnn1.execute(strsql)



set cnn1=nothing

%>
<body bgcolor="#FFFFFF">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
	  
		<div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>An email with your question has been sent out.</b></font></div> 
		
      </td>
    </tr>
  </table>
  <p> 
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.close()">
  </p>
</div>
