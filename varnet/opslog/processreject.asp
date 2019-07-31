
<body bgcolor="#FFFFFF">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
        <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>INVOICE 
          WAS REJECTED SUCCESSFULLY</b></font></div>
      </td>
    </tr>
  </table>
  <p> 
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.close()">
  </p>
</div>
<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
job=Request.form("job")
d=Request.form("lastinvdate")
user=request.form("user")
message=request.form("message")
'response.write d
'response.end
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
strsql = "Update invoice_submission Set submitted=0 where (jobno='"& job &"' and invoice_date='"&d&"')"
cnn1.execute strsql
strsql="sp_invoice_reject " & job & ",'" &  d & "','" & user & "','" & message & "'"

cnn1.execute strsql

set cnn1=nothing


%>