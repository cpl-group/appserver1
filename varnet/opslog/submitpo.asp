<%@Language="VBScript"%>
<%
poid = Request("poid")


Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

sqlstr = "update po set submitted=1 where id=" & poid

cnn1.execute sqlstr
Response.Write "<html><head><h2><br><center><b>" & vbCrLf
Response.Write "PO Submitted"
Response.Write "</head></html>" & vbCrLf 
%> 
