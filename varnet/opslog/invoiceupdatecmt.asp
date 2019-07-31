<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
invoice_comment=Request.form("invoicecomment")
invoicedate=Request.form("day")
job=Request.form("job")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")
strsql = "Update invoice_submission Set invoice_comment='"& invoice_comment &"' where ( jobno='"& job &"') and invoice_date ='" & invoicedate & "'"  



cnn1.execute strsql

set cnn1=nothing

Response.redirect "corpinvoice.asp"%>