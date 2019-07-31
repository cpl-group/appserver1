<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
invoice_comment=Request.form("invoicecomment")
cname = Request.form("invCname")
ctelephone = Request.form("invCtelephone")
cemail 	= Request.form("invCemail")
invoicedate=Request.form("invoice_date")
job=Request.form("job")
invoice_amt = trim(request("tot_amt"))
invoice_amt = cdbl(invoice_amt)
'tot_amt = trim(request("tot_amt"))
if (isempty(invoice_amt)) or (not isnumeric(invoice_amt)) then
	%><script>
		alert("Please enter a valid number for the total amount.  You entered '<%=invoice_amt%>'  .");
		window.history.back();		
	</script><%
	response.end
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")
strsql = "Update invoice_submission Set invoice_comment='"& invoice_comment &"', invoice_amt=" & invoice_amt &", cname = '"& cname &"', ctelephone = '"& ctelephone &"', cemail ='"& cemail &"' where ( jobno='"& job &"') and invoice_date ='" & invoicedate & "'"

cnn1.execute strsql

set cnn1=nothing
%><script>document.location = "<%=request("source")%>.asp"</script>