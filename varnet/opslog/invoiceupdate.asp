<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
job=Request.form("job")
invoice=Request.form("invoice")
invtype=request.Form("invtype")
user="ghnet\"&Session("login")

if invtype=1 then 
	tm=1
	cont=0
else
	tm=0
	cont=1
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")
if not isempty(invoice) then
strsql = "Update invoice_submission Set submitted=1, invoice_comment='"& invoice &"', invoice_date=convert(char, getdate(), 101),tmtype='"& tm &"',contract='"& cont &"' where ( jobno='"& job &"') and invoice_submission.date > (select last_invoice from [job log] where [entry id]='"&job&"')"

cnn1.execute strsql

end if
strsql = "Update [job log] Set last_invoice= convert(char, getdate(), 101) where ( [entry id]='"& job &"')"



cnn1.execute strsql
sql="sp_invoice_email '"& job &"'"

cnn1.Execute(sql)
set cnn1=nothing

Response.Write "<html><head><h2><br><center><b>" & vbCrLf
Response.Write "Invoice Submitted"
Response.Write "</head></html>" & vbCrLf 
%>