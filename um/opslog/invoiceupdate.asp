<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
job=Request.form("job")
tot_amt = trim(request("tot_amt"))
if (isempty(tot_amt)) or (not isnumeric(tot_amt)) then
	%><script>
		alert("Please enter a valid number for the total amount.  You entered '<%=tot_amt%>'  .");
		window.history.back();		
	</script><%
	response.end
end if
invoice=Request.form("invoice")
invtype=request.Form("invtype")
customcontact = request("desigContact")
if customcontact = "1" then 
	cName	= request("invCname")
	cTelephone = request("invCtele")
	cEmail = request("invCemail")
	user="ghnet\"&Session("login")
else 
	cName	= request("invoicecontact")
	cTelephone = "see customer record"
	cEmail = "see customer record"
end if
if invtype=1 then 
	tm=1
	cont=0
else
	tm=0
	cont=1
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

strsql="sp_invoice '"& job &"', " & tot_amt & ",'"&getxmlusername &"'"
cnn1.execute strsql




if not isempty(invoice) then
  strsql = "Update invoice_submission Set invoice_comment='"& invoice &"', tmtype='"& tm &"',contract='"& cont &"',customcontact='"& customcontact &"', cName='"& cName &"',cTelephone='"& cTelephone&"',cEmail='"& cEmail &"' where jobno='"& job &"' and submitted=1 and flag=0 and invoice_comment='no comment'"
end if

cnn1.execute strsql


sql="sp_invoice_email '"& job &"'"
cnn1.Execute(sql)

set cnn1=nothing

Response.Write "<html><head><h2><br><center><b>" & vbCrLf
Response.Write "Invoice Submitted"
Response.Write "</head></html>" & vbCrLf 
%>