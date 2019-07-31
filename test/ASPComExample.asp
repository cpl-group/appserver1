<%@ LANGUAGE = VBScript %>
<% 
'
'  ASP/COM Client Integration Example
'
'  This file must be used in conjunction with ASPCOMExample.htm.
'  Please copy both example files to a directory on your web server
'  the has script execute access.
'
'  You must install the Payflow Pro COM Client before you can
'  successfully run this example.
'
%>
<html>

<head>
<title>ASPComExample.asp File</title>
</head>

<body BGCOLOR="#FFFFFF">
<font FACE="ARIAL,HELVETICA"><%

' create the PNCOMClient component
Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")


'build the parameter list, such that we have a sale transaction and
'a credit card tender.
parmList = "TRXTYPE=S&TENDER=C&ZIP=12345&COMMENT1=ASP/COM Test Transaction"

'set the account form the html form
parmList = parmList + "&ACCT=" + request.form("cardNum") 

'set the password from the html form
parmList = parmList + "&PWD=" + request.form("password")

'set the user from the HTML form
parmList = parmList + "&USER=" + request.form("user")

'set the vendor from the HTML form
parmList = parmList + "&VENDOR=" + request.form("vendor")

'set the partner from the HTML form
parmList = parmList + "&PARTNER=" + request.form("partner")

'set the expiration date form the HTML form
parmList = parmList + "&EXPDATE=" + request.form("cardExp")

'set the amount from the HTML form
parmList = parmList + "&AMT=" + request.form("amount")


Ctx1 = client.CreateContext("test-payflow.verisign.com", 443, 30, "", 0, "", "")
curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
client.DestroyContext (Ctx1)

' handle the response
done = 0

'loop until we're done processing the entire string
Do while Len(curString) <> 0

	'get the next name value pair

	if InStr(curString,"&") Then
		varString = Left(curString, InStr(curString , "&" ) -1)
	else
		varString = curString
	end if
	
	Response.Write "<br>"
	
	'get the name part of the name/value pair
	name = Left(varString, InStr(varString, "=" ) -1)
	
	'get the value out of the name/value pair
	value = Right(varString, Len(varString) - (Len(name)+1))
	
	'write out the name/value pair in "name = value" format
	response.write name
	response.write " = "
	Response.Write value
	
	Response.Write "<br>"
	
	'skip over the &
	if Len(curString) <> Len(varString) Then 
		curString = Right(curString, Len(curString) - (Len(varString)+1))
	else
		curString = ""
	end if
Loop

%>
</font>
</body>
</html>
