<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%


quantity=secureRequest("qty")
unit=secureRequest("unit")
invnum=secureRequest("invnum")
unitprice=secureRequest("unitprice")
description=secureRequest("description")
tax=secureRequest("tax")
id=secureRequest("key")
pid=secureRequest("poid")

boolDelete = Request.Form("boolDelete")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

If boolDelete = "false" Then
	strsql = "Update po_item Set qty='" & quantity &"',unit=0, invnum='" & invnum&"',unitprice="&unitprice&",description='" & description &"' where id='"& id&"'"
	cnn1.execute strsql
	
	strsql="Select po_item.qty,po_item.unitprice,po.ship_amt,sum(po_item.qty*po_item.unitprice) as total,po.tax from po_item join po on po_item.poid=po.id where poid='" & pid&"' group by po_item.qty,po_item.unitprice,po.ship_amt,po.tax"
	rst1.Open strsql, cnn1, 0, 1, 1
	if not rst1.eof then
		ship_amt=rst1("ship_amt")
		sum=0
		do until rst1.eof
			sum=sum + rst1("total")
			rst1.movenext
		loop
		sum=sum+ship_amt
		strsql="Update po set po_Total=(("&sum&")*(1+po.tax)) where (id='"& pid&"')"
	end if
	cnn1.execute strsql
'finished update branch

ElseIf boolDelete = "true" then
	strsql = "DELETE FROM po_item WHERE id = '" & id & "'"
	cnn1.execute strsql
end if 'finished delete branch
		
set cnn1=nothing
tmpMoveFrame =  "location = 'poview.asp?poid="& pid & "'" & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>