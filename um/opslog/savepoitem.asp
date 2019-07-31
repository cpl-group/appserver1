<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

poid = secureRequest("poid")
qty=secureRequest("qty")
unit=secureRequest("unit")
invnum=secureRequest("invnum")
unitprice=secureRequest("unitprice")
description=secureRequest("description")
tax=secureRequest("tax")



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

strsql = "insert po_item (qty, unit, invnum, unitprice, description, poid) values (" & qty & ",0, '" & invnum& "', " &unitprice& ", '" & description & "', " & POID & ")"
cnn1.execute strsql

strsql="Select po_item.qty,po_item.unitprice,po.ship_amt,sum(po_item.qty*po_item.unitprice) as total,po.tax from po_item join po on po_item.poid=po.id where poid='" & poid&"' group by po_item.qty,po_item.unitprice,po.ship_amt,po.tax"

rst1.Open strsql, cnn1, 0, 1, 1

if not rst1.eof then

	ship_amt=rst1("ship_amt")
	if not rst1.eof then
	sum=0
	do until rst1.eof
		sum=sum + rst1("total")
		rst1.movenext
	loop
	sum=sum+ship_amt
	strsql="Update po set po_Total=(("&sum&")*(1+po.tax)) where (id='"& poid&"')"
end if
end if

cnn1.execute strsql

set cnn1=nothing

tmpMoveFrame = "document.location='poview.asp?poid="& poid & "'" & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>