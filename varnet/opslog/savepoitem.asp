<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%

poid = Request.Form("poid")
qty=Request.Form("qty")
unit=Request.Form("unit")
invnum=Request.Form("invnum")
unitprice=Request.Form("unitprice")
description=Request.Form("description")
tax=Request.Form("tax")



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")

strsql = "insert po_item (qty, unit, invnum, unitprice, description, poid) values (" & qty & ",'" & unit& "', '" & invnum & "', " &unitprice& ", '" & description & "', " & POID & ")"
'response.write strsql
'response.end
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

tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "poview.asp?poid="& poid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>