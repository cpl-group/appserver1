<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
jobnum=Request.Form("jobnum")
podate=Request.Form("podate")
vendor=Request.Form("vendor")
jobadd=Request.Form("jobaddr")
shipadd=Request.Form("shipaddr")
req=Request.Form("req")
poid=Request.Form("id1")
descr=Request.Form("description")
samt=Request.Form("ship_amt")
tax=Request.Form("tax1")
total=Request.Form("total")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

strsql = "Update po Set podate='" & podate & "', vendor='" & vendor & "', jobaddr='" & jobadd & "', shipaddr='" & shipadd & "', ship_amt=" & samt & ",tax=" & tax & ",Requistioner='" & req & "', description='" & descr & "'where id='"& poid&"'"

cnn1.execute strsql

strsql="Select po_item.qty,po_item.unitprice,po.tax,po.ship_amt,sum(po_item.qty*po_item.unitprice) as total from po_item join po on po_item.poid=po.id where poid='" & poid&"' group by po_item.qty,po_item.unitprice,po.tax,po.ship_amt"
	rst1.Open strsql, cnn1, 0, 1, 1
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


cnn1.execute strsql

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "poview.asp?poid="& poid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
set cnn1=nothing
%>


