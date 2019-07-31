<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
jobnum=secureRequest("jobnum")
podate=secureRequest("podate")
vendor=secureRequest("vendor")
jobadd=secureRequest("jobaddr")
shipadd=secureRequest("shipaddr")
req=secureRequest("req")
poid=Request("id1")
descr=secureRequest("description")
samt=secureRequest("ship_amt")
tax=secureRequest("tax1")
total=secureRequest("total")
vid=secureRequest("vid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

strsql = "Update po Set podate='" & podate & "', vendor='" & vendor & "', vid='" & vid & "', jobaddr='" & jobadd & "', shipaddr='" & shipadd & "', ship_amt=0,tax="&tax&",Requistioner='" & req & "', submittedby = '"&getXmlUserName()&"', description='" & descr & "' where id='"& poid&"'"
cnn1.execute strsql

strsql="Select po_item.qty,po_item.unitprice,po.tax,po.ship_amt,sum(po_item.qty*po_item.unitprice) as total from po_item join po on po_item.poid=po.id where poid='" & poid&"' group by po_item.qty,po_item.unitprice,po.tax,po.ship_amt"
	rst1.Open strsql, cnn1, 0, 1, 1
if not rst1.eof then
	ship_amt=rst1("ship_amt")
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


