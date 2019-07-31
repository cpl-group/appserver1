
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%

d=Request.Form("description")
manf=Request.Form("manf")
fixc=Request.Form("fixc")
b=Request.Form("ballast")
lqty=Request.Form("lqty")
zip=Request.Form("zip")
lwatts=Request.Form("lwatts")
lcnum=Request.Form("lcnum")
volts=Request.Form("volts")
estLL=Request.Form("estLL")
remarks=Request.Form("remarks")
blife=Request.Form("blife")
ballastqty=Request.Form("ballastqty")
cid=Request.Form("cid")
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

strsql = "insert fixture_types (description,manufacturer,fix_catalog,ballast_type,lamp_qty,lamp_watts,lamp_catalog,volts,remarks,avg_lamp_life,ballast_life,ballast_qty,client) values ('" &d & "', '" & manf & "', '" &fixc & "', '" & b& "','" &lqty & "','" &lwatts & "', '" & lcnum & "', '" & volts& "', '" & remarks& "', '" & estLL& "', '" & blife& "', '" & ballastqty& "', '" & cid& "')"
'response.write strsql
'response.end
cnn1.execute strsql


sqlstr = "select max(id) as id from fixture_types "

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1
maxid=rst1("id")


tmpMoveFrame = "location = " & Chr(34) & _
				  "fixtypesearch.asp?cid=" +cid& chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

rst1.close
set cnn1=nothing

%>