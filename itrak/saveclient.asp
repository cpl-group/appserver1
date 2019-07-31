<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!--#INCLUDE file="treenode_functions.asp"-->
<%
dim bldg, addr, city, state, phone, fax, zip, sqft, fl, c1, cp1, logo
bldg=Request.Form("bldgnum")
addr=Request.Form("address")
city=Request.Form("city")
state=Request.Form("state")
phone=Request.Form("phone")
fax=Request.Form("fax")
zip=Request.Form("zip")
sqft=Request.Form("sqft")
fl=Request.Form("fl")
c1=Request.Form("name1")
cp1=Request.Form("phone1")
logo=Request.Form("logourl")

dim cnn1, rst1, strsql, maxid, tmpMoveFrame
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")


'strsql = "INSERT nodes (label, clientid) values ('"&Clients&"') "
strsql = "insert clients (corp_name,address,city,state,zip,contact,contactphone,logo) values ('" & bldg& "', '" & addr & "', '" & city & "', '" & state & "', '" & zip & "','" & c1 & "', '" & cp1 & "', '" & logo & "')"
'response.write strsql
'response.end
cnn1.execute strsql


strsql = "select max(id) as id from clients"
rst1.Open strsql, cnn1, 0, 1, 1
maxid = rst1("id")
dim bnid

bnid = addBuildingNode(maxid, 0, maxid, "Select Region", "region", 1, 1, "", cnn1)
	addBuildingNode maxid, bnid, maxid, "Select City", "city", 1, 0, "", cnn1

'response.write sqlstr

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "manageaccounts.asp?cid="& maxid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

rst1.close
set cnn1=nothing

%>