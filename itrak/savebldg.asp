<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!--#INCLUDE file="treenode_functions.asp"-->
<%
dim bldg, addr, city, state, phone, fax, zip, sqft, fl, c1, cp1, c2, cp2, c3, cp3, nid, cid, scroll
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
c2=Request.Form("name2")
cp2=Request.Form("phone2")
c3=Request.Form("name3")
cp3=Request.Form("phone3")
nid=Request.Form("nid")
cid=Request.Form("cid")
scroll=Request("scroll")


dim cnn1, rst1, strsql, maxid
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")



strsql = "insert facilityinfo (bldgname,address,city,state,zip,sqft,contact1,contact_phone1,contact2,contact_phone2,contact3,contact_phone3,clientid) values ('" & bldg& "', '" & addr & "', '" & city & "', '" & state & "', '" & zip & "','" &sqft & "','" & c1 & "', '" & cp1 & "', '" & c2& "', '" & cp2& "', '" & c3& "', '" & cp3& "', '"&cid&"')"
'response.write strsql
'response.end
cnn1.execute strsql


strsql = "select max(id) as id from facilityinfo "

'response.write sqlstr
rst1.Open strsql, cnn1, 0, 1, 1
maxid=rst1("id")


'##### adding to tree #####'
if trim(nid)<>"" then
	dim bnid, facInfo, maintNid, lightingNid
	bnid = addBuildingNode(cid, nid, maxid, addr, "address", 1, 1, "", cnn1)
		facInfo = addBuildingNode(cid, bnid, maxid, "Facility Info", "service", 1, 1, "updatebldg.asp?id="&maxid, cnn1)
			addBuildingNode cid, facInfo, maxid, "Primary Contact", "service", 1, 0, "contactinfo.asp?bldg="&maxid, cnn1
		addBuildingNode cid, bnid, maxid, "Floor Plans", "service", 2, 0, "", cnn1
		addBuildingNode cid, bnid, maxid, "Furniture Plans", "service", 3, 0, "", cnn1
		addBuildingNode cid, bnid, maxid, "Reflected Ceiling", "service", 4, 0, "", cnn1
		addBuildingNode cid, bnid, maxid, "Voice/Data Plan", "service", 5, 0, "", cnn1
		maintNid = addBuildingNode(cid, bnid, maxid, "Maintenance", "service", 6, 1, "", cnn1)
			lightingNid = addBuildingNode(cid, maintNid, maxid, "Lighting", "service", 1, 1, "", cnn1)
				addBuildingNode cid, lightingNid, maxid, "Fixture Management", "service", 1, 0, "fixtypesearch.asp?cid="&cid, cnn1
				addBuildingNode cid, lightingNid, maxid, "Lighting Management", "service", 2, 0, "floorsearch.asp?bldg="&maxid, cnn1
				addBuildingNode cid, lightingNid, maxid, "Lighting Reports", "service", 3, 0, "reportingindex.asp?bldg="&maxid, cnn1
			addBuildingNode cid, maintNid, maxid, "Mechanical PGI", "service", 2, 0, "", cnn1
			addBuildingNode cid, maintNid, maxid, "Other", "service", 3, 0, "", cnn1
	
	cnn1.execute = "UPDATE facilityinfo SET nodeid="&bnid&" WHERE id="&maxid
end if
dim tmpMoveFrame
tmpMoveFrame = "document.location = ""managebldg.asp?cid=" & cid & """"

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

rst1.close
set cnn1=nothing
%>