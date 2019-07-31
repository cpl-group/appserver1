<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
bldgnum=Request.Form("bldgnum")
address=Request.Form("address")
sqft=Request.Form("sqft")
rev=Request.Form("rev")
if bldgnum = "" or address= "" or sqft="" or rev="" then
	response.redirect "capnewbldg.asp"
end if
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_capacity_db")
sql = "Insert into tlbldg (bldgnum, address, sqft, revision, rev_date) "_
	& "values ("_
	& "'" & bldgnum & "', "_
	& "'" & address & "', "_
	& "'" & sqft & "', "_
	& "'" & rev & "', convert(char, getdate(), 101))"


'response.write sql
cnn1.execute sql
'cnn1.execute strsql2
set cnn1=nothing
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				"capindex.asp"& chr(34) & vbCrLf 
'tmpMoveFrame =  "document.location = " & Chr(34) & _
'				"capbldginfo.asp?bldgnum="& bldgnum & _
'				"&address=" & address & "&sqft=" & sqft &_
'				"&rev="& rev & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>