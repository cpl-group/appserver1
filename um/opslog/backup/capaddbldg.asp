<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
bldgnum=Request.Form("bldgnum")
address=Request.Form("address")
sqft=Request.Form("sqft")
rev=Request.Form("rev")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=Capacity_db;"
sql = "Insert into tlbldg (bldgnum, address, sqft, revision) "_
	& "values ("_
	& "'" & bldgnum & "', "_
	& "'" & address & "', "_
	& "'" & sqft & "', "_
	& "'" & rev & "')"


'response.write sql
'cnn1.execute sql
'cnn1.execute strsql2
set cnn1=nothing
'tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
'				"capindex.asp"& chr(34) & vbCrLf 
tmpMoveFrame =  "document.location = " & Chr(34) & _
				"capbldginfo.asp?bldgnum="& bldgnum & _
				"&address=" & address & "&sqft=" & sqft &_
				"&rev="& rev & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>