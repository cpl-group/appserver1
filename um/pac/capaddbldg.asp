<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldgnum, address, sqft, rev, sql, cnn1, tmpmoveframe
bldgnum=split(Request.Form("bldgnum"),"|")(0)
address=split(Request.Form("bldgnum"),"|")(1)

sqft=Request.Form("sqft")
rev=Request.Form("rev")
if bldgnum = "" or address= "" or sqft="" or rev="" then
	response.redirect "capnewbldg.asp?msg=All fields must be completed"
end if
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"engineering")
sql = "Insert into tlbldg (bldgnum, address, sqft, revision, rev_date) "_
	& "values ("_
	& "'" & bldgnum & "', "_
	& "'" & address & "', "_
	& "'" & sqft & "', "_
	& "'" & rev & "', convert(char, getdate(), 101))"

logger(sql)
cnn1.execute sql
set cnn1=nothing
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				"capindex.asp"& chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>