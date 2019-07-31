<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
bldgnum=Request.form("bldgnum")
address=Request.form("address")
sqft=request.Form("sqft")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_capacity_db")
sql = "Update tlbldg Set address='"& address &"', sqft='"& sqft &"', rev_date=convert(char, getdate(), 101) where bldgnum='"& bldgnum &"'"

'response.write sql
cnn1.execute sql

set cnn1=nothing

Response.redirect "capbldginfo.asp?bldgnum="& bldgnum
%>