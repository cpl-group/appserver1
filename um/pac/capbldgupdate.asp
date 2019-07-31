<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="adovbs.inc" -->
<%
bldgnum=Request.form("bldgnum")
address=Request.form("address")
sqft=request.Form("sqft")
revdate=request.Form("revdate")
revnum=request.Form("revnum")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"engineering")
sql = "Update tlbldg Set address='"& address &"', sqft='"& sqft &"', rev_date='"&revdate &"' , revision='"&revnum &"'where bldgnum='"& bldgnum &"'"

'response.write sql
cnn1.execute sql

set cnn1=nothing

Response.redirect "capbldginfo.asp?bldgnum="& bldgnum
%>