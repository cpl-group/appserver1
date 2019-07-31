
<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")


strsql = "Update buildings Set bldgname='" & Request.Form("bldgname") & "', strt='" & Request.Form("strt") & "', city='" & Request.Form("city") &"', state='" & Request.Form("state") & "', zip='" & Request.Form("zip") & "', btbldgname='" & Request.Form("btbldgname") & "', btstrt='" & Request.Form("btstrt") & "', btcity='" & Request.Form("btcity") &"', btstate='" & Request.Form("btstate") & "', btzip='" & Request.Form("btzip") & "' where bldgnum='" & Request.Form("bldgnum") & "'"
'response.write strsql
'response.end

cnn1.execute strsql

set cnn1=nothing

dim tmp

tmp =  "title.asp?bldg=" & Request("bldgnum")
Response.redirect tmp
%>
