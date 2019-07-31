<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
company=Request.Form("CompanyName")
first=Request.Form("first")
last=Request.Form("last")
t=Trim(Request.Form("title"))
addr=Request.Form("addr")
city=Request.Form("city")
state=Request.Form("state")
phone=Request.Form("phone")
fax=Request.Form("fax")
zip=Request.Form("zip")
country=Request.Form("country")
email=Request.Form("email")
cid=Request.Form("cid")
org=Request.Form("org")
orgtype=Request.Form("orgtype")
industry=Request.Form("industry")
otherindustry=Request.Form("otherindustry")
ref=Request.Form("ref")
other=Request.Form("other")
otherorg=Request.Form("otherorg")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"Intranet")

if trim(otherorg)<>"" then
strsql="insert mkt_organizations (org) values ('"&otherorg&"')"
cnn1.execute strsql
rst1.open "select id from mkt_organizations where org='"&otherorg&"'", cnn1
org = rst1("id")
rst1.close
end if

if trim(otherindustry)<>"" then
strsql="insert mkt_industries (industry) values ('"&otherindustry&"')"
cnn1.execute strsql
rst1.open "select id from mkt_industries where industry='"&otherindustry&"'", cnn1
industry = rst1("id")
rst1.close
end if

if ref="Other" then
strsql="insert MKT_ref (referred) values ('"&other&"')"
cnn1.execute strsql

strsql = "insert contacts (Company,first_name,last_name,address,city,state,zip,country,title,phone,fax,EMAIL,org,orgtype,referredby,industry)values ('" & company & "', '" & first & "', '" &last & "', '" & addr & "', '" & city & "', '" & state & "', '" & zip & "','" &country & "','" & t & "', '" & phone & "', '" & fax & "', '" & EMAIL& "', '" & org& "', '" & orgtype& "', '" & other& "', '"& industry &"')"
cnn1.execute strsql

else
strsql = "insert contacts (Company,first_name,last_name,address,city,state,zip,country,title,phone,fax,EMAIL,org,orgtype,referredby,industry)values ('" & company & "', '" & first & "', '" &last & "', '" & addr & "', '" & city & "', '" & state & "', '" & zip & "','" &country & "','" & t & "', '" & phone & "', '" & fax & "', '" & EMAIL& "', '" & org& "', '" & orgtype& "', '" & ref& "', '"& industry &"')"
cnn1.execute strsql
end if
Set rst1 = Server.CreateObject("ADODB.recordset")
strsql="select max(id)as cid from contacts"
rst1.Open strsql, cnn1, 0, 1, 1
cid=rst1("cid")
'response.write job
'response.end
cnn1.execute strsql
set cnn1=nothing

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "contactview.asp?cid="& cid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>