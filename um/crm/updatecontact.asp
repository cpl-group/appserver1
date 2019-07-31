
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
ref=Request.Form("ref")
other=Request.Form("other")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"Intranet")

if trim(other)<>"" then
strsql="insert MKT_ref (referred) values ('"&other&"')"
cnn1.execute strsql
strsql = "update contacts set Company='" & company & "',first_name='" & first & "',last_name='" &last & "',address='" & addr & "',city='" & city & "',state='" & state & "',zip='" & zip & "',country='" &country & "',title='" & t & "',phone='" & phone & "',fax='" & fax&"',email='" & email & "',org='" & org & "',orgtype='" & orgtype& "',referredby='"& other &"', industry='"& industry &"' where id='" & cid & "'"
cnn1.execute strsql

else

strsql = "update contacts set Company='" & company & "',first_name='" & first & "',last_name='" &last & "',address='" & addr & "',city='" & city & "',state='" & state & "',zip='" & zip & "',country='" &country & "',title='" & t & "',phone='" & phone & "',fax='" & fax&"',email='" & email & "',org='" & org & "',orgtype='" & orgtype& "',referredby='" & ref& "', industry='" & industry & "' where id='" & cid & "'"

cnn1.execute strsql
end if

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "contactview.asp?cid="& cid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
set cnn1=nothing
%>