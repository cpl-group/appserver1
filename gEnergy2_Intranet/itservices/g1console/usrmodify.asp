<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
key=Request.Form("key")
choice=trim(Request.Form("choice"))
user=Request.Form("userlogin")
passwd=Request.Form("passwd")
name=Request.Form("name")
email=Request.Form("email")
custompage = trim(Request.Form("defaultpage"))
initial_page=Request.Form("page")
company=split(Request.Form("company_portfolio"), "_")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"dbCore")
if(choice = "Save") then
		
	strsql = "Insert into clients (username, paswd, name, telephone, email, company,portfolio_id, initial_page, custompage) "_
	& "values ("_
	& "'" & user & "', "_
	& "'" & passwd & "', "_
	& "'" & name & "', "_
	& "'" & telephone & "', "_
	& "'" & email & "', "_
	& "'" & company(0) & "', "_
	& "'" & company(1) & "', "_
	& "'" & initial_page & "', "_
	& "'" & custompage & "')"
'Response.Write strsql	
end if

if choice="Update" then
strsql = "Update clients Set username='" & user & "', paswd='" & passwd & "', name='" & name & "', telephone='" & telephone & "', email='" & email & "', company='" & company(0) & "', portfolio_id='"& company(1) & "', initial_page='"& initial_page &"', custompage = "&custompage&" where (clientkey='"& key &"')"
'Response.Write strsql
end if

if choice="Delete" then
strsql = "Delete from clients where (clientkey='"& key &"')"
strsql2 = "Delete from clientsetup where (userid='"& user &"')"
cnn1.execute strsql2
end if

cnn1.execute strsql
set cnn1=nothing
Response.Write "<script>" & vbCrLf

if choice="Delete" then
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
                "index.asp"& chr(34) & vbCrLf 
else
	tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "usrdetail.asp?userid=" & user & chr(34) & vbCrLf 
end if
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>