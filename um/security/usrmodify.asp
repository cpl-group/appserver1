<%@Language="VBScript"%>
<!-- #include file="../opslog/adovbs.inc" -->
<%
key=Request.Form("key")
choice=Request.Form("choice")
user=Request.Form("user")
passwd=Request.Form("passwd")
name=Request.Form("name")
telephone=Request.Form("telephone")
email=Request.Form("email")
initial_page=Request.Form("initial_page")
company=Request.Form("company")
regioncount=Request.Form("regioncount")


'response.write(key)
'response.write(choice)
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"dbCore")
if(choice = "Save") then
		
	strsql = "Insert into clients (username, paswd, name, telephone, email, company,  regioncount, initial_page) "_
	& "values ("_
	& "'" & user & "', "_
	& "'" & passwd & "', "_
	& "'" & name & "', "_
	& "'" & telephone & "', "_
	& "'" & email & "', "_
	& "'" & company & "', "& regioncount &" ,"_
	& "'" & initial_page & "')"
'Response.Write strsql	
end if

if choice="Update" then
strsql = "Update clients Set username='" & user & "', paswd='" & passwd & "', name='" & name & "', telephone='" & telephone & "', email='" & email & "', company='" & company & "',regioncount='" & regioncount & "', initial_page='"& initial_page &"' where (clientkey='"& key &"')"
'Response.Write strsql
end if

if choice="Delete" then
strsql = "Delete from clients where (clientkey='"& key &"')"
strsql2 = "Delete from clientsites where (userid='"& user &"')"
end if
cnn1.execute strsql
set cnn1=nothing
Response.Write "<script>" & vbCrLf

if choice="Delete" then
	'tmpMoveFrame2 = "parent.frames.site.location= "& Chr(34) & _
  '              "null.htm"& chr(34) & vbCrLf 
	tmpMoveFrame =  "document.location = " & Chr(34) & _
                "usrdetail.asp"& chr(34) & vbCrLf 
else
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "usrdetail.asp?username=" & user & chr(34) & vbCrLf 
'tmpMoveFrame2 = "parent.frames.site.location= "& Chr(34) & _
 '               "usrsite.asp?username="& user & chr(34) & vbCrLf 				  
'response.write user
end if
'Response.Write tmpMoveFrame2
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>