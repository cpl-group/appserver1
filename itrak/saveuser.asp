<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!--#INCLUDE file="treenode_functions.asp"-->
<%
dim cid, name, tel, email, userid, pass, olduser, action
cid=Request.Form("cid")
name=Request.Form("name")
tel=Request.Form("telephone")
email=Request.Form("email")
userid=Request.Form("userid")
pass=Request.Form("pass")
olduser=Request.Form("olduser")
action=Request.Form("action")

dim cnn1, rst1, strsql, tmpMoveFrame
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

if trim(action)="Update" then
  strsql = "UPDATE users SET name='" & name & "', telephone='" & tel & "', email='" & email & "', userid='" & userid & "', paswd='" & pass & "' WHERE userid like '" & olduser & "' AND clientid = " & cid
elseif trim(action)="Delete" then
  strsql = "DELETE from users WHERE userid like '" & userid & "' and clientid=" & cid
else
  strsql = "insert users (clientid,userid,name,telephone,email,paswd) values ('" & cid & "', '" & userid & "', '" & name & "', '" & tel & "', '" & email & "', '" & pass & "')"
end if

'response.write strsql
'response.end
cnn1.execute strsql

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "manageaccounts.asp?cid="& cid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

'rst1.close
set cnn1=nothing

%>