<!-- #include file="./adovbs.inc" -->
<%


fMsg = ""
iMsg = ""
loggedin = False

' Determine which button the Web visitor clicked and take
' appropriate action.

If Request("btnLogin") <> "" Then

  ProcessLogin
  
End If

Sub ProcessLogin()
' Verify that the Web visitor submitted both an email address
' and a password.

  If Request("userid") = "" Then
  
    	If Request("paswd") = "" Then
    
      		fMsg = "Enter USER ID and password."
      		Exit Sub
      
    		Else    
      		fMsg ="enter USER ID"
      		Exit Sub
      		
    	End If
    	
  Else
    If Request("paswd") = "" Then
      fMsg = "Enter password."
      Exit Sub
    End If
  End If

' Build ADO connection string.
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"


' Create and open an ADO recordset object.
  sql = "SELECT * " & _
         "FROM clients " & _
         "WHERE (username = '" & Request("userid") & "') ; "
  Set rsVis = Server.CreateObject("ADODB.Recordset")
  
  rsVis.Open sql, cnn1, adOpenDynamic, adLockOptimistic

 If rsVis.EOF Then
        
'     user ID address not in database, returning visitor
      fMsg = "USER ID not found."
      exit sub       
  Else  
       If Request("paswd") = Trim(rsVis("paswd")) Then
      	path= Trim(rsVis("initial_page"))
      	loggedin = True
		
      	else
      	fMsg = "Wrong Password"

      	exit sub   
       end if
       
End If
  
  rsVis.Close

  If fMsg <> "" Then
  Exit Sub
  End If

' if new user redirect to a welcome page

  	 
  If loggedin Then  	
    Session("loginemail") = Request("userid")
    
    'Response.redirect GoBack
    response.redirect "../g1_clients/index.asp"
   
  End If
  
End Sub
%>
<html>

<head>
<link rel="stylesheet" type="text/css" href="../holiday/holiform.css">

<title>GENERGY Registration and Login Page</title>

<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body BGCOLOR="#FFFFFF" BACKGROUND="http://www.genergy.com/bg_login_right.gif" LINK="#0000CC" VLINK="#003399" TEXT="#000033">

<h1 align="left"><FONT COLOR="#CC0000" SIZE="+2" FACE="Trebuchet MS,Arial,Helvetica">

<br>
<br>
Genergy Customer Login</FONT></h1>

<p class="msg">
<%
If Session("loginemail") = "" Then
  Response.Write "You are currently not logged in."
Else
  'Response.Write "You are currently logged in as " & Session("loginemail")
End If

%></p>

<p><span class="err"><%=fMsg%></span><span class="msg"><%=iMsg%></span>&nbsp;</p>

<form method="POST" action="login.asp" name="FrontPage_Form1">
 <table border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF"  width="540" height="385">
 <tr>
 <td align="center" width="528" height="92">
 <table border="0" cellpadding="0" cellspacing="3" width="390" height="123">
 <tr>
 <td FONT FACE="Trebuchet MS,Arial,Helvetica" width="100" height="25">USER ID:</FONT></td>
 <td width="274" height="25"><input type="text" name="userid" size="30" value="<%=Request("userid")%>"></td>
 </tr>
 <tr>
 <td FONT FACE="Trebuchet MS,Arial,Helvetica" width="100" height="25">PASSWORD:</font></td>
 <td width="274" height="25"><input type="password" name="paswd" size="15" value="<%=Request("paswd")%>"><input type="submit" value="Login" name="btnLogin"></td>
 </tr>
 </table>
 </td>
 </tr>
 </table>
</form>


</body>

</form>

</html>





