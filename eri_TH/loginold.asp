<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
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
cnn1.Open application("cnnstr_security")


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
      	path= Trim(rsVis("initial_page"))'&"?fc="&Oct(rnd*1000000)
      	loggedin = True
		session("userid") = request("userid")
		session("RoleID") = rsVis("roleid")
		Dim rsRole
		set rsRole = server.createobject("ADODB.Recordset")
		rsRole.open "SELECT Label FROM tblrole WHERE roleid=" & session("RoleID"), cnn1
		session("RoleName") = rsRole("Label")
		rsRole.close
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
	response.redirect path
  End If
  
End Sub
%>
<html>

<head>
<link rel="stylesheet" type="text/css" href="../holiday/holiform.css">

<title>gEnergyOne Login</title>

<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>


<body BGCOLOR="#FFFFFF" LINK="#0000CC" VLINK="#0000CC" TEXT="#000033" onload="document.forms['FrontPage_Form1'].userid.focus();">
<form method="POST" action="login.asp" name="FrontPage_Form1">
  <table width="100%" border="0" height="100%" align="center">
    <tr> 
      <td height="30"> 
        <div align="center"><img src="images/login_header.jpg" width="321" height="40"></div>
      </td>
    </tr>
    <tr> 
      <td height="137"> 
        <div align="center"> <font face="Arial, Helvetica, sans-serif" size="2"> 
          <%
If Session("loginemail") = "" Then
  Response.Write "You are currently not logged in."
Else
  'Response.Write "You are currently logged in as " & Session("loginemail")
End If
%>
          </font> 
          <p><font face="Arial, Helvetica, sans-serif">Username</font> 
            <input type="text" name="userid">
          </p>
          <p><font face="Arial, Helvetica, sans-serif">Password</font> 
            <input type="password" name="paswd">
          </p>
          <p> 
            <input type="submit" name="btnLogin" value="Login">
            <input type="reset" name="Reset" value="Reset">
          </p>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2"><span class="err"><%=fMsg%></span><span class="msg"><%=iMsg%></span>&nbsp;</font> 
          </b></div>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><font color="#FF0000"><img src="images/login_client_footer.jpg" width="319" height="32"></font></b></font></div>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><font color="#FF0000">NOTE: 
          </font><font face="Arial, Helvetica, sans-serif" size="1" color="#FF0000"><b>gEnergyOne 
          Rev. 1.1.2;</b></font><font color="#FF0000"> SOME ADVANCED FEATURES 
          REQUIRE INTERNET <a href="http://www.microsoft.com/windows/ie/downloads/ie6/default.asp" target="new">EXPLORER 
          +6.0</a></font></b></font> </div>
      </td>
    </tr>
    <tr>
      <td height="2">
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"></font></div>
      </td>
    </tr>
  </table>
  <div align="center"></div>
</form>


</body>

</form>

</html>





