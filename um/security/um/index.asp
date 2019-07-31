<%@Language="VBScript"%>

<html>
<head>
<title>Employee Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function loadpage(page) {
	opener.location = page
	window.close()
}
</script>
<%
'yoooooooooooooooooooooooooooooooo
	if isempty(Request.form("login")) then
		Session("NumAttempts") = 0
		TheMessage = "Welcome"
	else
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open getConnect(0,0,"dbCore")

		strsql = "SELECT * from employees where login='" & Request.Form("login") & "'"
		rst1.Open strsql, cnn1, 0, 1, 1
					
		if not rst1.eof then
			if rst1("loginattempts") > 3 then 
							
				TheMessage="Account Locked. Please Contact Your System Adminstrator."
				Session("NumAttempts") = 4
				
			Else
		
				rst1.close		
				strsql = "SELECT * from employees where login='" & Request.Form("login") & "' and password='" & Request.Form("password") & "'"
				rst1.Open strsql, cnn1, 0, 1, 1
		
				if rst1.eof then
							
							Session("NumAttempts") = Session("NumAttempts") + 1
							if Session("NumAttempts") = 0 then
								TheMessage = "Login not found, please try again.(" & Session("NumAttempts") & ")"
							else if Session("NumAttempts") > 3 then
								TheMessage = "Login not found. No more attempts allowed.(" & Session("NumAttempts") & ")"
								else
								TheMessage = "Login not found, please try again.(" & Session("NumAttempts") & ")"
								end if
							end if
							strsql = "UPDATE employees SET loginattempts="& Session("NumAttempts") & " where login = '" & Request.Form("login") & "'"
							cnn1.execute strsql
							rst1.close
										
				else
							Session("login") = rst1("login")
							Session("name") = rst1("name")
							Session("roleid")=4
							Session("um") = rst1("um")
							Session("eri") = rst1("eri")
							Session("opslog") = rst1("opslog")
							Session("ts") = rst1("ts")
							Session("corp") = rst1("corp")
							Session("it") = rst1("it")
							Session("admin") = rst1("admin")
							
							Session("fMessage") = "You are currenlty logged on."
												
							strsql = "UPDATE employees SET status=1 where login = '" & Request.Form("login") & "'"
							cnn1.execute strsql
							rst1.close
							set cnn1 = nothing
							
						%>
						<script>loadpage('http://appserver1.genergy.com/um/init.asp')</script>
						<%
				end if
							
			end if
		
		else

			rst1.close		
			strsql = "SELECT * from employees where login='" & Request.Form("login") & "' and password='" & Request.Form("password") & "'"
			rst1.Open strsql, cnn1, 0, 1, 1
	
			if rst1.eof then
						
						Session("NumAttempts") = Session("NumAttempts") + 1
						if Session("NumAttempts") = 0 then
							TheMessage = "Login not found, please try again.(" & Session("NumAttempts") & ")"
						else if Session("NumAttempts") > 3 then
							TheMessage = "Login not found. No more attempts allowed.(" & Session("NumAttempts") & ")"
							else
							TheMessage = "Login not found, please try again.(" & Session("NumAttempts") & ")"
							end if
						end if
						strsql = "UPDATE employees SET loginattempts=4 where login = '" & Request.Form("login") & "'"
						cnn1.execute strsql
						rst1.close
									
			else
						Session("login") = rst1("login")
						Session("name") = rst1("name")
						Session("um") = rst1("um")
						Session("eri") = rst1("eri")
						Session("opslog") = rst1("opslog")
						Session("ts") = rst1("ts")
						Session("fMessage") = "You are currenlty logged on."
						strsql = "UPDATE employees SET status=1 where login = '" & Request.Form("login") & "'"
						cnn1.execute strsql
						rst1.close
						set cnn1 = nothing
						response.redirect "http://appserver1.genergy.com/um/main.asp"
			end if
			
				
		end if





	end if					
%>


<body bgcolor="#FFFFFF" text="#000000" onload="document.forms['form1'].login.focus();">
<table width="100%" border="0" height="100%">
  <tr>
    <td align="center" valign="middle"> 
      <div align="center">
        <table width="100%" border="0" height="100%">
          <tr>
            <td height="30"> 
              <div align="center"><img src="images/login_header.jpg" width="321" height="40"></div>
            </td>
          </tr>
          <tr>
            <td height="137">
              <div align="center">
			   <% if Session("NumAttempts") > 3 then 
			   
			   		Response.write themessage 
			     
				  else 
			   %>
			    <form name="form1" method="post" action="index.asp">
                  <p><font face="Arial, Helvetica, sans-serif">Username</font> 
                    <input type="text" name="login">
                  </p>
                  <p><font face="Arial, Helvetica, sans-serif">Password</font> 
                    <input type="password" name="password">
                  </p>
                  <p> 
                    <input type="submit" name="Submit" value="Login">
                    <input type="reset" name="Reset" value="Reset">
                  </p>
				  <% Response.write themessage %>
                </form>
				<% End IF %>
              </div>
            </td>
          </tr>
          <tr>
            <td>
              <div align="center"><img src="images/login_footer.jpg" width="321" height="40"></div>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>
</body>
</html>
