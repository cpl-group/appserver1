<%@Language=VBscript%>
<%'for gathering ip addresses of users
dim userIP, conn, rsIP, IPcount, IPsql, email, out, userID
set conn = server.createobject("ADODB.connection")
set rsIP = server.createobject("ADODB.recordset")
conn.open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"

userIP = request.ServerVariables("REMOTE_ADDR")
email = request.form("email")
userID = session("userid")

rsIP.open "SELECT IP, email, [Count] FROM tblIP WHERE IP='"&userIP&"'", conn
if not(rsIP.EOF) then 'if the record set is not empty there has already been an IP entered --it tallies the count and redirects
    IPcount = rsIP("Count")+1
    conn.execute "UPDATE tblIP SET [Count]="&IPcount&" WHERE IP='"&userIP&"'", conn, adOpenStatic
    response.redirect "https://appserver1.genergy.com/g1_clients/boma/bomaindex.asp"
elseif (isValidEmail) then 'checks if there is a valid email if so inserts it and rediredts
    conn.execute "INSERT INTO tblIP (IP, [Count], Site, email) values ('"&userIP&"', 1, '"&userID&"', '"&email&"')", conn, adOpenStatic
    response.redirect "https://appserver1.genergy.com/g1_clients/boma/bomaindex.asp"
else 'if all esle fails it does nothing
    'nothing -- failed email login
end if

function isValidEmail()'valid email is not null "" and contains at least one "@" and "."
    isValidEmail = true
    if email="" then
        isValidEmail = false
    elseif ((InStr(email, "@")=0) or (InStr(email,".")=0)) then
        isValidEmail = false
    end if
end function
%>


<html>

<head>
<link rel="stylesheet" type="text/css" href="../holiday/holiform.css">

<title>GENERGY Registration and Login Page</title>

<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body BGCOLOR="#FFFFFF" LINK="#0000CC" VLINK="#003399" TEXT="#000033">

<form method="POST" action="email_login.asp" name="emaillogin">
  <table width="100%" border="0" height="100%" align="center">
    <tr> 
      <td height="30"> 
        <div align="center"><img src="/eri_th/images/login_header.jpg" width="321" height="40"></div>
      </td>
    </tr>
    <tr> 
      <td height="14"> 
        <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2">This 
          site is only available to Boma Members. </font></b></div>
      </td>
    </tr>
    <tr> 
      <td height="14"> 
        <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2">To 
          Verify your membership, we ask that you please enter your email address 
          below.</font></b></div>
      </td>
    </tr>
    <tr> 
      <td height="9"> 
        <div align="center"><font size="1"><b><font face="Arial, Helvetica, sans-serif">NOTE: 
          You will only have to enter it once from each computer you use to access 
          the site. </font></b></font></div>
      </td>
    </tr>
    <tr>
      <td height="2">
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
          </font> <%=out%> </div>
        <p align="center"><font face="Arial, Helvetica, sans-serif">Email Address</font> 
          <input type="text" name="email">
        </p>
          </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"> 
       
            <input type="submit" name="btnLogin" value="Submit">
            <input type="hidden" name="userID" value="<%=userID%>">
            <input type="reset" name="Reset" value="Reset">
            <font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font> 
        </div>
      </td>
    </tr>
    <tr> 
      <td> 
        <div align="center"><img src="/eri_th/images/login_client_footer.jpg" width="319" height="32"></div>
      </td>
    </tr>
  </table>
</form>


</body>

</form>

</html>





