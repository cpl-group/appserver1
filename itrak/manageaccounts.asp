<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
		if isempty(Session("name")) then
'			Response.Redirect "../index.asp"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
	
cid= Request.Querystring("cid")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

dim logo
rst1.open "SELECT distinct logo FROM clients WHERE id="&cid, cnn1
if not rst1.eof then logo = rst1("logo")
if trim(logo)="" then logo = "logos/genergy2.gif"
rst1.close

sqlstr = "select * from clients where id = '"&cid&"'"
rst1.Open sqlstr, cnn1, 0, 1, 1
clientname = rst1("corp_name")
%>
<script>

function newuser(){
  var temp="newuser.asp?cid=" + <%=cid%> + "&action=Add";
  location=temp;
}

function edituser(user){
  var temp="newuser.asp?cid=" + <%=cid%> + "&userid=" + user + "&action=Edit";
  location=temp;
}

</script>
<html>
<head>
<title>User Accounts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="clientview.asp?id=<%=cid%>">
  <table width="100%" border="0" cellpadding="3" cellspacing="1">
  <tr bgcolor="#0099ff">
    <td colspan="2">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr bgcolor="#0099ff">
      <td><span class="standardheader"><%=clientname%> | Account Manager</span></td>
      <td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=cid%>'" class="standard" disabled>&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=cid%>'" class="standard"></td>
    </tr>
    </table>
    </td>
  </tr>

  <tr bgcolor="#eeeeee" valign="top">
    <td width="18%" align="right"><span class="standard"></span></td>
    <td><span class="standard">
    <b><%=rst1("corp_name")%></b><br>
    </span>
    </td>
  </tr>
  <tr bgcolor="#eeeeee" valign="top">
    <td align="right"><span class="standard">Address</span></td>
    <td>
    <span class="standard">
    <%=rst1("address")%><br>
    <%=rst1("city")%>, <%=rst1("state")%> &nbsp;<%=rst1("zip")%>
    </span>
    </td>
  </tr>
  <tr bgcolor="#eeeeee" valign="top">
    <td align="right"><span class="standard">Main Contact</span></td>
    <td><span class="standard">
    <%=rst1("Contact")%><br>
    <%=rst1("contactphone")%><br>
    </span>
    </td>
  </tr>
  <tr bgcolor="#eeeeee" valign="top">
    <td align="right"><span class="standard">Logo</span></td>
    <td><span class="standard">
    <img src="<%=logo%>" width="225" height="40" border="0" style="margin-bottom:4px;"><br>
    <span style="font-size:7pt;">(<%=logo%>)</span>
    </span>
    </td>
  </tr>
  <tr bgcolor="#cccccc" valign="top">
   <td></td>
   <td><input type="submit" value="Edit" style="padding-left:4px;padding-right:4px;" class="standard"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td></td>
    <td>
    <table width="89%" border="0" cellpadding="3" cellspacing="1">
    <tr>
      <td><span class="standard"><b>User ID</b></span></td>
      <td><span class="standard"><b>Name</b></span></td>
      <td><span class="standard"><b>Telephone</b></span></td>
      <td><span class="standard"><b>Email</b></span></td>
    </tr>
  
    <%
    rst1.close
    sqlstr = "select * from users where clientid='"& cid &"'"
    rst1.Open sqlstr, cnn1, 0, 1, 1
  
    do until rst1.eof
    %>
      
    <tr bgcolor="#eeeeee" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'eeeeee'" onclick="edituser('<%=rst1("userid")%>');">
    <td><span class="standard"><%=rst1("userid")%></span></td>
    <td><span class="standard"><%=rst1("name")%></span></td>
    <td><span class="standard"><%=rst1("telephone")%></span></td>
    <td><span class="standard"><%=rst1("email")%></span></td>
    </tr>
  
    <%
    rst1.movenext
    loop
    
    %>
    </table>
    
    </td>
  </tr>
  <tr bgcolor="#cccccc">
    <td></td>
    <td>
    <input type="hidden" name="owner" value="<%=cid%>">
    <input type="hidden" name="cid" value="<%=cid%>">
    <input type="button" name="action" value="Add New User" onclick="newuser()" class="standard">
    </td>
  </tr>
  </table>
</form>
<%
rst1.close
set cnn1=nothing
%>
</body>
</html>