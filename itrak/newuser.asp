<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>

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

dim cid, userid, action
cid = request("cid")
userid = request("userid")
action = request("action")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

sqlstr = "select corp_name from clients where id = '"&cid&"'"
rst1.Open sqlstr, cnn1, 0, 1, 1
clientname = rst1("corp_name")
rst1.close

%>

<title>New User</title>
<script language="JavaScript">

function deleteuser(user){
  temp = "newuser.asp?cid=" & <%=cid%> & "&userid=" & <%=userid%> & "action=Delete";
  location = temp;
}

function confirmDelete(){
  retval = window.confirm("Are you sure you want to delete this item?");
  return retval;
}

</script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="saveuser.asp">
<input type="hidden" name="cid" value="<%=cid%>">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#0099ff">
  <td colspan="4">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#0099ff">
    <td><span class="standardheader"><%=clientname%> | User Info</span></td>
    <td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=cid%>'" class="standard"></td>
  </tr>
  </table>
  </td>
</tr>
<%
dim name, tel, email, pass, buttonval
if trim(action)="Edit" then
  sqlstr = "select * from users where userid like '"& userid &"' and clientid=" & cid
  rst1.Open sqlstr, cnn1, 0, 1, 1
  name = rst1("name")
  tel = rst1("telephone")
  email = rst1("email")
  pass = rst1("paswd")
  user = userid
  'initial_page = rst1("initial_page")
  buttonval = "Update"

elseif trim(action)="Add" then
  buttonval = "Save"

elseif trim(action)="Delete" then
  
end if
%>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">User ID</span></td> 
	<td><input type="text" name="userid" value="<%=user%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Password</span></td> 
	<td><input type="password" name="pass" value="<%=pass%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Name</span></td>
	<td><input type="text" name="name" value="<%=name%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Telephone</span></td>
	<td><input type="text" name="telephone" value="<%=tel%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Email</span></td>
	<td><input type="text" name="email" value="<%=email%>"></td>
</tr>
<!-- <tr bgcolor="#eeeeee"> -->
<!-- 	<td align="right"><span class="standard">Initial Page</span></td>  -->
<!-- 	<td><input type="text" name="pass" value="<% =initial_page%>" ></td> -->
<!-- </tr> -->
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	<td>
	<span class="standard">
	<input type="hidden" name="olduser" value="<%=userid%>">
	<input type="submit" name="action"  value="<%=buttonval%>" class="standard"> 
	&nbsp;<input type="reset" value="Cancel" onclick="location='manageaccounts.asp?cid=<%=cid%>'" class="standard"> 
	&nbsp;<input type="submit" name="action"  value="Delete" onclick="return confirmDelete();deleteuser('<%=userid%>');" class="standard">
	</span>
	</td>
</tr>
</table>
</form>
</body>
</html>
