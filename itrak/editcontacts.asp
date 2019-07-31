<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim name, email, title, phone, bldg, contactid, action,alert
bldg = request("bldg")
contactid = request("contactid")
action = request("action")

dim rst1, cnn1, cmd
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
cnn1.open getconnect(0,0,"engineering")
cmd.activeconnection = cnn1

if trim(contactid)<>"" and trim(action)="" then
	rst1.open "SELECT * FROM FacInfo WHERE id="&contactid, cnn1
	if not rst1.eof then
		name = rst1("name")
		email = rst1("email")
		title = rst1("title")
		phone = rst1("phone")
		alert = rst1("alert")
	end if
else
	name = request("name")
	email = request("email")
	title = request("title")
	phone = request("phone")
	alert = request("alert")
end if
if action = "Add" then
	cmd.commandtext = "INSERT INTO FacInfo (name, email, title, phone, alert, bldgnum) values ('"&name&"', '"&email&"', '"&title&"', '"&phone&"','"&alert&"','"&bldg&"')"
	cmd.execute
elseif action="Update" then
	cmd.commandtext = "UPDATE FacInfo set name='"&name&"', email='"&email&"', title='"&title&"', phone='"&phone&"', alert='"&alert&"' WHERE id="&contactid
	cmd.execute
elseif action="Delete" then
	cmd.commandtext = "DELETE FROM FacInfo WHERE id="&contactid
	cmd.execute
end if
if trim(action)<>"" or trim(action)="Cancel" then%>
<script>
document.location.href="contactinfo.asp?bldg=<%=bldg%>";
window.close();
</script>
<%end if%>

<html>
<head>
<title>Revision Management</title>
<script language="JavaScript">
<%if trim(contactid)<>"" then%>
try{top.applabel("Contact Management - Edit Contact");}catch(exception){}
<%else%>
try{top.applabel("Contact Management - Add New Contact");}catch(exception){}
<%end if%>

function confirmDelete(){
  retval = window.confirm("Are you sure you want to delete this item?");
  return retval;
}

function nevermind(){
  var temp = "contactinfo.asp?bldg=<%=bldg%>";
  location = temp;
}

</script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action="editcontacts.asp">

  <table border="0" width="100%" cellpadding="3" cellspacing="1">
    <tr align="right" bgcolor="#FFFFFF"> 
      <td colspan="2"> <span class="standardheader"> 
        <%if trim(contactid)<>"" then%>
        <input type="submit" name="action" value="Update" class="standard">
        <input type="submit" name="action" value="Delete" onClick="return confirmDelete();" class="standard">
        <input type="submit" name="action" value="Cancel" onClick="nevermind();" class="standard">
        <%else%>
        <input type="submit" name="action" value="Add" class="standard">
        <input type="submit" name="action" value="Cancel" onClick="nevermind();" class="standard">
        <%end if%>
        </span> </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Name</span></td>
      <td><input type="text" name="name" size="25" value="<%=name%>"></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Email</span></td>
      <td><input type="text" name="email" size="15" value="<%=email%>"> </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Title</span></td>
      <td><input type="text" name="title" size="15" maxlength="25" value="<%=title%>"></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Phone</span></td>
      <td><input type="text" name="phone" size="15" maxlength="20" value="<%=phone%>"></td>
    </tr>
    <tr bgcolor="#eeeeee">
      <td align="right">&nbsp;</td>
      <td> <input name="alert" type="checkbox" value="1" <%if alert then %> checked <%end if%>> <font size="1" face="Arial, Helvetica, sans-serif">Activate 
        Account to Receive Maintenance emails</font></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td><input type="hidden" name="bldg" value="<%=bldg%>">
        <input type="hidden" name="contactid" value="<%=contactid%>"></td>
      <td>&nbsp;</td>
    </tr>
  </table>

</form>
</body>
</html>
