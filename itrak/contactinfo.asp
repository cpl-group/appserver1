<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim bldg,DefaultBackgroundColor
bldg = request.querystring("bldg")

dim cnn1, rst1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getconnect(0,0,"engineering")
rst1.open "SELECT * FROM FacInfo WHERE bldgnum="&bldg, cnn1
%>
<html>
<head>
<title>Contact Info</title>
<link rel="Stylesheet" href="/genergy2/styles.css">
<script language="JavaScript">
function openpage(id){
  document.location = 'editcontacts.asp?bldg=<%=bldg%>&contactid='+id;
}
try{top.applabel("Contact Management");}catch(exception){}
</script>
</head>

<body bgcolor="#FFFFFF">

<table border="0" width="100%" cellpadding="3" cellspacing="1">
  <tr bgcolor="#ffffff"> 
    <td colspan="5"> <table border=0 cellpadding="0" cellspacing="0" width="100%">
        <tr valign="middle"> 
          <td align="right" colspan=2><input type="button" value="Add New Contact" onclick="document.location='editcontacts.asp?bldg=<%=bldg%>';" class="standard"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
  <td colspan="5">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" dwcopytype="CopyTableRow">
      <tr> 
  <tr bgcolor="#cccccc">
        <td width="17" bgcolor="#FFFFFF"><div align="center"><img src="images/note.gif" width="28" height="21"></div></td>
    <td nowrap><span class="standard"><b>Name</b></span></td>
    <td nowrap><span class="standard"><b>Email</b></span></td>
    <td nowrap><span class="standard"><b>Title</b></span></td>
    <td nowrap><span class="standard"><b>Phone</b></span></td>
  </tr>
  <%do until rst1.eof
%>
  <tr valign="top" bgcolor="white" onclick="openpage(<%=rst1("id")%>);" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = '<%=DefaultBackgroundColor%>'">
    <td width="13" bgcolor="#FFFFFF"> 
      <%if rst1("alert") then%>
      <div align="center"><img src="images/greencheck.gif" alt="User Configured to recieve maintenance emails" width="13" height="15" align="absmiddle"></div> 
        <%end if%>
      </td>
    <td><span class="standard"><%=rst1("name")%></span></td>
    <td><span class="standard"><%=rst1("email")%></span></td>
    <td><span class="standard"><%=rst1("title")%></span></td>
    <td><span class="standard"><%=rst1("phone")%></span></td>
  </tr>
  <%rst1.movenext
loop%>
</table>
</body>

</html>
