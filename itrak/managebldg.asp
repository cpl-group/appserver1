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
sqlstr = "select corp_name from clients where id = '"&cid&"'"
rst1.Open sqlstr, cnn1, 0, 1, 1
clientname = rst1("corp_name")
rst1.close

%>
<script>

function newbldg(owner){
  var temp= "newbldg.asp?owner=" + owner + "&cid=" + <%=cid%>;
	location=temp;
}

function editbldg(id){
  var temp = "updatebldg.asp?id=" + id + "&cid=" + <%=cid%>;
  location=temp;
}
</script>
<html>
<head>
<title>Facility Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="bldgupd.asp">
  <table width="100%" border="0" cellpadding="3" cellspacing="1">
  <tr bgcolor="#0099ff">
    <td colspan="3">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr bgcolor="#0099ff">
      <td><span class="standardheader"><%=clientname%> | Facility Information</span></td>
      <td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=cid%>'" class="standard" disabled>&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=cid%>'" class="standard"></td>
    </tr>
    </table>
    </td>
  </tr>
  <tr bgcolor="#cccccc">
	  <td><span class="standard"><b>Building Name</b></span></td>
	  <td><span class="standard"><b>Address, City, State Zip</b></span></td>
	  <td><span class="standard"><b>Square Feet</b></span></td>
  </tr>

  <%
  sqlstr = "select * from facilityinfo where clientid='"& cid &"'"
  rst1.Open sqlstr, cnn1, 0, 1, 1

  do until rst1.eof
  %>
    
  <tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="editbldg(<%=rst1("id")%>);">
	<td><span class="standard"><%=rst1("bldgname")%></span></td>
	<td><span class="standard"><%=rst1("address")%>, <%=rst1("city")%>, <%=rst1("state")%> &nbsp;<%=rst1("zip")%></span></td>
	<td><span class="standard"><%=rst1("sqft")%></span></td>
  </tr>

  <%
  rst1.movenext
  loop
  
  %>
  <tr bgcolor="#cccccc">
    <td colspan="3">
    <input type="hidden" name="owner" value="<%=cid%>">
    <input type="hidden" name="cid" value="<%=cid%>">
    <input type="button" name="AddNewBuilding" value="Add New Building" onclick="newbldg(owner.value)" class="standard">
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