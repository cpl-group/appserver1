
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
	
id= Request.Querystring("id")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

sqlstr = "select * from facilityinfo where id='"& id &"'"


'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then

%>
<html>
<head>
<title>Building Update</title>
<script language="JavaScript">
function nevermind(){
  var temp = "editbldg.asp?id=" + <%=id%>;
  location = temp;
}

function goEdit(){
  var temp = "editbldg.asp?id=" + <%=id%> + "&action='edit'";
  location = temp;
}

function checkfields(theform){
//note: this function is slightly different than the checkfields
//on other pages. It removes single quotes rather than replacing
//them with two single quotes
  retval = true;
  for (i=0;i<theform.length;i++){
    if (theform.elements[i].value.indexOf("'") > -1) {
      theform.elements[i].value = theform.elements[i].value.replace(/'/g,"" );
    }
  }
  return retval;
}

function checkNumber(thefield){
  re = /\D/;
    bad = re.test(document.forms['form2'].elements[thefield].value);
    if (bad) { 
      document.forms['form2'].elements[thefield].style.backgroundColor = "#ccccff";
      alert("Please only use numbers in this field.");
    } else {
      document.forms['form2'].elements[thefield].style.backgroundColor = "#ffffff"; 
    }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="bldgupd.asp" onsubmit="return checkfields(this);">
  <table width="100%" border="0" cellpadding="3" cellspacing="1">
  <tr bgcolor="#0099ff">
    <td colspan="2"><span class="standardheader">Facility Information</span></td>
  </tr>
  
	<input type="hidden" name="srcfile" value="editbldg">
  <% if request("action") <> "" then %>
  <tr bgcolor="#eeeeee"> 
    <td width="18%" align="right">
    <span class="standard">Building 
      Name<input type="hidden" name="bid" value="<%=rst1("id")%>">
    </span>
    </td>
    <td><input type="text" name="bldgnum" value="<%=rst1("bldgname")%>"></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Address</span></td>
      <td><input type="text" name="address" value="<%=rst1("address")%>"></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">City</span></td>
      <td><input type="text" name="city" value="<%=rst1("city")%>"></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">State</span></td>
      <td>
      <table border=0 cellpadding="0" cellspacing="0">
      <tr valign="middle">
        <td><span class="standard"><input type="text" name="state" value="<%=rst1("state")%>" size="4"></span></td>
        <td width="12"><span class="standard">&nbsp;</span></td>
        <td><span class="standard">Zip Code&nbsp;</span></td>
        <td><span class="standard"><input type="text" name="zip" value="<%=rst1("zip")%>" size="10"></span></td>
      </tr>
      </table>
      </td>
  </tr>
  <tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Square Feet</span></td>
	  <td><span class="standard"><input type="text" name="sqft" value="<%=rst1("sqft")%>" size="10" maxlength="10"> SQFT</span></td>
  </tr>
  <tr bgcolor="#cccccc">
	  <td>&nbsp;</td>
	  <td>
	  <input type="hidden" name="owner" value="<%=rst1("clientid")%>">
    <input type="submit" name="action"  value="Update" class="standard">
    <input type="button" name="cancel" value="Cancel" onclick="nevermind();" class="standard"><input type="hidden" name="cid" value="<%=request("cid")%>">
    </td>
  </tr>
  
  <% else %>
  <tr bgcolor="#eeeeee"> 
    <td width="18%" align="right">
    <span class="standard">Building 
      Name<input type="hidden" name="bid" value="<%=rst1("id")%>">
    </span>
    </td>
    <td><span class="standard"><%=rst1("bldgname")%></span></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Address</span></td>
      <td><span class="standard"><%=rst1("address")%></span></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">City</span></td>
      <td><span class="standard"><%=rst1("city")%></span></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">State</span></td>
      <td>
      <table border=0 cellpadding="0" cellspacing="0">
      <tr valign="middle">
        <td><span class="standard"><span class="standard"><%=rst1("state")%></span></span></td>
        <td width="12"><span class="standard">&nbsp;</span></td>
        <td><span class="standard">Zip Code&nbsp;</span></td>
        <td><span class="standard"><%=rst1("zip")%></span></td>
      </tr>
      </table>
      </td>
  </tr>
  <tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Square Feet</span></td>
	  <td><span class="standard"><span class="standard"><%=rst1("sqft")%></span> SQFT</span></td>
  </tr>
  <tr bgcolor="#cccccc">
	  <td>&nbsp;</td>
	  <td>
	  <input type="hidden" name="owner" value="<%=rst1("clientid")%>">
    <input type="button" name="action"  value="Edit" onclick="goEdit();" class="standard">
    </td>
  </tr>
	</table>
	<%end if%>
	<% else %>
	<html>
	<head>
	<title>No Records Found</title>
  <link rel="Stylesheet" href="/genergy2/styles.css">
	</head>
	<body bgcolor="#ffffff">
	<p class="standard" style="margin:20px;">No records found</p>
  <%end if%>
</form>
</body>
</html>