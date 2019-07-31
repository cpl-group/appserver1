<%option explicit%>
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
		dim user
		user=Session("name")

dim bldg, fid
bldg=request.querystring("bldg")
fid=request.querystring("fid")

dim floor, sqft

Set rst1 = Server.CreateObject("ADODB.recordset")
Set cnn1 = Server.CreateObject("ADODB.connection")
cnn1.Open getconnect(0,0,"engineering")
if trim(fid)<>"" then 'if coming for update not new node
	dim sqlstr, rst1, cnn1
	
	sqlstr= "select * from floor where bldg='"&bldg&"' and id="&fid
	rst1.Open sqlstr, cnn1
	if not rst1.eof then
	floor = rst1("floor")
	sqft =  rst1("sqft")
	end if
	rst1.close
end if
sqlstr = "select bldgname from facilityinfo where id=" & bldg
rst1.open sqlstr, cnn1
if not rst1.eof then
%>
<html>
<head>

<title>New Room</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<%if trim(fid) = "" then %>
try{top.applabel("Floor Management - Add Floor");}catch(exception){}
<%else%>
try{top.applabel("Floor Management - Edit Floor Setup");}catch(exception){}
<%end if%>
<!--
function checkfields(theform){
  retval = true;
  for (i=0;i<theform.length;i++){
    if (theform.elements[i].value.indexOf("'") > -1) {
      theform.elements[i].value = theform.elements[i].value.replace(/'/g,"''" );
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

function confirmDelete(){
  retval = window.confirm("Are you sure you want to delete this item?");
  return retval;
}

//-->
</script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="savefloor.asp" onsubmit="return checkfields(this);">
  <table width="100%" cellpadding="3" cellspacing="1" border="0">
  <tr align="right" bgcolor="#FFFFFF">
      <td colspan="2"><font face="Arial, Helvetica, sans-serif"><span class="standard">
        <input type="hidden" name="fid" value="<%=fid%>">
        <input type="hidden" name="bldg" value="<%=bldg%>">
        <%if trim(fid)<>"" then%>
        <input type="submit" name="Submit" value="Update" class="standard">
        <input type="submit" name="Submit" value="Delete" onClick="return confirmDelete();" class="standard">
        <%else%>
        <input type="submit" name="Submit" value="Save" class="standard">
        <%end if%>
        <input name="reset" type="reset" class="standard" onClick="location='floorsearch.asp?bldg=<%=bldg%>'" value="Cancel">
        </span></font><font face="Arial, Helvetica, sans-serif" color="#ffffff" size="2"><span class="standard"></span></font></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right" width="35%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Building Name</span></font></td>
      <td width="65%"><font face="Arial, Helvetica, sans-serif"><span class="standard"> 
        <%=rst1("bldgname")%></span></font></td>
	</tr>
	<tr bgcolor="#eeeeee">
      <td align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Floor Name</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><span class="standard"> 
        <input type="text" name="floor" value="<%=floor%>">
        </span></font></td>
	</tr>
	<tr bgcolor="#eeeeee">
      <td align="right"><font face="Aral, Helvetica, sans-serif" size="2"><span class="standard">SQFT</span></font></td>
      <td> <font face="Arial, Helvetica, sans-serif"><span class="standard"> 
        <input type="text" name="sqft" value="<%=sqft%>" onChange="checkNumber('sqft');">
        </span></font></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td align="right"><font face="Aral, Helvetica, sans-serif" size="2"><span class="standard"></span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><span class="standard"> </span></font></td>
     </tr>
  </table>
	
  
</form>
<%else%>
	<html>
	<head>
	<title>No Records Found</title>
  <link rel="Stylesheet" href="/genergy2/styles.css">
	</head>
	<body bgcolor="#ffffff">
	<p class="standard" style="margin:20px;">Unable to complete your request: the building record was not found. It is possible that the building you are trying to modify has been deleted in the Facilities Manager but has not been removed from the node tree.</p>
<% end if%>
</body>
</html>
