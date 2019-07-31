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
dim bldg, fid
bldg=request.querystring("bldg")
fid=request.querystring("fid")
rid=request.querystring("rid")
fl=request.querystring("floor")	
rid=request.querystring("rid")
room=request.querystring("room")


dim floor, sqft

if trim(rid)<>"" then 'if coming for update not new node
	dim sqlstr, rst1, cnn1
	Set cnn1 = Server.CreateObject("ADODB.connection")
	Set rst1 = Server.CreateObject("ADODB.recordset")
	cnn1.Open getconnect(0,0,"engineering")
	
	sqlstr= "select * from room where id="&rid
	rst1.Open sqlstr, cnn1
	name = rst1("room")
	sqft =  rst1("sqft")
	'est=rst1("est_hr_wk")
end if
%>

<title>New Room</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
<script language="JavaScript">
<%if trim(rid)<>"" then%>
try{top.applabel("Floor Management - View <%=name%> on <%=fl%> floor");}catch(exception){}
<%else%>
try{top.applabel("Floor Management - Add New Room on <%=fl%> floor");}catch(exception){}
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

</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="saveroom.asp" onsubmit="return checkfields(this);">
  <table width="100%" cellpadding="3" cellspacing="1" border="0">
    <tr bgcolor="#eeeeee">
      <td align="right" bgcolor="#ffffff" colspan=2><span class="standard"> 
        <%if trim(rid)<>"" then%>
        <input type="submit" name="Submit" value="Update" class="standard">
        <input type="submit" name="Submit" value="Delete" onClick="return confirmDelete();" class="standard">
        <%else%>
        <input type="submit" name="Submit" value="Save" class="standard">
        <%end if%>
        <input name="reset" type="reset" class="standard" onClick="location='roomsearch.asp?bldg=<%=bldg%>&fid=<%=fid%>'" value="Cancel">
        </span></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td width="35%" align="right"><span class="standard">Floor</span></td>
      <td width="65%"> <span class="standard"> 
        <input type="hidden" name="fid" value="<%=fid%>">
        <input type="hidden" name="rid" value="<%=rid%>">
        <input type="hidden" name="floor" value="<%=fl%>">
        <input type="hidden" name="bldg" value="<%=bldg%>">
        <%=fl%></span></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Room Name</span></td>
      <td><span class="standard"> 
        <input type="text" name="room" value="<%=name%>">
        </span></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><font face="Aral, Helvetica, sans-serif" size="2"><span class="standard">SQFT</span></td>
      <td><span class="standard"> 
        <input type="text" name="sqft" value="<%=sqft%>" onChange="checkNumber('sqft');">
        </span></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td></td>
      <td>&nbsp; </td>
    </tr>
  </table>
	
  
</form>
</body>
</html>
