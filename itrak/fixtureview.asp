<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<%
		
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getconnect(0,0,"engineering")
		
id=request.querystring("id")
bldg = request.querystring("bldg")

sqlstr = "select * from fixture_types where id='"& id &"'"
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then

%>
<title>New Fixture</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
<script language="JavaScript">
try{top.applabel("Lighting Catalogue - Edit / View Fixture");}catch(exception){}
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
<script src="messages.js" type="text/javascript" language="Javascript1.2"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="typeupd.asp" onsubmit="return checkfields(this);">
<table width="100%" cellpadding="3" cellspacing="1" border="0">
<tr align="right" bgcolor="#FFFFFF">
      <td colspan="2"><span class="standard">
        <input type="submit" name="submit" value="Update" class="standard">
        &nbsp;
        <input type="submit" name="submit" value="Delete" onClick="return confirmDelete();" class="standard">
        &nbsp;
        <input name="reset" type="reset" class="standard" onClick="location='fixtypesearch.asp?cid=<%=rst1("client")%>&bldg=<%=bldg%>'" value="Cancel">
        </span><span class="standard" style="color:#ffffff;"></span></td>
    </tr>
<tr><td colspan="2" bgcolor="#cccccc" height="18"><span class="standard"><b>Fixture</b></span></span>
<input type="hidden" name="CID" value="<%=rst1("client")%>" >
<input type="hidden" name="bldg" value="<%=bldg%>" >
</td></tr>

<tr valign="top" bgcolor="#eeeeee">
	<td style="padding-top:7px;" align="right"><span class="standard">Manufacturer</span></td>
	<td><span class="standard"><input type="text" name="manf" value="<%=rst1("manufacturer")%>" ><input type="hidden" name="fid" value="<%=rst1("id")%>""><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('manufacturer',event.x,event.y)"><img src="images/question.gif" width="13" height="13" hspace="4" border="0"></a></span></td>	  
</tr>
<tr valign="top" bgcolor="#eeeeee">
	<td style="padding-top:7px;" align="right"><span class="standard">Catalog Number</span></td>
	<td>
	<span class="standard">
  <input type="text" name="fixc" value="<%=rst1("fix_catalog")%>" size="50" >
  <a onMouseOut="closeHelpBox()" onMouseOver="helpbox('catalog_number',event.x,event.y)"><img src="images/question.gif" width="13" height="13" hspace="4" border="0"></a></span>
  </td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
	<td style="padding-top:7px;" align="right"><span class="standard">Fixture  Description</span></td>
	<td>
	<span class="standard"><textarea name="description"><%=rst1("description")%></textarea><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('fixture_description',event.x,event.y)"><img src="images/question.gif" width="13" height="13" hspace="4" border="0"></a>
	</span>
	</td>	  
</tr>

<tr><td colspan="2" bgcolor="#cccccc" height="18"><span class="standard"><b>Ballast</b></span></td></tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Ballast Type</span></td>
	<td><span class="standard"><input type="text" name="ballast" value="<%=rst1("ballast_type")%>" ><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('ballast_type',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a> &nbsp;Quantity:&nbsp;<input type="text" name="bqty" value="<%=rst1("ballast_qty")%>" onChange="checkNumber('bqty');" size="6"></span></td>	  
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Ballast 
        Life In Hours</span></td>
	<td><span class="standard"><input type="text" name="blife" value="<%=rst1("ballast_life")%>" onChange="checkNumber('blife');" ><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('ballast_life_in_hours',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	  
</tr>

<tr><td colspan="2" bgcolor="#cccccc" height="18"><span class="standard"><b>Lamp</b></span></td></tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Lamp Catalog Number</span></td>
	<td><span class="standard"><input type="text" name="lcnum" value="<%=rst1("lamp_catalog")%>" ><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('lamp_catalog_number',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a> &nbsp;Quantity:&nbsp;<input type="text" name="lqty" value="<%=rst1("lamp_qty")%>"  onChange="checkNumber('lqty');" size="6"></span></td>	  
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Lamp 
        Power (W)</span></td>
	<td><span class="standard"><input type="text" name="lwatts" value="<%=rst1("lamp_watts")%>"  onChange="checkNumber('lwatts');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('lamp_power',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	  
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Estimated 
        Lamp Life In Hours</span></td>
	<td><span class="standard"><input type="text" name="estLL" value="<%=rst1("est_lamp_life")%>"  onChange="checkNumber('estLL');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('estimated_lamp_life_in_hours',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	  
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Voltage 
        (V)</span></td>
	<td><span class="standard"><input type="text" name="volts" value="<%=rst1("volts")%>" onChange="checkNumber('volts');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('voltage',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	  
</tr>
<tr valign="top" bgcolor="#eeeeee">
	<td align="right" style="padding-top:7px;"><span class="standard"><b>General Remarks</b></span></td>
	<td>
	<span class="standard"><textarea name="remarks"><%=rst1("remarks")%></textarea><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('general_remarks',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span>
	</td>	  
</tr>
<tr bgcolor="#cccccc">
	<td align="right"><span class="standard"></span></td>
	  <td>&nbsp;</td>	  
</tr>
</table>


</form>
<!--#INCLUDE FILE="helpbox.htm"-->
</body>
</html>
<%END IF%>
