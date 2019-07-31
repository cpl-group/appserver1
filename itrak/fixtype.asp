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
cid=request.querystring ("cid")	
%>

<title>New Fixture</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">

<script language="JavaScript">
try{top.applabel("Lighting Catalogue - New Fixture");}catch(exception){}

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
//-->
</script>
<script src="messages.js" type="text/javascript" language="Javascript1.2"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="savefixtype.asp" onsubmit="return checkfields(this);">
<table width="100%" border="0" cellpadding="3" cellspacing="1" >
  <tr> 
    <td width="68%" align="right" bgcolor="#FFFFFF"> 
	  <input type="hidden" name="cid" value="<%=cid%>">
        <span class="standard">
        <input type="submit" name="choice22"  value="Save" class="standard">
        &nbsp;
        <input name="reset" type="reset" class="standard" onClick="location='fixtypesearch.asp?cid=<%=cid%>'" value="Cancel">
        </span> </td>
 	</tr>
</table>

  <table width="100%" cellpadding="3" cellspacing="1" border="0">
	<tr><td colspan="2" bgcolor="#cccccc" height="18"><span class="standard"><b>Fixture</b></span></td></tr>
    
	<tr bgcolor="#eeeeee">
      <td align="right"><span class="standard">Manufacturer</span></td>
      <td><span class="standard"><input type="text" name="manf"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('manufacturer',event.x,event.y)"><img src="images/question.gif" width="13" height="13" hspace="4" border="0"></a></span></td>	  
	</tr>
	<tr valign="top" bgcolor="#eeeeee">
      <td align="right" style="padding-top:7px;" height="49"><span class="standard">Fixture 
        Catalog Number</span></td>
      <td height="49"> <span class="standard"> 
        <input type="text" name="fixc" size="50"><a onMouseOut="helpup('help_catalog_number');" onMouseOver="helpdrop('help_catalog_number','catalog_number');"><img src="images/question-rt.gif" width="22" height="13" hspace="4" border="0" name="help_catalog_number_img"></a>
      <div id="help_catalog_number" class="standard" style="display:'none'; margin-right:40px;padding-top:6px;padding-bottom:6px;"></div>    
      </span>
      </td>	    
    </tr>
	<tr valign="top" bgcolor="#eeeeee">
      <td align="right" style="padding-top:7px;"><span class="standard">Fixture Description</span></td>
      <td><span class="standard"><textarea name="description"></textarea><a onMouseOut="helpup('help_fixture_description');" onMouseOver="helpdrop('help_fixture_description','fixture_description');"><img src="images/question-rt.gif" alt="?" title="" width="22" height="13" hspace="4" border="0" name="help_fixture_description_img"></a></span>
      <div id="help_fixture_description" class="standard" style="display:'none'; margin-right:40px;padding-top:6px;padding-bottom:6px;"></div>    
      </td>	    
    </tr>
	<tr><td colspan="2" bgcolor="#cccccc" height="18"><span class="standard"><b>Ballast</b></span></td></tr>
	<tr bgcolor="#eeeeee">
      <td width="35%" align="right"><span class="standard">Ballast Type</span></td>
      <td width="65%"><span class="standard"><input type="text" name="ballast"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('ballast_type',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a> &nbsp;Quantity <input type="text" name="ballastqty" size="6" onChange="checkNumber('ballastqty');"></span></td>	    
    </tr>
	<tr bgcolor="#eeeeee">
      <td align="right"><span class="standard">Ballast 
        Life In Hours</span></td>
      <td><span class="standard"><input type="text" name="blife" onChange="checkNumber('blife');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('ballast_life_in_hours',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	    
    </tr>
	<tr><td colspan="2" bgcolor="#cccccc" height="18"><span class="standard"><b>Lamp</b></span></td></tr>
	<tr bgcolor="#eeeeee">
		<td width="35%" align="right"><span class="standard">Lamp Catalog Number</span></td>
		<td width="65%"><span class="standard"><input type="text" name="lcnum"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('lamp_catalog_number',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a> &nbsp;Quantity: <input type="text" name="lqty" size="6" onChange="checkNumber('lqty');"></span></td>	    
	</tr>
	<tr bgcolor="#eeeeee">
		
      <td align="right"><span class="standard">Lamp 
        Power (W)</span></td>
		<td> <span class="standard"> <input type="text" name="lwatts" onChange="checkNumber('lwatts');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('lamp_power',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	    
	</tr>
	<tr bgcolor="#eeeeee">
		
      <td align="right"><span class="standard">Average 
        Lamp Life In Hours</span></td>
		<td> <span class="standard"><input type="text" name="estLL" onChange="checkNumber('estLL');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('estimated_lamp_life_in_hours',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	    
	</tr>
	<tr bgcolor="#eeeeee">
		
      <td align="right"><span class="standard">Voltage 
        (V) </span></td>
		<td><span class="standard"><input type="text" name="volts" onChange="checkNumber('volts');"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('voltage',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></span></td>	    
	</tr>
	<tr valign="top" bgcolor="#eeeeee">
		<td align="right" style="padding-top:7px;"><span class="standard"><b>General Remarks</b></span></td>
		<td><span class="standard"><textarea name="remarks"></textarea><a onMouseOut="helpup('help_general_remarks');" onMouseOver="helpdrop('help_general_remarks','general_remarks');"><img src="images/question-rt.gif" alt="?" title="" width="22" height="13" hspace="4" border="0" name="help_general_remarks_img"></a></span>
    <div id="help_general_remarks" class="standard" style="display:'none'; margin-right:40px;padding-top:6px;padding-bottom:6px;"></div>    
    </td>	    
	</tr>
	<tr bgcolor="#cccccc">
		<td align="right"><span class="standard"><b></b></span></td>
		
      <td><span class="standard"> </span></td>	    
	</tr>
	</table>
	
</form>
<!--#INCLUDE FILE="helpbox.htm"-->
</body>
</html>
