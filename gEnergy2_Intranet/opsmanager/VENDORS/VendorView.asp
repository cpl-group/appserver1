<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

dim company, vendor, name, address_1, address_2, city, state, zip, telephone, fax_number
company = trim(request("company"))
vendor = trim(request("vendor"))
if company="" then company = "EM"
%>
<html>

<title>Vendor Search</title>
<script language="JavaScript" type="text/javascript">
//<!--

function customerdetail(cid) {
  theURL="VendorEdit.asp?vendor=" + cid + "&company=" + '<%=company%>'
  openwin(theURL,600,475)
}

function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function screencompany(company) {
    document.location.href="cis_manage.asp?company="+company	
}

//display quickhelp
var helpIsOn = 0;
function toggleHelp(){
  if (helpIsOn) { 
    document.all.quickhelptext.style.display='none';
    helpIsOn = 0;
   } else { 
    document.all.quickhelptext.style.display='inline';
    helpIsOn = 1;
   }
}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}

//-->
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#eeeeee">
<form name="form1">
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #cccccc;">
<tr>
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp;<b>Vendors</b></td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <!-- <img src="/um/opslog/images/btn-new_contact.gif" alt="New Contact" border="0" onclick="opencontact('new');" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"> -->
  <input type="button" name="newvendor" value="New Vendor" onclick="openwin('VendorEdit.asp?company=<%=company%>',400,300)">
  <img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr>
	<td style="border-top:1px solid #ffffff;"> 
    	<select name="company" onchange="this.form.submit()">
		<%
        rst.Open "select * from companycodes where active = 1 and code <> 'AC' order by name", cnn
        if not rst.eof then
			if company="" then company = rst("code")
	        do until rst.eof
	        %><option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><%=rst("name")%></option><%
	        rst.movenext
	        loop
        end if
        rst.close%>
		</select>
	</td>
	<td align="right" style="border-top:1px solid #ffffff;"><a href="javascript:toggleHelp();" style="text-decoration:none;"><img src="/gEnergy2_Intranet/opsmanager/joblog/images/quick_help.gif" align="absmiddle" border="0">&nbsp;<b>Quick Help</b></a></td>
</tr>
<tr valign="top">
	<td colspan="2" height="255">
<!-- 	<div id="quickhelptext" style="display:none;">
<ul>
<li>Click the radio button next to a company to show its customers. Customers and contacts are maintained separately for all entities.
</ul>
</div>
-->
<div id="customers" style="overflow:auto;width:100%;height:245px;border:1px solid #cccccc;">
<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc"><%
rst.Open "SELECT distinct vendor, name FROM " & company & "_MASTER_APM_VENDOR order by name", cnn
do until rst.eof
	%><tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onclick="customerdetail('<%=rst("vendor")%>');"><td><%=rst("name")%></td></tr><%
	rst.movenext
loop%>
</table>  
</div>

  </td>
</tr>
<tr>
  <td style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;">
&nbsp;  <!-- <img src="/um/opslog/images/btn-new_contact.gif" alt="New Contact" border="0" onclick="opencontact('new');" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"> -->
  </td>
  <td align="right" style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
</form>
</body>
</html>