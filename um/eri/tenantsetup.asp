<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">

<script>
function openpopup(){
//configure "Open Logout Window
parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}
</script>

<script>
function addnew(){
// Load Add new tenant form
var bldg=document.choosebldg.bldg.value
//document.frames.title.location.href="null.htm";
document.frames.piclist.location.href="ti_add.asp?bldg=" + bldg;
document.frames.info.location.href="tenantlist.asp?bldg=" + bldg;
}
function lookup(){
	document.frames.title.location.href="title.asp?bldg="+document.forms[0].bldg.value
}
function  open_bldginfo(){
  document.frames.title.location.href = "buildinginfo.asp?bldgnum="+document.forms[0].bldg.value;
}
function open_download(){
  document.frames.title.location.href = "eri_download.asp?bldgnum="+document.forms[0].bldg.value+"&pid="+document.forms[0].portfolio.value;
  document.frames.info.location.href = "/null.htm";
}
</script>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>

</head>
<body bgcolor="#eeeeee" text="#000000">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr> 
    <td bgcolor="#6699cc"><span class="standardheader">ERI Manager | Tenant Setup</span></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr bgcolor="#eeeeee"> 
    <td valign="top" nowrap style="border-bottom:1px solid #cccccc;"> 
      <%
		dim cnn1, rst1, sqlStr
		
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.open getConnect(0,0,"engineering")
		sqlStr = "select * from buildings order by strt"
		rst1.Open sqlStr, cnn1
		%>
      <form target=title method="POST" action="tenantsetup.asp" name="choosebldg">
          <select name="bldg">
            <%
		  while not rst1.EOF
		  %><option <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=rst1("bldgnum")%>"><%=rst1("strt")%></option><%
		  rst1.movenext
		  wend  
		  rst1.close  
		  set cnn1 = nothing
		  %>
          </select>
          <input type="button" name="Submit" value="Lookup" onclick="lookup()">
		  <input type="button" name="bldgadd" value="Edit Building" onclick="open_bldginfo()"> 
    </td>
	<td align="right" style="border-bottom:1px solid #cccccc;"><a href="#" onclick="open_download()">Select Download from</a>
		<select name="portfolio">
			<%
			rst1.open "select distinct management, Owner_id FROM buildings WHERE Owner_id<>'999' ORDER BY management"
			do until rst1.eof
				%><option value="<%=rst1("Owner_id")%>"><%=rst1("management")%> (<%=rst1("Owner_id")%>)</option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select>
	</td>
  </tr>
  <tr>
    <td style="border-top:1px solid #ffffff;" colspan="2">
<IFRAME name="title" width="100%" height="240" src="null.htm" scrolling="no" frameborder=0 border=0></IFRAME> 
<!--[[IFRAME name="piclist" width="100%" height="120" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0]][[/IFRAME]] -->
<IFRAME name="info" width="100%" height="210" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
    </td>
  </tr>
</table>
</form>
</body>
</html>
