<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>G1 Console QA Module</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		

<script>
function viewconsole(version){
        
	var userid 
	if (version == "" || version == null || version == "undefined")
	{
	    alert("Please select a console version and try again");
	    return false;
	}
	userid = document.chooseversion.userid.value;
	document.frames.info.location.href="launchconsole.asp?version=" + version + "&userid=" + userid;
	/*if (consolelink.indexOf("/g1_clients/index2.asp")>=0)
	{
		document.frames.info.location.href = "/genergy2_intranet/itservices/g1console/usrdetail.asp?userid=" + userid 
	}
	else
	{
	document.frames.info.location.href = "/um/security/usrdetail.asp?username=" + userid 	
	
	}
	//TOOK OUT COMMENTED BLOCK IN ORDER TO ACCEPT ALL CONSOLELINK AND NOT JUST INDEX2.ASP.  MICHELLE T.
	*/
}
function edit_client(){
	var userid, consolelink
	userid = document.chooseversion.userid.value;
	consolelink = (document.chooseversion.consolelink.value).toString();
	document.frames.info.location.href = "/genergy2_intranet/itservices/g1console/usrdetail.asp?userid=" + userid 
	
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

</script>

</head>
<% 	
'2/6/2009 N.Ambo restricted list of clients to show only clients which are active

lid = request("lid")	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set userdetails = Server.CreateObject("ADODB.recordset")
Set rs = Server.CreateObject("ADODB.recordset")

cnn1.open getConnect(0,0,"dbCore")

strsql = "SELECT  username, company, initial_page FROM clients where isActive = 1 order by company, username"

rs.Open strsql, cnn1, adOpenStatic
if not rs.EOF then 

%>
<script language="javascript">
var uidlist = new Array()
var linklist = new Array()
 
<% 		x = 0	
		while not rs.EOF 
			userid = rs("username")
		%>
		uidlist[<%=x%>] = '<%=userid%>' 
		linklist[<%=x%>] = '<%=trim(rs("initial_page"))%>'
<% 
		x=x+1
		rs.movenext
		wend
		rs.movefirst
%>
function updateuid(){
	document.chooseversion.userid.value = uidlist[document.chooseversion.uidselect.selectedIndex-1]
	document.chooseversion.consolelink.value = linklist[document.chooseversion.uidselect.selectedIndex-1]
}
</script>
<% end if %>
<body bgcolor="#eeeeee" text="#000000">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
	  <tr> 
		
    <td bgcolor="#666699"><span class="standardheader">Console Manager</span></td>
	  </tr>
	</table>
  <form method="POST" name="chooseversion">
	
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> Enter Username then select a console version: </td>
    </tr>
    <tr> 
      <td style="border-top:1px solid #ffffff;">  
      <input type="text" name="userid">
        <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <%
		if not rs.EOF then 
		%>
        <select name="uidselect" onChange="updateuid()">
          <option value="">NONE</option>
          <% 
		while not rs.eof 
		%>
          <option value="<%=rs("username")%>"><%=rs("company")%> (<%=rs("username")%>)</option>
          <% 
		rs.movenext
		wend
		%>
        </select>
        <%
			end if
		%>
        <input name="consolelink" type="text" size="50" style="background-color:#eeeeee" disabled>
        </span> </td>
    </tr>
    <tr> 
      <td width="21%" style="border-top:1px solid #ffffff;"> 
	  	<input type="button" name="Submit" value="Launch Console" onClick="return viewconsole(document.chooseversion.consolelink.value)">
        <% if checkgroup("IT Services,GenergyCorporateExec") then %>
        <input type="button" name="editclient" value="Edit Client Setup" onClick="edit_client()">
        <input type="button" name="newuser" value="Add New Account" onClick="document.frames.info.location.href='/genergy2_intranet/itservices/g1console/usrdetail.asp'"> 
        <% end if %>
      </td>
    </tr>
  </table>
  </form>
<IFRAME name="info" width="100%" height="90%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0> </IFRAME> 
</body>
</html>
