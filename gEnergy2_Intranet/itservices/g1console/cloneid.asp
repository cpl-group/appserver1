<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<% 
targetid = request("target")
sourceid = request("source")

if trim(targetid) <> "" and trim(sourceid) <> "" then 
		Set cnn = Server.CreateObject("ADODB.Connection")
		set cmd = server.createobject("ADODB.Command")
		cnn.open getConnect(0,0,"dbCore")
		cnn.CursorLocation = adUseClient
		cmd.CommandText = "sp_cloneid"
		cmd.CommandType = adCmdStoredProc
		
		Set prm = cmd.CreateParameter("userid", adVarChar, adParamInput, 100)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("useridTOclone", adVarChar, adParamInput, 100)
		cmd.Parameters.Append prm
		cmd.Name = "clone"
		Set cmd.ActiveConnection = cnn

		cnn.clone targetid, sourceid
		set cnn = nothing
		%>
		<script>
		opener.window.location = opener.window.location
		window.close()
		</script>	
<%
else
%>
<html>
<head>
<title>Clone gEnergyOne User</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
</head>
<script>
function confirmclone(){
	var target
	var source
	
	target = document.chooseclone.target.value
	source = document.chooseclone.source.value
	  if (confirm("Are you sure you want to clone the setup for user " +source+ " to user " +target+ "?")) {
              document.location= "cloneid.asp?target="+target+"&source="+source
       }
}

</script>
<body bgcolor="#eeeeee" text="#000000">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
	  <tr> 
		
    <td bgcolor="#6699cc"><span class="standardheader">Clone User Accounts</span></td>
	  </tr>
	</table>
  <form method="POST" name="chooseclone" action="cloneid.asp">
	
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> Source USERID</td>
      <td valign="top" style="border-bottom:1px solid #cccccc;"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
		<% 	
		uid = request("uid")	
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set userdetails = Server.CreateObject("ADODB.recordset")
		Set rs = Server.CreateObject("ADODB.recordset")
		
		cnn1.open getConnect(0,0,"dbCore")
		
		strsql = "SELECT  username, company, initial_page FROM clients where initial_page like '%index2.asp' order by company, username"
		
		rs.Open strsql, cnn1, adOpenStatic
		
		if not rs.EOF then 
		%>
        <select name="source">
          <option value="">NONE</option>
          <% 
		while not rs.eof 
		%>
          <option value="<%=trim(rs("username"))%>"><%=rs("company")%> (<%=rs("username")%>)</option>
          <% 
		rs.movenext
		wend
		rs.movefirst
		%>
        </select>
        <%
			end if
		%>
        </span></td>
    </tr>
    <tr> 
      <td style="border-top:1px solid #ffffff;">Target USERID</td>
      <td style="border-top:1px solid #ffffff;"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">
        <%
		if not rs.EOF then 
		%>
        <select name="target">
          <option value="">NONE</option>
          <% 
		while not rs.eof 
		%>
          <option value="<%=trim(rs("username"))%>" <%if trim(rs("username")) = trim(uid) then %>selected<%end if%>><%=rs("company")%> (<%=rs("username")%>)</option>
          <% 
		rs.movenext
		wend
		rs.close
		%>
        </select>
        <%
			end if
		%>
        </span></td>
    </tr>
    <tr> 
      <td width="21%" colspan="2" style="border-top:1px solid #ffffff;"> <input type="button" name="Button" value="Clone Account" onclick="confirmclone()">
        <input type="button" name="close" value="Cancel" onClick="window.close()"> 
      </td>
    </tr>
  </table>
  </form>
<IFRAME name="info" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0> </IFRAME> 
</body>
</html>
<% end if %>