<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
sub closewindow()
	%>
	<script>
		window.close();
	</script>
	<%
	response.end
end sub

if 	not( _
	checkgroup("Genergy Users")<>0 _
	or checkgroup("clientOperations")<>0 _
	) then%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, sql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim id, action, ratecode, description
id = trim(request("id"))
action = trim(request("action"))
ratecode = trim(request("ratecode"))
description = trim(request("description"))
if trim(action)<>"" then
	if trim(action)="Add Point" then
		sql = "INSERT INTO RateCodes (RateCode, Description) VALUES ('"&RateCode&"', '"&Description&"')"
	elseif trim(action)="Update Point" then
		sql =  "UPDATE RateCodes SET RateCode='"&RateCode&"', Description='"&Description&"' WHERE id="&id
	elseif trim(action)="Delete Point" then
		sql =  "DELETE FROM RateCodes WHERE id='"&id&"'"
	end if
  'Logging Update
'  response.write sql
'  response.end
  logger(sql)
  'end Log
	if sql<>"" then cnn1.execute sql
  ratecode = ""
  description = ""
  id = ""
end if
%>
<html>
<head>
<title>OUC Rate Code Entry</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
</head>
<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="oucRateCode.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">OUC Rate Codes</span></td>
</tr>
</table>
<%
dim rowcolor
rst1.open "SELECT * FROM ratecodes", cnn1
if not rst1.EOF then%>
<!-- <table width="100%" border="0" cellpadding="3" cellspacing="0">
</table> -->
<div style="height:165;overflow:auto;border-bottom:1px solid #cccccc;">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#dddddd">
	<td><span class="standard"><b>Rate&nbsp;Code</b></span></td>
	<td><span class="standard"><b>Description</b></span></td>
	<td>&nbsp;</td>
</tr>
<%do until rst1.EOF
      if id=trim(rst1("id")) then
        rowcolor = "#ccffcc"
        ratecode = rst1("ratecode")
        description = rst1("description")
      else
        rowcolor = "white"
      end if
      %>
  		<tr bgcolor="<%=rowcolor%>" onmouseover="this.style.backgroundColor='lightgreen'" onmouseout="this.style.backgroundColor='<%=rowcolor%>'" onclick="document.location='oucRateCode.asp?id=<%=rst1("id")%>';">
  			<td width="5%"><span class="standard"><%=rst1("ratecode")%>&nbsp;</span></td>
  			<td width="5%"><span class="standard"><%=rst1("description")%>&nbsp;</span></td>
  			<td width="70%">&nbsp;</td>
  		</tr>
  		<%
      rst1.movenext
loop
  %>
</table>
</div>
<!-- </table> -->
<%end if
rst1.close%>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #ffffff;">
<tr>
  <td>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr><td align="right">Rate&nbsp;Code&nbsp;</td>
      <td><input type="text" name="ratecode" size="10" value="<%=ratecode%>"></td></tr>
  <tr><td align="right">Description&nbsp;</td>
      <td><textarea cols="15" rows="3" name="description"><%=description%></textarea></td></tr>
  <tr>
    <td>&nbsp;</td>
    <td>
    <%if id<>"" then%>
    <input type="submit" name="action" value="Update Point" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    <input type="submit" name="action" value="Delete Point" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    <input type="button" value="Cancel" style="border:1px outset #ddffdd;background-color:ccf3cc;" onclick="document.location='oucRateCode.asp';">
    <%else%>
    <input type="submit" name="action" value="Add Point" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    <%end if%>
    <input type="hidden" name="id" value="<%=id%>">
    </td>
  </tr>
  </table>
  </td>
</tr>
</table>

</form>
</body>
</html>