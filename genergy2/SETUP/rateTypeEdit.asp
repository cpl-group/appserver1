<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, rtid
rid = secureRequest("rid")
rtid = secureRequest("rtid")

dim rtype,rcity
if trim(rtid)<>"" then
	rst1.Open "SELECT regions.city as rcity,type FROM ratetypes join regions on ratetypes.regionid=regions.id WHERE ratetypes.id='" & rtid & "'", cnn1
	if not rst1.EOF then
		rtype = rst1("type")
		rcity = rst1("rcity")
	end if
	rst1.close
else
  rst1.open "select city as rcity from regions where id='" & rid & "'", cnn1
  if not rst1.EOF then rcity = rst1("rcity")
  rst1.close
end if

%>
<html>
<head>
<title>Edit Rate Type</title>
<script>
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="rateTypeSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#000000">
	<td colspan="2"><span class="standardheader">
    <a href="index.asp" target="main"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0"></a> Utility Manager Setup
	</span></td>
</tr>
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
		<%if trim(rtid)<>"" then%>
			Update Rate Type | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=rcity%> Region</a>
		<%else%>
			Add New Rate Type | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=rcity%> Region</a>
		<%end if%>
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td width="30%" align="right"><span class="standard">Rate Type</span></td>
	<td width="70%"><input type="text" name="rtype" value="<%=rtype%>"></td>
</tr>
<tr bgcolor="#eeeeee"> 
	<td style="border-bottom:1px solid #999999;"><span class="standard">&nbsp;</span></td>
	<td style="border-bottom:1px solid #999999;">
		<%if trim(rtid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="button" name="action" value="Cancel" onclick="history.go(-1);" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="button" name="action" value="Cancel" onclick="history.go(-1);" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%end if%><span class="standard"><br>&nbsp;</span>
	</td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">
<input type="hidden" name="rtid" value="<%=rtid%>">


</form>
</body>
</html>
