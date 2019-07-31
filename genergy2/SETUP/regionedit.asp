<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql, rst2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, copy
copy = false
rid = secureRequest("rid")
if trim(rid)="copy" then
	copy = true
	rid = ""
end if

dim city, citycode, action
if trim(rid)<>"" then
	rst1.Open "SELECT * FROM regions WHERE id='"&rid&"'", cnn1
	if not rst1.EOF then
		city = rst1("city")
		citycode = rst1("city_code")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Region View</title>
<script>
function rateTypeEdit(rtid)
{	document.location.href = 'rateTypeEdit.asp?rid=<%=rid%>&rtid='+rtid;
}
function seasonEdit(seasonid)
{	document.location.href = 'seasonedit.asp?rid=<%=rid%>&seasonid='+seasonid;
}
function ratePeakEdit(rPid)
{	document.location.href = 'ratePeakEdit.asp?rid=<%=rid%>&rPid='+rPid;
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="regionsave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#000000">
<%
dim showWeirdBlackBar
showWeirdBlackBar = false
if allowGroups("Genergy Users") AND showWeirdBlackBar then
%>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><span class="standardheader"><a href="index.asp" target="main" class="breadcrumb" style="text-decoration:none;"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0">Utility Manager Setup</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="frameset.asp" target="main" class="breadcrumb" style="text-decoration:none;">Update Meters</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="portfolioview.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Portfolios</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="regionView.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Rates</a></span></td>
  </tr>
  </table>
<%end if%>
  </td>
</tr>
<tr bgcolor="#3399cc">
  <td style="border-bottom:1px solid #ffffff;"><span class="standard" style="color:#ffffff">&nbsp;<a href="regionView.asp" style="color:#ffffff;font-weight:bold;text-decoration:none;">Rate Setup</a> | <%=city%> Region | <a href="seasonView.asp?rid=<%=rid%>" style="color:#ffffff">Seasons &amp; Rate Peaks</a> | <a href="rateTypeView.asp?rid=<%=rid%>" style="color:#ffffff">Rate Types</a></span></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#eeeeee">
	<td valign="top" style="border-bottom:1px solid #999999;padding-left:12px;">
	<span class="standard"><b>
  <%if copy then%>
    Copy Region:
  <%elseif trim(rid)<>"" then%>
    Update Region:
  <%else%>
    Add Region:
  <%end if%>
	</b></span>
	</td>
	<td style="border-bottom:1px solid #999999;">
  <table border=0 cellpadding="3" cellspacing="0">
<%if copy then%>
  <tr>
    <td align="right"><span class="standard">Copy City</span></td> 
    <td>
		<select name="copyfrom">
		<%
			rst1.open "SELECT * FROM regions ORDER BY city", cnn1
				do until rst1.eof
					response.write "<option value="""&rst1("id")&""">"&rst1("city")&"</option>"
					rst1.movenext
				loop
			rst1.close
		%>
		</select>
	</td>
  </tr>
<%end if%>
  <tr>
    <td align="right"><span class="standard">City</span></td> 
    <td><input type="text" name="city" value="<%=city%>"></td>
  </tr>
  <tr>
    <td align="right"><span class="standard">City Code</span></td>
    <td><input type="text" name="citycode" value="<%=citycode%>"></td>
  </tr>
  <tr>
    <td></td>
    <td><%if copy then%>
			<input type="submit" name="action" value="Copy" class="standard" style="width:90px;background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%elseif trim(rid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="width:90px;background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="width:90px;background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%end if%>
    </td>
  </tr>
  </table><br>
  <span class="standard">

<!--
  [[ol type="1"]]
  [[li]]Enter Rate Seasons
  [[li]]Enter Rate Peaks[[br]]For each season, define peak and off-peak time periods
  [[li]]Add Rate Type[[br]]Rate types hold the rates that will be applied based on rate peak
  [[li]]Enter Rate.
  [[/ol]]
-->
  </span>
	</td>
</tr>
<tr>
  <td colspan="2" bgcolor="#dddddd"><input type="button" value="Set Up Rate Seasons" onclick="document.location='seasonView.asp?rid=<%=rid%>';" class="standard">&nbsp;<input type="button" value="Set Up Rate Types" onclick="document.location='rateTypeView.asp?rid=<%=rid%>';" class="standard">&nbsp;<input type="button" value="Set Up Holidays" onclick="document.location='holidayView.asp?rid=<%=rid%>';" class="standard"><input type="button" value="Rate Builder" onclick="document.location='ratebuilder/editcomponents.asp';" class="standard"><input type="button" value="Fuel Sheet Adjustments" onclick="document.location='ratebuilder/monthlyadjustments.asp';" class="standard"></td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">

</form>
</body>
</html>
