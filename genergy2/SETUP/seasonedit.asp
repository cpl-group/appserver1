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

dim rid, seasonid
rid = secureRequest("rid")
seasonid = secureRequest("seasonid")

dim season,smonth,sday,syear,emonth,eday,eyear, effective_date
if trim(seasonid)<>"" then
	rst1.Open "SELECT * FROM rateseasons WHERE id='" & seasonid & "'", cnn1
	if not rst1.EOF then
		season = rst1("season")
		smonth = rst1("smonth")
		sday = rst1("sday")
		emonth = rst1("emonth")
		eday = rst1("eday")
		effective_date = rst1("effective_date")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Rate Seasons</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="seasonsave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
		<%if trim(seasonid)<>"" then%>
			Update Season
		<%else%>
			Add New Season
		<%end if%>
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Season</span></td> 
	<td><input type="text" name="season" value="<%=season%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Start Date</span></td>
	<td><input type="text" name="smonth" value="<%=smonth%>" maxlength="2" size="3">/<input type="text" name="sday" value="<%=sday%>" maxlength="2" size="3"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">End Date</span></td>
	<td><input type="text" name="emonth" value="<%=emonth%>" maxlength="2" size="3">/<input type="text" name="eday" value="<%=eday%>" maxlength="2" size="3"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Effective Date</span></td>
	<td><input type="text" name="effective_date" value="<%=effective_date%>" size="5"></td>
</tr>
<tr bgcolor="#dddddd"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if trim(seasonid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="button" name="action" value="Cancel" onclick="location='seasonView.asp?rid=<%=rid%>';" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="button" name="action" value="Cancel" onclick="location='seasonView.asp?rid=<%=rid%>';" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">
<input type="hidden" name="seasonid" value="<%=seasonid%>">

</form>
</body>
</html>






