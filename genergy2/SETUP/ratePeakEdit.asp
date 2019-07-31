<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, rPid
rid = secureRequest("rid")
rPid = secureRequest("rPid")

dim sweekday,stime,eweekday,etime, seasonid, label, peakname
if trim(rPid)<>"" then
	rst1.Open "SELECT * FROM ratepeak WHERE id='" & rPid & "'", cnn1
	if not rst1.EOF then
		sweekday = rst1("sweekday")
		stime = rst1("stime")
		eweekday = rst1("eweekday")
		etime = rst1("etime")
		seasonid = rst1("seasonid")
		label = rst1("label")
		peakname = rst1("peakname")
	end if
	rst1.close
end if

dim city
if trim(rid)<>"" then
	rst1.Open "SELECT * FROM regions WHERE id='"&rid&"'", cnn1
	if not rst1.EOF then
		city = rst1("city")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Edit Rate Peak</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="ratePeakSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#000000">
	<td colspan="2"><span class="standardheader">
    <a href="index.asp" target="main"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0"></a> Utility Manager Setup
	</span></td>
</tr>
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
		<%if trim(rPid)<>"" then%>
			Update Rate Peak | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=city%> Region</a>
		<%else%>
			Add New Rate Peak | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=city%> Region</a>
		<%end if%>
	</span></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Rate Peak Label</span></td>
	<td><input type="text" name="label" value="<%=label%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Peak Type</span></td>
	<td>
		<select name="peakname">
			<option value="1" <%if trim(peakname)="1" then response.write "SELECTED"%>>On Peak</option>
			<option value="2" <%if trim(peakname)="2" then response.write "SELECTED"%>>Off Peak</option>
			<option value="3" <%if trim(peakname)="3" then response.write "SELECTED"%>>Int Peak</option>
			<option value="0" <%if trim(peakname)="0" then response.write "SELECTED"%>>n/a</option>
		</select>
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Weekdays</span></td> 
	<td>
		<select name="sweekday">
			<option value="2" <%if sweekday=2 then response.write "SELECTED"%>>Monday</option>
			<option value="3" <%if sweekday=3 then response.write "SELECTED"%>>Tuesday</option>
			<option value="4" <%if sweekday=4 then response.write "SELECTED"%>>Wednesday</option>
			<option value="5" <%if sweekday=5 then response.write "SELECTED"%>>Thursday</option>
			<option value="6" <%if sweekday=6 then response.write "SELECTED"%>>Friday</option>
			<option value="7" <%if sweekday=7 then response.write "SELECTED"%>>Saturday</option>
			<option value="1" <%if sweekday=1 then response.write "SELECTED"%>>Sunday</option>
		</select>
		to
		<select name="eweekday">
			<option value="2" <%if eweekday=2 then response.write "SELECTED"%>>Monday</option>
			<option value="3" <%if eweekday=3 then response.write "SELECTED"%>>Tuesday</option>
			<option value="4" <%if eweekday=4 then response.write "SELECTED"%>>Wednesday</option>
			<option value="5" <%if eweekday=5 then response.write "SELECTED"%>>Thursday</option>
			<option value="6" <%if eweekday=6 then response.write "SELECTED"%>>Friday</option>
			<option value="7" <%if eweekday=7 then response.write "SELECTED"%>>Saturday</option>
			<option value="1" <%if eweekday=1 then response.write "SELECTED"%>>Sunday</option>
		</select>
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Start Time</span></td>
	<td><input type="text" name="stime" value="<%=stime%>" size="4"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">End Time</span></td>
	<td><input type="text" name="etime" value="<%=etime%>" size="4"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Rate Season</span></td>
	<td>
		<select name="seasonid">
			<%
			rst1.open "SELECT * FROM rateseasons WHERE regionid="&rid&" ORDER BY sday desc, smonth desc", cnn1
			do until rst1.eof
				%><option value="<%=rst1("id")%>"<%if trim(seasonid)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("season")%> (<%=rst1("smonth")%>/<%=rst1("sday")%>) to (<%=rst1("emonth")%>/<%=rst1("eday")%>)</option><%
				rst1.movenext
			loop
			%>
		</select>
	</td>
</tr>
<tr bgcolor="#dddddd"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if trim(rPid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		  <input type="button" value="Cancel" class="standard" onclick="document.location='seasonView.asp?rid=<%=rid%>';" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		  <input type="button" value="Cancel" class="standard" onclick="document.location='seasonView.asp?rid=<%=rid%>';" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">
<input type="hidden" name="rPid" value="<%=rPid%>">

</form>
</body>
</html>