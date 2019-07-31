<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Rate Setup")) then '("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, rtid, rateid
rid = secureRequest("rid")
rtid = secureRequest("rtid")
rateid = secureRequest("rateid")

dim rate, peak, utility, ratefrom, rateto, Itemtype, linecharge, monthstart, monthend, startdate, enddate
if trim(secureRequest("ratedup"))<>"" then
	rid = secureRequest("rid")
	rtid = secureRequest("rtid")
	rateid = secureRequest("rateid")
	rate = secureRequest("rate")
	peak = secureRequest("peak")
	utility = secureRequest("utility")
	ratefrom = secureRequest("ratefrom")
	rateto = secureRequest("rateto")
	Itemtype = secureRequest("Itemtype")
	linecharge = secureRequest("linecharge")
	monthstart = secureRequest("monthstart")
	monthend = secureRequest("monthend")
	startdate = secureRequest("startdate")
	enddate = secureRequest("enddate")
elseif trim(rateid)<>"" then
	rst1.Open "SELECT *  FROM rate WHERE id='" & rateid & "'", cnn1
'	response.write "SELECT *  FROM rate WHERE id='" & rateid & "'"
'	response.end
	if not rst1.EOF then
		rate = rst1("rate")
		peak = rst1("peak")
		utility = rst1("utility")
		ratefrom = rst1("ratefrom")
		rateto = rst1("rateto")
		Itemtype = rst1("Itemtype")
		linecharge = rst1("linecharge")
		monthstart = rst1("monthstart")
		monthend = rst1("monthend")
		startdate = rst1("startdate")
		enddate = rst1("enddate")
	end if
	rst1.close
end if

dim rtype
if trim(rtid)<>"" then
  rst1.open "select type from ratetypes where id='" & rtid & "'", cnn1
  if not rst1.EOF then
		rtype = rst1("type")
  end if
  rst1.close
end if

dim city
if trim(rid)<>""  then
	rst1.Open "SELECT city FROM regions WHERE id='"&rid&"'", cnn1
	if not rst1.EOF then city = rst1("city")
  rst1.Close
end if

%>
<html>
<head>
<title>Edit Rate</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
<script>
function confirmClose(){
	 if (confirm("Cancel changes?")){
      window.close()
    }
}

function confirmDelete(){
	 if (confirm("Really Delete?")){
		doDelete()
      window.close()
    }
}
</script>
</head>

<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="rateSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td colspan="2" bgcolor="#000000">
<%if allowGroups("Genergy Users") then%>
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
	<td colspan="2"><span class="standardheader">
		<%if trim(rateid)<>"" then%>
			Update Rate | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=city%> Region</a>
		<%else%>
			Add New Rate | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=city%> Region</a>
		<%end if%>
	</span></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Rate Type:</span></td>
	<td><span class="standard"><b><%=rtype%></b></span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Rate $</span></td>
	<td><input type="text" name="rate" value="<%=rate%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Rate Peak</span></td>
	<td>
		<select name="peak">
			<option value="0">Non-Applicable</option>
			<%
			rst1.open "SELECT rp.id as rpid, * FROM ratepeak rp INNER JOIN rateSeasons rs ON rs.id=rp.seasonid WHERE rs.regionid='"&rid&"'", cnn1
			do until rst1.eof
				%><option value="<%=rst1("rpid")%>"<%if trim(peak)=trim(rst1("rpid")) then response.write " SELECTED"%>><%=rst1("label")%>: <%=weekdayname(cint(rst1("sweekday")))%>-<%=weekdayname(cint(rst1("eweekday")))%>, <%=rst1("stime")%>-<%=rst1("etime")%> (<%=rst1("season")%>)</option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select> <input type="button" value="Edit Rate Peaks" onclick="document.location='seasonView.asp?rid=<%=rid%>';" class="standard">
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Utility</span></td>
	<td>
		<select name="utility">
		<%rst1.open "SELECT * FROM tblutility", cnn1
		do until rst1.eof%>
			<option value="<%=trim(rst1("utilityid"))%>" <%if lcase(utility)=trim(rst1("utilityid")) then response.write "SELECTED"%>><%=rst1("utilitydisplay")%></option>
		<%rst1.movenext
		loop
		rst1.close
		%>
		</select> <input type="button" value="Edit Utilities" onclick="location='utilityEdit.asp';" class="standard">
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Level</span></td>
	<td>From: <input type="text" name="ratefrom" value="<%=ratefrom%>"> To: <input type="text" name="rateto" value="<%=rateto%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Item Type</span></td>
	<td>
		<select name="itemtype">
			<option value="Energy" <%if lcase(Itemtype)="energy" then response.write "SELECTED"%>>Energy</option>
			<option value="Demand" <%if lcase(Itemtype)="demand" then response.write "SELECTED"%>>Demand</option>
			<option value="Static" <%if lcase(Itemtype)="static" then response.write "SELECTED"%>>Static</option>
		</select>
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Line Charge</span></td>
	<td><select name="linecharge">
			<%
			rst1.open "SELECT * FROM rateDescription", cnn1
			do until rst1.eof
				%><option value="<%=rst1("id")%>"<%if trim(linecharge)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("description")%></option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select> <input type="button" value="Edit Rate Peaks" onclick="document.location='seasonView.asp?rid=<%=rid%>';" class="standard">
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Starting Month</span></td>
	<td><input type="text" name="monthstart" value="<%=monthstart%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Ending Month</span></td>
	<td><input type="text" name="monthend" value="<%=monthend%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Start Date</span></td>
	<td><input type="text" name="startdate" value="<%=startdate%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">End Date</span></td>
	<td><input type="text" name="enddate" value="<%=enddate%>"></td>
</tr>
<tr bgcolor="#dddddd"> 
	<td><span class="standard">&nbsp;</span></td>
	<td>
		<%if trim(rateid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="submit" name="action" value="Delete" onclick="confirmDelete()" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		  <input type="button" value="Cancel" class="standard" onclick="confirmClose()" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		  <input type="button" value="Cancel" class="standard" onclick="confirmClose();" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">
<input type="hidden" name="rtid" value="<%=rtid%>">
<input type="hidden" name="rateid" value="<%=rateid%>">
<%if trim(secureRequest("ratedup"))<>"" then%><span class="standard"><font color="#990000">*This was not saved because it would duplicate another entry.</font></span><%end if%>
</form>
</body>
</html>
