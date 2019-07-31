<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql, rst2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid
rid = secureRequest("rid")

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

function showSeasonEdit(seasonid){	
  document.all['txt'+seasonid].style.display = 'none';
  document.all['edit'+seasonid].style.display = 'inline';
  document.all['row'+seasonid].style.backgroundColor = '#ddffdd';
}

function hideSeasonEdit(seasonid){	
  document.all['txt'+seasonid].style.display = 'inline';
  document.all['edit'+seasonid].style.display = 'none';
  document.all['row'+seasonid].style.backgroundColor = '#eeeeee';
}

function visibilityChange(rate,labels){
	try{
		state = labels.innerHTML;
		state = (state=="[-]"?"[+]":"[-]");
		document.all[rate].style.display=(state=="[+]"?'none':'inline');
		labels.innerHTML = state;
	}catch(exception){};
}

</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<form>
<tr>
  <td colspan="2" bgcolor="#000000">
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
  <td><span class="standard" style="color:#ffffff">&nbsp;<a href="regionView.asp" style="color:#ffffff;font-weight:bold;text-decoration:none;">Rate Setup</a> | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff"><%=city%> Region</a> | Seasons &amp; Rate Peaks | <a href="rateTypeView.asp?rid=<%=rid%>" style="color:#ffffff">Rate Types</a></span></td>
  <td align="right">
  <input type="button" value="Add Season" onclick="document.all['newseason'].style.display='inline';" id=1 name=1>
  <input type="button" value="Manage Holidays" onclick="document.location = 'holidayView.asp?rid=<%=rid%>';" id=1 name=1>
  </td>
</tr>
</form>
</table>
<div id="newseason" style="display:'none';">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<form name="newseasonform" method="post" action="seasonSave.asp">
<input type="hidden" name="rid" value="<%=rid%>">
<tr>
  <td align="center" colspan="2" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td><span class="standard"><b>New Season:</b></span></td>
    <td><input type="text" name="season" value=""></td>
    <td><span class="standard"><b>Start Date:</b></span></td>
    <td><input type="text" name="smonth" value="" maxlength="2" size="3">/<input type="text" name="sday" value="" maxlength="2" size="3"></td>
    <td><span class="standard"><b>End Date:</b></span></td>
    <td><input type="text" name="emonth" value="" maxlength="2" size="3">/<input type="text" name="eday" value="" maxlength="2" size="3"></td>
	<td><span class="standard"><b>Effective Date:</b></span><input type="text" name="effective_date" value="" size="5"></td>
    <td><input type="submit" name="action" value="Save" class="standard"></td>
    <td><input type="button" value="Cancel" onclick="document.all['newseason'].style.display='none';" class="standard"></td>
  </tr>
  </table>        
  </td>
</tr>
</form>
</table>
</div>
<%
dim hasSeasons, hasRatepeaks
'#######season list#######'
if trim(rid)<>"" then
	rst1.Open "SELECT rs.id as rsid, * FROM rateseasons rs INNER JOIN regions ON rs.regionid=regions.id WHERE regionid='"&rid&"' ORDER BY effective_date desc", cnn1
	if not rst1.EOF then hasSeasons = true
	if not rst1.EOF then%>
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
<%do until rst1.EOF%>
    <form name="form2" method="post" action="seasonSave.asp">
    <input type="hidden" name="rid" value="<%=rid%>">
    <input type="hidden" name="seasonid" value="<%=rst1("rsid")%>">
    <tr>
      <td bgcolor="#eeeeee" colspan="2" style="padding:12px;">
      <table border=0 cellpadding="0" cellspacing="0" width="100%" style="border:1px solid #cccccc;">            
      <tr id="row<%=rst1("rsid")%>" bgcolor="#eeeeee">
        <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;padding:3px;"><label onclick="visibilityChange('panel<%=rst1("rsid")%>', this)">[+]</label>
        <span id="txt<%=rst1("rsid")%>" class="standard" style="display:'inline';"><b><%=rst1("season")%></b> (<%=rst1("Smonth")%>/<%=rst1("Sday")%> - <%=rst1("EMonth")%>/<%=rst1("Eday")%>)  Effective:<%=rst1("effective_date")%></span>
        <span id="edit<%=rst1("rsid")%>" class="standard" style="display:'none';">
        <table border=0 cellpadding="3" cellspacing="0">
        <tr>
          <td><span class="standard"><b>Season:</b></span></td>
          <td><input type="text" name="season" value="<%=rst1("season")%>"></td>
          <td><span class="standard"><b>Start Date:</b></span></td>
          <td><input type="text" name="smonth" value="<%=rst1("smonth")%>" maxlength="2" size="3">/<input type="text" name="sday" value="<%=rst1("sday")%>" maxlength="2" size="3"></td>
          <td><span class="standard"><b>End Date:</b></span></td>
          <td><input type="text" name="emonth" value="<%=rst1("emonth")%>" maxlength="2" size="3">/<input type="text" name="eday" value="<%=rst1("eday")%>" maxlength="2" size="3"></td>
		  <td><span class="standard"><b>Effective Date:</b></span><input type="text" name="effective_date" value="<%=rst1("effective_date")%>" size="5"></td>
		  <td>
          <td><input type="submit" name="action" value="Update" class="standard"></td>
          <td><input type="button" value="Cancel" onclick="hideSeasonEdit(<%=rst1("rsid")%>);" class="standard"></td>
        </tr>
        </table>        
        </span>
        </td>
        <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;padding:3px;" align="right"><input type="button" value="Edit Season" onclick="showSeasonEdit(<%=rst1("rsid")%>);" class="standard"></td>
        <!--[[td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;padding:3px;" align="right"]][[input type="button" value="Edit Season" onclick="seasonEdit([[%=rst1("id")%]]);" class="standard"]][[/td]]-->
      </tr>
      <tr valign="top" bgcolor="#eeeeee">
        <td colspan="2" bgcolor="#ffffff">
      <%dim sqlstr
      sqlstr = "SELECT rp.id as rpid, * FROM ratepeak rp INNER JOIN rateSeasons rs ON rs.id=rp.seasonid WHERE rs.regionid='"&rid&"' and rs.id='"&rst1("rsid")&"'"
      rst2.Open sqlstr, cnn1
      if not rst2.EOF then hasRatepeaks = true
      if not rst2.EOF then%>
        <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#cccccc" id="panel<%=rst1("rsid")%>" style="display:none">
        <tr bgcolor="#dddddd">
          <td width="20%"><span class="standard"><b>Rate Peak Label</b></span></td>
          <td width="20%"><span class="standard"><b>Rate Peak Type</b></span></td>
          <td width="20%"><span class="standard"><b>Days in Week</b></span></td>
          <td width="60%"><span class="standard"><b>Time of Day</b></span></td>
        </tr>
        <%do until rst2.EOF%>
        <tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="ratePeakEdit(<%=rst2("rpid")%>);">
          <td><span class="standard"><%=rst2("label")%></span></td>
          <td><span class="standard">
		  <%
		  if trim(rst2("peakname"))="1" then
		  	response.write "On Peak"
		  elseif trim(rst2("peakname"))="2" then
		  	response.write "Off Peak"
		  elseif trim(rst2("peakname"))="3" then
		  	response.write "Int Peak"
		  else
		  	response.write "N/A"
		  end if
		  %>
		  </span></td>
          <td><span class="standard"><%=weekdayname(cint(rst2("sweekday")))%>-<%=weekdayname(cint(rst2("eweekday")))%></span></td>
          <td><span class="standard"><%=rst2("stime")%>-<%=rst2("etime")%></span></td>
        </tr>
        
        <%rst2.movenext
        loop%>
        <tr bgcolor="#ffffff">
          <td colspan="4"><input type="button" value="Add Rate Peak" onclick="ratePeakEdit('');" id=1 name=1></td>
        </tr>
        </table>
      <%
      else%>
        <table border=0 cellpadding="3" cellspacing="0" width="100%" id="panel<%=rst1("rsid")%>" style="display:none">
        <tr bgcolor="#dddddd"><td><span class="standard"><b>Rate Peaks</b></span></td></tr>
        <tr><td><span class="standard">No rate peaks have been set up for this season.</span></td></tr>
        <tr><td><input type="button" value="Add Rate Peak" onclick="ratePeakEdit('');" id=1 name=1></td></tr>
        </table>
      <% 
      end if
      rst2.close
       %>
      </table>
      </td>
    </tr>
    </form>
		<%rst1.movenext
		loop%>
  	<tr><td colspan="2" bgcolor="#eeeeee" style="border-bottom:1px solid #999999;" height="10"><span class="standard">&nbsp;</span></td></tr>
    </table>
  <br><br>
	<%
	else %>
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#dddddd"><td><span class="standard"><b>Seasons</b></span></td></tr>
  <tr bgcolor="#ffffff"><td><span class="standard">No rate seasons have been set up for this region.</span></td></tr>
  <tr bgcolor="#ffffff"><td><input type="button" value="Add Season" onclick="document.all['newseason'].style.display='inline';" id=1 name=1></td></tr>
  <tr bgcolor="#ffffff"><td height="500">&nbsp;</td></tr>
  </table>

	<% end if
	rst1.close

end if

%>
</form>
</body>
</html>
