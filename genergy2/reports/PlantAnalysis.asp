<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if

dim m, y, building, pid, pdf, groupname, groupid, tmonth, tday, tyear, fmonth, fday, fyear, action, patrend, pascatter
if trim(request("pdf"))="yes" then pdf = true else pdf = false
pid = request("pid")
m = request("m")
y = request("y")
groupid = request("groupid")
building = request("building")
tmonth = request("tmonth")
tday = request("tday")
tyear = request("tyear")
fmonth = request("fmonth")
fday = request("fday")
fyear = request("fyear")
action = request("action")
if tmonth = "" or tday = "" or tyear = "" or fmonth = "" or fday = "" or fyear = "" then
  tmonth = datepart("m",date())
  tday = datepart("d",date())
  tyear = datepart("yyyy",date())
  fmonth = datepart("m",date())
  fday = datepart("d",date())
  fyear = datepart("yyyy",date())
end if
if trim(action)="Confirm" then
  if trim(request("patrend"))="true" then session("patrend") = true else session("patrend") = false
  if trim(request("pascatter"))="true" then session("pascatter") = true else session("pascatter") = false
end if
if trim(session("patrend"))<>"" then patrend = session("patrend") else patrend = request("patrend")
if trim(session("pascatter"))<>"" then pascatter = session("pascatter") else pascatter = request("pascatter")

if lcase(patrend)="true" then patrend = true else patrend = false
if lcase(pascatter)="true" then pascatter = true else pascatter = false

if not(patrend) and not(pascatter) then
  patrend = true
  session("patrend") = true
end if

if datediff("d",datevalue(tmonth&"/"&tday&"/"&tyear),datevalue(fmonth&"/"&fday&"/"&fyear))<0 then 
  fmonth = tmonth
  fday = tday
  fyear = tyear
end if
dim rst1, cnn1, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,building,"billing")
cnn1.CursorLocation = adUseClient
if groupid<>"" and groupid<>"0" then
	rst1.open "SELECT groupname FROM [group] WHERE id='"&groupid&"'", cnn1
	if not rst1.eof then groupname = rst1("groupname")
	rst1.close
end if
dim selected
%>
<html><head><title>Plant Analysis Report</title>
<script>
function loadReport(id)
{ document.location = 'PlantAnalysis.asp?groupid='+id.value
}

</script>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" link="#0000FF" vlink="#0000FF" alink="#0000FF" onload="document.all['loadFrame1'].style.visibility='hidden'">
<form name="weather">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td<%if pdf then%> align="center"<%end if%>><span class="standardheader"><font color="white">Plant Analysis Report<%if pdf then response.write " : "&groupname%></font></span></td>
  <%
  if not(pdf) then
  dim qs
  qs = request.servervariables("SCRIPT_NAME")&"?"&request.servervariables("QUERY_STRING")
  if trim(action)<>"Confirm" and trim(request("patrend"))="" and trim(request("pascatter"))="" then qs = qs & "&patrend="&patrend&"&pascatter="&pascatter
  %><td align="right"><input type="button" value="Preferences" onclick="if(document.all['preferences'].style.display=='none'){document.all['preferences'].style.display='inline';}else{document.all['preferences'].style.display='none';}"><input type="button" value="Print Current View" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport.asp?devIP=<%=request.servervariables("SERVER_NAME")%>&qs=<%=server.URLEncode(qs)%>','','width=600,height=400,scrollbars=no,resizable=yes');"></td><%end if%>
</tr>
</table>
<table width="100%" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"<%if pdf then%> align="center"<%end if%>><%if not pdf then%>Show report for:<%end if%>&nbsp;
  <%
  if not pdf then 
    response.write "<select name=""groupid"" ><option value=""0"">Select a Input Group</option>"
    rst1.open "SELECT * FROM [group] WHERE type=5 ORDER BY grouplabel", cnn1
    do until rst1.eof
      selected=""
      if trim(groupid) = trim(rst1("id")) then selected=" SELECTED"
      response.write "<option value="""&rst1("id")&""""&selected&">"&rst1("grouplabel")&"</option>"
      rst1.movenext
    loop
    rst1.close
    response.write "</select>"
  end if
  %>

<b>Dates:</b>
  <%if not pdf then%>
  <select name="tmonth">
    <%dim i
      for i = 1 to 12
        %><option value="<%=i%>"<%if cint(tmonth)=i then response.write " SELECTED"%>><%=left(monthname(i),3)%></option><%
      next
    %>
  </select>
  <select name="tday">
    <%
      for i = 1 to 31
        %><option value="<%=i%>"<%if cint(tday)=i then response.write " SELECTED"%>><%=i%></option><%
      next
    %>
  </select>
  <select name="tyear">
    <%
      rst1.open "SELECT distinct datepart(year,datestart) as year FROM billyrperiod b WHERE datestart is not null ORDER BY datepart(year,datestart) desc"
      do until rst1.eof
        %><option value="<%=rst1("year")%>"<%if cint(tyear)=cint(rst1("year")) then response.write " SELECTED"%>><%=rst1("year")%></option><%
        rst1.movenext
      loop
      rst1.close
    %>
  </select>
  To:
  <select name="fmonth">
    <%
      for i = 1 to 12
        %><option value="<%=i%>"<%if cint(fmonth)=i then response.write " SELECTED"%>><%=left(monthname(i),3)%></option><%
      next
    %>
  </select>
  <select name="fday">
    <%
      for i = 1 to 31
        %><option value="<%=i%>"<%if cint(fday)=i then response.write " SELECTED"%>><%=i%></option><%
      next
    %>
  </select>
  <select name="fyear" onchange="submit();">
    <%
      rst1.open "SELECT distinct datepart(year,datestart) as year FROM billyrperiod b WHERE datestart is not null ORDER BY datepart(year,datestart) desc"
      do until rst1.eof
        %><option value="<%=rst1("year")%>"<%if cint(fyear)=cint(rst1("year")) then response.write " SELECTED"%>><%=rst1("year")%></option><%
        rst1.movenext
      loop
      rst1.close
    %>
  </select>
  <input type="submit" value="GO">
  <%else%>
  <%=monthname(tmonth)&" "&tday&", "&tyear%> to <%=monthname(fmonth)&" "&fday&", "&fyear%>
  <%end if%>
  </td>
</tr>
<tr>
  <td bgcolor="#eeeeee">
  <%if not pdf then%>
  <div id="preferences" style="display:none;">
  <br><b>Plant Analysis Report Preferences</b>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr valign="top">
    <td>
    <!-- begin usage/degree prefs -->
    Show on top graph:<br>
    <table border=0 cellpadding="3" cellspacing="0">
    <tr>
      <td><input type="checkbox" name="pascatter" value="true"<%if session("pascatter") then response.write " CHECKED"%>></td>
      <td>Scatter Points</td>
    </tr>
    <tr>
      <td><input type="checkbox" name="patrend" value="true"<%if session("patrend") then response.write " CHECKED"%>></td>
      <td>Trend Line</td>
    </tr>
    </table>
    <!-- end usage/degree prefs -->
    </td>
    <td width="20">&nbsp;</td>
    <td>
    <!-- begin load profile prefs -->
    <!-- end load profile prefs -->
    </td>
  </tr>
  </table>
    <input type="submit" name="action" value="Confirm">

  </div>
  <%end if%>
  </td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
  <td>
  <table border="0" cellpadding="3" cellspacing="1" width="600" align="center">
  <tr>
    <td align="center"><b><font size="2">Plant Efficiency Report <%=groupname%>, <%=monthname(tmonth)&" "&tday&", "&tyear%> to <%=monthname(fmonth)&" "&fday&", "&fyear%></font></b></td>
  </tr>
  </table>
  </td>
</tr>
<%if trim(groupid)<>"" and trim(groupid)<>"0" then%>
<tr>
  <td align="center"><img src="PlantAnalysisInOut.asp?groupid=<%=groupid%>&groupname=<%=groupname%>&startdate=<%=tmonth%>/<%=tday%>/<%=tyear%>&enddate=<%=fmonth%>/<%=fday%>/<%=fyear%>&patrend=<%=patrend%>&pascatter=<%=pascatter%>" width="600" height="310"></td>
</tr>
<tr>
  <td align="center"><img src="PlantAnalysisRunHours.asp?groupid=<%=groupid%>&groupname=<%=groupname%>&startdate=<%=tmonth%>/<%=tday%>/<%=tyear%>&enddate=<%=fmonth%>/<%=fday%>/<%=fyear%>" width="600" height="350"></td>
</tr>
<tr>
<td align="center">
  <table border="0" cellpadding="3" cellspacing="1" width="600" align="center">
  <tr>
    <td width="400"><%if pdf then%><img src="/genergy2/invoices/invoice-logo-1.jpg" alt="" width="202" height="143" border="0"><%end if%></td>
    <td>
      <b>Point Sources:</b><br>
      Input
      <table cellpadding="1" cellspacing="1" bgcolor="#000000" width="140">
      <tr bgcolor="#eeeeee"><td>Meter</td><td>Type</td></tr>
      <%
      rst1.open "SELECT distinct meternum, utilitydisplay FROM meters m INNER JOIN tblleasesutilityprices lup ON m.leaseutilityid=lup.leaseutilityid INNER JOIN tblutility u ON u.utilityid=lup.utility WHERE meterid in (SELECT typeid FROM groupitems WHERE groupid in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid AND g.id='"&groupid&"' AND typecode='i'))", cnn1
      do until rst1.eof
        %><tr bgcolor="white"><td nowrap><%=rst1("meternum")%></td><td nowrap><%=rst1("utilitydisplay")%></td></tr><%
        rst1.movenext
      loop
      rst1.close
      %>
      </table>
      Output
      <table cellpadding="1" cellspacing="1" bgcolor="#000000" width="140">
      <tr bgcolor="#eeeeee"><td>Meter</td><td>Type</td></tr>
      <%
      rst1.open "SELECT distinct meternum, utilitydisplay FROM meters m INNER JOIN tblleasesutilityprices lup ON m.leaseutilityid=lup.leaseutilityid INNER JOIN tblutility u ON u.utilityid=lup.utility WHERE meterid in (SELECT typeid FROM groupitems WHERE groupid in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid AND g.id='"&groupid&"' AND typecode='o'))", cnn1
      do until rst1.eof
        %><tr bgcolor="white"><td nowrap><%=rst1("meternum")%></td><td nowrap><%=rst1("utilitydisplay")%></td></tr><%
        rst1.movenext
      loop
      rst1.close
      %>
      </table>
    </td>
  </tr>
  </table>
</td>
</tr>
<%end if%>
</table>
<br>
<%if not pdf then%>
<div id="loadFrame1" style="visibility:visible; position:absolute;left:320;top:150;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
<%end if%>
</body>
</html>
