<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if

dim m, y, building, pid, utilityid, ypid, startdate, enddate, currentdate, dutcdate, pdf, action, uddusage, udddegree, utusage, uttemp, uthumid, wddtrend, wddscatter

dutcdate = request("dutcdate")
if trim(request("pdf"))="yes" then pdf = true else pdf = false
pid = request("pid")
m = request("m")
y = request("y")
action = request("action")
utilityid = request("utilityid")
if trim(utilityid)="" then utilityid=2
building = request("building")
ypid = request("ypid")
if ypid = "0" or ypid="" or instr(ypid,"|")=0 then
  enddate = date()
  startdate = dateadd("d",-30,enddate)
  ypid = 0
else
  startdate = trim(split(ypid,"|")(0))
  enddate = trim(split(ypid,"|")(1))
    if datediff("d",startdate,enddate)<0 then enddate = dateadd("d",30,startdate)
end if
dim DBlocalIP
if trim(building)<>"" then DBlocalIP = ""


'set sessions
if trim(action)="Confirm" then
  if trim(request("uddusage"))="True" then session("uddusage") = true else session("uddusage") = false
  if trim(request("udddegree"))="True" then session("udddegree") = true else session("udddegree") = false
  if trim(request("utusage"))="True" then session("utusage") = true else session("utusage") = false
  if trim(request("uttemp"))="True" then session("uttemp") = true else session("uttemp") = false
  if trim(request("uthumid"))="True" then session("uthumid") = true else session("uthumid") = false
  if trim(request("wddscatter"))="True" then session("wddscatter") = true else session("wddscatter") = false
  if trim(request("wddtrend"))="True" then session("wddtrend") = true else session("wddtrend") = false
end if
if trim(session("utusage"))<>"" then utusage = session("utusage") else utusage = request("utusage")
if trim(session("uttemp"))<>"" then uttemp = session("uttemp") else uttemp = request("uttemp")
if trim(session("uthumid"))<>"" then uthumid = session("uthumid") else uthumid = request("uthumid")
if trim(session("udddegree"))<>"" then udddegree = session("udddegree") else udddegree = request("udddegree")
if trim(session("uddusage"))<>"" then uddusage = session("uddusage") else uddusage = request("uddusage")
if trim(session("wddscatter"))<>"" then wddscatter = session("wddscatter") else wddscatter = request("wddscatter")
if trim(session("wddtrend"))<>"" then wddtrend = session("wddtrend") else wddtrend = request("wddtrend")

if lcase(utusage)="true" then utusage = true else utusage = false
if lcase(uttemp)="true" then uttemp = true else uttemp = false
if lcase(uthumid)="true" then uthumid = true else uthumid = false
if lcase(udddegree)="true" then udddegree = true else udddegree = false
if lcase(uddusage)="true" then uddusage = true else uddusage = false
if lcase(wddscatter)="true" then wddscatter = true else wddscatter = false
if lcase(wddtrend)="true" then wddtrend = true else wddtrend = false

if not(uddusage) and not(udddegree) and not(utusage) and not(uttemp) and not(uthumid) then
  uddusage = true
  uttemp = true
  utusage = true
  udddegree = true
  session("uddusage") = true
  session("uttemp") = true
  session("utusage") = true
  session("udddegree") = true
end if


if not(wddtrend) and not(wddscatter) then
  wddtrend = true
  session("wddtrend") = true
end if

dim rst1, cnn1, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,building,"billing")
cnn1.CursorLocation = adUseClient
dim selected
%>
<html><head><title>Weather Report</title>
<script>
function loadbuilding(building)
{ 
document.location = 'weatherReport.asp?pid=<%=pid%>&building='+building+'&utilityid=<%=utilityid%>'
}

function loadypid(ypid)
{ document.location = 'weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid='+ypid
}

function pickDate(dutcdate)
{ var frm = document.forms[0];
  var startdate
  var enddate
  sfmonth = frm.sfmonth.value
  sfday = frm.sfday.value
  sfyear = frm.sfyear.value
  if((sfmonth!='')&&(sfday!='')&&(sfyear!=''))
  { stmonth = frm.stmonth.value
    stday = frm.stday.value
    styear = frm.styear.value
    startdate = sfmonth+'/'+sfday+'/'+sfyear
    if((stmonth!='')&&(stday!='')&&(styear!=''))
    { enddate = stmonth+'/'+stday+'/'+styear
      //alert(enddate)
    }else
    { enddate = new Date(parseFloat(sfyear), parseFloat(sfmonth)-1, parseFloat(sfday), 0, 0, 0, 0);
      enddate.setDate(enddate.getDate() + 30);
      enddate = (enddate.getMonth()+1)+'/'+enddate.getDate()+'/'+enddate.getYear()
    }
    document.location = 'weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid='+startdate+'|'+enddate+'&dutcdate='+dutcdate
  }
}
</script>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form name="weather">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Weather Tracking Report</span></td>
  <%dim qs
  qs = request.servervariables("SCRIPT_NAME")&"?"&request.servervariables("QUERY_STRING")
  if trim(action)<>"Confirm" then qs = qs & "&uddusage="&uddusage&"&udddegree="&udddegree&"&utusage="&utusage&"&uttemp="&uttemp&"&uthumid="&uthumid&"&wddtrend="&wddtrend&"&wddscatter="&wddscatter
  %>
  <%if not(pdf) then%><td align="right"><input type="button" value="Preferences" onclick="if(document.all['preferences'].style.display=='none'){document.all['preferences'].style.display='inline';}else{document.all['preferences'].style.display='none';}"><input type="button" value="Print Current View" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport.asp?devIP=<%=request.servervariables("SERVER_NAME")%>&qs=<%=server.URLEncode(qs)%>','','width=600,height=400,resizable=yes');"></td><%end if%>
</tr>
</table>
<table width="100%" cellpadding="3" cellspacing="0" border="0">
<tr>
  <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%if not pdf then%>Show report for:<%else%><b>Show report for:</b><%end if%>&nbsp;
  <%
  if not pdf then response.write "<select name=""building"" onchange=""loadbuilding(this.value);""><option value=""0"">Select a building</option>"
  rst1.open "SELECT isnull(date_offline,'') as dateoffline, * FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
  do until rst1.eof
    selected=""
    if trim(building) = trim(rst1("bldgnum")) then selected=" SELECTED"
    if not(pdf) then
      %><option value="<%=rst1("bldgnum")%>" <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> <%=selected%>><%=rst1("strt")%> (<%=rst1("bldgnum")%>)</option><%
    elseif selected=" SELECTED" then
      %><%=rst1("strt")%> (<%=rst1("bldgnum")%>)<%if isBuildingOff("02") then%> offline <%=rst1("dateoffline")%><%end if%><%
    end if
    
    rst1.movenext
  loop
  rst1.close
  if not pdf then response.write "</select>"
  %>
  <%
  dim custom_dates
  custom_dates = true
  if not pdf then
    if trim(building)<>"" then
      rst1.open "SELECT datestart, dateend, billyear, billperiod, * FROM billyrperiod WHERE bldgnum='"&building&"' AND datestart<getdate() and utility="&utilityid&" ORDER BY datestart desc", cnn1
      response.write "<select name=""ypid"" onchange=""loadypid(this.value);""><option value=""0"">Last 30 Days</option>"
      do until rst1.eof
        selected=""
        if trim(ypid) = rst1("datestart")&"|"&rst1("dateend") then
          selected=" SELECTED"
          custom_dates = false
          'byear = rst1("billyear")
          'bperiod = rst1("billperiod")
        end if
        response.write "<option value="""&rst1("datestart")&"|"&rst1("dateend")&""""&selected&">"&rst1("datestart")&" - "&rst1("dateend")&"</option>"
        rst1.movenext
      loop
      if custom_dates and ypid<>"0" then response.write "<option selected>Custom Dates</option>"
      rst1.close
      response.write "</select>"
    end if
  else
    response.write "&nbsp;<b>Dates:</b>&nbsp;" & startdate &" - " & enddate
  end if

if not pdf then
  %>
  <input type="button" value="Electricity" style="border:1px outset #ffffff;background-color:#dddddd;padding-left:4px;padding-right:4px;<%if utilityid="2" then%>font-weight:bold<% end if %>" onclick="location='weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=2';">
  <input type="button" value="Gas" style="border:1px outset #ffffff;background-color:#dddddd;width:65px;padding-left:4px;padding-right:4px;<%if utilityid="3" then%>font-weight:bold<% end if %>" onclick="location='weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=3';">
  <input type="button" value="Steam" style="border:1px outset #ffffff;background-color:#dddddd;width:65px;padding-left:4px;padding-right:4px;<%if utilityid="1" then%>font-weight:bold<% end if %>" onclick="location='weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=1';">
  <input type="button" value="Chilled W." style="border:1px outset #ffffff;background-color:#dddddd;width:65px;padding-left:4px;padding-right:4px;<%if utilityid="6" then%>font-weight:bold<% end if %>" onclick="location='weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=6';">
<%
else
  response.write " <b>Utility:</b>&nbsp;"
  if utilityid="1" then response.write "Steam"
  if utilityid="2" then response.write "Gas"
  if utilityid="3" then response.write "Steam"
  if utilityid="6" then response.write "Chilled Water"
end if
%>
  </td>
</tr>
<tr>
  <td bgcolor="#eeeeee">
  <%if not pdf then%>
  <div id="preferences" style="display:none;">
  <br><b>Weather Tracking Report Preferences</b>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr valign="top">
    <td>
    <!-- begin usage/degree prefs -->
    Show on usage/degree day graph:<br>
    <table border=0 cellpadding="3" cellspacing="0">
    <tr>
      <td><input type="checkbox" name="uddusage" value="True"<%if session("uddusage") then response.write " CHECKED"%>></td>
      <td>Usage Series</td>
    </tr>
    <tr>
      <td><input type="checkbox" name="udddegree" value="True"<%if session("udddegree") then response.write " CHECKED"%>></td>
      <td>Degree Day</td>
    </tr>
    </table>
    <!-- end usage/degree prefs -->
    </td>
    <td width="20">&nbsp;</td>
    <td>
    <!-- begin load profile prefs -->
    Show on load profile/temperature graph:<br>
    <table border=0 cellpadding="3" cellspacing="0">
    <tr>
      <td><input type="checkbox" name="utusage" value="True"<%if session("utusage") then response.write " CHECKED"%>></td>
      <td>Usage</td>
    </tr>
    <tr>
      <td><input type="checkbox" name="uttemp" onclick="if(this.checked==true){document.weather.uthumid.checked=false;}" value="True"<%if session("uttemp") then response.write " CHECKED"%>></td>
      <td>Temperature</td>
    </tr>
    <tr>
       <td><input type="checkbox" name="uthumid" onclick="if(this.checked==true){document.weather.uttemp.checked=false;}" value="True"<%if session("uthumid") then response.write " CHECKED"%>></td>
     <td>Humidity</td>
    </tr>
    </table>
    <!-- end load profile prefs -->
    </td>
    <td width="20">&nbsp;</td>
    <td>
    <!-- begin load profile prefs -->
    Degree Day Scatter graph:<br>
    <table border=0 cellpadding="3" cellspacing="0">
    <tr>
      <td><input type="checkbox" name="wddscatter" value="True"<%if session("wddscatter") then response.write " CHECKED"%>></td>
      <td>Scatter points</td>
    </tr>
    <tr>
      <td><input type="checkbox" name="wddtrend" value="True"<%if session("wddtrend") then response.write " CHECKED"%>></td>
      <td>Trend Line</td>
    </tr>
    </table>
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
<%if trim(pid)<>"" and trim(building)<>"" then%>
<%
dim peakd, peakddate, previouspd, currentpd, columntitle
rst1.open "SELECT measure FROM tblutility WHERE utilityid="&utilityid, cnn1
columntitle = trim(rst1("measure"))
rst1.close
cmd.ActiveConnection = cnn1
cmd.CommandText = "sp_LMPDATA"
cmd.CommandType = adCmdStoredProc
Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 12)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 12)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("code", adVarChar, adParamInput, 1)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("string", adVarChar, adParamInput, 30)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("interval", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("title", adVarChar, adParamOutPut, 30)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("max", adDouble, adParamOutPut, 18,2)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("sum", adDouble, adParamOutPut, 18,2)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("pdemand", adInteger, adParamOutPut)
cmd.Parameters.Append prm

cmd.Parameters("from")		= startdate
cmd.Parameters("to")		= enddate
cmd.Parameters("code")		= "b"
cmd.Parameters("string")		= building
cmd.Parameters("utility")		= utilityid
cmd.Parameters("interval")		= 0
'response.write "exec sp_LMPDATA '"&startdate&"','"&enddate&"','b','"&building&"',"&utilityid&",0,0,0,0"
'response.end
set rst1 = cmd.execute

currentdate = startdate
peakddate = currentdate
peakd = 0
previouspd = 0
dim hasdata
hasdata = false
on error resume next
do until datediff("d",enddate,currentdate)>0
	if not rst1.eof then
    hasdata = true
		if year(currentdate)=year(rst1("date")) and month(currentdate)=month(rst1("date")) and day(currentdate)=day(rst1("date")) then
			currentpd = cDbl(rst1(columntitle))
			rst1.movenext
		else
			currentpd = 0
		end if
	else
		currentpd = 0
	end if
	if peakd < previouspd+currentpd then
		peakddate = currentdate
		peakd = previouspd+currentpd
	end if
	previouspd = currentpd
	currentdate = dateadd("n",15,currentdate)
loop
rst1.close

peakddate = month(peakddate)&"/"&day(peakddate)&"/"&year(peakddate)
if trim(dutcdate) = "" then dutcdate = peakddate
%>
<tr>
  <td align="center"><img src="usageBarChart.asp?building=<%=building%>&startdate=<%=startdate%>&enddate=<%=enddate%>&utilityid=<%=utilityid%>&dutcdate=<%=dutcdate%>&uddusage=<%=uddusage%>&udddegree=<%=udddegree%>" width="650" height="310"></td>
</tr>
<tr>
  <td align="center">
  <%if not pdf then%>
  <div style="width:650px;background-color:#eeeeee;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;padding:4px;">
  <table border=0 cellpadding="1" cellspacing="0" bgcolor="#eeeeee" width="100%">
  <tr valign="middle">
    <td align="right">From:</td>
    <td>
    <select name="sfmonth">
    <option value="">Month</option>
    <%
    dim i
    for i = 1 to 12
    %><option value="<%=i%>"<%if i=cdbl(split(startdate,"/")(0)) then response.write " SELECTED"%>><%=i%></option><%
    next
    %>
    </select>
    </td>
    <td>
    <select name="sfday">
    <option value="">Day</option>
    <%
    for i = 1 to 31
    %><option value="<%=i%>"<%if i=cdbl(split(startdate,"/")(1)) then response.write " SELECTED"%>><%=i%></option><%
    next
    %>
    </select>
    </td>
    <td>
    <select name="sfyear">
    <option value="">Year</option>
    <%
    for i = 1996 to 2003
    %><option value="<%=i%>"<%if i=cdbl(split(startdate,"/")(2)) then response.write " SELECTED"%>><%=i%></option><%
    next
    %>
    </select>
    </td>
    <td width="12">&nbsp;</td>
    <td align="right">To:</td>
    <td>
    <select name="stmonth">
    <option value="">Month</option>
    <%
    for i = 1 to 12
    %><option value="<%=i%>"<%if i=cdbl(split(enddate,"/")(0)) then response.write " SELECTED"%>><%=i%></option><%
    next
    %>
    </select>
    </td>
    <td>
    <select name="stday">
    <option value="">Day</option>
    <%
    for i = 1 to 31
    %><option value="<%=i%>"<%if i=cdbl(split(enddate,"/")(1)) then response.write " SELECTED"%>><%=i%></option><%
    next
    %>
    </select>
    </td>
    <td>
    <select name="styear">
    <option value="">Year</option>
    <%
    for i = 1996 to 2003
    %><option value="<%=i%>"<%if i=cdbl(split(enddate,"/")(2)) then response.write " SELECTED"%>><%=i%></option><%
    next
    %>
    </select>
    </td>
    <td><input type="button" value="Show Graph" onclick="pickDate('')"></td>
  </tr>
  <tr>
    <td colspan="6">&nbsp;</td>
    <td colspan="4"><small>(optional, defaults to 30 days)</small></td>
  </tr>
  </table>
  </div>
  <%end if%>
  </td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
  <td align="center"><img src="dailyUsageTempChart.asp?building=<%=building%>&startdate=<%=dutcdate%>&enddate=<%=dateadd("d",1,dateValue(dutcdate))%>&utilityid=<%=utilityid%>&ispeakday=<%if datediff("d",dutcdate,peakddate)=0 then response.write "(Peak%20Day)"%>&utusage=<%=utusage%>&uttemp=<%=uttemp%>&uthumid=<%=uthumid%>" width="650" height="310"><br>
      <%if not pdf then%>
    	<select name="dutcdate">
        <%
        currentdate = startdate
        do until datediff("d",enddate,currentdate)>0
          %><option value="<%=currentdate%>"<%if datediff("d",currentdate,dutcdate)=0 then response.write " SELECTED"%>><%=currentdate%><%if datediff("d",currentdate,peakddate)=0 then response.write "(Peak Day)"%></option><%
        	currentdate = dateadd("d",1,currentdate)
        loop
        %>
      </select>
      <input type="button" value="Show Graph" onclick="pickDate(document.forms[0].dutcdate.value)">
      <%end if%>
  </td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
  <td align="center"><table width="650"><tr><td><img src="degreeDayScatter.asp?building=<%=building%>&startdate=<%=startdate%>&enddate=<%=enddate%>&utilityid=<%=utilityid%>&wddtrend=<%=wddtrend%>&wddscatter=<%=wddscatter%>" width="400" height="310">
<%
do until cmd.Parameters.Count = 0
  cmd.Parameters.Delete 0
loop
cmd.CommandText = DBlocalIP&"sp_BldgStats"
cmd.CommandType = adCmdStoredProc
Set prm = cmd.CreateParameter("string", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("datefrom", adVarChar, adParamInput, 30)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("dateto", adVarChar, adParamInput, 30)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("d1", adVarChar, adParamOutput, 11)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("d2", adVarChar, adParamOutput, 11)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("d3", adVarChar, adParamOutput, 11)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("d4", adVarChar, adParamOutput, 11)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("d5", adVarChar, adParamOutput, 11)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("sqft", adInteger, adParamOutput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("LF", adVarChar, adParamOutput, 18)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("occp", adVarChar, adParamOutPut, 500)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("factype", adVarChar, adParamOutPut, 75)
cmd.Parameters.Append prm

cmd.Parameters("string") = building
cmd.Parameters("datefrom") = startdate
cmd.Parameters("dateto") = enddate
cmd.Parameters("utility") = utilityid
'response.write "sp_BLDGstats '"&building&"', '"&startdate&"', '"&enddate&" 23:59:00', "&utilityid
'response.end

if hasdata then cmd.execute
%>
  </td><td>
  <table border=0 cellpadding="2" cellspacing="0" width="200">
  <tr>
    <td colspan="2"><b>Satistical Information</b></td>
  </tr>
  <tr>
    <td valign="top" align="">Five&nbsp;Highest Demand&nbsp;Days:</td>
    <td><a <%if not pdf and isdate(cmd.Parameters("d1")) then%>href="weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid=<%=ypid%>&dutcdate=<%=datevalue(cmd.Parameters("d1"))%>"<%end if%>><%=cmd.Parameters("d1")%></a><br>
        <a <%if not pdf and isdate(cmd.Parameters("d2")) then%>href="weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid=<%=ypid%>&dutcdate=<%=datevalue(cmd.Parameters("d2"))%>"<%end if%>><%=cmd.Parameters("d2")%></a><br>
        <a <%if not pdf and isdate(cmd.Parameters("d3")) then%>href="weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid=<%=ypid%>&dutcdate=<%=datevalue(cmd.Parameters("d3"))%>"<%end if%>><%=cmd.Parameters("d3")%></a><br>
        <a <%if not pdf and isdate(cmd.Parameters("d4")) then%>href="weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid=<%=ypid%>&dutcdate=<%=datevalue(cmd.Parameters("d4"))%>"<%end if%>><%=cmd.Parameters("d4")%></a><br>
        <a <%if not pdf and isdate(cmd.Parameters("d5")) then%>href="weatherReport.asp?pid=<%=pid%>&building=<%=building%>&utilityid=<%=utilityid%>&ypid=<%=ypid%>&dutcdate=<%=datevalue(cmd.Parameters("d5"))%>"<%end if%>><%=cmd.Parameters("d5")%></a>
    </td>
  </tr>
  <tr>
    <td valign="top" align="">Load&nbsp;Factor:</td>
    <td><%=cmd.Parameters("LF")%></td>
  </tr>
  <tr>
    <td valign="top" align="">Square&nbsp;Footage:</td>
    <td><%=formatnumber(cmd.Parameters("sqft"),0)%> sqft</td>
  </tr>
  <tr>
    <td valign="top" align="">Occupancy:</td>
    <td>
    <%dim index
      if ubound(split(cmd.Parameters("occp"),"|"))>-1 then
        if split(cmd.Parameters("occp"),"|")(0)<>0 then
          response.write "<table cellspacing=""0"" cellpadding=""0"">"
          for each i in split(cmd.Parameters("occp"),"|")
            response.write "<tr><td align="""">"
            select case index
            case 0
            response.write "Occupancy:"
            case 1
            response.write "Number&nbsp;Submeter Tenants:"
            case 2
            response.write "Submeter&nbsp;Tenants sqft:"
            case 3
            response.write "Number&nbsp;ERI Tenants:"
            case 4
            response.write "ERI&nbsp;Tenants&nbsp;sqft:"
            case 5
            response.write "Total&nbsp;Tenants:"
            case 6
            response.write "ERI/Submeter Tenants&nbsp;sqft:"
            end select
            response.write "</td><td valign=""bottom"">"
            if index=0 then response.write "&nbsp;"&formatpercent(i,0)&"</td>" else response.write "&nbsp;"&formatnumber(i,0)&"</td></tr>"
            index = index + 1
          next
          response.write "</table>"
        else
          response.write "N/A"
        end if
      else
        response.write "N/A"
      end if
    %>
    </td>
  </tr>
  <tr>
    <td valign="top" align="">Type&nbsp;of&nbsp;Facility:</td>
    <td valign="bottom"><%=cmd.Parameters("factype")%></td>
  </tr>
  <tr>
    <td valign="top" align="">Contact&nbsp;Info:</td>
    <td valign="bottom">
    <%
    i = 0
	
    rst1.open "SELECT * FROM contacts WHERE bldgnum='"&building&"' ORDER BY name"
    do until rst1.eof
      response.write rst1("name")&"<br>"
      rst1.movenext
    loop
    %>
    </td>
  </tr>
  <tr>
    <td valign="top" align="">Type&nbsp;of&nbsp;Meter:</td>
    <td>
    <%
    select case utilityid
      case 1
        response.write "Steam"
      case 2
        response.write "Electric"
      case 6
        response.write "Chilled Water"
    end select
    %>
    </td>
  </tr>
  </table>
  </td></tr></table>
  
  
  </td>
</tr>
<%end if%>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="utilityid" value="<%=utilityid%>">
</form>
<br>
</body>
</html>
