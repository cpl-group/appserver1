<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim startdate, enddate, cstartdate, cenddate, currentdate, tmonth, tday, tyear, fmonth, fday, fyear, pdf, selected, groupname, cmonthst, cyearst, cmonthed, cyeared, weekdays, weekends, holiday
pdf = request("pdf")
if trim(pdf) = "yes" then pdf = true
groupname = request("groupname")
cmonthst = request("cmonthst")
cyearst = request("cyearst")
cmonthed = request("cmonthed")
cyeared = request("cyeared")
weekdays = server.htmlencode(request("weekdays"))
weekends = server.htmlencode(request("weekends"))
holiday = server.htmlencode(request("holiday"))
if cmonthst = "" or cyearst = "" or cmonthed = "" or cyeared = "" then
  cmonthst = datepart("m",date())
  cyearst = datepart("yyyy",date())
  cmonthed = datepart("m",date())
  cyeared = datepart("yyyy",date())
end if

tmonth = request("tmonth")
tday = request("tday")
tyear = request("tyear")
fmonth = request("fmonth")
fday = request("fday")
fyear = request("fyear")
if tmonth = "" or tday = "" or tyear = "" or fmonth = "" or fday = "" or fyear = "" then
  tmonth = datepart("m",date())
  tday = datepart("d",date())
  tyear = datepart("yyyy",date())
  fmonth = datepart("m",date())
  fday = datepart("d",date())
  fyear = datepart("yyyy",date())
end if
if datediff("d",datevalue(tmonth&"/"&tday&"/"&tyear),datevalue(fmonth&"/"&fday&"/"&fyear))<0 then 
  fmonth = tmonth
  fday = tday
  fyear = tyear
end if
cstartdate = cmonthst&"/1/"&cyearst
cenddate = cmonthed&"/"&day(dateadd("d",-1,dateadd("m",1,datevalue(cmonthed&"/1/"&cyeared))))&"/"&cyeared
currentdate = cstartdate

dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

%>
<html>
<head>
<title>Average Daily Usage Analysis</title>
</head>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form name="aduanalysis" method="get" action="AverageDailyUsage.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Average Daily Usage Analysis Report<%if pdf then response.write " (Page 1)"%></span></td>
  <%if not(pdf) then%><td align="right"><input type="button" value="Print Current View" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport.asp?qs=<%=server.URLEncode(request.servervariables("SCRIPT_NAME")&"?"&request.servervariables("QUERY_STRING"))%>','','width=600,height=400,resizable=yes');"></td><%end if%>
</tr>
</table>
<table width="100%" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%if not pdf then%>Show report for:<%else%><b>Show report for:</b><%end if%>&nbsp;
  <%
  dim i
  if not pdf then response.write "<select name=""groupname"" onchange=""submit();""><option value=""0"">Select a Input Group</option>"
  rst1.open "SELECT * FROM [group] WHERE type=1", cnn1
  do until rst1.eof
    selected=""
    if trim(groupname) = trim(rst1("groupname")) then selected=" SELECTED"
    if not(pdf) then
      response.write "<option value="""&rst1("groupname")&""""&selected&">"&rst1("groupname")&"</option>"
    elseif selected=" SELECTED" then
      response.write rst1("groupname")
    end if
    rst1.movenext
  loop
  rst1.close
  if not pdf then response.write "</select>"
  %>
<%if not(pdf) then%> 
Dates:
  <select name="cmonthst" onchange="submit();">
    <%
      for i = 1 to 12
        %><option value="<%=i%>"<%if cint(cmonthst)=i then response.write " SELECTED"%>><%=left(monthname(i),3)%></option><%
      next
    %>
  </select>
  <select name="cyearst" onchange="submit();">
    <%
      rst1.open "SELECT distinct datepart(year,datestart) as year FROM billyrperiod ORDER BY datepart(year,datestart) desc"
      do until rst1.eof
        %><option value="<%=rst1("year")%>"<%if cint(cyearst)=cint(rst1("year")) then response.write " SELECTED"%>><%=rst1("year")%></option><%
        rst1.movenext
      loop
      rst1.close
    %>
  </select>&nbsp;to&nbsp;<select name="cmonthed" onchange="submit();">
    <%
      for i = 1 to 12
        %><option value="<%=i%>"<%if cint(cmonthed)=i then response.write " SELECTED"%>><%=left(monthname(i),3)%></option><%
      next
    %>
  </select>
  <select name="cyeared" onchange="submit();">
    <%
      rst1.open "SELECT distinct datepart(year,datestart) as year FROM billyrperiod ORDER BY datepart(year,datestart) desc"
      do until rst1.eof
        %><option value="<%=rst1("year")%>"<%if cint(cyeared)=cint(rst1("year")) then response.write " SELECTED"%>><%=rst1("year")%></option><%
        rst1.movenext
      loop
      rst1.close
    %>
  </select>
<%else%>
  <%if cmonthst=cmonthed and cyeared=cyearst then%>
    <b>Date:</b>&nbsp;<%=monthname(cmonthst)%>&nbsp;<%=cyearst%>
  <%else%>
    <b>Dates:</b>&nbsp;<%=monthname(cmonthst)%>&nbsp;<%=cyearst%>&nbsp;to&nbsp;<%=monthname(cmonthed)%>&nbsp;<%=cyeared%>
  <%end if%>
<%end if%>
  </td>
</tr>
<tr>
  <td bgcolor="#eeeeee">
  </td>
</tr>

</table>
<%if trim(groupname)<>"" then%>
<center><font size="3"><b>Usage Calendar</b></font></center>
<TABLE BORDER="1" CELLSPACING="0" CELLPADDING="0" bordercolor="#000000" BGCOLOR="#99CCFF" width="650" align="center">
<TR style="background-color:#0000CC;color:#FFFFFF; font-family:Arial, Helvetica, sans-serif; font-size:12"> 
  <TD width="13%" ALIGN="center"><B>Sun</B></TD>
  <TD width="13%" ALIGN="center"><B>Mon</B></TD>
  <TD width="13%" ALIGN="center"><B>Tue</B></TD>
  <TD width="13%" ALIGN="center"><B>Wed</B></TD>
  <TD width="13%" ALIGN="center"><B>Thu</B></TD>
  <TD width="13%" ALIGN="center"><B>Fri</B></TD>
  <TD width="13%" ALIGN="center"><B>Sat</B></TD>
  <TD width="1%" ALIGN="center">&nbsp;&nbsp;</TD>
</TR>
<tr><td colspan="8">
<div style="overflow: auto; width:666; height:280" align="center">
<TABLE BORDER="1" CELLSPACING="0" CELLPADDING="0" bordercolor="#000000" BGCOLOR="#99CCFF" width="650" align="center">
<%
dim buffers, groupCol, calendararray(11), index, max
rst1.open "SELECT top 1 * FROM ["&groupname&"]", cnn1
groupCol = rst1.fields.Item(0).Name
rst1.close
rst1.open "SELECT sum(["&groupCol&"]) as usage, convert(datetime,left(date,11)) as d, (datepart(hour,date)-(datepart(hour,date) % 2))/2 as h FROM ["&groupname&"] WHERE date>='"&cstartdate&"' and date<='"&cenddate&"' GROUP BY convert(datetime,left(date,11)), (datepart(hour,date)-(datepart(hour,date) % 2))/2 ORDER BY d, h", cnn1
do until datediff("d",cenddate,currentdate)>0
  if day(currentdate)=1 then max = 0
  for index=0 to 11 
    calendararray(index) = 0
  next
  if not rst1.eof then
    do while not rst1.eof and index<>-1
      if datediff("d",datevalue(rst1("d")),currentdate)=0 then 
        calendararray(int(rst1("h"))) = rst1("usage")
        if cdbl(rst1("usage"))>max then max = cdbl(rst1("usage"))
        rst1.movenext
      else 
        index = -1
      end if
    loop
  end if
  if day(currentdate)=1 then
    response.write "<TR bgcolor=""#000000""><TD ALIGN=""center"" COLSPAN=""7""><B><font face=""arial"" color=""white"">"&monthname(month(currentdate))&"&nbsp;"&year(currentdate)&"</font></B></TD></TR>"
    for buffers = 2 to weekday(currentdate)
      response.write "<TD BGCOLOR=""#FFFFFF"">&nbsp;</TD>"
    next
  elseif weekday(currentdate)=1 then
    response.write "<TR>"
  end if
	response.write "<TD BGCOLOR=""#FFFFFF"" align=""center""><img width=""70"" height=""40"" src=""/genergy2/eri_th/lmp/MakeMiniLmp.asp?scale="&max&"&data="&join(calendararray,",")&"&day="&day(currentdate)&"""></TD>"
  if weekday(currentdate)=7 then response.write "</TR>"
 	currentdate = dateadd("d",1,currentdate)
loop
rst1.close
%>
</tr></table>
</div>
</td></tr></table>
<WxPrinter PageBreak>


<%if pdf then%>


<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Average Daily Usage Analysis Report (Page 2)</span></td>
</tr>
</table>
<table width="100%" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%if not pdf then%>Show report for:<%else%><b>Show report for:</b><%end if%>&nbsp;
  <%=groupname%>
  <%if cmonthst=cmonthed and cyeared=cyearst then%>
    <b>Date:</b>&nbsp;<%=monthname(cmonthst)%>&nbsp;<%=cyearst%>
  <%else%>
    <b>Dates:</b>&nbsp;<%=monthname(cmonthst)%>&nbsp;<%=cyearst%>&nbsp;to&nbsp;<%=monthname(cmonthed)%>&nbsp;<%=cyeared%>
  <%end if%>
  </td>
</tr>
<tr>
  <td bgcolor="#eeeeee">
  </td>
</tr>
</table>




<%end if%>
<%if not(pdf) then%><center><input type="button" value="Load Averages" onclick="submit()"></center><%end if%>
<table width="650" align="center">
<tr>
    <td align="right" valign="top">
    <%if not(pdf) or trim(weekdays)<>"" then%>Select Weekday:<%end if%><br>
<%if not(pdf) then%>
    Monday<input name="weekdays" type="checkbox" value="2"<%if instr(weekdays,"2")<>0 then response.write " CHECKED"%>><br>
    Tuesday&nbsp;<input name="weekdays" type="checkbox" value="3"<%if instr(weekdays,"3")<>0 then response.write " CHECKED"%>><br>
    Wednesday&nbsp;<input name="weekdays" type="checkbox" value="4"<%if instr(weekdays,"4")<>0 then response.write " CHECKED"%>><br>
    Thursday&nbsp;<input name="weekdays" type="checkbox" value="5"<%if instr(weekdays,"5")<>0 then response.write " CHECKED"%>><br>
    Friday&nbsp;<input name="weekdays" type="checkbox" value="6"<%if instr(weekdays,"6")<>0 then response.write " CHECKED"%>><br>
<%else
      for each i in split(weekdays,",")
        response.write weekdayname(i)&"<br>"
      next
end if%>
    </td>
<%if trim(weekdays)<>"" then%>
    <td <%if not(pdf) then%>align="right"<%end if%>><img src="AverageDailyLMP.asp?startdate=<%=cstartdate%>&enddate=<%=cenddate%>&day=<%=server.urlencode(weekdays)%>&groupname=<%=groupname%>" border="0" width="525" height="195"></td>
<%end if%>
</tr>
<tr><td align="right" valign="top">
    <%if not(pdf) or trim(weekends)<>"" then%>Selected Weekend:<%end if%><br>
<%if not(pdf) then%>
    Saturday&nbsp;<input name="weekends" type="checkbox" value="7"<%if instr(weekends,"7")<>0 then response.write " CHECKED"%>><br>
    Sunday&nbsp;<input name="weekends" type="checkbox" value="1"<%if instr(weekends,"1")<>0 then response.write " CHECKED"%>><br>
<%else
      for each i in split(weekends,",")
        response.write weekdayname(i)&"<br>"
      next
end if%>
    </td>
<%if trim(weekends)<>"" then%>
    <td <%if not(pdf) then%>align="right"<%end if%>><img src="AverageDailyLMP.asp?startdate=<%=cstartdate%>&enddate=<%=cenddate%>&day=<%=server.urlencode(weekends)%>&groupname=<%=groupname%>" border="0" width="525" height="195"></td>
<%end if%>
</tr>
<tr><td align="right" valign="top">
    <%if not(pdf) or trim(holiday)<>"" then%>Select Holiday:<%end if%><br>
<%if not(pdf) then%>
    <%
      rst1.open "SELECT distinct holiday from holidaysch where date>='"&cstartdate&"' and date<='"&cenddate&"' order by holiday", cnn1
      do until rst1.eof
        %><%=trim(rst1("holiday"))%>&nbsp;<input name="holiday" value="<%=trim(rst1("holiday"))%>" type="checkbox"<%if instr(holiday,trim(rst1("holiday")))<>0 then response.write " CHECKED"%>><br><%
        rst1.movenext
      loop
      rst1.close
else
      for each i in split(holiday,",")
        response.write i&"<br>"
      next
end if%>
    </td>
<%if trim(holiday)<>"" then%>
    <td <%if not(pdf) then%>align="right"<%end if%>><img src="AverageDailyLMP.asp?startdate=<%=cstartdate%>&enddate=<%=cenddate%>&holiday=<%=server.urlencode(holiday)%>&groupname=<%=groupname%>" border="0" width="525" height="195"></td>
<%end if%>
</tr>
<tr><td></td><td width="525"></td></tr>
</table>
<%end if'from if trim(groupname)<>""%>
</form>
</body>
</html>