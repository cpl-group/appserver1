<%option explicit%>
<%
dim currentdate, workdate, monthi
currentdate = now()
workdate = Dateserial(year(currentdate)-1, month(currentdate), 1)
monthi=1
%>
<html>
<head>
<title>Calendar</title>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>

<body>
<table>
<%
  do until datediff("d", currentdate, workdate)>0
    if monthi=1 then response.write "<tr>"
    %><td><%makemonth month(workdate), year(workdate)%></td><%
    if monthi=3 then 
      response.write "</tr>"
      monthi = 1
    else
      monthi = monthi + 1
    end if
    workdate = dateadd("m", 1, workdate)
  loop
%>
</table>
</body>
</html>



<%
Sub makemonth(m, y)
  dim monthworkdate, i, rows
  monthworkdate = DateSerial(y,m,1)
  rows = 0
  %>
  <table width="200" bgcolor="#3399cc" cellpadding="1" cellspacing="0" border="0"><tr><td>
    <table bgcolor="#FFFFFF" width="100%">
    <tr><td colspan="7" align="center"><b><%=monthName(m)%>&nbsp;<%=y%></b></td></tr>
    <tr><td width="14%">S</td><td width="14%">M</td><td width="14%">Tu</td><td width="14%">W</td><td width="14%">Th</td><td width="14%">F</td><td width="14%">S</td></tr>
    <%
    i=1
    do until i = weekday(monthworkdate)
      response.write "<td>&nbsp;</td>"
      i = i + 1
    loop
    
    do until month(monthworkdate) <> m
      if weekday(monthworkdate)=1 then response.write "<tr>"
      response.write "<td>"&day(monthworkdate)&"</td>"
      if weekday(monthworkdate)=7 then 
        response.write "</tr>"
        rows = rows + 1
      end if
      monthworkdate = dateadd("d", 1, monthworkdate)
    loop
    if weekday(monthworkdate)<>1 then rows = rows + 1
    if rows<6 then response.write "<tr><td colspan=""7"">&nbsp;</td>"
    %>
    </tr>
    </table>
  </td></tr></table>
<%end Sub%>
