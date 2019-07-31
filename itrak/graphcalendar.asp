<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
function senddata(mday)
	dim ki, tempstr, monthstr, daystr
	tempstr = "<img onclick=""loadlmpchart('"&month(cdate(dDate))&"/"&mday&"/"&year(cdate(dDate))&"')"" width=""70"" height=""40"" src=""MakeMiniLmp.asp?scale="&maxscale&"&data="
	for i = 0 to 11
		if montharray(mday,i)="" then montharray(mday,i) = 0
		tempstr=tempstr&montharray(mday,i)&","
	next
	tempstr=left(tempstr,len(tempstr)-1)&"&day="&mday&""""
	if trim(luid)="" then
    tempstr = tempstr&" alt=""Day peak demand: "&formatnumber(daypeakdemands(mday,2),1)&", kwh: "&formatnumber(daypeakdemands(mday,1),1)&",  at "&mid(daypeakdemands(mday,3),instr(daypeakdemands(mday,3)," ")+1)&"."
  	if cint(MdayOne) = mday then tempstr = tempstr&" Peak demand for "&MperiodOne&" is "&MpeakOne&"kw."
  	if cint(MdayTwo) = mday then tempstr = tempstr&" Peak demand for "&MperiodTwo&" is "&MpeakTwo&"kw."
    tempstr = tempstr&""""
  end if
	tempstr = tempstr&">"
	senddata = tempstr
end function

Function GetDaysInMonth(iMonth, iYear)
	Dim dTemp
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	GetDaysInMonth = Day(dTemp)
End Function
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
	Dim dTemp
	dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
	GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function

Function SubtractOneMonth(dDate)
	SubtractOneMonth = DateAdd("m", -1, dDate)
End Function

Function AddOneMonth(dDate)
	AddOneMonth = DateAdd("m", 1, dDate)
End Function
'############################
'# End Function Declaration #
'############################

dim rst1, rst2, cnn1, strsql, prm, cmd
set cmd = server.createobject("ADODB.Command")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.connection")
cnn1.open application("cnnstr_genergy1")

Dim building, meterid, luid, tenantmeter, theL
Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table
const interval = 2 'sets the interval of the display

building = request.querystring("b")
meterid = request.querystring("m")
luid = request.querystring("luid")
tenantmeter = request.querystring("tenantmeter")
if tenantmeter<>"1" then theL = "L"

' Get selected date.  There are two ways to do this.
' First check if we were passed a full date in RQS("date").
' If so use it, if not look for seperate variables, putting them togeter into a date.
' Lastly check if the date is valid...if not use today
If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
		' The elegant solution for those of you running IIS4
		'If Request.QueryString.Count <> 0 Then Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
	End If
End If

'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
'this will be for meters -->if supplied with meter and building and no luid
dim strsqlmax
if trim(luid)<>"" and tenantmeter<>"1" then
	strsql = "SELECT datepart(hour,p.date) as hour, datepart(day,p.date) as day, Strt, p.date as date, sum(p.kwh) as kwh FROM ([pulse_"& building &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.leaseutilityid="&luid&" GROUP BY datepart(hour,p.date), datepart(day,p.date), Strt, p.date ORDER BY date"
	strsqlmax = "SELECT top 1 isnull(sum(p.kwh),10) as kwh FROM ([pulse_"& building &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.leaseutilityid="&luid&" GROUP BY datepart(hour,p.date), datepart(day,p.date), Strt, p.date ORDER BY kwh desc"
else
	strsql = "SELECT meters.meterid, datepart(hour,p.date) as hour, datepart(day,p.date) as day, Meters.MeterNum, Strt, Meters.MeterId, p.date as date, (datepart(hour,p.date)*100)+datepart(minute,p.date) as time, (meters.multiplier) as multiplier, p.kwh FROM ([pulse_"& building & theL &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.meterid="&meterid&" ORDER BY date"
	strsqlmax = "SELECT isnull(max(p.kwh),10) as kwh FROM ([pulse_"& building & theL & "] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.meterid="&meterid
end if
rst1.open strsqlmax, cnn1
if not rst1.eof then
	maxscale = cDbl(rst1("kwh"))*4
else
	maxscale = 0
end if
rst1.close
'response.write strsql
'response.end
rst1.open strsql, cnn1

'get the peak demands for every day
dim peakdemandDay(31)
cnn1.CursorLocation = adUseClient
' specify stored procedure to run
cmd.CommandText = "sp_lmp_getdate"
cmd.CommandType = adCmdStoredProc
'input params
Set prm = cmd.CreateParameter("meterid", adVarChar, adParamInput, 5)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("BY", adVarChar, adParamInput, 10)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("month", adVarChar, adParamInput, 3)
cmd.Parameters.Append prm
'output params
Set prm = cmd.CreateParameter("d1", adVarchar, adParamOutput, 21)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bpt1", adVarchar, adParamOutput, 18)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("p1", adVarchar, adParamOutput, 18)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bp1i", adVarchar, adParamOutput, 50)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("d2", adVarchar, adParamOutput, 21)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bpt2", adVarchar, adParamOutput, 18)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("p2", adVarchar, adParamOutput, 18)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bp2i", adVarchar, adParamOutput, 50)
cmd.Parameters.Append prm
'set connection
Set cmd.ActiveConnection = cnn1
'set input parameters
cmd.Parameters("meterid")	= meterid
cmd.Parameters("BY")		= year(dDate)
cmd.Parameters("month")		= month(dDate)
'cmd.Parameters("p2")		= month(dDate)
'response.write meterid&"|"&year(dDate)&"|"&month(dDate)&"<BR>"
'response.end
'execution
dim daypeakdemands(31,3)
if luid="" then 
  set rst2 = cmd.execute
  do until rst2.EOF
  	daypeakdemands(cint(rst2("day")),3) = rst2("time")
  	daypeakdemands(cint(rst2("day")),1) = rst2("KWH")
  	daypeakdemands(cint(rst2("day")),2) = rst2("KW")
  	rst2.movenext
  loop
  dim MpeakOne, Mtimeone, MdayOne, MpeakTwo, Mtimetwo, Mdaytwo, MperiodOne, MperiodTwo
  MdayOne = cmd.Parameters("d1")
  MtimeOne = cmd.Parameters("bpt1")
  MpeakOne = cmd.Parameters("p1")
  MperiodOne = cmd.Parameters("bp1i")
  MdayTwo = cmd.Parameters("d2")
  MtimeTwo = cmd.Parameters("bpt2")
  MpeakTwo = cmd.Parameters("p2")
  MperiodTwo = cmd.Parameters("bp2i")
  dim maxscale'this is the static scale size
  'if MpeakOne>MpeakTwo then
  '	maxscale=MpeakOne
  'else
  '	maxscale=MpeakTwo
  'end if
  rst2.close
else
  MdayOne = 0
  MtimeOne = 0
  MpeakOne = 0
  MperiodOne = 0
  MdayTwo = 0
  MtimeTwo = 0
  MpeakTwo = 0
  MperiodTwo = 0
end if
%>
<html>
<head>
<title>Calendar</title>
</head>
<script>
function loadlmpchart(cdate)
{	var d = cdate
	var pd = parent.dateAddDays(cdate,-1);
	var nd = parent.dateAddDays(cdate,1);
	if(document.forms['popups'].popups.checked)
	{	var m = parent.document.forms['form1'].m.value;
		var b = parent.document.forms['form1'].b.value;
		var l = parent.document.forms['form1'].luid.value;
		var i = parent.document.forms['form1'].zoom.value;
		var lmp = parent.document.forms['form1'].lmp.value;
		var tenantmeter = parent.document.forms['form1'].tenantmeter.value;
		temp="lmpload2.asp?m="+m+"&d="+d+"&b="+b+"&s=&e=&luid="+l+"&lmp="+lmp+"&tenantmeter="+tenantmeter+"&i="+i;
		window.open(temp,"","statusbar=0,menubar=0,scrollbars=no,HEIGHT=310,WIDTH=600");
	}else
	{	parent.document.forms[0].d.value = d;
		parent.document.forms[0].pd.value = pd;
		parent.document.forms[0].nd.value = nd;
		parent.loadchart();
	}
}
</script>
<body link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" onload="parent.shownav('calnav');parent.closeLoadBox('loadFrame1');">
<table border="0" cellspacing="0" cellpadding="0" align="center">
<TR><TD><TABLE BORDER="1" CELLSPACING="0" CELLPADDING="0" bordercolor="#000000" BGCOLOR="#99CCFF" width="650" height="277">
        <TR bgcolor="#000000"> 
          <TD ALIGN="center" COLSPAN="7">
            <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
              <TR> 
                <TD ALIGN="center"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><B><%= MonthName(Month(dDate)) & "  " & Year(dDate) %></B></font></TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
        <TR style="background-color:#0000CC;color:#FFFFFF; font-family:Arial, Helvetica, sans-serif; font-size:12"> 
          <TD ALIGN="center"><B>Sun</B></TD>
          <TD ALIGN="center"><B>Mon</B></TD>
          <TD ALIGN="center"><B>Tue</B></TD>
          <TD ALIGN="center"><B>Wed</B></TD>
          <TD ALIGN="center"><B>Thu</B></TD>
          <TD ALIGN="center"><B>Fri</B></TD>
          <TD ALIGN="center"><B>Sat</B></TD>
        </TR>
        <%
' Write spacer cells at beginning of first row if month doesn't start on a Sunday.
If iDOW <> 1 Then
	Response.Write vbTab & "<TR>" & vbCrLf
	iPosition = 1
	Do While iPosition < iDOW
		Response.Write vbTab & vbTab & "<TD BGCOLOR=#FFFFFF>&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
End If

'getting graph data and putting into 2D array
dim breakloop, i, dayarray(), tdColor, montharray(31,11), daypeakarray(31)
redim dayarray(cINT(24/interval)-1)
breakloop = false
do until rst1.eof
	for i = 0 to 23 Step interval
		dim tempkwh1, tempkwh2
		if not(rst1.eof) and (cINT(i)=cINT(rst1("hour"))) then
			montharray(cINT(rst1("day")), (i/2)) = trim(rst1("kwh"))
			rst1.movenext
		else
			montharray(cINT(rst1("day")), (i/2)) = 0
		end if
		if rst1.eof then i = 23
	next
'	daypeakarray(cINT(rst1("day")) = tempdaypeak
loop

iCurrent = 1
iPosition = iDOW
Do While iCurrent <= iDIM
	' If we're at the begginning of a row then write TR
	If iPosition = 1 Then
		Response.Write vbTab & "<TR>" & vbCrLf
	End If
	
	' If the day we're writing is the selected day then highlight it somehow.
'	If iCurrent = Day(dDate) Then
	If ((cint(MdayOne) = iCurrent) or (cint(MdayTwo) = iCurrent)) Then
		tdColor = "#990000"
	Else
		tdColor = "#FFFFFF"
	End If
	Response.Write vbTab & vbTab & "<TD BGCOLOR="""& tdcolor &""" align=""center"">"&senddata(iCurrent)&"</TD>" & vbCrLf
	
	' If we're at the endof a row then write /TR
	If iPosition = 7 Then
		Response.Write vbTab & "</TR>" & vbCrLf
		iPosition = 0
	End If
	
	' Increment variables
	iCurrent = iCurrent + 1
	iPosition = iPosition + 1
Loop

If iPosition <> 1 Then
	Do While iPosition <= 7
		Response.Write vbTab & vbTab & "<TD BGCOLOR=#FFFFFF>&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
	Response.Write vbTab & "</TR>" & vbCrLf
End If
%>
</tr></table>
<table border="0" cellspacing="0" cellpadding="0"><tr><td><form name="popups"></td></tr></table>
<span style="font-family:Arial, Helvetica, sans-serif; font-size:9">Shows popups 
  <input type="checkbox" name="popups" value="1"<%if request.querystring("popups")="true" then response.write " checked"%> checked>
  </span>
<table border="0" cellspacing="0" cellpadding="0"><tr><td></form></td></tr></table>
</body>
</html>
