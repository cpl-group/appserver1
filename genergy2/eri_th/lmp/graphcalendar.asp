<%option explicit
server.scripttimeout = 180
%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Calendar</title>
</head>
<%
'COMMENTS
'1/29/2008 N.Ambo corrected propblem for timeout expired error by setting a timeout for the cmd 
'when SP "sp_lmp_getdate" is called

function getTitle()
	dim sql
	if meterid<>"" and meterid<>"0" then 
		sql = "SELECT meternum as title FROM meters WHERE meterid='"&meterid&"'"
	elseif billingid<>"" and billingid<>"0" then 
		sql = "SELECT tName as title FROM tblleases WHERE billingid='"&billingid&"'"
	elseif bldg<>"" and bldg <>"0" then 
		sql = "SELECT strt as title FROM buildings WHERE bldgnum='"&bldg&"'"
	elseif pid<>"" and pid <>"0" then 
		sql = "SELECT name as title FROM portfolio WHERE id="&pid
	else 
	end if
	rst1.open sql, cnn1
	if not rst1.eof then getTitle = rst1(0)
	rst1.close
end function

function senddata(mday)
	dim ki, tempstr, monthstr, daystr
	tempstr = "<img onclick=""loadlmpchart('"&month(cdate(dDate))&"/"&mday&"/"&year(cdate(dDate))&"')"" width=""70"" height=""40"" src=""MakeMiniLmp.asp?scale="&maxscale&"&data="
	for i = 0 to 11
		if isnull(montharray(mday,i)) or trim(montharray(mday,i))="" then montharray(mday,i) = 0
		tempstr=tempstr&montharray(mday,i)&","
	next
	tempstr=left(tempstr,len(tempstr)-1)&"&day="&mday
  if trim(billingid)="" and meterid<>"0" then
  	'tempstr = tempstr&""" alt=""Day peak demand: "&formatnumber(daypeakdemands(mday,2),2)&", "&usage&": "&formatnumber(cdbl(daypeakdemands(mday,1)),0)&",  at "&mid(daypeakdemands(mday,3),instr(daypeakdemands(mday,3)," ")+1)&"."
  	'if cint(MdayOne) = mday then tempstr = tempstr&" Peak demand for "&MperiodOne&" is "&MpeakOne&"kw."
  	'if cint(MdayTwo) = mday then tempstr = tempstr&" Peak demand for "&MperiodTwo&" is "&MpeakTwo&"kw."
  end if
	tempstr = tempstr&""">"
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

Dim bldg, meterid, tenantmeter, pulsetable, utility, usage, units, groupname, lmptype, lmpcode, indiwindow, billingid,luid, title, pid
Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table
dim calendarend
const interval = 1 'sets the interval of the display

bldg = trim(request("bldg"))
meterid = trim(request("meterid"))
billingid = trim(request("billingid"))
tenantmeter = trim(request("tenantmeter"))
utility = trim(request("utility"))
indiWindow = trim(request("indiWindow"))
pid = trim(request("pid"))

dim rst1, rst2, cnn1, strsql, prm, cmd,Andorst,cnn666
set cmd = server.createobject("ADODB.Command")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
set Andorst = server.createobject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.connection")
'set cnn666 = server.createobject("ADODB.connection")

cnn1.open getConnect(pid,bldg,"billing")
cnn1.CursorLocation = adUseClient
cnn1.commandtimeout = 300 '1/29/2008 nambo modified length of timeout
dim divisor
rst1.open "SELECT * FROM tblutility WHERE utilityid="&utility, getConnect(pid,bldg,"dbCore")
if not rst1.eof then usage = rst1("measure") else usage = "KWH"
rst1.close
if cint(utility) = 17 then
	usage="pulse"
end if

title = getTitle()

if trim(meterid)<>"" then
    lmptype="m"
    lmpcode = meterid
elseif trim(billingid)<>"" then
    lmptype="L"
    rst1.open "SELECT BillingName, leaseutilityid FROM tblLeases l , tblleasesutilityprices lup WHERE l.billingid=lup.billingid AND lup.utility="&utility&" AND l.billingId="&billingid, cnn1
    if not(rst1.eof) then
'        tenantname = rst1("BillingName")
        luid = cint(rst1("leaseutilityid"))
        rst1.close
    end if
    lmpcode=luid
elseif trim(bldg)<>"" then
    lmptype="b"
    lmpcode=bldg
    rst1.open "SELECT meterid FROM meters m, tblleasesutilityprices lup WHERE lup.leaseutilityid=m.leaseutilityid and lmp=1 and lup.utility="&utility&" and bldgnum='"&bldg&"'", cnn1
    if not rst1.eof then meterid = rst1("meterid") else meterid=0
    rst1.close
elseif trim(pid)<>"" then
    lmptype="p"
    lmpcode=pid
end if

' Get selected date.  There are two ways to do this.
' First check if we were passed a full date in RQS("date").
' If so use it, if not look for seperate variables, putting them togeter into a date.
' Lastly check if the date is valid...if not use today

If IsDate(Request("date")) Then
	dDate = CDate(Request("date"))
Else
	dDate = Date()
End If
dDate = month(dDate)&"/1/"&year(dDate)
calendarend = dateadd("m",1,dDate)
'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
'this will be for meters -->if supplied with meter and building and no luid

dim strsqlmax
if trim(groupname)<>"" then
  if trim(billingid)<>"" and tenantmeter<>"1" then
  	strsql = "SELECT datepart(hour,p.date) as hour, datepart(day,p.date) as day, Strt, p.date as date, sum(p."&usage&") as kwh FROM (["& pulsetable &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.leaseutilityid="&billingid&" GROUP BY datepart(hour,p.date), datepart(day,p.date), Strt, p.date ORDER BY date"
  	strsqlmax = "SELECT top 1 sum(p."&usage&") as kwh FROM (["& pulsetable &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.leaseutilityid="&billingid&" GROUP BY datepart(hour,p.date), datepart(day,p.date), Strt, p.date ORDER BY kwh desc"
  else
  	strsql = "SELECT meters.meterid, datepart(hour,p.date) as hour, datepart(day,p.date) as day, Meters.MeterNum, Strt, Meters.MeterId, p.date as date, (datepart(hour,p.date)*100)+datepart(minute,p.date) as time, (meters.multiplier) as multiplier, p."&usage&" FROM (["& pulsetable &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.meterid="&meterid&" ORDER BY date"
  	strsqlmax = "SELECT max(p."&usage&") as kwh FROM (["& pulsetable &"] p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum WHERE datepart(month,p.date)="&month(dDate)&" AND datepart(year,p.date)="&year(dDate)&" AND ((datepart(hour,p.date)*100)+datepart(minute,p.date)) % "&interval*100&"=0 AND meters.meterid="&meterid
  end if
  'response.write strsqlmax
  'response.end
  rst1.open strsqlmax, cnn1
end if

if trim(groupname)="" then
		Do While cmd.Parameters.Count > 0
		cmd.Parameters.Delete(0)
		Loop
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
    Set prm = cmd.CreateParameter("utility", adInteger, adParamInput, 2)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("interval", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("title", adVarChar, adParamOutPut, 30)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("max", adDouble, adParamOutPut, 18,2)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("sum", adDouble, adParamOutPut, 18,2)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("peakdemand", adInteger, adParamOutPut)
    cmd.Parameters.Append prm
    
    cmd.Parameters("from")		= ddate
    cmd.Parameters("to")		= dateadd("m",1,ddate)
    cmd.Parameters("code")		= lmptype
    cmd.Parameters("string")		= lmpcode
    cmd.Parameters("utility")		= utility
    cmd.Parameters("interval")		= 1
    'response.write "exec sp_LMPDATA '"&ddate&"','"&dateadd("m",1,ddate)&"','"&lmptype&"','"&lmpcode&"',"&utility&",1,0<br>"
    'response.end
    set rst1 = cmd.execute
end if

if isnumeric(cmd.Parameters("max")) then maxscale = cdbl(cmd.Parameters("max")) else maxscale = 1

'get the peak demands for every day
dim peakdemandDay(31)
cnn1.CursorLocation = adUseClient
' specify stored procedure to run
cmd.CommandText = "sp_lmp_getdate"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 300 '1/29/2008 n.ambo added timeout for command
'cnn666.open getConnect(pid,bldg,"billing")
'Set cmd.ActiveConnection = cnn666


Do While cmd.Parameters.Count > 0
cmd.Parameters.Delete(0)
Loop
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




'set input parameters
cmd.Parameters("meterid")	= meterid
cmd.Parameters("BY")		= year(dDate)
cmd.Parameters("month")		= month(dDate)
'response.write "exec sp_lmp_getdate "&meterid&","&year(dDate)&","&month(dDate)&"<BR>"
'response.write cnn1
'response.end
'execution

dim daypeakdemands(31,3)
'if trim(billingid)="" and meterid<>"0" and not(rst1.eof) then
if meterid<>"0" and meterid<>"" and not(rst1.eof) then
  set rst2 = cmd.execute
 
  
if rst2.state then 
  while not rst2.eof
  	daypeakdemands(cint(rst2("day")),3) = rst2("time")
  	daypeakdemands(cint(rst2("day")),1) = rst2("KWH")
  	daypeakdemands(cint(rst2("day")),2) = rst2("KW")
  	rst2.movenext
  wend 
  dim maxscale'this is the static scale size
  'rst2.close
  
  end if
end if
%>
<script>
function lloadlmpchart(cdate)
{	temp="lmpload.asp?meterid=<%=meterid%>&startdate="+cdate+"&bldg=<%=bldg%>&billingid=<%=billingid%>&tenantmeter=<%=tenantmeter%>&interval=1&utility=<%=utility%>";
	if(document.popups.popups.checked){
   window.open(temp,"","statusbar=0,menubar=0,scrollbars=no,HEIGHT=310,WIDTH=600");
   }else
   {  lloadlmpchart(cdate)
   }
}

function loadlmpchart(cdate)
{	if(document.forms['popups'].popups.checked)
	{	temp="lmpload.asp?meterid=<%=meterid%>&startdate="+cdate+"&bldg=<%=bldg%>&billingid=<%=billingid%>&tenantmeter=<%=tenantmeter%>&interval=0&utility=<%=utility%>";
		window.open(temp,"","statusbar=0,menubar=0,scrollbars=no,HEIGHT=310,WIDTH=600");
	}else
	{	parent.document.forms[0].startdate.value = cdate;
		parent.loadchart();
	}
}
</script>
<body link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" <%if indiWindow="" then%>onload="parent.shownav('calnav');parent.closeLoadBox('loadFrame1');"<%end if%>>
<table border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td style="font-family:arial;font-size:11px;" align="right"><b><%=title%></b></td></tr>
<TR><TD><TABLE BORDER="1" CELLSPACING="0" CELLPADDING="0" bordercolor="#000000" BGCOLOR="#99CCFF" width="650" height="277">
        <TR bgcolor="#000000"> 
          <TD ALIGN="center" COLSPAN="7">
            <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
              <TR style="font-family:arial;font-size:11px;color:white;">
			  	
            <td>&nbsp;</td>
                <TD ALIGN="center"><B><%= MonthName(Month(dDate)) & "  " & Year(dDate) %></B></TD>
			  	
            <td align="right">&nbsp;</td>
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
dim breakloop, i, dayarray(), tdColor, montharray(31,23), daypeakarray(31)
redim dayarray(cINT(24/interval)-1)
breakloop = false
do until rst1.eof
	for i = 0 to 23 Step interval
		dim tempkwh1, tempkwh2
		if not(rst1.eof) and (cINT(i)=hour(rst1("date"))) then
			montharray(day(rst1("date")), (i/2)) = trim(rst1(usage))
			rst1.movenext
		else
			montharray(day(rst1("date")), (i/2)) = 0
		end if
		if rst1.eof then i = 23
	next
'	daypeakarray(cINT(rst1("day")) = tempdaypeak
loop


if bldg<>"" and bldg<>"0" then
	Do While cmd.Parameters.Count > 0
	cmd.Parameters.Delete(0)
	Loop
	cmd.ActiveConnection = cnn1
	cmd.CommandText = "sp_Period_PDEMAND"
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("meters", adVarChar, adParamInput, 400)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("lease", adVarChar, adParamInput, 500)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 200)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("util", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("coincident", adInteger, adParamOutput)
	cmd.Parameters.Append prm
	cmd.Parameters("from") = dDate
	cmd.Parameters("to") = calendarend
	cmd.Parameters("meters") = meterid
	cmd.Parameters("lease") = Luid
	cmd.Parameters("bldg") = bldg
	cmd.Parameters("util") = utility
	if trim(cmd.Parameters("meters"))="" then cmd.Parameters("meters") = "0"
	if trim(cmd.Parameters("bldg"))="" then cmd.Parameters("bldg") = "0"
	if trim(cmd.Parameters("lease"))="" then cmd.Parameters("lease") = "0"
	'response.write "exec sp_Period_PDEMAND '"&cmd.Parameters("from")&"', '"&cmd.Parameters("to")&"', "&cmd.Parameters("meters")&", "&cmd.Parameters("lease")&", '"&cmd.Parameters("bldg")&"', "&cmd.Parameters("util")&", 0"
	'response.end
	Set rst2 = cmd.Execute
	
end if


iCurrent = 1
iPosition = iDOW
Do While iCurrent <= iDIM
	' If we're at the begginning of a row then write TR
	If iPosition = 1 Then
		Response.Write vbTab & "<TR>" & vbCrLf
	End If
	
	' If the day we're writing is the selected day then highlight it somehow.
	'	If iCurrent = Day(dDate) Then
	if not rst2.eof then
		If day(rst2("date")) = iCurrent Then
			tdColor = "#990000"
			rst2.movenext
		Else
			tdColor = "#FFFFFF"
		End If
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
rst1.close

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
