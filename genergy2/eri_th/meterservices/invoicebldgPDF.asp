<!-- <html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000"> -->
<%
leaseid = trim(Request("l"))
ypid = trim(request("y"))
building = trim(request("building"))
pid = trim(request("pid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))

dim cnn1, rst1, rst2, bldgrs
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")
%>

<%
dim templid, tempypid, totalput
totalout = ""
if pid<>"" and byear<>"" and bperiod<>"" then
	set bldgrs = Server.CreateObject("ADODB.Recordset")
	bldgrs.open "select ypid, leaseutilityid from buildings b inner join tblbillbyperiod bbp on bbp.bldgnum=b.bldgnum where b.portfolioid='"&pid&"' and bbp.billyear="&byear&" and bbp.billperiod="&bperiod&" and bbp.posted=1 order by ypid, leaseutilityid", cnn1
	do until bldgrs.eof
		templid = trim(bldgrs("leaseutilityid"))
		tempypid = trim(bldgrs("ypid"))
		totalout = totalout& showtenantbill(templid, tempypid)
		bldgrs.movenext
	loop
elseif building<>"" then
	set bldgrs = Server.CreateObject("ADODB.Recordset")
	if ypid<>"" then
		bldgrs.open "select leaseutilityid from tblbillbyperiod where bldgnum='"&building&"' and ypid="&ypid&" and posted=1", cnn1
		do until bldgrs.eof
			templid = trim(bldgrs("leaseutilityid"))
			totalout = totalout&showtenantbill(templid, ypid)
			bldgrs.movenext
		loop
	elseif byear<>"" and bperiod<>"" then
		bldgrs.open "select leaseutilityid, ypid from tblbillbyperiod where bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod&" and posted=1", cnn1
		do until bldgrs.eof
			templid = trim(bldgrs("leaseutilityid"))
			tempypid = trim(bldgrs("ypid"))
			totalout = totalout & showtenantbill(templid, tempypid)
			bldgrs.movenext
		loop
	end if
elseif leaseid<>"" and ypid<>"" then
	showtenantbill leaseid, ypid
end if
set cnn1 = nothing


' Create the object
Set APW = Server.CreateObject("APWebGrabber.Object")
APW.URL = "http://www.genergy.com"
APW.IEHeader "<table width=""100%"" border=""0""><tr><td height=""68""><img src=""invoice%20logo.jpg"" width=""202"" height=""143""></td></tr></table>"
APW.CreateFromHTMLText = totalout
' Tell WebGrabber to GO
R= APW.DoPrint("127.0.0.1",64320)
' Now wait for a result
Result = APW.Wait("127.0.0.1",64320,60,"")
' To get the name of the PDF, you have to use the activePDF Server object
If Result = "019" Then ' This was a good request
Set PDF = Server.CreateObject("APServer.Object")
' Get the
PDF.FromString APW.Prt2DiskSettings
PDFFileName = PDF.OutputDirectory & "\" & PDF.NewUniqueID & ".pdf"
Else
Response.Write("Error! " & Result)
End If
call APW.Cleanup("127.0.0.1",64320)

%>










<%
'### begin of showtenantbill, is rest of file ###

function showtenantbill(leaseid, ypid)
Set rst2 = Server.CreateObject("ADODB.recordset")

rst2.open "select * from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid & " and posted=1", cnn1
if not rst2.eof then
%>
<!-- <div style="page-break-before:always"> -->
<!-- header -->
<!-- <table width="100%" border="0"><tr><td height="68"><img src="invoice%20logo.jpg" width="202" height="143"></td></tr></table> -->


<!-- header -->
<%dim outputstring
outputstring = ""
outputstring = outputstring &"<table width=""100%"" border=""0"" height=""100%"">"
outputstring = outputstring &"<tr><td height=""485"" valign=""top"">"
outputstring = outputstring &"<table width=""80%"" border=""0"" align=""center"" bordercolor=""#FFFFFF"" style=""font-family: Arial, Helvetica, sans-serif; font-size:10px"">"
outputstring = outputstring &"<tr><td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td width=""30%"" bgcolor=""#CCCCCC"" bordercolor=""#FFFFFF"" align=""center"">Invoice Number</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr><td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"	<td width=""30%"" bgcolor=""#CCCCCC"" bordercolor=""#FFFFFF"" align=""center"">EL." & rst2("billperiod") & Right(rst2("billyear"),2)&  "." & rst2("tenantnum") &"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"</table>"
outputstring = outputstring &"<table width=""80%"" border=""1"" align=""center"" bordercolor=""#000000"" cellspacing=""0"" style=""font-family: Arial, Helvetica, sans-serif; font-size:10px"">"
outputstring = outputstring &"<tr bordercolor=""#FFFFFF"" bgcolor=""#CCCCCC"">"
outputstring = outputstring &"	<td width=""13%"" align=""center"">Period</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">From</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">To</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center""></td>"
outputstring = outputstring &"	<td width=""15%"" align=""center""></td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">CONSUMPTION</td>"
outputstring = outputstring &"	<td width=""12%"" align=""center""></td>"
outputstring = outputstring &"	<td width=""30%"" align=""center"">DEMAND</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr bordercolor=""#FFFFFF"" bgcolor=""#CCCCCC""> "
outputstring = outputstring &"	<td width=""13%"" align=""center"">"&rst2("billyear")&"/"&rst2("billperiod")&"</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">"&rst2("datestart")-1&"</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">"&rst2("dateend")&"</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">METER</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">On Peak</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">Off Peak</td>"
outputstring = outputstring &"	<td width=""12%"" align=""center"">KWHR</td>"
outputstring = outputstring &"	<td width=""30%"" align=""center"">KW</td>"
outputstring = outputstring &"</tr>"

Set rst1 = Server.CreateObject("ADODB.recordset")
rst1.open "select * from tblmetersbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid, cnn1

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
	outputstring = outputstring &"<tr bordercolor=""#FFFFFF""> "
	outputstring = outputstring &"	<td width=""13%""></td>"
	outputstring = outputstring &"	<td width=""15%"">&nbsp;</td>"
	outputstring = outputstring &"	<td width=""15%"">&nbsp;</td>"
	outputstring = outputstring &"	<td width=""15%"" bordercolor=""#FFFFFF"">"&rst1("Meternum")&"</td>"
	outputstring = outputstring &"	<td width=""15%"" align=""center"" bordercolor=""#FFFFFF"">"&Formatnumber(rst1("onpeak"),0)&"</td>"
	outputstring = outputstring &"	<td width=""15%"" align=""center"" bordercolor=""#FFFFFF"">"&Formatnumber(rst1("offpeak"),0)&"</td>"
	outputstring = outputstring &"	<td width=""12%"" align=""center"" bordercolor=""#FFFFFF"">"&Formatnumber(rst1("kwhused"),0)&"</td>"
	outputstring = outputstring &"	<td width=""30%"" align=""center"" bordercolor=""#FFFFFF"">"&Formatnumber(rst1("demand_P"),0)&"</td>"
	outputstring = outputstring &"</tr>"

	tot_onpeak = tot_onpeak + rst1("onpeak")
	tot_offpeak= tot_offpeak+ rst1("offpeak")
	tot_kwhused= tot_kwhused + rst1("kwhused")
	tot_demand_p= tot_demand_p + rst1("demand_P")
	
	rst1.movenext
wend

else
end if
outputstring = outputstring &"<tr bordercolor=""#FFFFFF""> "
outputstring = outputstring &"	<td width=""13%""></td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">&nbsp;</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"">&nbsp;</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"" bgcolor=""#CCCCCC"" bordercolor=""#FFFFFF"" align=""center"">Totals</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"" bordercolor=""#FFFFFF"" bgcolor=""#CCCCCC"" align=""right"">"&Formatnumber(tot_onpeak,0)&"</td>"
outputstring = outputstring &"	<td width=""15%"" align=""center"" bordercolor=""#FFFFFF"" bgcolor=""#CCCCCC"">"&Formatnumber(tot_offpeak,0)&"</td>"
outputstring = outputstring &"	<td width=""12%"" align=""center"" bordercolor=""#FFFFFF"" bgcolor=""#CCCCCC"" align=""right"">"&Formatnumber(tot_kwhused,0)&"</td>"
outputstring = outputstring &"	<td width=""30%"" align=""center"" bordercolor=""#FFFFFF"" bgcolor=""#CCCCCC"" align=""right"">"&FormatNumber(tot_demand_P,0)&"</td>"
outputstring = outputstring &"	</tr>"

outputstring = outputstring &"</table>"
outputstring = outputstring &"<table width=""80%"" border=""0"" align=""center"" bordercolor=""#FFFFFF"" cellspacing=""0"" style=""font-family: Arial, Helvetica, sans-serif; font-size:10px"">"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td width=""1%"">&nbsp;</td>"
outputstring = outputstring &"	<td width=""1%"">&nbsp;</td>"
outputstring = outputstring &"	<td width=""1%"">&nbsp;</td>"
outputstring = outputstring &"	<td width=""70%"">&nbsp;</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td rowspan = ""6"" colspan=""4"" valign=""bottom"">&nbsp;</td>"
outputstring = outputstring &"	<td width=""10%"" align=""right""><b>Admin Fee</b></td>"
outputstring = outputstring &"	<td width=""7%"" align=""right"">"&FormatCurrency((clng(rst2("Adminfee"))*(clng(rst2("energy"))+clng(rst2("demand"))+ clng(rst2("Addonfee"))+((clng(rst2("energy"))+clng(rst2("demand"))+clng(rst2("Addonfee")))* clng(rst2("Adminfee"))))),2)&"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td width=""7%"" align=""right""><font face=""Arial, Helvetica, sans-serif"" size=""1""><b>Service Fee</b></td>"
outputstring = outputstring &"	<td width=""10%"" align=""right"">"&FormatCurrency((rst2("Addonfee")*metercount),2)&"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td width=""7%"" bgcolor=""#CCCCCC"" align=""right""><b>Sub Total</b></td>"
outputstring = outputstring &"	<td width=""10%"" bgcolor=""#CCCCCC"" align=""right"">"&FormatCurrency((clng(rst2("energy"))+clng(rst2("demand"))+clng(rst2("Addonfee"))+((clng(rst2("energy"))+clng(rst2("demand"))+clng(rst2("Addonfee")))* clng(rst2("Adminfee")))),2)&"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td width=""7%"" bgcolor=""#CCCCCC"" align=""right""><b>Sales Tax</b></td>"
outputstring = outputstring &"	<td width=""10%"" bgcolor=""#CCCCCC"" align=""right"">"&FormatCurrency(rst2("tax"),2)&"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td width=""7%"" bgcolor=""#CCCCCC"" align=""right""><b>Total Amt</b></td>"
outputstring = outputstring &"	<td width=""10%"" bgcolor=""#CCCCCC"" align=""right"">"&FormatCurrency(rst2("totalamt"),2)&"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"</table>"
outputstring = outputstring &"<table width=""80%"" border=""0"" align=""center"" bordercolor=""#FFFFFF"" cellspacing=""0"">"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td width=""10%""><img src=""MakeChartyrly.asp?lid=" & leaseid & "&by=" & rst2("billyear") & "&bp="&rst2("billperiod")&""" width=""600"" height=""175""></td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr> "
outputstring = outputstring &"	<td></td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"</table>"


outputstring = outputstring &"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"<tr><td valign=""top""><hr width=""80%"" align=""center"">"
outputstring = outputstring &"	<table width=""80%"" border=""0"" align=""center"" style=""font-family: Arial, Helvetica, sans-serif; size:10px"">"
outputstring = outputstring &"	<tr style=""font-size:12px"">"
outputstring = outputstring &"		<td>Tenant Name and Address:</td>"
outputstring = outputstring &"		<td>Make Check Payable To:</td>"
outputstring = outputstring &"	</tr>"
outputstring = outputstring &"	<tr style=""font-size:10px""><td><b>"&rst2("tenantname")&" ("&rst2("tenantnum")&")</b></td>"
outputstring = outputstring &"		<td><b>"&rst2("btbldgname")&"</b></td>"
outputstring = outputstring &"	</tr>"
outputstring = outputstring &"	<tr style=""font-size:10px""><td><b>"&rst2("tstrt")&"</b></td>"
outputstring = outputstring &"		<td><b>"&rst2("btstrt")&"</b></td>"
outputstring = outputstring &"	</tr>"
outputstring = outputstring &"	<tr style=""font-size:10px""><td><b>"&rst2("tcity")&", "&rst2("tstate")&" "&rst2("tzip")&"</b></td>"
outputstring = outputstring &"		<td><b>"&rst2("btcity")&", "&rst2("btstate")&" "&rst2("btzip")&"</b></td>"
outputstring = outputstring &"	</tr>"
outputstring = outputstring &"	</table>"
outputstring = outputstring &"<p><font size=""2""></font></p>"

outputstring = outputstring &"</td>"
outputstring = outputstring &"</tr>"
outputstring = outputstring &"</table>"
'response.write outputstring
%>

<!-- </div> -->
<%

rst2.close
set rst2 = nothing
showtenantbill = outputstring
end function
%>