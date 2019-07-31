<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
 Server.ScriptTimeout = 300
sub closewindow()
	%>
	<script>
		window.close();
	</script>
	<%
	response.end
end sub

if 	not( _
	checkgroup("Genergy Users")<>0 _
	or checkgroup("clientOperations")<>0 _
	) then%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim pid, bldg, customsrc, action, tid, lid, id, byear, bperiod
pid = request("pid")
bperiod = request("bperiod")
byear = request("byear")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
customsrc = request("customsrc")

dim cnn1, rst1, sql, cmd
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
cnn1.open getLocalConnect(bldg)

dim bldgname, address, billingname, basecapchg, baseconchg, basecondt, basecpi, baseelectric, currentcpi, currentelectric, mintons, custchrg, credit, deltat, flatchg, flatcap, exp_date, script, cpiadj, epiadj, cpiunder, epiunder, rateid, dtadj, smonth, emonth, sday, eday, mincapchg, cpyepicpi
basecapchg = request("basecapchg")
baseconchg = request("baseconchg")
basecondt = request("basecondt")
basecpi = request("basecpi")
baseelectric = request("baseelectric")
currentcpi = request("currentcpi")
currentelectric = request("currentelectric")
mintons = request("mintons")
custchrg = request("custchrg")
credit = request("credit")
if trim(credit)="" then credit = 0
deltat = request("deltat")
id = request("id")
flatchg = request("flatchg")
if trim(flatchg) = "" then flatchg="0"
flatcap = request("flatcap")
if trim(flatcap) = "" then flatcap="0"
exp_date = request("exp_date")
cpiadj = request("cpiadj")
epiadj = request("epiadj")
epiunder = request("epiunder")
cpiunder = request("cpiunder")
rateid = request("rateid")
if trim(request("dtadj"))="" then dtadj = 0 else dtadj = request("dtadj")
if trim(request("smonth"))="" then smonth = 0 else smonth = request("smonth")
if trim(request("emonth"))="" then emonth = 0 else emonth = request("emonth")
if trim(request("sday"))="" then sday = 0 else sday = request("sday")
if trim(request("eday"))="" then eday = 0 else eday = request("eday")
mincapchg = request("mincapchg")
action = request("action")
cpyepicpi = request("cpyepicpi")
'errors
dim printerr
if dtadj>0 then
  if not(isdate(smonth&"/"&sday&"/"&year(date()))) or not(isdate(emonth&"/"&eday&"/"&year(date()))) then printerr = printerr & "Delta T adjustment needs proper start and end dates.<br>"
end if
dim insertSET
if trim(printerr)="" then
	
	insertSET = customsrc&" SET basecapchg='"&basecapchg&"', baseconchg='"&baseconchg&"', basecondt='"&basecondt&"', basecpi='"&basecpi&"', baseelectric='"&baseelectric&"', currentcpi='"&currentcpi&"', currentelectric='"&currentelectric&"', mintons='"&mintons&"', custchrg='"&custchrg&"', credit='"&credit&"', deltat='"&deltat&"', flatchg='"&flatchg&"', flatcap='"&flatcap&"', exp_date='"&exp_date&"', cpiadj='"&cpiadj&"', epiadj='"&epiadj&"', cpiunder='"&cpiunder&"', epiunder='"&epiunder&"', rateid='"&rateid&"', dtadj='"&dtadj&"', smonth='"&smonth&"', emonth='"&emonth&"', sday='"&sday&"', eday='"&eday&"', mincapchg='"&mincapchg&"'"
	if trim(action)="Save" then
  		sql = "INSERT INTO "&customsrc&" (leaseutilityid, billyear, billperiod, basecapchg, baseconchg, basecondt, basecpi, baseelectric, currentcpi, currentelectric, mintons, custchrg, credit, deltat, flatchg, flatcap, exp_date, cpiadj, epiadj, cpiunder, epiunder, rateid, dtadj, smonth, emonth, sday, eday, mincapchg) VALUES ('"&lid&"', '"&byear&"', '"&bperiod&"', '"&basecapchg&"', '"&baseconchg&"', '"&basecondt&"', '"&basecpi&"', '"&baseelectric&"', '"&currentcpi&"', '"&currentelectric&"', '"&mintons&"', '"&custchrg&"', '"&credit&"', '"&deltat&"', '"&flatchg&"', '"&flatcap&"', '"&exp_date&"', '"&cpiadj&"', '"&epiadj&"', '"&cpiunder&"', '"&epiunder&"', '"&rateid&"', '"&dtadj&"', '"&smonth&"', '"&emonth&"', '"&sday&"', '"&eday&"', '"&mincapchg&"')"
  		'closewindow()
	elseif trim(action)="Update" then
		sql = "UPDATE "&insertSET&" WHERE leaseutilityid="&lid&" and billyear="&byear&" and billperiod="&bperiod
		' closewindow()
	end if
	
	
	'response.write sql
	if sql<>"" then
		cnn1.execute sql
		%><script>alert("Data Saved.")</script><%
		if cpyEPICPI = "allport" or cpyEPICPI = "allbldg" then
			cmd.CommandTimeout = 300
			cmd.CommandType = adCmdStoredProc
			cmd.CommandText = "sp_Copy_Custom_OUC"
			cmd.Parameters.Append cmd.CreateParameter("pid", adInteger, adParamInput)
			cmd.Parameters.Append cmd.CreateParameter("bldg", adInteger, adParamInput)
			cmd.Parameters.Append cmd.CreateParameter("bldgnum", adVarchar, adParamInput, 20)
			cmd.Parameters.Append cmd.CreateParameter("lid", adInteger, adParamInput)
			cmd.Parameters.Append cmd.CreateParameter("by", adInteger, adParamInput)
			cmd.Parameters.Append cmd.CreateParameter("bp", adInteger, adParamInput)
			cmd.Parameters("pid") = pid
			cmd.Parameters("bldgnum") = bldg
			cmd.Parameters("lid") = lid
			cmd.Parameters("by") = byear
			cmd.Parameters("bp") = bperiod
			if cpyEPICPI = "allport" then
				cmd.Parameters("bldg") = 0'[0=portfolio copy|1=building copy]
				cmd.ActiveConnection = getConnect(pid,bldg,"billing")
			else' cpyEPICPI = "allbldg" then
				cmd.Parameters("bldg") = 1'
				cmd.ActiveConnection = cnn1
		end if
			'response.write "sp_Copy_Custom_OUC "&cmd.Parameters("pid")&", "&cmd.Parameters("bldg")&", '"&cmd.Parameters("bldgnum")&"', "&cmd.Parameters("lid")&", "&cmd.Parameters("by")&", "&cmd.Parameters("bp")
			'response.end
			cmd.execute
		end if
	closewindow
	end if
	
	if trim(bldg)<>"" then
		dim temp
		temp = "SELECT b.bldgname, b.strt, b.customsrc, l.billingname FROM buildings b INNER JOIN tblLeases l ON l.bldgnum=b.bldgnum WHERE b.bldgnum='"&bldg&"'"
		rst1.Open temp , cnn1
		if not rst1.EOF then
			bldgname = rst1("bldgname")
			address = rst1("strt")
			customsrc = "custom_ouc1"
			billingname = rst1("billingname")
		end if
		rst1.close		

		if trim(tid)<>"" and trim(customsrc)<>"" and trim(byear)<>"" and trim(bperiod)<>"" then
			dim tempvar
			tempvar = "SELECT * FROM ["&customsrc&"] WHERE leaseutilityid='"&lid&"' and billyear="&byear&" and billperiod="&bperiod
			'response.write(tempvar)
			rst1.open tempvar, cnn1
			if not rst1.eof then 
				basecapchg = rst1("basecapchg")
				baseconchg = rst1("baseconchg")
				basecondt = rst1("basecondt")
				basecpi = rst1("basecpi")
				baseelectric = rst1("baseelectric")
				currentcpi = rst1("currentcpi")
				currentelectric = rst1("currentelectric")
				mintons = rst1("mintons")
				custchrg = rst1("custchrg")
				credit = rst1("credit")
				deltat = rst1("deltat")
				id = rst1("id")
				flatchg = rst1("flatchg")
				flatcap = rst1("flatcap")
				exp_date = rst1("exp_date")
				cpiadj = rst1("cpiadj")
				epiadj = rst1("epiadj")
				cpiunder = rst1("cpiunder")
				epiunder = rst1("epiunder")
				rateid = rst1("rateid")
				dtadj = rst1("dtadj")
				smonth = rst1("smonth")
				emonth = rst1("emonth")
				sday = rst1("sday")
				eday = rst1("eday")
				mincapchg = rst1("mincapchg")
			end if
			rst1.close
			if trim(lid)<>"" and trim(id)="" then
				rst1.open "SELECT top 1 * FROM [custom_ouc1] WHERE leaseutilityid='"&lid&"' ORDER BY billyear desc, billperiod desc, id desc", cnn1
				if rst1.eof then
					rst1.close
					rst1.open "SELECT top 1 * FROM [custom_ouc1] ORDER BY billyear desc, billperiod desc, id desc", cnn1
				end if
				if not rst1.eof then
					script = "frm.basecapchg.value = '"&rst1("basecapchg")&"';"&_
								"frm.baseconchg.value = '"&rst1("baseconchg")&"';"&_
								"frm.basecondt.value = '"&rst1("basecondt")&"';"&_
								"frm.basecpi.value = '"&rst1("basecpi")&"';"&_
								"frm.baseelectric.value = '"&rst1("baseelectric")&"';"&_
								"frm.mintons.value = '"&rst1("mintons")&"';"&_
								"frm.custchrg.value = '"&rst1("custchrg")&"';"&_
								"frm.credit.value = '"&rst1("credit")&"';"&_
								"frm.deltat.value = '"&rst1("deltat")&"';"&_
								"if('"&rst1("flatchg")&"'=='True') frm.flatchg.checked = 'true';"&_
								"frm.flatcap.value = '"&rst1("flatcap")&"';"&_
								"frm.exp_date.value = '"&rst1("exp_date")&"';"&_
								"frm.cpiadj.value = '"&rst1("cpiadj")&"';"&_
								"frm.epiadj.value = '"&rst1("epiadj")&"';"&_
								"frm.dtadj.value = '"&rst1("dtadj")&"';"&_
								"frm.smonth.value = '"&rst1("smonth")&"';"&_
								"frm.emonth.value = '"&rst1("emonth")&"';"&_
								"frm.sday.value = '"&rst1("sday")&"';"&_
								"frm.eday.value = '"&rst1("eday")&"';"&_
								"frm.mincapchg.value = '"&rst1("mincapchg")&"';"&_
								"if('"&rst1("cpiunder")&"'=='True') frm.cpiunder.checked = 'true';"&_
								"if('"&rst1("epiunder")&"'=='True') frm.epiunder.checked = 'true';"
				end if
				rst1.close
			end if
		end if
	end if
	if trim(flatcap) = "0" then flatcap=""
	if trim(exp_date) = "1/1/1900" then exp_date=""
end if

%>
<html>
<head>
<title>Building View</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
<script>
function fillblanks()
{	frm = document.form2
	<%=script%>
}

function calCPIadj()
{	frm = document.form2
	frm.cpiadj.value = roundNumber(frm.currentcpi.value/frm.basecpi.value,4)
	if(frm.cpiunder.checked&&(frm.cpiadj.value<1)) frm.cpiadj.value=1;
}

function calEPIadj()
{	frm = document.form2;
	frm.epiadj.value = roundNumber(frm.currentelectric.value/frm.baseelectric.value,6);
	if(frm.epiunder.checked&&(frm.epiadj.value<1)) frm.epiadj.value=1;
}

function roundNumber(num,precision)
{	var i,sZeros = '';
	for(i = 0;i < precision;i++)
		sZeros += '0';
	i = Number(1 + sZeros);
	return Math.round(num * i) / i;
}

function checkadjmorethanone()
{	var frm = document.form2;
	
}

function checkescoBox(option)
{	if(option=='0')
	{	document.forms['form2'].credit.disabled=true;
		document.forms['form2'].credit.value='0';
	}else
	{	document.forms['form2'].credit.disabled=false;
	}
}

function confirmBeforeSubmit(action){
	var frm = document.forms['form2']
	frm.action.value = action
	if ((frm.cpyEPICPI[0].checked) && (confirm("Are you sure you would like to copy this EPI and CPI information for all tenants in this portfolio?  Only tenants in buildings with billperiods set up for <%=byear%>/<%=bperiod%> will be copied."))){
		frm.submit();
	}
	if ((frm.cpyEPICPI[1].checked) && (confirm("Are you sure you would like to copy this EPI and CPI information for all tenants in this building?"))){
		frm.submit();
	}
	if (frm.cpyEPICPI[2].checked){
		frm.submit();
	}
}
</script>
</head>

<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 <%if trim(byear)<>"" and trim(bperiod)<>"" then%>onload="checkescoBox(document.form2.rateid.value)"<%end if%>>
<%if trim(lid)<>"" and trim(byear)<>"" and trim(bperiod)<>"" then%>
<script> window.resizeTo(550,695);</script>
<form name="form2" method="post" action="oucleasefields.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<%if 1=0 then 'if checkgroup("clientOperations")=0 then%>
<tr><td bgcolor="#000000">
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
</td></tr>
<%end if%>
<tr bgcolor="#3399cc"><td>
	<table border=0 cellpadding="0" cellspacing="0" width="100%">
	<tr><td><span class="standardheader">
			Update OUC Custom Fields | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%'=portfolioname%></a> &gt; <%=bldgname%> &gt; <%=billingname%></span>
			</span></td></tr>
	</table></td>
</tr>
<tr bgcolor="#eeeeee">
	<td style="border-bottom:1px solid #999999">
<font color="red"><%if trim(printerr)<>"" then
    response.write "Data Not saved due to error:<br>"&printerr
end if%></font>
<table border="0" cellpadding="3" cellspacing="0">
<tr><td align="right" valign="bottom"><span class="standard">Building Name</span></td>
	<td valign="bottom"><span class="standard"><%=bldgname%></span>&nbsp;&nbsp;&nbsp;<%if trim(id)="" then%><input type="button" onclick="fillblanks()" value="Load Numbers" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%></td>
</tr>
<tr bgcolor="#eeeeee" class="standard">
    <td align="right">Bill Year</td>
    <td><%=byear%></td>
</tr>
<tr bgcolor="#eeeeee" class="standard">
    <td align="right">Bill Period</td>
    <td><%=bperiod%></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Contracted Capacity</span></td>
    <td><input type="text" name="mintons" value="<%=mintons%>"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Base Capacity Charge</span></td>
    <td><input type="text" name="basecapchg" value="<%=basecapchg%>">&nbsp;Never&nbsp;Less&nbsp;Than<input type="text" name="mincapchg" value="<%=mincapchg%>"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Base Consumption</span></td>
    <td><input type="text" name="baseconchg" value="<%=baseconchg%>"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Adjustable Consumption Base Price</span></td>
    <td><input type="text" name="basecondt" value="<%=basecondt%>"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Base CPI</span></td>
    <td><input type="text" name="basecpi" value="<%=basecpi%>" onkeyup="calCPIadj()"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Current CPI</span></td>
    <td><input type="text" name="currentcpi" value="<%=currentcpi%>" onkeyup="calCPIadj()"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">CPI Adjustment</span></td>
    <td><input type="text" name="cpiadj" value="<%=cpiadj%>" READONLY>&nbsp;CPI not less than 1&nbsp;<input type="checkbox" name="cpiunder" onclick="calCPIadj()" value="1"<%if trim(cpiunder)="True" then response.write " CHECKED"%>></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Base Electric</span></td>
    <td><input type="text" name="baseelectric" value="<%=baseelectric%>" onkeyup="calEPIadj()"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Current Electric</span></td>
    <td><input type="text" name="currentelectric" value="<%=currentelectric%>" onkeyup="calEPIadj()"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">EPI Adjustment</span></td>
    <td><input type="text" name="epiadj" value="<%=epiadj%>" READONLY>&nbsp;EPI not less than 1&nbsp;<input type="checkbox" name="epiunder" onclick="calEPIadj()" value="1"<%if trim(epiunder)="True" then response.write " CHECKED"%>></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Additional Meter Charge</span></td>
    <td><input type="text" name="custchrg" value="<%=custchrg%>"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Credit</span></td>
    <td><input type="text" name="credit" value="<%=credit%>">&nbsp;
        <select name="rateid" onchange="checkescoBox(this.value)">
          <option value="0">No Rate Code</option>
          <%
          rst1.open "SELECT * FROM ratecodes ORDER BY ratecode", getConnect(pid,bldg,"billing")
          do until rst1.eof
            %><option value="<%=trim(rst1("id"))%>"<%if trim(rateid)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("ratecode")%></option><%
            rst1.movenext
          loop
          rst1.close
          %>
        </select>
    </td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Base Delta T</span></td>
    <td><input type="text" name="deltat" value="<%=deltat%>"></td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right" valign="top"><span class="standard">Delta T Adjustment</span></td>
    <td>
    <table border=0 cellpadding="0" cellspacing="0">
    <tr>
      <td colspan="3"><input type="text" name="dtadj" value="<%=dtadj%>" not></td>
    </tr>
    <tr>
      <td class="standard">From:&nbsp;</td>
      <td width="20">&nbsp;</td>
      <td class="standard">To:&nbsp;</td>
    </tr>
    <tr>
      <td><input type="text" name="smonth" value="<%=smonth%>" size="2" maxlength="2">/<input type="text" name="sday" value="<%=sday%>" size="2" maxlength="2"></td>
      <td width="20">&nbsp;</td>
      <td><input type="text" name="emonth" value="<%=emonth%>" size="2" maxlength="2">/<input type="text" name="eday" value="<%=eday%>" size="2" maxlength="2"></td>
    </tr>
    </table>
    </td>
</tr>
<tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Flat Capacity Charge</span></td>
    <td><input type="checkbox" name="flatchg" value="1"<%if trim(flatchg)="True" then response.write " CHECKED"%>> <input type="text" name="flatcap" value="<%=flatcap%>"> Expiration Date <input type="text" name="exp_date" size="10" value="<%=exp_date%>"></td>
</tr>
<tr>
	<td></td>
	<td colspan="7">
		<input type="radio" name= "cpyEPICPI" value="allport">Copy EPI and CPI info to all tenants in this portfolio
	</td>
</tr>
<tr>
	<td></td>
	<td colspan="7">
		<input type="radio" name= "cpyEPICPI" value="allbldg">Copy EPI and CPI info to all tenants in this building
	</td>
</tr>
<tr>
	<td></td>
	<td colspan="7">
		<input type="radio" name ="cpyEPICPI" value="no" checked>Don't copy any info
	</td>
</tr>
<tr bgcolor="#eeeeee">
    <td></td>
    <td>
      <%if trim(id)<>"" then%>
        <input type="button" name="updateButton" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;" onclick="javascript:confirmBeforeSubmit('Update')">
      <%else%>
        <input type="button" name="saveButton" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;" onclick="javascript:confirmBeforeSubmit('Save')">
      <%end if%>
	</td>
</tr>

</table>
	</td>
</tr>
</table>
<input type="hidden" name="action">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="customsrc" value="<%=customsrc%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="byear" value="<%=byear%>">
<input type="hidden" name="bperiod" value="<%=bperiod%>">
</form>
<%else%>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#dddddd">
<td width="18%"><b>Bill Year</b></td>
<td width="23%"><b>Bill Period</b></td>
<td width="17%"><b>Start Date</b></td>
<td width="28%"><b>End Date</b></td>
</tr>
<%
rst1.open "SELECT * FROM billyrperiod WHERE utility=(SELECT utility FROM tblleasesutilityprices WHERE leaseutilityid="&lid&") and bldgnum='"&bldg&"' ORDER BY billyear desc, billperiod desc", cnn1
if rst1.eof then response.write "<tr valign=""top"" colspan=""4""><td>There are no bill periods setup.</td></tr>"
do until rst1.eof%>
	<tr bgcolor="#ffffff" valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='<%="oucleasefields.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&byear="&rst1("billyear")&"&bperiod="&rst1("billperiod")%>'">
	<td><%=rst1("billyear")%></td>
	<td><%=rst1("billperiod")%></td>
	<td><%=rst1("datestart")%></td>
	<td><%=rst1("dateend")%></td>
	</tr>
<%
rst1.movenext
loop
%>
</table>
<%end if%>
</body>
</html>