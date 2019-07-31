<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
'COMMENTS
'12/7/2007 N.Ambo changed Usage3MonthLowLimit=0.0 to Usage3MonthHighLimit=0.0 since  Usage3MonthLowLimit was being repeated 
'1/16/2008 N.Ambo modified to add option for user to enter a charge code associated with the meter for PA meters
'1/31/2008 N.Ambo modifed to include a new filed for recording ntoes on the functionality of the meter (functionality_desc field)

'dim xml
'Set xml = getXmlSession()
'response.Write(xml.xml)
'response.End()
if not(allowGroups("Genergy Users,clientOperations,IT Services")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, meterid
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
meterid = request("meterid")

dim cnn1, rst1, strsql
dim rst2, rst3, cnn2

set cnn1 = server.createobject("ADODB.connection")
set cnn2 = server.CreateObject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
set rst3 = server.CreateObject("ADODB.recordset")
cnn1.open getConnect(pid,bldg,"billing")
cnn2.open getConnect(pid,bldg,"dbCore")

dim Meternum, StartDate, DateOff, DateRead, TimeRead, Factor, Usagex, Capacityx, Location, Riser, online, note, Cumulative, variance, lmp, floor, datasource, customsrc, nobill, refmeterid, lastPCperiod, lastPCyear, manualentry, readorder, lmchannel, lmnum, units, Meter_CT_Ratio, Meter_Model, Meter_Voltage, metertype, cavee_monitor, data_frequency, deduct, extusg, meteraddonID,opentickets,totaltickets, ticketcount, masterticketid, share,Meter_JobID, Functionality
dim Sprinkler, chargecode, minValue, maxValue, estimateOn, gatewayDevice, gatewayIp, gateComm, ipidAddress, readerNotes
meteraddonID = 0

' Added by Tarun 1/28/2008
Dim MeterSerialNumber, StartUpDate, GatewayDeviceId, GatewayDeviceIdmnum, Manufacturer
Dim LocationNotes 'Tarun 2/21/2008
Dim active 'RSM
Dim metersetupcharge 'RSM 11/10/2015

if trim(meterid)<>"" then
	rst1.Open "SELECT m.*, a.*, isnull(mp.addonfee,0) as meteraddonID, (SELECT top 1 cavee_monitor FROM cavee_setup cs WHERE cs.meterid="&meterid&") as cavee_monitor FROM meters m LEFT JOIN (SELECT top 1 c.billyear, c.billperiod, c.meterid FROM consumption c, peakdemand p WHERE p.BillYear=c.BillYear and p.BillPeriod=c.BillPeriod and p.meterid=c.meterid and c.meterid="&meterid&" ORDER BY c.billyear desc, c.billperiod desc) a ON m.meterid=a.meterid LEFT JOIN MeterPrices mp ON mp.meterid=m.meterid WHERE m.meterid="&meterid, cnn1
	if not rst1.EOF then
		Meternum = rst1("Meternum")
		StartDate = rst1("DateStart")
		DateOff = rst1("DateOffLine")
		DateRead = rst1("DateLastRead")
		TimeRead = rst1("TimeLastRead")
		Factor = rst1("multiplier")
		Usagex = rst1("manualmultiplier")
		Capacityx = rst1("Demandmultiplier")
		Location = rst1("Location")
		Riser = rst1("Riser")
		online = rst1("online")
		note = rst1("metercomments")
		Cumulative = rst1("Cumulative")
		variance = rst1("variance")
		lmp = rst1("lmp")
		floor = rst1("floor")
		datasource = rst1("datasource")
		customsrc = rst1("customsrc")
		nobill = rst1("nobill")
		refmeterid = rst1("refmeterid")
		lastPCperiod = rst1("billperiod")
		lastPCyear = rst1("billyear")
		manualentry = rst1("manualentry")
		readorder	= rst1("readorder")
		lmchannel = rst1("lmchannel")
		lmnum = rst1("lmnum")
		units = rst1("calculate")
		Meter_CT_Ratio = rst1("ct_ratio")
		Meter_Model = rst1("model")
		Meter_Voltage = rst1("voltage")
		metertype = trim(rst1("category"))
		cavee_monitor = rst1("cavee_monitor")
		data_frequency = rst1("data_frequency")
		deduct = rst1("deduct")
		extusg = rst1("extusg")
		meteraddonID = rst1("meteraddonID")
		share = rst1("shared")
		Meter_JobID = trim(rst1("job_id"))
		Functionality = rst1("functionality_desc") '1/31/2008 N.Ambo added
		Manufacturer = rst1("Manufacturer")
		LocationNotes = rst1("LocationNotes")
		gatewayDevice = rst1("gatewayDevice")
		gatewayIp = rst1("gatewayIp")
		gateComm = rst1("gateCommunication")
		ipidAddress = rst1("modBusIdentifier")
		readerNotes = rst1("reader_notes")
        active = rst1("active")  'RSM
        metersetupcharge = rst1("metersetupcharge")  'RSM 11/10/2015
	end if
	rst1.close
	
	rst1.open "Select * from meterthreshold where meterid = " + meterid, cnn1
	
	if not rst1.EOF then
	    minValue = rst1("minUsagethresh")
	    maxValue = rst1("maxUsagethresh")
	    estimateOn = 1
    else
        minValue = 0
        maxValue = 999999
        estimateOn = 0
    end if
	rst1.close
	
	if cint(pid) = 108 then
		rst2.Open "SELECT MeterId, Sprinkler FROM tblPASprinklerMeters WHERE MeterId = " & meterid, cnn1
		if not rst2.EOF then
			Sprinkler = rst2("Sprinkler")
		else
			Sprinkler = "False"
		end if
		 'rst1.Open "SELECT meterchargecodeid,chargecode FROM tblPAmeterchargecodes WHERE meternum = '" & meternum & "'", cnn1
		 rst1.Open "SELECT meterchargecodeid,chargecode FROM tblPAmeterchargecodes WHERE meternum = '" & meternum & "' and bldgnum='" & bldg & "'" , cnn1
		if not rst1.EOF then
			chargecode = rst1("chargecode")
		end if
		rst1.Close
	end if
	
	rst1.Open "SELECT MeterId, SerialNumber, StartUpDate, GatewayDeviceId FROM tblMeterextDetails WHERE MeterId = " & meterid , cnn1
	if not rst1.EOF then
		MeterSerialNumber = rst1("SerialNumber")
		StartUpDate = rst1("StartUpDate")
		GatewayDeviceId= rst1("GatewayDeviceId")
	end if
	rst1.Close 
end if

dim bldgname, portfolioname
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
	end if
	rst1.close
end if

dim billingname
if trim(bldg)<>"" then
  rst1.open "SELECT billingname FROM tblleases WHERE billingid='"&tid&"'", cnn1
	if not rst1.EOF then
		billingname = rst1("billingname")
	end if
	rst1.close
end if

dim utilitydisplay, utilityid
if trim(lid)<>"" then
	dim sqlStatemunt
	sqlStatemunt = "select * from tblutility tu join tblleasesutilityprices tlup on tu.utilityid=tlup.utility where leaseutilityid="&lid
	rst1.open sqlStatemunt, cnn1
	utilitydisplay = rst1("utilitydisplay")
	utilityid = rst1("utilityid")
	rst1.close
else
	rst1.open "SELECT * FROM tblutility tu JOIN tblleasesutilityprices tlup ON tu.utilityid=tlup.utility JOIN meters m ON tlup.leaseutilityid=m.leaseutilityid WHERE m.meterid="&meterid, cnn1
	utilitydisplay = rst1("utilitydisplay")
	utilityid = rst1("utilityid")
	lid = rst1("leaseutilityid")
	rst1.close
end if

dim hasDatasource, othermeters
rst1.open "SELECT count(*) FROM sysobjects WHERE name='"&datasource&"'", getLocalConnectCom(bldg)
if not rst1.eof then
	if cint(rst1(0))>0 then hasDatasource = true else hasDatasource = false
end if
rst1.close

othermeters = ""
rst1.open "SELECT meternum FROM meters WHERE online=1 and meterid<>'"&meterid&"' and bldgnum='"&bldg&"'", getConnect(0,bldg,"billing")
do until rst1.eof
	if len(othermeters)>0 then othermeters = othermeters & ","
	othermeters = othermeters & """" & rst1("meternum") & """"
	rst1.movenext
loop
rst1.close

if trim(meterid) <> "" and bldg <> "" then 
	dim ticket
	set ticket = New tickets
	ticket.Label="Meter"
	ticket.Note="Master Ticket for Meter ID " &split(getBuildingIP(bldg),"\")(1)&"-"& meterid
	ticket.ccuid  = "rbdept"
	ticket.client = 1
	if meterid<>"0" then ticket.findtickets "meterid", split(getBuildingIP(bldg),"\")(1)&"-"&meterid
end if
%>
<html>
<head>
<title>Building View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}

function meterEdit(meterid)
{	document.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
}

function reloadFilter(bldg,transfermeter)
{	document.location = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&bldgfilter='+bldg+'&transfermeter='+transfermeter
}

function checkFields(frm){
	if((frm.deduct.checked)&&(frm.refmeterid.value=='0')){
		alert("Meters with deduct selected must have a meter reference selected.");
		return(false);
	}
	//if((frm.shared.checked)&&(frm.refmeterid.value=='0')){
	//	alert("Meters with shared selected must have a meter reference selected.");
	//	return(false);
	//}
	var othermeters = Array(<%=othermeters%>);
	for(i=0;i<othermeters.length;i++){
		if(othermeters[i].toUpperCase()==frm.Meternum.value.toUpperCase()){
			if(frm.refmeterid.value=='0'){
				frm.shared.checked='true';
				alert("The specified meter name matches another meter in this building making it a shared meter. Please select a meter reference for the source meter.");
				return(false);
			}
		}
	}
	if(frm.meteraddonID.value=='0'){
		return(confirm("Are you sure this meter has no addon fee?"));
	}
}

    function estimateChanged(checked)
    {
        if(!checked)
        {
            document.form2.minValue.disabled = true;
            document.form2.maxValue.disabled = true;
        }
        else
        {
            document.form2.minValue.disabled = false;
            document.form2.maxValue.disabled = false;
        }
    }
    
    function selectChanged(ip)
    {
        document.form2.gatewayIP.value = ip;
    }
    
    function changedCommType(commType)
    {
        if (commType.toString().toLowerCase().indexOf("modbus", 0) == -1)
        {
            document.form2.ip_id.disabled=true;
        }
        else
        {
            document.form2.ip_id.disabled=false;
        }
    }
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#3399cc"> 
    <td colspan="2"> <span class="standardheader"> 
      <%if trim(meterid)<>"" then%>
      Update <%=utilitydisplay%> Meter <br><span style="font-weight:normal;"> <%=billingname%> (ID# SVR<%=split(getBuildingIP(bldg),"\")(1)%>-<%=meterid%>)</span> 
      <%else%>

      Add <%=utilitydisplay%> Meter | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> 
      &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> 
      &gt; <a href="tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;"><%=billingname%></a></span> 
      <%end if%>
      </span></td>
    <td colspan="3" align="right">
	<%

	if not(isBuildingOff(bldg)) and meterid <> "" then ticket.MakeButton

	%>
	<button id="qmark2" onClick="openCustomWin('help.asp?page=<%if trim(meterid)<>"" then%>meteredit<%else%>meteradd<%end if%>','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) 
      Quick Help</button></td>
  </tr>
  <% if trim(meterid)="" then %>
  <tr>
    <td colspan="5" bgcolor="#dddddd"><b>Transfer An Existing Meter</b></td>
  </tr>
  <tr> 
    <td colspan="5" bgcolor="#eeeeee"> 
      <!-- begin meter transfer -->
      <form name="form1" method="post" action="meterTransferSave.asp">
        <table border=0 cellpadding="3" cellspacing="1">
          <tr bgcolor="#eeeeee"> 
            <td align="right"><span class="standard">Select meter to transfer:</span></td>
            <td> 
				<%
				dim transfermeter
				transfermeter = request("transfermeter")
				
				dim bldgfilter
				bldgfilter = request("bldgfilter")
				
				dim mwhere
				if trim(bldgfilter)<>"" then
					mwhere = "and b.bldgnum='"&bldgfilter&"'"
				end if
				
				rst1.open "SELECT * FROM meters m INNER JOIN buildings b on b.bldgnum=m.bldgnum WHERE leaseutilityid in (SELECT lup.leaseutilityid FROM tblleasesutilityprices lup INNER JOIN tblLeases l ON lup.billingid=l.billingid WHERE bldgnum='"&bldg&"') "&mwhere&" and online=1 order by meternum", cnn1
				dim nometers
				nometers = true
				if not rst1.eof then
					nometers = false
					%>
					<select name="transfermeter" onChange="reloadFilter('<%=bldgfilter%>',this.value)">
					<option value="">Select Meter</option>
					<%
					do until rst1.eof
						%>
						<option value="<%=rst1("meterid")%>"<%if trim(rst1("meterid"))=trim(transfermeter) then response.write " SELECTED"%>><%=rst1("meternum")%> (<%=rst1("meterid")%>)</option>
						<%
						rst1.movenext
					loop
				else
					response.write "<i>No meters available for transfer</i>"
				end if
				rst1.close
				%>
				</select> 
			</td>
            <%if trim(transfermeter)<>"" then%>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td align="right"><span class="standard">Include data back to bill 
              period:</span></td>
            <td><select name="bybp">
                <option value="0|0">Entire History</option>
                <%
        rst1.open "SELECT distinct billyear, billperiod FROM billyrperiod byp INNER JOIN meters m ON byp.bldgnum=m.bldgnum WHERE m.meterid="&transfermeter&" and byp.datestart<getdate() ORDER BY billyear desc, billperiod desc", cnn1
        do until rst1.eof
          %>
                <option value="<%=rst1("billyear")%>|<%=rst1("billperiod")%>"><%=rst1("billyear")%> 
                period <%=rst1("billperiod")%></option>
                <%
          rst1.movenext
        loop
        rst1.close
        %>
              </select> </td>
            <%end if%>
            <td> 
              <%if "l"<>"" and nometers=false then%>
              <input type="submit" name="action" value="Transfer" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"> 
              <%end if%>
            </td>
          </tr>
        </table>
        <input type="hidden" name="pid" value="<%=pid%>">
        <input type="hidden" name="bldg" value="<%=bldg%>">
        <input type="hidden" name="tid" value="<%=tid%>">
        <input type="hidden" name="lid" value="<%=lid%>">
      </form>
      <!-- end meter transfer -->
    </td>
  </tr>
  <% end if %><form name="form2" method="post" action="metersave.asp" onSubmit="return(checkFields(this))">
    <tr bgcolor="#eeeeee">
      <td valign="middle" align="left" colspan="5" style="border-bottom:1px solid #999999"> 
	  <% if meterid <> "" then ticket.Display pid,true, true, false%>
        </td>
    </tr>
  <tr bgcolor="#eeeeee"> 
    <td colspan="5" style="border-bottom:1px solid #ffffff;"> 
      <table border=0 cellpadding="3" cellspacing="1">
        <tr> 
          <td align="right"><span class="standard">Meter Name</span></td>
          <td><input name="Meternum" type="text" value="<%=left(Meternum,20)%>" maxlength="20">
          <td align="right"><span class="standard">Serial #</span></td>
          <td><input name="SerialNum" type="text" value="<%=left(MeterSerialNumber,20)%>" maxlength="20">          
            <input type="checkbox" value="1" name="online" <%if online="1" then Response.Write "CHECKED"%>>
            <span class="standard">On Line</span>&nbsp;
            <input type="checkbox" value="1" name="nobill" <%if nobill="True" then Response.Write "CHECKED"%>>
            <span class="standard">No&nbsp;Billing</span>
            <input type="checkbox" value="1" name="active" <%if active="True" then Response.Write "CHECKED"%>> 
            <span class="standard">Active</span>
            <input type="checkbox" value="1" name="metersetupcharge" <%if metersetupcharge="True" then Response.Write "CHECKED"%>> 
            <span class="standard">Meter Setup Charge</span>
          </td>
        </tr>
      </table></td>
  </tr>
  <tr valign="top"> 
    <td bgcolor="#eeeeee" width="30%" style="border-bottom:1px solid #cccccc;"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="50%">
	      <table border="0" cellpadding="3" cellspacing="1">
	        <tr bgcolor="#eeeeee"> 
	          <td align="right"><span class="standard">UM Start Date</span></td>
	          <td><input type="text" name="StartDate" value="<%=StartDate%>" size="14"></td>
	        </tr>
	        <tr bgcolor="#eeeeee"> 
	          <td align="right"><span class="standard">UM Date Off</span></td>
	          <td><input type="text" name="DateOff" value="<%=DateOff%>" size="14"></td>
	        </tr>
	        <tr bgcolor="#eeeeee"> 
	          <td align="right"><span class="standard">StartUp Date</span></td>
	          <td><input type="text" name="StartUpDate" value="<%=StartUpDate%>" size="14"></td>
	        </tr>	        
	        <tr bgcolor="#eeeeee"> 
	          <td align="right"><span class="standard">Date&nbsp;Last&nbsp;Read</span></td>
	          <td>
	            <%if trim(meterid)="" then%>
	            N/A
	            <%end if%>
	            <%=DateRead%></td>
	        </tr>
	      </table>
	  </td><td align="center" width="50%" valign="top">
	      <table border="0" cellpadding="3" cellspacing="1" align="center">
	        <tr bgcolor="#CCCCCC"><td>
				<%if not(isBuildingOff(bldg)) then%><a href="#" onClick="window.open('AddonFeeQuick.asp?bldg=<%=bldg%>','addonfee','scrollbars=no,width=280,height=30')">Global Add on Fee</a>:<br><%end if%>
				<select name="meteraddonID">
				<option value="0">No Add-on Fee</option>
				<%
					rst1.open "SELECT * FROM building_addonfee WHERE bldgnum='"&bldg&"'", cnn1
					do until rst1.eof%>
						<option value="<%=cint(rst1("id"))%>"<%if cint(meteraddonID) = cint(rst1("id")) then response.write " selected" end if%>><%=rst1("addonfee")%></option><%
						rst1.movenext
					loop
					rst1.close%>
				</select>
			</td></tr>
		  </table>
	  </td></tr></table>
	</td>
    <td width="30" rowspan="2" bgcolor="#eeeeee" style="border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;">&nbsp;</td>
    <td width="30%" rowspan="2" bgcolor="#eeeeee" style="border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
      <table border="0" cellpadding="3" cellspacing="1">
        <tr bgcolor="#eeeeee"> 
          <td align="right"><span class="standard">Factor</span></td>
          <td><input type="text" name="Factor" value="<%=Factor%>" size="8"></td>
        </tr>
        <tr bgcolor="#eeeeee"> 
          <td align="right"><span class="standard">Usage.x</span></td>
          <td><input type="text" name="Usagex" value="<%=Usagex%>" size="8"></td>
        </tr>
        <tr bgcolor="#eeeeee"> 
          <td align="right"><span class="standard">Capacity.x</span></td>
          <td><input type="text" name="Capacityx" value="<%=Capacityx%>" size="8"></td>
        </tr>
        <%if trim(meterid)="" then meterid = "0"%>
        <%rst1.open "SELECT * FROM meters WHERE online=1 and  bldgnum='"&bldg&"' and meterid<>"&meterid&" ORDER BY meternum", getConnect(0,bldg,"billing")
  if not rst1.eof then
  %>
        <tr bgcolor="#eeeeee"> 
          <td align="right"><span class="standard">Meter Reference</span></td>
          <td> <select name="refmeterid">
              <option value="0">N/A</option>
              <%
			do until rst1.eof
				%>
              <option value="<%=rst1("meterid")%>"<%if refmeterid=rst1("meterid") then response.write " SELECTED"%>><%=rst1("meternum")%></option>
              <%
				rst1.movenext
			loop
			%>
            </select> </td>
        </tr>
        <%
  end if
  rst1.close
  %>
<!--     </table> -->
		<!-- meter make info -->
<!-- 		<table> -->
		<%if allowGroups("Techsrv Admin,IT Services") then%>
		<tr>
			<td align="right">CT&nbsp;Ratio</td>
			<td><select name="Meter_CT_Ratio">
					<option value="0">Not Available</option><%
						rst1.open "SELECT * FROM Meter_CT_Ratio", getConnect(pid,bldg,"billing")
						do until rst1.eof%>
							<option value="<%=rst1("id")%>" <%if rst1("id")=Meter_CT_Ratio then response.write "SELECTED"%>><%=rst1("CT_Ratio")%></option>
						<%rst1.movenext
						loop
						rst1.close
				%></select>
			</td>
		</tr>
		<tr>
			<td align="right">Manufacturer</td>
			<td><select name="Manufacturer">
					<option value="0">Not Available</option><%
						rst1.open "SELECT * FROM Meter_Manufacturer", getConnect(pid,bldg,"billing")
						do until rst1.eof%>
							<option value="<%=rst1("id")%>" <%if rst1("id")=Manufacturer then response.write "SELECTED"%>><%=rst1("Manufacturer")%></option>
						<%rst1.movenext
						loop
						rst1.close
				%></select>
			</td>
		</tr>		
		<tr>
			<td align="right">Model</td>
			<td><select name="Meter_Model">
					<option value="0">Not Available</option><%
						rst1.open "SELECT mo.id as mid, * FROM Meter_Model mo, Meter_Manufacturer ma WHERE mo.Manufacturer_id=ma.id", getConnect(pid,bldg,"billing")
						do until rst1.eof%>
							<option value="<%=rst1("mid")%>" <%if rst1("mid")=Meter_Model then response.write "SELECTED"%>><%=rst1("Model")%> (<%=rst1("manufacturer")%>)</option>
						<%rst1.movenext
						loop
						rst1.close
					%>
					</select>
			</td>
		</tr>
		<tr>
			<td align="right">Voltage</td>
			<td><select name="Meter_Voltage">
					<option value="0">Not Available</option><%
						rst1.open "SELECT * FROM Meter_Voltage", getConnect(pid,bldg,"billing")
						do until rst1.eof%>
							<option value="<%=rst1("id")%>" <%if rst1("id")=Meter_Voltage then response.write "SELECTED"%>><%=rst1("voltage")%></option>
						<%rst1.movenext
						loop
						rst1.close
				%></select>
			</td>
		</tr>
	  <tr bgcolor="#eeeeee"> 
          <td align="right"><span class="standard">Job ID</span></td>
          <td><input type="text" name="Meter_JobID" value="<%=Meter_JobID%>" size="8" <%if not(allowGroups("gTS_Admins,IT Services")) then response.Write("disabled")%>></td>
        </tr>
		<%else%>
				<input type="hidden" name="Meter_CT_Ratio" value="<%=Meter_CT_Ratio%>">
				<input type="hidden" name="Manufacturer" value="<%=Manufacturer%>">
				<input type="hidden" name="Meter_Model" value="<%=Meter_Model%>">
				<input type="hidden" name="Meter_Voltage" value="<%=Meter_Voltage%>">
				<% '3/23/2009 N.Ambo this section to allow billers to view the job number field as requested by Davide G. and Rosa Basso 
				if allowgroups("gReadingandBilling") then %>
					<tr bgcolor="#eeeeee"> 
					<td align="right"><span class="standard">Job ID</span></td>
					<td><input type="text" name="Meter_JobID" value="<%=Meter_JobID%>" size="8" <% response.Write("disabled")%> ID="Text2"></td>
				<%end if%>
		<%end if%>
		<tr>
            <td align="right"><span class="standard">Meter Field Functionality</span></td>
           <td><TEXTAREA rows="4" cols="50" name="Functionality"><%=Functionality%></TEXTAREA></td>
        </tr>
		<tr>
            <td align="right"><span class="standard">Meter Location Notes</span></td>
           <td><TEXTAREA rows="4" cols="50" name="LocationNotes"><%=LocationNotes%></TEXTAREA></td>
        </tr>
        <tr>
            <td align="right"><span class="standard">Meter Reader Notes</span></td>
            <td><textarea rows="4" cols="50" name="readerNotes"><%=readerNotes%></textarea></td>
        </tr>
		</table>
		<!-- meter make info end -->
		</td>
    <td width="1" rowspan="2" bgcolor="#eeeeee" style="border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;">&nbsp;</td>
    <td rowspan="2" bgcolor="#eeeeee" style="border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
      <table border="0" cellpadding="3" cellspacing="1">
        <tr bgcolor="#eeeeee"> 
          <td align="right">Location</td>
          <td><input type="text" name="Location" value="<%=Location%>" size="4"></td>
        </tr>
        <tr bgcolor="#eeeeee"> 
          <td align="right">Floor</td>
          <td>
		  	<%if utilityid = 2 then %>
					<%
					dim rstFloor
					set rstFloor = server.createobject("ADODB.recordset")
					rstFloor.open "select fl_name as floor from tblfloor where bldgnum = '" & bldg & "' order by orderno, id", getConnect(pid,bldg,"engineering")
					if not rstFloor.eof then
					%>
				<select name="floor">
					<option value="N/A">N/A</option>
					<%
						while not rstFloor.eof
							%>
							<option value="<%=rstFloor("floor")%>" <%if trim(Floor) = trim(rstFloor("floor")) then response.write "selected" end if%>>
								<%=rstFloor("floor")%>
							</option>
							<%
							rstFloor.moveNext
						wend
					%>
				</select>
			<%else%>
	   			<input type="text" name="floor" value="<%=floor%>" size="4">
			<%				
				end if
			else
			%>
	   			<input type="text" name="floor" value="<%=floor%>" size="4">
	   		<%end if %>  	
		  </td>
        </tr>
		<%if utilityid = 2 then %>
		<tr bgcolor="#eeeeee"> 
			<td align="right"><span class="standard">Riser</span></td>
			<td>
				<select name="Riser">
					<option value="N/A">N/A</option>
					<%
					dim rstRiser
					set rstRiser = server.createobject("ADODB.recordset")
					rstRiser.open "select riser_name as riser from tblriser where bldgnum = '" & bldg & "'", getConnect(pid,bldg,"engineering")
					if not rstRiser.eof then
						while not rstRiser.eof
							%>
							<option value="<%=rstRiser("riser")%>" <%if riser = rstRiser("riser") then response.write "selected" end if%>><%=rstRiser("riser")%></option>
							<%
							rstRiser.moveNext
						wend
					end if
					rstRiser.close
					%>
				</select>
			</td>
		</tr>
		<%else%>
			<input type="hidden" name="Riser" value="<%'=Riser%>">
		<%end if%>
		<tr><td>Meter&nbsp;Type</td>
				<td>
				<select name="metertype">
				<option value="0">Sub Meter</option>
					<%
					rst1.open "SELECT * FROM meter_category", getConnect(pid,bldg,"billing")
					do until rst1.eof%>
						<option value="<%=rst1("id")%>"<%if metertype = trim(rst1("id")) then response.write " selected" end if%>><%=rst1("description")%></option>
						<%
						rst1.movenext
					loop
					rst1.close%>
				</select>
				</td>
		</tr>
		<%if not manualentry then%>
		<tr><td>Data&nbsp;Fequency</td>
			<td>
				<select name="data_frequency">
<!-- 				<option value="0">Sub Meter</option> -->
					<%
					rst1.open "SELECT * FROM data_frequency ORDER BY id", getConnect(pid,bldg,"billing")
					do until rst1.eof%>
						<option value="<%=rst1("id")%>"<%if trim(data_frequency) = trim(rst1("id")) then response.write " selected" end if%>><%=rst1("data_frequency")%></option>
						<%
						rst1.movenext
					loop
					rst1.close%>
				</select>
			</td>
		</tr>
		<%end if%>
        <% if trim(meterid)<>"0" then %>
<!--         <tr bgcolor="#eeeeee"> 
          <td valign="top" align="right"></td>
          <td valign="top">
          </td>
        </tr>
 -->        <tr bgcolor="#eeeeee"> 
          <td colspan="2">
<img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:openCustomWin('caveeSetup.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid=<%=meterid%>','customlink', 'width=350,height=425;')">CAVEE Settings</a><br>
            <%
if trim(customsrc)<>"" then
	response.write "*Contains Custom fields"
end if

rst1.open "SELECT * FROM custom_links WHERE code=4 and unitid='"&pid&"'", cnn1
do while not rst1.eof
	response.write "<img src=""images/aro-rt.gif"" align=""absmiddle"" hspace=""2"" border=""0""><a href=""javascript:openCustomWin('"&rst1("link")&"?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&meterid="&meterid&"','customlink', 'width="&rst1("width")&",height="&rst1("height")&"');"">"&rst1("label")&"</a><br>"
	rst1.movenext
loop
rst1.close
%>
          </td>
        </tr>
        <% end if %>
        <tr bgcolor="#eeeeee"> 
          <td colspan="2" align="center"><span class="standard">Review/Edit Variance % Limits</span></td>
		</tr>
		<%
			Dim Usage3MonthLowLimit,Usage3MonthHighLimit,Demand3MonthLowLimit,Demand3MonthHighLimit
			Dim UsageLastMonthLowLimit, UsageLastMonthHighLimit, DemandLastMonthLowLimit, DemandLastMonthHighLimit
			Dim UsageLastYrPeriodLowLimit,UsageLastYrPeriodHighLimit,DemandLastYrPeriodLowLimit,DemandLastYrPeriodHighLimit
			
				Usage3MonthLowLimit = 0.0
				Usage3MonthHighLimit = 0.0
				Demand3MonthLowLimit = 0.0  
				Demand3MonthHighLimit = 0.0
				
				UsageLastMonthLowLimit = 0.0
				UsageLastMonthHighLimit = 0.0
				DemandLastMonthLowLimit = 0.0
				DemandLastMonthHighLimit = 0.0				
				
				UsageLastYrPeriodLowLimit = 0.0
				UsageLastYrPeriodHighLimit = 0.0
				DemandLastYrPeriodLowLimit = 0.0
				DemandLastYrPeriodHighLimit = 0.0				
								
				
			
			rst1.open "SELECT * FROM tblMeterVarianceLimits WHERE MeterId=" & meterid , cnn1
			If not rst1.eof then
			
				Usage3MonthLowLimit = rst1("Usage3MonthLowLimit")
				Usage3MonthHighLimit = rst1("Usage3MonthHighLimit")'12/7/2007 N.ambo amended
				Demand3MonthLowLimit = rst1("Demand3MonthLowLimit")
				Demand3MonthHighLimit = rst1("Demand3MonthHighLimit")
				
				UsageLastMonthLowLimit = rst1("UsageLastMonthLowLimit")
				UsageLastMonthHighLimit = rst1("UsageLastMonthHighLimit")
				DemandLastMonthLowLimit = rst1("DemandLastMonthLowLimit")
				DemandLastMonthHighLimit = rst1("DemandLastMonthHighLimit")
				

				UsageLastYrPeriodLowLimit = rst1("UsageLastYrPeriodLowLimit")
				UsageLastYrPeriodHighLimit = rst1("UsageLastYrPeriodHighLimit")
				DemandLastYrPeriodLowLimit = rst1("DemandLastYrPeriodLowLimit")
				DemandLastYrPeriodHighLimit = rst1("DemandLastYrPeriodHighLimit")
								
				
			End If
			rst1.close%>		
        <tr bgcolor="#eeeeee"> 
<!--           <td align="right" valign="top"></td> -->
		  <td>Usage 3 Month Avg Variance</td>
		  <td align="left">Low <input type="text" name="Usg3MonthLowLimit" value="<%=Usage3MonthLowLimit%>" size="5"></td>
		  <td align="left">High <input type="text" name="Usg3MonthHighLimit" value="<%=Usage3MonthHighLimit%>" size="5"></td>		
<!--           <td align="center"><input type="text" name="variance" value="<%=variance%>" size="8"></td> -->
        </tr>
        <tr bgcolor="#eeeeee"> 
<!--           <td align="right" valign="top"></td> -->
		  <td>Demand 3 Month Avg Variance</td>
		  <td align="left">Low <input type="text" name="dmd3MonthLowLimit" value="<%=Demand3MonthLowLimit%>" size="5"></td>
		  <td align="left">High <input type="text" name="dmd3MonthHighLimit" value="<%=Demand3MonthHighLimit%>" size="5"></td>		
<!--           <td align="center"><input type="text" name="variance" value="<%=variance%>" size="8"></td> -->
        </tr>
        <tr bgcolor="#eeeeee"> 
<!--           <td align="right" valign="top"></td> -->
		  <td>Usage Last Month Variance</td>
		  <td align="left">Low <input type="text" name="UsgLastMonthLowLimit" value="<%=UsageLastMonthLowLimit%>" size="5"></td>
		  <td align="left">High <input type="text" name="UsgLastMonthHighLimit" value="<%=UsageLastMonthHighLimit%>" size="5"></td>		
<!--           <td align="center"><input type="text" name="variance" value="<%=variance%>" size="8"></td> -->
        </tr>   
        <tr bgcolor="#eeeeee"> 
<!--           <td align="right" valign="top"></td> -->
		  <td>Demand Last Month Variance</td>
		  <td align="left">Low <input type="text" name="dmdLastMonthLowLimit" value="<%=DemandLastMonthLowLimit%>" size="5"></td>
		  <td align="left">High <input type="text" name="dmdLastMonthHighLimit" value="<%=DemandLastMonthHighLimit%>" size="5"></td>		
<!--           <td align="center"><input type="text" name="variance" value="<%=variance%>" size="8"></td> -->
        </tr>
        <tr bgcolor="#eeeeee"> 
<!--           <td align="right" valign="top"></td> -->
		  <td>Usage Last Year Period Variance</td>
		  <td align="left">Low <input type="text" name="UsgLastYrPeriodLowLimit" value="<%=UsageLastYrPeriodLowLimit%>" size="5"></td>
		  <td align="left">High <input type="text" name="UsgLastYrPeriodHighLimit" value="<%=UsageLastYrPeriodHighLimit%>" size="5"></td>		
<!--           <td align="center"><input type="text" name="variance" value="<%=variance%>" size="8"></td> -->
        </tr>   
        <tr bgcolor="#eeeeee"> 
<!--           <td align="right" valign="top"></td> -->
		  <td>Demand Last Year Period Variance</td>
		  <td align="left">Low <input type="text" name="dmdLastYrPeriodLowLimit" value="<%=DemandLastYrPeriodLowLimit%>" size="5"></td>
		  <td align="left">High <input type="text" name="dmdLastYrPeriodHighLimit" value="<%=DemandLastYrPeriodHighLimit%>" size="5"></td>		
<!--           <td align="center"><input type="text" name="variance" value="<%=variance%>" size="8"></td> -->
        </tr>                      
<!--        <tr bgcolor="#eeeeee"> 
		  <td colspan="2"><span style="font-size:9px;">* the above variance will be adjusted by 8% when analyzing usage.</span></td>
        </tr> -->
      </table></td>
  </tr>
  <tr valign="top">
    <td valign="top" bgcolor="#eeeeee" style="border-bottom:1px solid #cccccc;">
    <table border="0" align="left" cellpadding="3" cellspacing="0">
        <tr bgcolor="#CCCCCC"> 
          <td colspan="2" style="border-right:1px solid #eeeeee;"><b>Manual Entry Setup</b></td>
          <td colspan="2"><b>Real Time Setup</b></td>
        </tr>
        <tr bgcolor="#CCCCCC"> 
          <td align="right">Manual&nbsp;Meter</td>
          <td style="border-right:1px solid #eeeeee;">
							<input type="checkbox" name="manualentry" value="1" <%if manualentry = true then %> checked <%end if%>>
						</td>
          <td align="right">Data Coll.&nbsp;Channel</td>
          <td><%if allowGroups("IT Services,gTechnicalServices") then%>
								<input type="text" name="lmchannel" value="<%=lmchannel%>" size="10">
							<%else%>
								<%=lmchannel%>
								<input type="hidden" name="lmchannel" value="<%=lmchannel%>">
							<%end if%>
					</td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td align="right">Extended&nbsp;Usage</td>
          <td style="border-right:1px solid #eeeeee;"><input type="checkbox" name="extusg" value="1" <%if extusg = true then %> checked <%end if%>></td>
          <td align="right">Data Coll.&nbsp;Number</td>
          <td><%if allowGroups("IT Services,gTechnicalServices") then%>
								<input type="text" name="lmnum" value="<%=lmnum%>" size="10">
							<%else%>
								<%=lmnum%>
								<input type="hidden" name="lmnum" value="<%=lmnum%>">
							<%end if%>
					</td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td align="right">Read&nbsp;Order</td>
          <td style="border-right:1px solid #eeeeee;"><input type="text" name="readorder" value="<%=trim(readorder)%>" size="3"></td>
          <td align="right">Gateway&nbsp;Device&nbsp;</td>
          <td style="border-right:1px solid #eeeeee;">
			<%if allowGroups("IT Services,gTechnicalServices") then%>
			    <%
                    strsql = "select ip, deviceName from RM where bldgnum='"&bldg&"'"
                    rst3.open strsql, cnn2
                    
                 if NOT rst3.EOF then
			     %>
			        <select name="gatewayDevice" onchange="selectChanged(this.value);">
			        <option value="" selected="selected">Select a gateway device</option>
			            <% do until rst3.eof %>
                    <option value="<%=rst3("ip")%>" <%if (trim(gatewayDevice)=trim(rst3("ip"))) then %> selected="SELECTED" <% end if %> > <%=rst3("deviceName")%> </option>
                    <%
                        response.Write("IP : " + rst3("ip") + "<br />")
                        response.Write("gatewayDevice : " + gatewayDevice)
                        rst3.movenext
                        loop 
                    %>
			        </select>
			        <input type="text" name="gatewayIP" value="<%=trim(gatewayDevice)%>" size="20" disabled /></td>
    		<%end if
    		
          end if%>
        </tr>
        <tr bgcolor="#CCCCCC">
            <td align="right">Type of Communication</td>
            <td style="border-right:1px solid #eeeeee;">
                <select name="gateComm" onchange="changedCommType(this.value)" >
                    <option value="" selected="selected">Select Meter Type</option>
                    <option value="Pulse" <% if trim(gateComm) = "Pulse" then %> selected <% end if %> >Pulse</option>
                    <option value="Modbus RTU" <% if trim(gateComm) = "Modbus RTU" then %> selected <% end if %> >Modbus RTU</option>
                    <option value="Modbus TCP" <% if trim(gateComm) = "Modbus TCP" then %> selected <% end if %> >Modbus TCP</option>
            </select></td>
            <td align="right">Modbus IP / ID address </td>
            <td style="border-right:1px solid #eeeeee;"><input type="text" name="ip_id" value="<%=ipidAddress%>" /></td>
        </tr>
       <%	   
			  rst1.open "select * from conversions where utilityid = " & utilityid &" or utilityid = 0", cnn1
			  if not rst1.EOF then 
				%>
					<tr bgcolor="#CCCCCC">
						<td align="right">Data Source</td>
						<td nowrap>
						  <%if allowGroups("IT Services,Techsrv Admin") then%>
						  <input type="text" name="datasource" value="<%=datasource%>" size="20">
						  <%else%>
						  <input type="hidden" name="datasource" value="<%=datasource%>">
						  <%=datasource%>
						  <%end if%>
				        <br><%if Not hasDatasource then%><font color="red">invalid&nbsp;datasource</font><%end if%>
						</td>
						<td nowrap>Data Source Units</td>
						<td>
							<select name="units">
								<option value="0" <%if trim(units) = "0" then%>Selected<%end if%>>Default for Utility</option>
								<%while not rst1.eof %> 
								<option value="<%=rst1("id")%>" <%if trim(units) = trim(rst1("id")) then%>Selected<%end if%>><%=rst1("conversion_Label")%></option>
								<%
								rst1.movenext
								wend
								%>
							</select>
						</td>
					</tr>
					<tr bgcolor="#CCCCCC">
			        <td>Estimate Interval data on import? </td>
			        <td style="border-right:1px solid #eeeeee;"><input type="checkbox" name="estimate" <%if estimateOn = 1 then%>checked="checked" <%end if%> onclick="estimateChanged(this.checked)" /></td>
			       <td>Min Usage Threshold : <input type="text" id="minValue" name="minValue" value="<%=minValue %>" <%if estimateOn = 0 then%>disabled="true" <%end if%> /></td>
			       <td>Max Usage Threshold : <input type="text" id="maxValue" name="maxValue" value="<%=maxValue %>" <%if estimateOn = 0 then%>disabled="true" <%end if%> /></td>
			</tr>
				<%
				rst1.close
			else
				rst1.close
			  	rst1.open "select * from conversions where id = " & cint(units), cnn1
			
			%>
			<tr bgcolor="#CCCCCC">
			<td></td>
			<td style="border-right:1px solid #eeeeee;"></td>
			<td colspan=2>
			<%if not rst1.EOF then %> 
				<input type="hidden" name="units" value="<%=units%>"><%=rst1("conversion_label")%>
			<%else%>
				<input type="hidden" name="units" value="0">No Conversion Selected / Required	
			<%end if %>
			</td>
			</tr>
			<%
			rst1.close
		end if
		%>
				<tr bgcolor="#CCCCCC">
				    <td></td>
				    <td style="border-right:1px solid #eeeeee;" valign="top"></td><br />
				    <td></td>
          <td valign="top">
            <%if not(isBuildingOff(bldg)) then%><img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="datafieldedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&meterid=<%=meterid%>&tid=<%=tid%>&lid=<%=lid%>">Define&nbsp;Data&nbsp;Fields</a><%end if%>
            <%
        		rst1.open "SELECT * FROM datasource WHERE datasource='"&datasource&"' and meterid='"&meterid&"' "
          
        		if not rst1.EOF then
        			if (rst1("fieldname1")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname1") & "<br>"
        			if (rst1("fieldname2")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname2") & "<br>"
        			if (rst1("fieldname3")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname3") & "<br>"
        			if (rst1("fieldname4")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname4") & "<br>"
        			if (rst1("fieldname5")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname5") & "<br>"
        			if (rst1("fieldname6")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname6") & "<br>"
        			if (rst1("fieldname7")<>"") then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rst1("fieldname7") & "<br>"
        		end if
        		rst1.close
        		%>
          </td>
				</tr>
      </table>
  </td></tr>
  <tr bgcolor="#eeeeee"> 
    <td colspan="5" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;"> <table border=0 cellpadding="3" cellspacing="0">
        <% if Cint(pid) = 108 then %>
                <tr><td align="left">Charge Code <input type="text" name="Chargec"  value="<%=chargecode%>" size="20" ID="Text1"></td></tr>
				<tr><td><input type="checkbox" name="Sprinkler" <%if Sprinkler="True" then Response.Write "CHECKED"%> ID="Checkbox1">&nbsp;<span class="standard">Sprinkler</span></td></tr>	
		<% end if%>
          <tr bgcolor="#eeeeee"> 
            <td colspan=2>
				<input type="checkbox" value="1" name="Cumulative" <%if Cumulative="True" then Response.Write "CHECKED"%>>&nbsp;<span class="standard">Cumulative</span> &nbsp;&nbsp;&nbsp;
				<input type="checkbox" value="1" name="deduct" onClick="if(this.checked){this.form.lmchannel.value='';this.form.lmnum.value='';this.form.lmchannel.disabled=true;this.form.lmnum.disabled=true;}else{this.form.lmchannel.disabled=false;this.form.lmnum.disabled=false;}" <%if deduct="True" then Response.Write "CHECKED"%>>&nbsp;<span class="standard">Deduct</span> &nbsp;&nbsp;&nbsp; 
				<input type="checkbox" value="1" name="shared" <%if share="True" then Response.Write "CHECKED"%>>&nbsp;<span class="standard">Shared</span> &nbsp;&nbsp;&nbsp; 
				<input type="checkbox" value="1" name="lmp" <%if lmp="True" then Response.Write "CHECKED"%>>&nbsp;<span class="standard">LMP</span></td> 
          </tr>
        </table></td>
  </tr>
  <tr bgcolor="#cccccc"> 
    <td colspan="2" bgcolor="#cccccc" style="border-top:1px solid #999999;"> 
	<%if not(isBuildingOff(bldg)) then%>
      <%if trim(meterid)<>"0" then%>
      <input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
      <!-- 			<input type="submit" name="action" value="Delete" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> -->
      <input type="reset" name="action" value="Cancel" class="standard" onClick="document.location='meternull.htm';" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
			<%if allowGroups("Techsrv Admin,IT Services") and hasDatasource then%><input type="submit" name="action" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;" value="DPR SYNC"> <font color="#336699">*Record must be updated before enacting DPR SYNC.</font><%end if%>&nbsp;
      <%else%>
      <input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
      <input type="reset" name="action" value="Cancel" class="standard" onClick="document.location='tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>';" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
      <%end if%>
	<%end if%>&nbsp;
    </td>
    <td colspan="3" align="right" style="border-top:1px solid #999999;">
	
      <%if trim(meterid)<>"0" and trim(datasource)<>"" then
		        %>
		<input type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" value="Cavee Monitor" onClick="openCustomWin('/genergy2/umreports/cavee/caveeMonitorLog.asp?meterid=<%=meterid%>&bldg=<%=bldg%>&pid=<%=pid%>&tid=<%=tid%>&lid=<%=lid%>','LMPPopup','width=800,height=475')"> 
		<input type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" value="Meter LMP" onClick="openCustomWin('/genergy2/eri_th/lmp/lmp.asp?hideOptions=true&meterid=<%=meterid%>&bldg=<%=bldg%>&utility=<%=utilityid%>&indiWindow=true','LMPPopup','width=800,height=475')"> 
		<input type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" value="Interval Data" onClick="openCustomWin('/genergy2/UMreports/meterPulseReport.asp?meterid=<%=meterid%>&bldg=<%=bldg%>','','width=680,height=400,scrollbars=yes,resizable=yes')"> 
      <%
      end if%>
      <%if isnumeric(lastPCperiod) and isnumeric(lastPCyear) then%>
      <input type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" value="History" onClick="open('/genergy2/validation/update_billentry.asp?meterid=<%=meterid%>&byear=<%=lastPCyear%>&bperiod=<%=lastPCperiod%>&tname=<%=Server.URLEncode(billingname)%>&tnumber=&building=<%=bldg%>&pid=<%=pid%>&utilityid=<%=utilityid%>&posted=True', 'update_billentry','left=8,top=8,scrollbars=yes,width=770, height=380, status=no');">
      <%end if%>
      &nbsp;</td>
  </tr><input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="meterid" value="<%=meterid%>">
</form>
</table>

</body>
</html>
