<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%
dim pid, bldg, tid, lid, meterid
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
meterid = request("meterid")

dim cnn1, rst1, strsql
dim rst2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")

cnn1.open getConnect(pid,bldg,"billing")

dim Meternum, StartDate, DateOff, DateRead, TimeRead, Factor, Usagex, Capacityx, Location, Riser, online, note, Cumulative, variance, lmp, floor, datasource, customsrc, nobill, refmeterid, lastPCperiod, lastPCyear, manualentry, readorder, lmchannel, lmnum, units, Meter_CT_Ratio, Meter_Model, Meter_Voltage, metertype, cavee_monitor, data_frequency, deduct, extusg, meteraddonID,opentickets,totaltickets, ticketcount, masterticketid, share,Meter_JobID, Functionality
dim Sprinkler, chargecode
meteraddonID = 0

' Added by Tarun 1/28/2008
Dim MeterSerialNumber, StartUpDate, GatewayDeviceId, GatewayDeviceIdmnum, Manufacturer
Dim LocationNotes 'Tarun 2/21/2008

if trim(meterid)<>"" then
	rst1.Open "SELECT m.*, a.*, isnull(mp.addonfee,0) as meteraddonID, (SELECT top 1 cavee_monitor FROM cavee_setup cs WHERE cs.meterid="&meterid&") as cavee_monitor FROM meters m LEFT JOIN (SELECT top 1 c.billyear, c.billperiod, c.meterid FROM consumption c, peakdemand p WHERE p.BillYear=c.BillYear and p.BillPeriod=c.BillPeriod and p.meterid=c.meterid and c.meterid="&meterid&" ORDER BY c.billyear desc, c.billperiod desc) a ON m.meterid=a.meterid LEFT JOIN MeterPrices mp ON mp.meterid=m.meterid WHERE m.meterid="&meterid, cnn1
	if not rst1.EOF then
		Meternum= rst1("Meternum")
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
	end if
	rst1.close
	
	end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="setup.css" type="text/css" rel="stylesheet" />
<title>Untitled Document</title>
</head>

<body>
	<table style="border-top: 1px groove #EEEEEE; border-left: 1px groove #EEEEEE; border-right: 1px solid #000000" cellspacing="0" cellpadding="3" border="0" width="660">
  		<tr class="blueBack">
        	<td style="text-align: center"><h1 style="color: #FFFFFF">METER DETAILS - <%=meterid%> </h1></td>
        </tr>
        <tr class="greyBack">
        	<td style="border-bottom: 1px solid #000000; height: 20px">&nbsp;</td>
        </tr>
        <tr class="greyBack">
        	<td style="border-bottom: 1px solid #CCCCCC"><table cellspacing="2" cellpadding="2" border="0" align="center">
            	<tr>
                	<td><label>Meter Name:</label></td>
                    <td><%=Meternum%></td>
                    <td width="10"></td>
                    <td><label>Serial Number:</label></td>
                    <td><% %></td>
                </tr>
            </table>
        </tr>
        <tr class="greyBack">
        	<td style="border-bottom: 1px solid #000000;"><table style="margin-top: 10px; margin-bottom: 10px" cellspacing="2" cellpadding="2" border="0" align="center">
            		<tr>
                    	<td style="text-align: right"><label>Charge Code:</label></td>
                        <td><% %></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Factor:</label></td>
                        <td><label><%=Factor%></label></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Usage.x:</label></td>
                        <td><%=UsageX%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Capacity.x:</label></td>
                        <td><%=CapacityX%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Meter Reference:</label></td>
                        <td><% %></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>CT Ratio</label></td>
                        <td><%=Meter_CT_Ratio%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Manufacturer:</label></td>
                        <td><%=Manufacturer%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Model:</label></td>
                        <td><%=Meter_Model%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Voltage:</label></td>
                        <td>120</td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Location:</label></td>
                        <td><%=Location%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Floor:</label></td>
                        <td><%=Floor%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Meter Type:</label></td>
                        <td><%=MeterType%></td>
                    </tr>
            </table></td>
        </tr>
    </table>
</body>
</html>