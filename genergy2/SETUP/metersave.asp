<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'COMMENTS
'1/16/2008 N.Ambo modified to allow for updates, inserts, and deletes to be made to table tblPAMeterChargeCodes
'for charge codes associated with the meter (this is in regrards to the option that has now been added for the 
'user to enter those chargecodes on screen for PA meters)
'1/31/2008 N.Ambo modifed to include a new field for recording ntoes on the functionality of the meter (functionality_desc field)
%>
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim bldg, action, tid, pid, lid, meterid
meterid = request("meterid")
lid = request("lid")
tid = request("tid")
pid = request("pid")
bldg = request("bldg")
action = request("action")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql, strsqlDS, strsqlAO, strsqlCM, strsqlTh
dim strsqlSM
Dim strSQLML
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim Meternum, StartDate, DateOff, DateRead, TimeRead, Factor, Usagex, Capacityx, Functionality
Dim Location, Riser, online, note, Cumulative, variance, lmp, floor, datasource
Dim nobill, refmeterid, readorder, manualentry, lmchannel, lmnum, units, creatinguser
Dim Meter_CT_Ratio, Meter_Model, Meter_Voltage, metertype, data_frequency, deduct, extusg
Dim meteraddonID, share,Meter_JobID,rst2,sql
Dim Sprinkler, chargecode
Dim minValue, maxValue, estimate, gatewayIp, gatewayDevice, gateComm, ipidAddress, modbus, readerNotes

Dim Usage3MonthLowLimit,Usage3MonthHighLimit,Demand3MonthLowLimit,Demand3MonthHighLimit
Dim UsageLastMonthLowLimit, UsageLastMonthHighLimit, DemandLastMonthLowLimit, DemandLastMonthHighLimit
Dim UsageLastYrPeriodLowLimit,UsageLastYrPeriodHighLimit,DemandLastYrPeriodLowLimit,DemandLastYrPeriodHighLimit

' Added by Tarun 1/28/2007

Dim MeterSerialNumber, StartupDate, GateWayDeviceId, strSQLMXD, Manufacturer
Dim LocationNotes
Dim active 'RSM
Dim metersetupcharge 'RSM 11/10/2015

creatinguser = getXMLUserName
Meternum = request("Meternum")
StartDate = request("StartDate")
DateOff = request("DateOff")
DateRead = request("DateRead")
TimeRead = request("TimeRead")
Factor = request("factor")
Usagex = request("Usagex")
Capacityx = request("Capacityx")
Location = request("Location")
Riser = request("Riser")
online = request("online")
note = request("note")
Cumulative = request("Cumulative")
variance = request("variance")
lmp = request("lmp")
floor = request("floor")
datasource = request("datasource")
nobill = request("nobill")
refmeterid = request("refmeterid")
readorder = request("readorder")
manualentry = request("manualentry")
lmnum = trim(request("lmnum"))
lmchannel = trim(request("lmchannel"))
units = request("units")
Meter_CT_Ratio = request("Meter_CT_Ratio")
Meter_Model = request("Meter_Model")
Meter_Voltage = request("Meter_Voltage")
metertype = request("metertype")
data_frequency = request("data_frequency")
deduct = request("deduct")
extusg = request("extusg")
meteraddonID = request("meteraddonID")
share = request("shared")
Meter_JobID = request("Meter_JobID")
estimate = request("estimate")
minValue = request("minValue") '5/11/2009 KCheng added
maxValue = request("maxValue") '5/11/2009 KCheng added
gatewayIp = request("gatewayIP")'5/18,2009 KCheng added
gatewayDevice = request("gatewayDevice")'5/18/2009 Kcheng added
gateComm = request("gateComm")
ipidAddress = request("ip_id")
Functionality = request("Functionality") '1/31/2008 N.Ambo added
readerNotes = request("readerNotes")
active = request("active")  'rsm
metersetupcharge = request("metersetupcharge")  'rsm 11/10/2015

if (lcase(gateComm)="pulse" OR gateComm = "") then
    modbus = 0
else
    modbus = 1
end if

if (request("minValue") = "" ) then minValue = 0
if (request("maxValue") = "" ) then maxValue = 999999


' Meter Variance Limits
Usage3MonthLowLimit = request("Usg3MonthLowLimit")
Usage3MonthHighLimit = request("Usg3MonthHighLimit")
Demand3MonthLowLimit = request("dmd3MonthLowLimit")
Demand3MonthHighLimit = request("dmd3MonthHighLimit")
				
UsageLastMonthLowLimit = request("UsgLastMonthLowLimit")
UsageLastMonthHighLimit = request("UsgLastMonthHighLimit")
DemandLastMonthLowLimit = request("dmdLastMonthLowLimit")
DemandLastMonthHighLimit = request("dmdLastMonthHighLimit")
				

UsageLastYrPeriodLowLimit = request("UsgLastYrPeriodLowLimit")
UsageLastYrPeriodHighLimit = request("UsgLastYrPeriodHighLimit")
DemandLastYrPeriodLowLimit = request("dmdLastYrPeriodLowLimit")
DemandLastYrPeriodHighLimit = request("dmdLastYrPeriodHighLimit")

' Added by Tarun 1/28/2007
MeterSerialNumber = request("SerialNum")
StartupDate = request("StartUpDate")
GateWayDeviceId = request("GateWayDeviceId")
Manufacturer = request("Manufacturer")

'If GateWayDeviceId = "" then 
'	GateWayDeviceId = "0"
'End IF

LocationNotes = request("LocationNotes")

if Cint(pid) = 108 then
	Sprinkler = Request("Sprinkler")
	chargecode = request("Chargec")
	if  Sprinkler = "on" then
		Sprinkler = 1 
	else
		Sprinkler = 0 
	end if
end if 

if trim(metertype) = "" then metertype = 0
if trim(manualentry) <> "1" then manualentry = 0
if trim(deduct) <> "1" then deduct = 0
if trim(data_frequency) = "" then data_frequency = 0
if trim(extusg) <> "1" then extusg = 0
if trim(share) <> "1" then share = 0
if trim(meternum)="" then
	response.write "No meter number given."
	response.end
end if

if instr(trim(action),"Confirm")=0 and trim(online)<>"1" then%>
  <link rel="Stylesheet" href="setup.css" type="text/css">
  <form method="post" action="metersave.asp" class="standard">
  Are you sure you want to make this meter offline? <input type="submit" name="action" value="Confirm <%=trim(action)%>">&nbsp;<input type="button" value="Back" onclick="history.back();"><br>
  <input type="hidden" name="creatinguser" value="<%=creatinguser%>">
  <input type="hidden" name="Meternum" value="<%=Meternum%>">
  <input type="hidden" name="StartDate" value="<%=StartDate%>">
  <input type="hidden" name="DateOff" value="<%=DateOff%>">
  <input type="hidden" name="DateRead" value="<%=DateRead%>">
  <input type="hidden" name="TimeRead" value="<%=TimeRead%>">
  <input type="hidden" name="Factor" value="<%=Factor%>">
  <input type="hidden" name="Usagex" value="<%=Usagex%>">
  <input type="hidden" name="Capacityx" value="<%=Capacityx%>">
  <input type="hidden" name="Location" value="<%=Location%>">
  <input type="hidden" name="Riser" value="<%=Riser%>">
  <input type="hidden" name="online" value="<%=online%>">
  <input type="hidden" name="note" value="<%=note%>">
  <input type="hidden" name="Cumulative" value="<%=Cumulative%>">
  <input type="hidden" name="variance" value="<%=variance%>">
  <input type="hidden" name="lmp" value="<%=lmp%>">
  <input type="hidden" name="floor" value="<%=floor%>">
  <input type="hidden" name="datasource" value="<%=datasource%>">
  <input type="hidden" name="nobill" value="<%=nobill%>">
  <input type="hidden" name="refmeterid" value="<%=refmeterid%>">
  <input type="hidden" name="meterid" value="<%=meterid%>">
  <input type="hidden" name="lid" value="<%=lid%>">
  <input type="hidden" name="tid" value="<%=tid%>">
  <input type="hidden" name="pid" value="<%=pid%>">
  <input type="hidden" name="bldg" value="<%=bldg%>">
  <input type="hidden" name="readorder" value="<%=readorder%>">
  <input type="hidden" name="manualentry" value="<%=trim(manualentry)%>">
  <input type="hidden" name="lmnum" value="<%=lmnum%>">
  <input type="hidden" name="lmchannel" value="<%=lmchannel%>">
  <input type="hidden" name="units" value="<%=units%>">
  <input type="hidden" name="Meter_CT_Ratio" value="<%=Meter_CT_Ratio%>">
  <input type="hidden" name="Meter_Model" value="<%=Meter_Model%>">
  <input type="hidden" name="Meter_Voltage" value="<%=Meter_Voltage%>">
  <input type="hidden" name="metertype" value="<%=metertype%>">
  <input type="hidden" name="data_frequency" value="<%=data_frequency%>">
  <input type="hidden" name="deduct" value="<%=deduct%>">
  <input type="hidden" name="extusg" value="<%=extusg%>">
  <input type="hidden" name="meteraddonID" value="<%=meteraddonID%>">
  <input type="hidden" name="share" value="<%=share%>">
  <input type="hidden" name="Meter_JobID" value="<%=Meter_JobID%>">
  <input type="hidden" name="Sprinkler" value="<%=Sprinkler%>">
  <input type="hidden" name="Chargec" value="<%=chargecode%>">
  <input type="hidden" name="Functionality" value="<%=Functionality%>">
  <input type="hidden" name="minValue" value="<%=minValue%>" />
  <input type="hidden" name="maxValue" value="<%=maxValue%>" />
  <input type="hidden" name="gatewayIp" value="<%=gatewayIp%>" />
  <input type="hidden" name="gatewayDevice" value="<%=gatewayDevice%>" />
  <input type="hidden" name="gateComm" value="<%=gateComm %>" />
  <input type="hidden" name="ipidAdress" value="<%=ipidAddress%>" />
  <input type="hidden" name="modbus" value="<%=modbus%>" />
  <input type="hidden" name="readerNotes" value="<%=readerNotes %>" />
  
  <input type="hidden" name="Usg3MonthLowLimit" value="<%=Usage3MonthLowLimit%>">
  <input type="hidden" name="Usg3MonthHighLimit" value="<%=Usage3MonthHighLimit%>">
  <input type="hidden" name="dmd3MonthLowLimit" value="<%=Demand3MonthLowLimit%>">
  <input type="hidden" name="dmd3MonthHighLimit" value="<%=Demand3MonthHighLimit%>">
  
   <input type="hidden" name="UsgLastMonthLowLimit" value="<%=UsageLastMonthLowLimit%>">
  <input type="hidden" name="UsgLastMonthHighLimit" value="<%=UsageLastMonthHighLimit%>">
  <input type="hidden" name="dmdLastMonthLowLimit" value="<%=DemandLastMonthLowLimit%>">
  <input type="hidden" name="dmdLastMonthHighLimit" value="<%=DemandLastMonthHighLimit%>"> 
  
  <input type="hidden" name="UsgLastYrPeriodLowLimit" value="<%=UsageLastYrPeriodLowLimit%>">
  <input type="hidden" name="UsgLastYrPeriodHighLimit" value="<%=UsageLastYrPeriodHighLimit%>">
  <input type="hidden" name="dmdLastYrPeriodLowLimit" value="<%=DemandLastYrPeriodLowLimit%>">
  <input type="hidden" name="dmdLastYrPeriodHighLimit" value="<%=DemandLastYrPeriodHighLimit%>"> 

  <input type="hidden" name="SerialNum" value="<%=MeterSerialNumber%>">
  <input type="hidden" name="StartUpDate" value="<%=StartupDate%>">
  <input type="hidden" name="GatewayDeviceId" value="<%=GatewayDeviceId%>"> 
   <input type="hidden" name="Manufacturer" value="<%=Manufacturer%>"> 
   
   <input type="hidden" name="LocationNotes" value="<%=LocationNotes%>"> 
   <input type="hidden" name="active" value="<%=active%>"> 
   <input type="hidden" name="metersetupcharge" value="<%=metersetupcharge%>"> 
  </form>
  <%
  response.end
end if
if instr(trim(action),"Confirm")>0 then action = mid(trim(action),9)

if trim(action)="Save" then
	strsql = "INSERT INTO meters ([user], Meternum, DateStart, DateOffLine, DateLastRead, TimeLastRead, multiplier, manualmultiplier, " & _
				"Demandmultiplier, Location, Riser, online, metercomments, Cumulative, leaseutilityid, bldgnum, gatewayDevice, gateCommunication, modBusIdentifier, modbus, " & _
				" variance, lmp, floor, datasource, nobill, refmeterid, manualentry, readorder, lmnum, lmchannel, reader_notes, " & _
				" calculate, CT_Ratio, Manufacturer, Model, voltage, category, data_frequency, deduct, extusg, shared,job_id, functionality_desc, LocationNotes, active, metersetupcharge) " & _
			  " values ('"&creatinguser&"','"&Meternum&"', '"&StartDate&"', '"&DateOff&"', '"&DateRead&"', '"&TimeRead&"', '"&Factor&"', '"&Usagex & _
				"', '"&Capacityx&"', '"&Location&"', '"&Riser&"', '"&online&"', '"&note&"', '"&Cumulative&"', '"&lid&"', '"&bldg&"', '"&gatewayDevice&"', '"&gateComm&"', '"&ipidAddress&"', '"&modbus&"', '" & _
				variance&"', '"&lmp&"', '"&floor&"', '"&datasource&"', '"&nobill&"', '"&refmeterid&"', '"&trim(manualentry)&"', '"&readorder&"', '"&lmnum&"', '"&lmchannel&"', '"&readerNotes&"', '" & _
				units&"', '"&Meter_CT_Ratio & "', '" & Manufacturer & "', '" & Meter_Model&"', '"&Meter_Voltage&"', '"&metertype&"', '"&data_frequency&"', '"&deduct&"', "&extusg&", "&share&",'"&trim(Meter_JobID)&"','"&functionality&"','"&LocationNotes&"','"&active&"','"&metersetupcharge&"')"
			  
elseif trim(action)="Delete" then
	strsql = "DELETE FROM meters WHERE meterid="&meterid
elseif trim(action)="DPR SYNC" then
	strsql = "exec sp_Sync_DPR "&meterid
else
	strsql = "UPDATE meters set Meternum='"&Meternum&"', DateStart='"&StartDate & _
					"', DateOffLine='"&DateOff&"', DateLastRead='"&DateRead & _
				    "', gateWayIP='"&gatewayIP&"', gatewayDevice='"&gatewayDevice & _
				    "', gateCommunication='"&gateComm&"', modBusIdentifier='"&ipidAddress&"', modbus='"&modbus & _
					"', TimeLastRead='"&TimeRead&"', multiplier='"&Factor&"', reader_notes='"&readerNotes & _
					"', manualmultiplier='"&Usagex&"', Demandmultiplier='"&Capacityx & _
					"', Location='"&Location&"', Riser='"&Riser&"', online='"&online& _
					"', metercomments='"&note&"', Cumulative='"&Cumulative&"', leaseutilityid='"&lid & _
					"', bldgnum='"&bldg&"', variance='"&variance&"', lmp='"&lmp&"', floor='"&floor & _
					"', datasource='"&datasource&"', nobill='"&nobill&"', refmeterid='"&refmeterid & _
					"', readorder='"&readorder&"', manualentry='"&trim(manualentry)&"', lmnum='"&lmnum & _
					"', lmchannel='"&lmchannel&"', calculate='"&units&"', CT_Ratio='"&Meter_CT_Ratio & _
					"', Manufacturer='" & Manufacturer & "', Model='"&Meter_Model & "', voltage='"&Meter_Voltage&"', category='"&metertype & _
					"', data_frequency='"&data_frequency&"', deduct='"&deduct&"', extusg="&extusg & _
					", shared="&share&", job_id = '"&trim(Meter_JobID)&"', functionality_desc = '"&functionality&"', LocationNotes = '" & LocationNotes & "', active='"&active&"', metersetupcharge='"&metersetupcharge&"' WHERE meterid="&meterid
end if

'response.Write(strsql)
'response.end
'Logging Update
logger(strsql)
'end Log

cnn1.Execute strsql

if trim(action)="Save" then
	rst1.open "SELECT max(meterid) as id FROM meters", cnn1
	meterid = rst1("id")
	rst1.close
	if Cint(pid) = 108 then
		strsqlSM = "INSERT INTO tblPASprinklerMeters (meterid, Sprinkler) VALUES ('"&meterid&"', '"&Sprinkler&"')"
		strsqlCM = "INSERT INTO tblPAMeterChargeCodes (bldgnum,meternum, chargecode) VALUES ('"&bldg&"', '"&meternum&"', '"&chargecode&"')"
	end if
	strSqlTH = "INSERT INTO MeterThreshold (meterid, MinUsagethresh, MaxUsagethresh) VALUES ('" & meterid & "', '" & minValue & "', '" & maxValue &"')"
	strsqlDS = "INSERT INTO datasource (meterid, datasource) VALUES ('"&meterid&"', '"&datasource&"')"
	strsqlAO = "UPDATE MeterPrices SET AddonFee="&meteraddonID&" WHERE meterid='"&meterid&"'"
	strSQLML = "INSERT INTO tblMeterVarianceLimits (meterid, [Usage3MonthLowLimit] ,[Usage3MonthHighLimit]" & _
					",[Demand3MonthLowLimit], [Demand3MonthHighLimit], [UsageLastMonthLowLimit]" & _
					",[UsageLastMonthHighLimit], [DemandLastMonthLowLimit], [DemandLastMonthHighLimit] " & _
					",[UsageLastYrPeriodLowLimit], [UsageLastYrPeriodHighLimit], [DemandLastYrPeriodLowLimit],[DemandLastYrPeriodHighLimit]) " & _
			   "Values (" & meterid & "," & Usage3MonthLowLimit & "," & Usage3MonthHighLimit & ", " & Demand3MonthLowLimit & ", " & _
						Demand3MonthHighLimit & ", " & UsageLastMonthLowLimit & ", " & UsageLastMonthHighLimit& ", " &  _
						DemandLastMonthLowLimit & "," & DemandLastMonthHighLimit & "," & UsageLastYrPeriodLowLimit &  ", " & _
						UsageLastYrPeriodHighLimit & ", " & DemandLastYrPeriodLowLimit & ", " & DemandLastYrPeriodHighLimit & ")"
						
	' Added by Tarun 1/28/2008
	strSQLMXD = "INSERT INTO tblMeterExtDetails (MeterId, SerialNumber, StartUpdate, GatewayDeviceId) " & _
				" VALUES (" & meterid & ",'" & MeterSerialNumber & "','" & StartupDate & "','" & GateWayDeviceId & "')"
				
elseif trim(action)="Delete" then
	if Cint(pid) = 108 then
		strsqlSM = "DELETE FROM tblPASprinklerMeters WHERE meterid='"&meterid&"'"
		strsqlCM = "DELETE FROM tblPAmeterChargeCodes WHERE meternum= '" &meternum& "'"
	end if
	strsqlDS = "DELETE FROM datasource WHERE meterid='"&meterid&"'"
	strsqlAO = "DELETE FROM MeterPrices WHERE meterid='"&meterid&"'"
	strsqlTH = "DELETE FROM MeterThreshold where meterid ='"&meterid&"'"

	
	strSQLML = "DELETE FROM tblMeterVarianceLimits WHERE meterid='"&meterid&"'"
	
elseif trim(action)="DPR SYNC" then
	strsqlDS = ""
else
	sql = "select * from meterprices where meterid = '"&meterid&"'" 
	rst2.open sql,cnn1
	if rst2.EOF then 
	strsqlDS = "UPDATE datasource SET datasource='"&datasource&"' WHERE meterid='"&meterid&"'"
	strsqlAO = "INSERT INTO MeterPrices (meterid, Addonfee) VALUES ('"&meterid&"', '"&meteraddonID&"')"
	
	else
		strsqlDS = "UPDATE datasource SET datasource='"&datasource&"' WHERE meterid='"&meterid&"'"
		strsqlAO = "UPDATE MeterPrices SET AddonFee="&meteraddonID&" WHERE meterid='"&meterid&"'"
		strSQLML = "DELETE FROM tblMeterVarianceLimits WHERE meterid=" & meterid & _
					" INSERT INTO tblMeterVarianceLimits (meterid, [Usage3MonthLowLimit] ,[Usage3MonthHighLimit]" & _
					",[Demand3MonthLowLimit], [Demand3MonthHighLimit], [UsageLastMonthLowLimit]" & _
					",[UsageLastMonthHighLimit], [DemandLastMonthLowLimit], [DemandLastMonthHighLimit] " & _
					",[UsageLastYrPeriodLowLimit], [UsageLastYrPeriodHighLimit], [DemandLastYrPeriodLowLimit],[DemandLastYrPeriodHighLimit]) " & _
			   "Values (" & meterid & "," & Usage3MonthLowLimit & "," & Usage3MonthHighLimit & ", " & Demand3MonthLowLimit & ", " & _
						Demand3MonthHighLimit & ", " & UsageLastMonthLowLimit & ", " & UsageLastMonthHighLimit & ", " &  _
						DemandLastMonthLowLimit & "," & DemandLastMonthHighLimit & "," & UsageLastYrPeriodLowLimit &  ", " & _
						UsageLastYrPeriodHighLimit & ", " & DemandLastYrPeriodLowLimit & ", " & DemandLastYrPeriodHighLimit & ")"
						
	' Added by Tarun 1/28/2008
	strSQLMXD = "DELETE FROM tblMeterExtDetails WHERE meterid=" & meterid & _
				" INSERT INTO tblMeterExtDetails (MeterId, SerialNumber, StartUpdate, GatewayDeviceId) " & _
				" VALUES (" & meterid & ",'" & MeterSerialNumber & "','" & StartupDate & "','" & GateWayDeviceId & "')"
		
	'response.Write(strSQLML)					
	end if
	
	strsqlTH = "UPDATE meterThreshold SET MinUsagethresh='"&minValue&"',  MaxUsagethresh='"&maxValue&"' WHERE meterid='"&meterid&"'"
	
	if Cint(pid) = 108 then
		strsqlSM = " DELETE FROM tblPASprinklerMeters WHERE meterid='"&meterid&"'  INSERT INTO tblPASprinklerMeters (meterid, Sprinkler) VALUES ('"&meterid&"', '"&Sprinkler&"')"
		strsqlCM = " DELETE FROM tblPAMeterChargeCodes WHERE meternum='"&meternum&"'  INSERT INTO tblPAMeterChargeCodes (bldgnum,meternum, chargecode) VALUES ('"&bldg&"', '"&meternum&"', '"&chargecode&"')"
	end if
	
end if
'response.Write strsqlDS &"<BR>"
'response.Write strsqlAO  
'Response.Write strSQLML
'response.End

'5/12/2009 KCheng added
if (estimate = "on") then
    'response.Write(strsqlTH)
    if trim(strsqlTH)<>"" then logger(strsqlTH)
    if trim(strsqlTH)<>"" then cnn1.Execute strsqlTH
end if

'Logging Update
if trim(strsqlDS)<>"" then logger(strsqlDS)
if trim(strsqlAO)<>"" then logger(strsqlAO)
if trim(strsqlSM)<>"" then logger(strsqlSM)
if trim(strsqlML)<>"" then logger(strsqlML)
if trim(strsqlCM)<>"" then logger(strsqlCM)
if trim(strSQLMXD)<>"" then logger(strSQLMXD)
'end Log

if trim(strsqlDS)<>"" then cnn1.Execute strsqlDS
if trim(strsqlAO)<>"" then cnn1.Execute strsqlAO
if trim(strsqlSM)<>"" then cnn1.Execute strsqlSM
if trim(strsqlML)<>"" then cnn1.Execute strsqlML
if trim(strsqlCM)<>"" then cnn1.Execute strsqlCM
if trim(strsqlMXD)<>"" then cnn1.Execute strsqlMXD

if trim(action)="Delete" or trim(action)="Update" or trim(action)="DPR SYNC" then
	%>
	<script>
	parent.document.location = 'contentfrm.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&action=meterview&meterid=<%=meterid%>'
	</script>
	<%
else
	Response.Redirect "contentfrm.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&meterid="&meterid&"&action=meterview"
end if
%>