<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%


dim tenantnum, flr, sqft, taxexempt, billingname, leaseexpired, interm, intermcharges, startdate, bldg, action, tid, pid, tName, tStrt, tCity, tState, tZip, tCountry, lmepExempt, onlinebill, ibsexempt, bsexempt, accounttype, lookup
dim corpStreet, corpCity, corpState, corpZip, corpCountry, corpAddressSameAsBillingCheck
dim leaseNumber, sequenceNumber
dim AcctCode
dim WattsperSqFtLow, WattsperSqFtHigh

' Added by Tarun 01/25/2008 :
dim TenantMoveInDate, strsql5, TenantEmail

tenantnum = secureRequest("tenantnum")
leaseNumber = secureRequest("leasenum")
sequenceNumber = secureRequest("seqnum")

AcctCode = secureRequest("AcctCode")

tid = secureRequest("tid")
pid = secureRequest("pid")
flr = secureRequest("flr")
sqft = secureRequest("sqft")
taxexempt = secureRequest("taxexempt")
billingname = secureRequest("billingname")
leaseexpired = secureRequest("leaseexpired")
interm = secureRequest("interm")
intermcharges = secureRequest("intermcharges")
startdate = secureRequest("startdate")
bldg = secureRequest("bldg")
action = secureRequest("action")
tName = secureRequest("tName")
tStrt = secureRequest("tStrt")
tCity = secureRequest("tCity")
tState = secureRequest("tState")
tZip = secureRequest("tZip")
tCountry = secureRequest("tCountry")
lmepExempt = secureRequest("lmepExempt")
onlinebill = secureRequest("onlinebill")
ibsexempt = secureRequest("ibsexempt")
bsexempt = secureRequest("bsexempt")
accounttype = secureRequest("accounttype")
corpAddressSameAsBillingCheck = secureRequest("corpAddressSameAsBillingCheck")
corpStreet = secureRequest("corpStreet")
corpCity = secureRequest("corpCity")
corpState = secureRequest("corpState")
corpZip = secureRequest("corpZip")
corpCountry = secureRequest("corpCountry")

' Watts per Sqft Variance Limits
WattsperSqFtLow =secureRequest("WattsPerSqFtLowLimit")
WattsperSqFtHigh =secureRequest("WattsPerSqFtHighLimit")

' Added by Tarun 01/25/2008 :
TenantMoveInDate = secureRequest("TenantMoveInDate")
TenantEmail = secureRequest("TenantEmail")

if not (corpAddressSameAsBillingCheck <> "on") then
	corpStreet = null
	corpCity = null
	corpState = null
	corpZip = null
	corpCountry = null
end if 

dim cnn1, rst1, strsql, strsql2, strsql3
dim strsql4
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim returnpage
returnpage = "tenantedit_pa.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&edit=0"
if instr(trim(action),"Confirm")=0 and trim(leaseexpired)="1" then%>
<link rel="Stylesheet" href="setup.css" type="text/css">
<form method="post" action="tenantsave.asp" class="standard">
Are you sure you want to make this tenant offline? <input type="submit" name="action" value="Confirm <%=trim(action)%>">&nbsp;<input type="button" value="Back" onclick="history.back();"><br>
<input type="hidden" name="tenantnum" value="<%=tenantnum%>">
<input type="hidden" name="leasenum" value="<%=leaseNumber%>">
<input type="hidden" name="AcctCode" value="<%=AcctCode%>">
<input type="hidden" name="seqnum" value="<%=sequenceNumber%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="flr" value="<%=flr%>">
<input type="hidden" name="sqft" value="<%=sqft%>">
<input type="hidden" name="taxexempt" value="<%=taxexempt%>">
<input type="hidden" name="billingname" value="<%=billingname%>">
<input type="hidden" name="leaseexpired" value="<%=leaseexpired%>">
<input type="hidden" name="interm" value="<%=interm%>">
<input type="hidden" name="intermcharges" value="<%=intermcharges%>">
<input type="hidden" name="startdate" value="<%=startdate%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tName" value="<%=tName%>">
<input type="hidden" name="tStrt" value="<%=tStrt%>">
<input type="hidden" name="tCity" value="<%=tCity%>">
<input type="hidden" name="tState" value="<%=tState%>">
<input type="hidden" name="tZip" value="<%=tZip%>">
<input type="hidden" name="tCountry" value="<%=tCountry%>">
<input type="hidden" name="lmepExempt" value="<%=lmepExempt%>">
<input type="hidden" name="onlinebill" value="<%=onlinebill%>">
<input type="hidden" name="ibsexempt" value="<%=ibsexempt%>">
<input type="hidden" name="bsexempt" value="<%=bsexempt%>">
<input type="hidden" name="accounttype" value="<%=accounttype%>">
<input type="hidden" name="corpAddressSameAsBillingCheck" value="<%=corpAddressSameAsBillingCheck%>">
<input type="hidden" name="corpStreet" value="<%=corpStreet%>">
<input type="hidden" name="corpCity" value="<%=corpCity%>">
<input type="hidden" name="corpState" value="<%=corpState%>">
<input type="hidden" name="corpZip" value="<%=corpZip%>">
<input type="hidden" name="corpCountry" value="<%=corpCountry%>">

<input type="hidden" name="WattsPerSqFtLowLimit" value="<%=WattsperSqFtLow%>">
<input type="hidden" name="WattsPerSqFtHighLimit" value="<%=WattsperSqFtHigh%>">

<input type="hidden" name="TenantMoveInDate" value="<%=TenantMoveInDate%>">
<input type="hidden" name="TenantEmail" value="<%=TenantEmail%>">

</form>
<%
response.end
end if
if instr(trim(action),"Confirm")>0 then action = mid(trim(action),9)
if trim(action)="Save" then
	strsql = "INSERT INTO tblleases (tenantnum, flr, sqft, taxexempt, billingname, leaseexpired, interm, intermcharges, startdate, bldgnum, tName, tStrt, tCity, tState, tZip, tCountry, lmepExempt, onlinebill, ibsexempt, billsummaryexempt, accounttype, corpStreet, corpCity, corpState, corpZip, corpCountry) values ('"&tenantnum&"', '"&flr&"', '"&sqft&"', '"&taxexempt&"', '"&billingname&"', '"&leaseexpired&"', '"&interm&"', '"&intermcharges&"', '"&startdate&"', '"&bldg&"', '"&tName&"', '"&tStrt&"', '"&tCity&"', '"&tState&"', '"&tZip&"', '"&tCountry&"', '"&lmepExempt&"', '"&onlinebill&"', '"&ibsexempt&"','"&bsexempt&"', '"&accounttype&"', '"&corpStreet&"', '"&corpCity&"', '"&corpState&"', '"&corpZip&"', '"&corpCountry&"')"
	lookup = "SELECT top 1 billingid FROM tblleases ORDER BY billingid desc"
	returnpage = "buildingedit.asp?pid="&pid&"&bldg="&bldg&"&edit=0"
elseif trim(action)="Delete" then
	strsql = "DELETE FROM tblleases WHERE billingid="&tid
	strsql2 = "DELETE FROM custom_PABT WHERE acctnumber='"&tid&"'"
	strsql3 = "DELETE FROM custom_SL WHERE acctnumber='"&tid&"'"
	strsql4 = "DELETE FROM tblTenantVarianceLimits WHERE billingid="&tid
else
	strsql = "UPDATE tblleases set tenantnum='"&tenantnum&"', flr='"&flr&"', sqft='"&sqft&"', taxexempt='"&taxexempt&"', billingname='"&billingname&"', leaseexpired='"&leaseexpired&"', interm='"&interm&"', intermcharges='"&intermcharges&"', startdate='"&startdate&"', bldgnum='"&bldg&"', tName='"&tName&"', tStrt='"&tStrt&"', tCity='"&tCity&"', tState='"&tState&"', tZip='"&tZip&"', tCountry='"&tCountry&"', lmepExempt='"&lmepExempt&"', onlinebill='"&onlinebill&"', ibsexempt='"&ibsexempt&"',billsummaryexempt='"&bsexempt&"', accounttype='"&accounttype&"', corpStreet='"&corpStreet&"', corpCity='"&corpCity&"', corpState='"&corpState&"', corpZip='"&corpZip&"', corpCountry='"&corpCountry&"' WHERE billingid="&tid
	strsql2 = "UPDATE custom_PABT set leasenumber='"&leaseNumber&"', seqnumber='"&sequenceNumber&"' WHERE acctnumber='"&tid&"'"
	strsql3 = "UPDATE custom_SL set AcctCode='"&AcctCode&"' WHERE acctnumber='"&tid&"'"
	strsql4 = "DELETE FROM tblTenantVarianceLimits WHERE billingid="&tid & _
			  " INSERT INTO tblTenantVarianceLimits (BillingId, WattsPerSqFtLowLimit, WattsPerSqFtHighLimit) " & _
			  " VALUES (" & tid & "," & WattsperSqFtLow & "," & WattsperSqFtHigh & ")"
	'Added by Tarun 1/25/2008 :
	strSql5 = " DELETE FROM tblTenantExtDetails WHERE billingid = " & tid & _
			  " INSERT INTO tblTenantExtDetails (BillingId,MoveInDate, TenantEmail) " & _
			  " VALUES (" & tid & ", '" & TenantMoveInDate & "', '" & TenantEmail & "')" 		  
	
end if

if not trim(action)= "Transfer Info To New Account" then
	'on error resume next

'Logging Update
logger(strsql)
'end Log
	cnn1.Execute strsql
	if err.number = -2147217900 then 
		response.write "<span style=""font-family:Arial,Helvetica,sans-serif; font-size:8pt;"">Accounts with existing lease utilities may not be deleted. If the account is no longer active, please mark the checkbox labeled ""Lease Expired"".<br><input type=button value=""Back"" onclick=""history.back();"" style=""font-family:Arial,Helvetica,sans-serif; font-size:8pt;cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;""></span>"
		response.end
	end if
	if lookup<>"" then
		rst1.open lookup, cnn1
		if not rst1.eof then tid = rst1(0)
		if tid<>"" then 
			strsql2 = "INSERT INTO custom_PABT (acctnumber, leasenumber, seqnumber) values ('"&tid&"', '"&leaseNumber&"', '"&sequenceNumber&"')"
			strsql3 = "INSERT INTO custom_SL (acctnumber, AcctCode) values ('"&tid&"', '"&AcctCode&"')"
			strsql4 = " INSERT INTO tblTenantVarianceLimits (BillingId, WattsPerSqFtLowLimit, WattsPerSqFtHighLimit) " & _
					  " VALUES (" & tid & "," & WattsperSqFtLow & "," & WattsperSqFtHigh & ")"
			'Added by Tarun 1/25/2008 :
			strSql5 = " INSERT INTO tblTenantExtDetails (BillingId,MoveInDate, TenantEmail) " & _
						" VALUES (" & tid & ", '" & TenantMoveInDate & "', '" & TenantEmail & "')" 	  
						  			
			returnpage = "leaseutilityedit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&edit=0"
		end if
	end if
	
	'custom port authority data
	if pid = 108 then
		cnn1.Execute strsql2
	end if
	if pid = 45 then
		cnn1.Execute strsql3
	end if
	cnn1.Execute strsql4
	
	cnn1.Execute strsql5
end if

if trim(action)="Transfer Info To New Account" then
'  dim newtid
'	rst1.Open "SELECT max(billingid) as id FROM "&DBlocalIP&"tblleases", cnn1
'	if not rst1.eof then newtid = rst1("id")
	returnpage = "tenantTransfer.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid
end if

  Response.Redirect returnpage

%>