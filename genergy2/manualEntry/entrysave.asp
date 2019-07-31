<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

dim bldg, action, pid, meterid, byear, bperiod, wyear, wperiod, rawprevious, rawonpeak, rawoffpeak, rawintpeak, rawcurrent, rawused, addperiod, addyear, estimatedc, usernotec, datepeak, estimatedp, usernotep, rawdemand, rawprev, scroll,rawpreviousoff,rawpreviousint, rawcurrentoff, rawcurrentint, rawusedoff, rawusedint, datepeak_off, rawdemand_off, rawprev_off, datepeak_int, rawdemand_int, rawprev_int
meterid = request("meterid")
pid = request("pid")
bldg = request("building")
action = request("action")
byear = request("byear")
bperiod = request("bperiod")
wyear = request("workingyear")
wperiod = request("workingperiod")
scroll = request("scroll")
if instr(request("addperiod"),"|")>0 then 
	addperiod = split(request("addperiod"),"|")(0)
	addyear = split(request("addperiod"),"|")(1)
end if
'consumption fields
if isnumeric(request("rawonpeak")) and trim(request("rawonpeak"))<>"" then rawonpeak = request("rawonpeak") else rawonpeak = 0
if isnumeric(request("rawoffpeak")) and trim(request("rawoffpeak"))<>"" then rawoffpeak = request("rawoffpeak") else rawoffpeak = 0
if isnumeric(request("rawintpeak")) and trim(request("rawintpeak"))<>"" then rawintpeak = request("rawintpeak") else rawintpeak = 0
if isnumeric(request("rawprevious")) and trim(request("rawprevious"))<>"" then rawprevious = request("rawprevious") else rawprevious = 0
if isnumeric(request("rawcurrent")) and trim(request("rawcurrent"))<>"" then rawcurrent = request("rawcurrent") else rawcurrent = 0
if isnumeric(request("rawcurrentoff")) and trim(request("rawcurrentoff"))<>"" then rawcurrentoff = request("rawcurrentoff") else rawcurrentoff = 0
if isnumeric(request("rawcurrentint")) and trim(request("rawcurrentint"))<>"" then rawcurrentint = request("rawcurrentint") else rawcurrentint = 0
if isnumeric(request("rawpreviousoff")) and trim(request("rawpreviousoff"))<>"" then rawpreviousoff = request("rawpreviousoff") else rawpreviousoff = 0
if isnumeric(request("rawpreviousint")) and trim(request("rawpreviousint"))<>"" then rawpreviousint = request("rawpreviousint") else rawpreviousint = 0
if isnumeric(request("rawused")) and trim(request("rawused"))<>"" then rawused = request("rawused") else rawused = 0
if isnumeric(request("rawusedoff")) and trim(request("rawusedoff"))<>"" then rawusedoff = request("rawused") else rawusedoff = 0
if isnumeric(request("rawusedint")) and trim(request("rawusedint"))<>"" then rawusedint = request("rawused") else rawusedint = 0


'datepeak_off, rawdemand_off, rawprev_off, datepeak_int, rawdemand_int, rawprev_int

if isnumeric(request("estimatedc")) and trim(request("estimatedc"))<>"" then estimatedc = request("estimatedc")
usernotec = request("usernotec")
'demand fields
if isnumeric(request("rawprev")) and trim(request("rawprev"))<>"" then rawprev = request("rawprev") else rawprev = 0
if isdate(request("datepeak")) and trim(request("datepeak"))<>"" then datepeak = request("datepeak") else datepeak = date()
if isnumeric(request("estimatedp")) and trim(request("estimatedp"))<>"" then estimatedp = request("estimatedp") else estimatedp = 0
usernotep = request("usernotep")
if isnumeric(request("rawdemand")) and trim(request("rawdemand"))<>"" then rawdemand = request("rawdemand") else rawdemand = 0
'extended demand fields
if isnumeric(request("rawprev_off")) and trim(request("rawprev_off"))<>"" then rawprev_off = request("rawprev_off") else rawprev_off = 0
if isdate(request("datepeak_off")) and trim(request("datepeak_off"))<>"" then datepeak_off = request("datepeak_off") else datepeak_off = date()
if isnumeric(request("rawdemand_off")) and trim(request("rawdemand_off"))<>"" then rawdemand_off = request("rawdemand_off") else rawdemand_off = 0
if isnumeric(request("rawprev_int")) and trim(request("rawprev_int"))<>"" then rawprev_int = request("rawprev_int") else rawprev_int = 0
if isdate(request("datepeak_int")) and trim(request("datepeak_int"))<>"" then datepeak_int = request("datepeak_int") else datepeak_int = date()
if isnumeric(request("rawdemand_int")) and trim(request("rawdemand_int"))<>"" then rawdemand_int = request("rawdemand_int") else rawdemand_int = 0

dim cnn1, rst1, strsqlc, strsqlp
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,bldg,"billing") 

if trim(action)="Save" then
	if trim(rawcurrent)="" then rawcurrent=0
		strsqlc = "INSERT INTO consumption (meterid, billyear, billperiod, rawonpeak, rawoffpeak, rawintpeak, rawprevious, rawcurrent, rawused,rawpreviousoff, rawcurrentoff, rawusedoff,rawpreviousint, rawcurrentint, rawusedint, estimated, usernote) values ('"&meterid&"', '"&wyear&"', '"&wperiod&"', '"&rawonpeak&"', '"&rawoffpeak&"', '"&rawintpeak&"', '"&rawprevious&"', '"&rawcurrent&"', '"&rawused&"', '"&rawpreviousoff&"', '"&rawcurrentoff&"', '"&rawusedoff&"', '"&rawpreviousint&"', '"&rawcurrentint&"', '"&rawusedint&"', '"&estimatedc&"', '"&usernotec&"')"
		strsqlp = "INSERT INTO peakdemand (meterid, billyear, billperiod, datepeak, estimated, usernote, rawdemand, rawprev, datepeak_off, datepeak_int, rawdemand_off, rawdemand_int, rawprev_off, rawprev_int) values ("&meterid&", "&wyear&", "&wperiod&", '"&datepeak&"', '"&estimatedp&"', '"&usernotep&"', '"&rawdemand&"', '"&rawprev&"', '"&datepeak_off&"', '"&datepeak_int&"', '"&rawdemand_off&"', '"&rawdemand_int&"', '"&rawprev_off&"', '"&rawprev_int&"')"
	elseif trim(action)="Delete" then
		strsqlc = "DELETE FROM consumption WHERE meterid="&meterid&" and billyear="&wyear&" and billperiod="&wperiod
		strsqlp = "DELETE FROM peakdemand WHERE meterid="&meterid&" and billyear="&wyear&" and billperiod="&wperiod
		bperiod = "0"
	elseif trim(action)="Update" then
		strsqlc = "UPDATE consumption set rawonpeak="&rawonpeak&", rawoffpeak="&rawoffpeak&", rawintpeak="&rawintpeak&", rawprevious="&rawprevious&", rawcurrent="&rawcurrent&", rawused="&rawused&", rawpreviousoff="&rawpreviousoff&", rawcurrentoff="&rawcurrentoff&", rawusedoff="&rawusedoff&", rawpreviousint="&rawpreviousint&", rawcurrentint="&rawcurrentint&", rawusedint="&rawusedint&", estimated='"&estimatedc&"', usernote='"&usernotec&"' WHERE meterid="&meterid&" and billyear="&wyear&" and billperiod="&wperiod
		strsqlp = "UPDATE peakdemand set datepeak='"&datepeak&"', estimated='"&estimatedp&"', usernote='"&usernotep&"', rawdemand="&rawdemand&", rawprev='"&rawprev&"', datepeak_off='"&datepeak_off&"', datepeak_int='"&datepeak_int&"', rawdemand_off='"&rawdemand_off&"', rawdemand_int='"&rawdemand_int&"', rawprev_off='"&rawprev_off&"', rawprev_int='"&rawprev_int&"' WHERE meterid="&meterid&" and billyear="&wyear&" and billperiod="&wperiod
	elseif trim(action)="Update Date" then
		strsqlc = "UPDATE meters SET datelastread ='"&request("datelastread")&"' WHERE meterid="&request("meterid")
	elseif trim(action)="Update All Dates" then
		strsqlc = "UPDATE meters SET datelastread ='"&request("datelastread")&"' WHERE bldgnum='"&bldg&"'"
	end if
response.Write strsqlc&"<br>"
response.Write strsqlp&"<br>"
'response.End

cnn1.Execute strsqlc
'response.end()
if not( trim(action)="Update Date" OR trim(action)="Update All Dates" ) then cnn1.Execute strsqlp
'if trim(action)="Save" then 'need to find the bldg for the building just added
'	rst1.Open "SELECT max(meterid) as id FROM meters", cnn1
'	if not rst1.eof then meterid = rst1("id")
'end if
'cnn1.Execute strsqlDS
'TK - 05/16/2006
	on error resume next
	set rst1 = nothing
	If cnn1.State = 1 then 
		cnn1.Close 
	End If
	set cnn1 = nothing
'#TK 
Response.Redirect "entry_select.asp?pid="&pid&"&building="&bldg&"&byear="&byear&"&bperiod="&bperiod&"&meterid="&meterid&"&scroll="&scroll
%>
