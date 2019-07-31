<%option explicit%>
<% 
'response.end%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")

dim bldg, bldgname, address, city, state, zip, sqft, action, pid, region, facilityType, btbldgname, btstrt, btcity, btstate, btzip, Submeters15, Submeters1, offline, ContactName, ContactPhone
bldgname = secureRequest("bldgname")
address = secureRequest("address")
city = secureRequest("city")
state = secureRequest("state")
zip = secureRequest("zip")
sqft = secureRequest("sqft")
action = secureRequest("action")
pid = secureRequest("pid")
bldg = secureRequest("bldg")
region = secureRequest("region")
facilityType = secureRequest("facilityType")
btbldgname = secureRequest("btbldgname")
btstrt = secureRequest("btstrt")
btcity = secureRequest("btcity")
btstate = secureRequest("btstate")
btzip = secureRequest("btzip")
Submeters15 = secureRequest("Submeters15")
Submeters1 = secureRequest("Submeters1")
offline = secureRequest("offline")
ContactName = secureRequest("bldgcontactname")
ContactPhone = secureRequest("bldgcontactphone")

'response.write offline
'response.end
if not isnumeric(sqft) then sqft = 0
if trim(offline)="on" then offline = 1 else offline = 0

setBuildingOffline bldg, offline
if trim(action)="Save" then
  if not isnumeric(Submeters15) then Submeters15 = 0
  if not isnumeric(Submeters1) then Submeters1 = 0
	if trim(bldg)<>"" then
    cnn1.CursorLocation = adUseClient
    cmd.activeConnection = getConnect(pid,bldg,"billing")
    cmd.CommandText =  "sp_NewBldg_Pid"
    cmd.CommandType = adCmdStoredProc
    Set prm = cmd.CreateParameter("name", adVarChar, adParamInput, 15)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pid", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("sub", adBoolean, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("15min", adSmallInt, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("min", adSmallInt, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bldgname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("strt", adVarChar, adParamInput, 100)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("city", adVarChar, adParamInput, 35)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("state", adChar, adParamInput, 2)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("zip", adVarChar, adParamInput, 9)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("sqft", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("region", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("facilitytype", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("btbldgname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("btSTRT", adVarChar, adParamInput, 100)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("btcity", adVarChar, adParamInput, 35)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("btstate", adChar, adParamInput, 2)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("btzip", adVarChar, adParamInput, 11)
    cmd.Parameters.Append prm
   ' Set prm = cmd.CreateParameter("ContactName", adVarChar, adParamInput, 30)
   ' cmd.Parameters.Append prm
  '  Set prm = cmd.CreateParameter("ContactPhone", adVarChar, adParamInput, 20)
    'cmd.Parameters.Append prm
    
	Set prm = cmd.CreateParameter("OIP", adVarChar, adParamInput, 25)
    cmd.Parameters.Append prm

    cmd.Parameters("name") = bldg
    cmd.Parameters("pid") = pid
    cmd.Parameters("sub") = 1
    cmd.Parameters("15min") = Submeters15
    cmd.Parameters("min") = Submeters1
    cmd.Parameters("bldgname") = bldgname
    cmd.Parameters("STRT") = address
    cmd.Parameters("city") = city
    cmd.Parameters("state") = state
    cmd.Parameters("zip") = zip
    cmd.Parameters("sqft") = sqft
    cmd.Parameters("region") = region
    cmd.Parameters("facilitytype") = facilitytype
    cmd.Parameters("btbldgname") = btbldgname
    cmd.Parameters("btSTRT") = btSTRT
    cmd.Parameters("btcity") = btcity
    cmd.Parameters("btstate") = btstate
	cmd.Parameters("OIP")  = getPidIP(PID)
	'cmd.Parameters("ContactName")  = ContactName
	'cmd.Parameters("ContactPhone")  = ContactPhone
	
	
    cmd.Parameters("btzip") = btzip
    
 'response.write "exec sp_NewBldg_Pid '"&bldg&"', '"&pid&"', 1, '"&Submeters15&"', '"&Submeters1&"', '"&bldgname&"', '"&address&"', '"&city&"', '"&state&"', '"&zip&"', '"&sqft&"', '"&region&"', '"&facilitytype&"', '"&btbldgname&"', '"&btSTRT&"', '"&btcity&"', '"&btstate&"', '"&btzip&"', 0" 
 'response.end
    cmd.execute()
	'response.Write(trim(cmd.Parameters("OIP")))
   'response.end

    setBuilding trim(bldg), Application("coreIP"),trim(pid),address,"1433"
  else
    response.write "<span style=""font-family:Arial,Helvetica,sans-serif; font-size:8pt;"">Please enter a building number.<br><input type=button value=""Back"" onclick=""history.back();"" style=""font-family:Arial,Helvetica,sans-serif; font-size:8pt;cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;""></span>"
    response.end 
  end if
'  response.write  strsql
elseif trim(action)="Delete" then
	strsql = "DELETE FROM buildings WHERE bldgnum='"&bldg&"'"
else
	strsql = "UPDATE buildings set bldgname='"&bldgname&"', strt='"&address&"', city='"&city&"', state='"&state&"', zip='"&zip&"', sqft='"&sqft&"', region='"&region&"', facilityType='"&facilityType&"', btbldgname='"&btbldgname&"', btstrt='"&btstrt&"', btcity='"&btcity&"', btstate='"&btstate&"', btzip='"&btzip&"', offline="&offline&", ContactName='"&ContactName&"', ContactPhone='"&ContactPhone&"' WHERE bldgnum='"&bldg&"'"
end if
'response.Write strsql
'response.End
'on error resume next

cnn1.open getLocalConnect(bldg)
if trim(strsql)<>"" then
  'Logging Update
  logger(strsql)
  'end Log
  cnn1.Execute strsql
end if
if err.number = -2147217900 then 
'  response.write "("&err.description&")"
'  response.end
	response.write "<span style=""font-family:Arial,Helvetica,sans-serif; font-size:8pt;"">An error has occurred. It is possible that you have tried to enter a building number that already exists in the database, or you have not entered one at all. Please go back and try entering a unique number for the building.<br><input type=button value=""Back"" onclick=""history.back();"" style=""font-family:Arial,Helvetica,sans-serif; font-size:8pt;cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"">"
	response.end
end if

if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT bldgnum FROM buildings WHERE id in (SELECT max(id) FROM buildings)", cnn1
	if not rst1.eof then bldg = rst1("bldgnum")
end if
if trim(action)="Delete" then
	Response.redirect "portfolioedit.asp?pid="&pid
else
	Response.redirect "buildingedit.asp?pid="&pid&"&bldg="&bldg
end if
%>