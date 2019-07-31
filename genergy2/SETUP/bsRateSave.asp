<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if



dim buildingNum, portfolioId, utility, season, amount, dateFrom, dateTo, by, bp, notes, micID,user, rateid, desc
buildingNum = request("buildingNum")
portfolioId = request("portfolioId")

utility = request("utility")
amount = request("amount")
notes = request("notes")
micID = request("micid")
user = request("user")
rateid = request("rateid")
desc = request("description")
'response.write rateid
''response.end

'if season was not supplied, it should be inserted as a zero, in order to overwrite old values if necessary.
season = request("season")
if not ne(season) then
	season = 0
end if

'if dates were not supplied, the table should be updated with nulls.  the sql statement will be changed to not have
' quotes around dateFrom and dateTo - the variables themselves will provide them.
dateFrom = request("dateFrom")
dateTo = request("dateTo")

if ne(dateFrom) AND ne(dateTo) then
	dateFrom = "'" & dateFrom & "'"
	dateTo = "'" & dateTo & "'"
else
	dateFrom = "null"
	dateTo = "null"
end if

'response.write (dateFrom & "-" & dateTo)
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")

cnn.open getLocalConnect(buildingNum)

'if by or bp was not supplied, it both should be inserted as a zero, in order to overwrite old values if necessary.
dim ypid
by = request("by")
bp = request("bp")

if ne(by) and ne(bp) then
	dim ypidSql, ypidRst
	set ypidRst = server.createobject("ADODB.recordset")
	ypidSql = "select ypid from billYrPeriod where billyear = '" & by & "' and billperiod = '" & bp & "' and utility = '" & utility & "' and bldgnum = '" & buildingNum & "'"
	ypidRst.open ypidSql, cnn
	if ypidRst.eof then
		ypid = 0
	else
		ypid = ypidRst("ypid")
	end if
	ypidRst.close
	
else
	by = 0
	bp = 0
	ypid = 0
end if


dim entryExists
strsql = "select bldgNum from misc_inv_credit where id = '" & micid & "'"
rst.open strsql, cnn
if rst.eof then
	entryExists = false
else
	entryExists = true
end if
rst.close

if entryExists then
	strsql = "update misc_inv_credit set utilityid = '" &utility&"',  note = '" &notes & "', amt = '" &amount& "', seasonid = '" & season & "', dateFrom = " & dateFrom & ", dateTo = " & dateTo & ", billYear = '" &by& "', [description] = '" &desc& "', billPeriod = '" &bp& "', ypid = '" & ypid &"', [user] = '" & user &"', rateid = '" & rateid &"' where id = '" & micid & "'"
	
else
	strsql = "insert into misc_inv_credit(bldgNum,pid,utilityid,note,amt,ypid, [user], rateid, [description]"
	if ne(season) then
		strsql = strsql & ",seasonid"
	end if
	if ne(dateFrom) then
		strsql = strsql & ",datefrom"
	end if
	if ne(dateTo) then
		strsql = strsql & ",dateto"
	end if
	if ne(by) then
		strsql = strsql &",billyear"
	end if
	if ne(bp) then
		strsql = strsql & ",billperiod"
	end if
	
	
	strsql = strsql  & ") values ('"&buildingNum&"','"&portfolioId&"','"&utility&"','"&notes&"','"&amount&"','"&ypid&"','"&user&"','"&rateid&"','"&desc&"'"
	if ne(season) then
		strsql = strsql & ",'" & season & "'"
	end if
	if ne(dateFrom) then
		strsql = strsql & ","&datefrom
	end if
	if ne(dateTo) then
		strsql = strsql & ","&dateto
	end if
	if ne(by) then
		strsql = strsql &",'"&by & "'"
	end if
	if ne(bp) then
		strsql = strsql & ",'"&bp & "'"
	end if
	strsql = strsql & ")"
end if

'response.write strsql
'response.end
rst.open strsql, cnn

function ne(someString)
	if isNull(someString) or isempty(someString) or someString = "" then
		ne = false
	else
		ne = true
	end if
end function
%>
<script>
alert("Information saved.");
window.close();
</script>