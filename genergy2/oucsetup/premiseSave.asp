<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

dim action, pid, premise, pointName, meterid, ptype, keyName, oldkey, oldDataSourceName, oldmeterid
action = request("action")
pid = request("pid")
premise = request("premise")
pointName = request("pointName")
meterid = request("meterid")
ptype = request("ptype")
keyname = request("keyname")
oldkey = request("oldkey")
oldDataSourceName = request("oldDataSourceName")
oldmeterid = request("oldmeterid")

if trim(action)="Save" then
	strsql = "INSERT INTO oucDataKey (premise, pointName, meterid, type, keyName) values ('"&premise&"', '"&pointName&"', '"&meterid&"', '"&ptype&"', '"&keyName&"')"
elseif trim(action)="Delete" then
	strsql = "DELETE FROM oucDatakey WHERE id="&pid
else
	strsql = "UPDATE oucDataKey set premise='"&premise&"', pointName='"&pointName&"', meterid='"&meterid&"', type='"&ptype&"', keyName='"&keyName&"' WHERE id="&pid
end if
'response.Write strsql
'response.End
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bldg for the building just added
	rst1.Open "SELECT max(id) as id FROM OUCDatakey", cnn1
	if not rst1.eof then pid = rst1("id")
	rst1.close
end if

'####### check for datasource table
dim dataSourceName
strsql = "SELECT datasource FROM meters WHERE meterid="&meterid
rst1.open strsql, cnn1
if not rst1.eof then
	dataSourceName = rst1("datasource")
	rst1.close
	strsql = "SELECT * FROM datasource WHERE datasource='"&dataSourceName&"' and meterid="&meterid
	rst1.open strsql, cnn1
	dim fieldindex, fieldname
	if not rst1.eof then
		for fieldindex = 1 to 7
			if trim(rst1("fieldname"&fieldindex))=oldkey then 
				fieldname = "fieldname"&fieldindex
				fieldindex = 7
			end if
		next
		if fieldname="" then
			for fieldindex = 1 to 7
				if isnull(rst1("fieldname"&fieldindex)) or trim(rst1("fieldname"&fieldindex))="" then
					fieldname = "fieldname"&fieldindex
					fieldindex = 7
				end if
				response.write fieldindex & " : " & isnull(rst1("fieldname"&fieldindex)) & " : " & "fieldname"&fieldindex & "|" & trim(rst1("fieldname"&fieldindex)) &"|<br>"
			next
		end if
		rst1.close
		
		'checking for old name to clear if the meter has been switched
		dim oldfieldname
		if meterid<>oldmeterid then
			strsql = "SELECT * FROM datasource WHERE meterid='"&oldmeterid&"'"
			rst1.open strsql
			if not rst1.eof then
				for fieldindex = 1 to 7
					if trim(rst1("fieldname"&fieldindex))=oldkey then 
						oldfieldname = "fieldname"&fieldindex
						fieldindex = 7
					end if
				next
			end if
			rst1.close
		end if
		
		if oldfieldname<>"" then
			strsql = "UPDATE datasource SET "&oldfieldname&"='' WHERE meterid='"&oldmeterid&"'"
			cnn1.execute strsql
		end if
		if fieldname<>"" then
			strsql = "UPDATE datasource SET "&fieldname&"='"&keyname&"' WHERE meterid='"&meterid&"'"
			cnn1.execute strsql
		end if
	else
'		strsql = "INSERT into datasource (datasource, meterid, fieldname1) values ('"&dataSourceName&"', '"&meterid&"', '"&keyname&"')"
'		cnn1.execute strsql
	end if
end if
'#######



if trim(action)="Delete" then
	Response.Redirect "premiseSetup.asp"
else
	Response.Redirect "premiseEdit.asp?pid="&pid
end if
%>