<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

'*******************************************************
'12/28/2007 N.Ambo changed else to elseif line 105
'*******************************************************
 
dim portfolio, portfolioname, action, pid, buildings15, buildings1, billtemplate, paymentterm, offline
dim cnn1, rst1, strsql, cmd, prm,strsql1
pid = secureRequest("pid")

set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,0,"dbCore")


action = secureRequest("action")

portfolio = secureRequest("portfolio")
portfolioname = secureRequest("portfolioname")
buildings15 = secureRequest("buildings15")
buildings1 = secureRequest("buildings1")
billtemplate = secureRequest("billtemplate")
paymentterm = secureRequest("paymenttext")
offline = secureRequest("offline")
dim fs,f, fdir
fdir = "D:\WebSites\isabella\genergyonline.com\pdfmaker\"&ucase(portfolio)
set fs=Server.CreateObject("Scripting.FileSystemObject")
if not isnumeric(buildings15) then buildings15 = 0
if not isnumeric(buildings1) then buildings1 = 0
if not isnumeric(billtemplate) then billtemplate = 0

if trim(action)="Save" then
    cnn1.CursorLocation = adUseClient
    cmd.activeConnection = getConnect(pid,0,"billing")
	
    cmd.CommandText = "sp_NewBldg_Pid"
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
    Set prm = cmd.CreateParameter("btzip", adVarChar, adParamInput, 9)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("OIP", adVarChar, adParamInput, 25)
    cmd.Parameters.Append prm
    
    cmd.Parameters("name") = portfolio
    cmd.Parameters("pid") = 0
    cmd.Parameters("sub") = 0
    cmd.Parameters("15min") = buildings15
    cmd.Parameters("min") = buildings1
    cmd.Parameters("bldgname") = portfolioname
    cmd.Parameters("STRT") = 0
    cmd.Parameters("city") = 0
    cmd.Parameters("state") = 0
    cmd.Parameters("zip") = 0
    cmd.Parameters("sqft") = 0
    cmd.Parameters("region") = 0
    cmd.Parameters("facilitytype") = 0
    cmd.Parameters("btbldgname") = 0
    cmd.Parameters("btSTRT") = 0
    cmd.Parameters("btcity") = 0
    cmd.Parameters("btstate") = 0
    cmd.Parameters("btzip") = 0
    cmd.Parameters("OIP") = Application("CoreIP")
    ''response.write "exec sp_NewBldg_Pid '"&portfolio&"', 0, 0, '"&buildings15&"', '"&buildings1&"', '"&portfolioname&"', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0" 
    'response.end
    
    cmd.execute()

  	rst1.Open "SELECT max(id) as id FROM portfolio", cnn1
  	if not rst1.eof then pid = rst1("id")
    setPortfolio pid, Application("CoreIP"),portfolio,""
	set f=fs.CreateFolder(fdir)
	set f=nothing
	set fs=nothing	
elseif trim(action)="Update" then '12/28/2007 N.Ambo changed from 'else' to 'elseif' specifying update option; if action = 'nosave' then no insert or update is done
	strsql = "UPDATE portfolio set name='"&portfolioname&"', templateid='"&billtemplate&"', paymentterm='"&paymentterm&"', offline='" &offline&"' WHERE id="&pid
end if

'response.End
if trim(strsql)<>"" then cnn1.Execute strsql
if trim(strsql1)<>"" then cnn1.Execute strsql1


Response.redirect "portfolioedit.asp?pid="&pid
%>