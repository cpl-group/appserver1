<%option explicit%>
<!--#INCLUDE Virtual="/genergy2/secure.inc"-->
<%
if  not allowgroups("Department Supervisors,IT Services")  then 
	setKeyValue "fMessage","Sorry, the module you attempted to access is unavailable to you."
	Response.Redirect "/genergy2/main.asp"
end if	

dim cocode, name, cmd, cnn
if instr(request("company"),"|")>0 then
	cocode = split(request("company"),"|")(0)
	name = split(request("company"),"|")(1)
	set cnn = server.createobject("adodb.connection")
	set cmd = server.createobject("adodb.command")
	cnn.open getConnect(0,0,"intranet")
	set cmd.ActiveConnection = cnn
	cmd.commandText = "exec sp_RUN_PAYROLL '"&cocode&"'"
	'response.write cmd.commandText
	'response.end
	cmd.execute
end if

response.redirect "admintimesheet.asp?name="&server.urlencode(name)
%>