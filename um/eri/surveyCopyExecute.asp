<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="./adovbs.inc" -->
<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">
<%
if isempty(Session("name")) then
	%><script>top.location="../index.asp"</script><%
	'	Response.Redirect "http://www.genergyonline.com"
end if	

dim bldg_no, tenant_copy_to, tenant_copy_from, cnnERI, cmd

bldg_no = Request("bldg")
tenant_copy_to = Request("tenant_no")
tenant_copy_from = Request("tenant_copy_from")

Set cnnERI = Server.CreateObject("ADODB.Connection")
cnnERI.open getConnect(0,0,"Engineering")

set cmd = server.createobject("ADODB.Command")
cmd.CommandText = "sp_copysurvey"
cmd.CommandType = adCmdStoredProc
cmd.ActiveConnection = cnnERI

dim prm
Set prm = cmd.CreateParameter("t_old", adVarChar, adParamInput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("t_new", adVarChar, adParamInput,20)
cmd.Parameters.Append prm

cmd.Parameters("t_old") = tenant_copy_from
cmd.Parameters("t_new") = tenant_copy_to

'response.write ("sp_copysurvey '" & tenant_copy_from & "' '" & tenant_copy_to & "'")

cmd.execute

'sp_copysurvey @t_old varchar(20),@t_new varchar(20) ... @t_old is the tenant # you are copying from, @t_new is the tenant # you are copying to. The new tenant # should be created by the user before copying a survey to it. 
%>
<script>
	alert("Survey Copied.")
	window.close
</script>	