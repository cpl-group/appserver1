<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
tenantnum=Request.Querystring("tnum")
passwd=Request.Querystring("pwd")

if not isempty(tenantnum) then

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set cnn2 = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
cnn2.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

	strsql = "select tenantnum from tenants where tenantnum='" & tenantnum & "'" 
	rst.Open strsql, cnn2, 0, 1, 1

if rst.eof then 
	msg = "The supplied tenant number was not found on our system. Please verify and try again."
	button="<input type='submit' name='Submit' value='Return to Register' onclick=javascript:document.location='register.htm'><input type='submit' name='Submit' value='Close Window' onclick='javascript:window.close()'>"
	Set cnn2 = nothing
	rst.close
else
	Set cnn2 = nothing
	rst.close
	
	strsql = "select tenantnum from tenant_access where tenantnum='" & tenantnum & "'"
	
	rst.Open strsql, cnn1, 0, 1, 1
	
	if not rst.eof then 
	
		msg="The supplied Tenant Number,  " & tenantnum & ", has already been registered. Please contact accountmaster@genergy.com"
		button="<input type='submit' name='Submit' value='Return to Register' onclick=javascript:document.location='register.htm'><input type='submit' name='Submit' value='Close Window' onclick='javascript:window.close()'>"
	
	else
	
	strsql = "insert tenant_access (tenantnum,password)values ('" & tenantnum & "','" & passwd & "')"
	
	cnn1.execute strsql
	
		msg="Thank you for registering. Please return to the login screen to login."
		button="<input type='submit' name='Submit' value='Close Window' onclick='javascript:window.close()'>"
		
	end if
	
	rst.close
	set cnn1=nothing
	end if
	end if
%>

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF" align="center">
  <tr>
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>Tenant 
        Registration</b></font></div>
    </td>
  </tr>
</table>
<p align="center"><%=msg%></p>
<p align="center"><%=button%></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF" align="center">
  <tr> 
    <td> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>

