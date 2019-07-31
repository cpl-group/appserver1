<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="./adovbs.inc" -->
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<%
if isempty(Session("name")) then
	%><script>top.location="../index.asp"</script><%
	'	Response.Redirect "http://www.genergyonline.com"
end if	

dim bldg_no, tenant_copy_to, tenant_copy_from, cnnERI, rst

bldg_no = Request("bldg")
tenant_copy_to = Request("tenant_no")
tenant_copy_from = Request("tenant_copy_from")

Set cnnERI = Server.CreateObject("ADODB.Connection")
cnnERI.open getConnect(0,0,"Engineering")
Set rst = Server.CreateObject("ADODB.Recordset")
%>
<body bgcolor="#eeeeee" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
	<tr bgcolor="#339999"> 
		<td><span class="standardheader">ERI Manager | Tenant Survey | Surver Copier</span></td>
	</tr>
</table>
<form name="tenant_copy_from_picker" action="surveyCopyExecute.asp" method="post">
<input name="bldg" type="hidden" value="<%=bldg_no%>">
<input name="tenant_no" type="hidden" value="<%=tenant_copy_to%>">
<table cellpadding="5">
	<tr>
		
   <td>Copy survey from:</td>
		<td>
			<select name="tenant_copy_from" onchange="tenant_copy_from_picker.action='tenant_survey_copier.asp';submit()">
				<%
				dim sql, countSelect
				countSelect = 0
				sql = "select TenantName, Tenant_No, Bldg_No from tenant_info where tenant_no in (select tenant_no from tblTenantSurvey) and Bldg_no = '" & bldg_no & "'"
				rst.open sql, cnnERI
				if not rst.eof then
					do while not rst.eof
						dim tenantNo, selected
						tenantNo = rst("Tenant_No")
						if tenant_copy_to <> tenantNo  then
							selected = ""
							countSelect = countSelect + 1
							if tenantNo = request("tenant_copy_from") then selected = "selected" end if
							%><option value="<%=tenantNo%>" <%=selected%>><%=rst("TenantName")%> , <%=tenantNo%></option><%	
						end if
						rst.movenext
					loop
				end if		
				%>
			</select>
			<%
			if countSelect < 1 then
				%><script>alert("There are no valid tenants to copy to.");window.close();</script><%
				response.End()
			end if
			%>
		</td>
	</tr>
	<tr height= "20">
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" align = "center">
			<input name="submitButton" type="submit" value="Execute">
		</td>
	</tr>
</table>
</form>
	