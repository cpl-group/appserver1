<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.open getConnect(0,0,"Engineering")

if request("submit")="Update" then
	strsql = "Update tenant_info Set tenantname='" & request("tname") & "', sqft=" & request("sqft") & ", last_sur_kwh=" & request("last_sur_kwh") & ", last_sur_kw=" & request("last_sur_kw") & ", base_hours=" & request("base_hours") & ", notes='" & request("notes") & "' where tenant_no='" & request("tenant_id") & "'"
else
	strsql = "exec sp_Delete_Tenant_Survey '"&request("tenant_id")&"'"
end if
cnn1.execute strsql
logger(strsql)
set cnn1=nothing

'Response.Write strsql

Response.Redirect "tenant_survey.asp?tenant_no=" & Request("tenant_id") & "&bldg=" & Request("bldg_no")
			
%>
