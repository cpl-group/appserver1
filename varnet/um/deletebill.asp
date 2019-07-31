<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("UM") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")

bldgnum=Request.QueryString("bldg")
luid=Request.QueryString("luid")
ypid=Request.QueryString("ypid")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

strsql= "Delete from tblBillByPeriod where LeaseUtilityId = '"& luid &"' and ypid='" & ypid & "' and bldgnum ='" & bldgnum & "'"
cnn1.execute strsql

set cnn1=nothing
%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#6699CC">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>TENANTS 
        BILL HAS BEEN DELETED</b></font></div>
    </td>
  </tr>
</table>