<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #INCLUDE VIRTUAL="/includes/ChartConst.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim cnn1, rst1, strsql, bldg, cmd, prm
Dim meterid, lmpstart, lmpend,lmpdate, accountname, tenantmeter, ishourly, interval, chartTimeInterval,l, pulsetable, utility, usage, units, groupname, lmptype, lmpcode, billingid, luid, total, part, datasource,pid

bldg=Request.Querystring("bldg")
meterid=Request.QueryString("meterid")
lmpdate=Request.QueryString("startdate")
billingid = Request.QueryString("billingid")
interval=Request.QueryString("interval")
tenantmeter = request.querystring("tenantmeter")
utility = request.querystring("utility")
groupname = request("groupname")
pid = Request.QueryString("pid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.command")
Set rst1 = Server.CreateObject("ADODB.recordset")

if trim(bldg) <> "" then 
	cnn1.open getLocalConnect(bldg) 
else 
	Response.write "<Div align='center'>BUILDING INFORMATION NOT FOUND</DIV>"
	response.end
end if
cnn1.CursorLocation = adUseClient
rst1.open "select utility from dbo.tblLeasesUtilityPrices where leaseutilityid in (select leaseutilityid from meters where meterid ='"&meterid&"')",cnn1
if not rst1.eof then 
	utility= trim(rst1("utility")) 
	response.redirect "/genergy2/eri_th/lmp/lmpload.asp?meterid=" & meterid&"&bldg="&bldg&"&lmp=1&utility="&utility&"&interval=0"
else 
	Response.write "<Div align='center'>METER UTILITY INFORMATION NOT FOUND</DIV>"
	response.end
end if 
rst1.close
%>

