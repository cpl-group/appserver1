<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, building, byear, ypid, lid, bperiod, utilityid, utilitydisplay, bldgname
pid = request("pid")
building = request("building")
if instr(request("bperiod"),"/")>0 then
	byear = split(request("bperiod"),"/")(1)
	bperiod = split(request("bperiod"),"/")(0)
end if
lid = request("lid")
utilityid = request("utilityid")
bldgname = ""

if utilityid = "" then utilityid = 0
if byear = "" then byear = 0
if bperiod = "" then bperiod = 0
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(building)

dim billlink
strsql = "SELECT location FROM portfolio p, billtemplates bt WHERE bt.id=p.templateid AND p.id="&pid
'response.write strsql
'response.end
rst1.open strsql, getConnect(pid,building,"Billing")
if not rst1.eof then billlink = rst1("location")
rst1.close

strsql = "SELECT b.id, bd.BldgName, b.ypid, b.TenantNum as TenantNumber, u.utilitydisplay as utilitytype, u.utilityid, lup.leaseutilityid, b.billingname, lup.billingid, b.totalamt, reject_date FROM tblbillbyperiod b INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=b.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid INNER JOIN dbo.tblutility u ON u.utilityid=lup.utility INNER JOIN buildings bd ON b.bldgnum=bd.bldgnum WHERE b.reject = 1 and b.bldgnum='"&building&"' and b.billperiod="&bperiod&" and b.billyear="&byear&" and lup.utility="&utilityid&" ORDER BY tenantname, reject_date"
rst1.open strsql, cnn1
if not rst1.eof then bldgname = rst1("bldgname")
%>
<html>
<head>
	<title>Deleted Bills</title>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
<script language="JavaScript1.2">
function viewBill(billid, rstlid, ypid, detailed, utility, buildpdf){
	buildpdf = (buildpdf!=false?true:false);
  var url = 'http://pdfmaker.genergyonline.com<%=billlink%>genergy2=true&devIP=<%=request.ServerVariables("SERVER_NAME")%>&building=<%=building%>&lid=' + rstlid + '&byear=<%=byear%>&bperiod=<%=bperiod%>&y='+ypid+'&ypid='+ypid+'&l=' + rstlid+'&detailed='+detailed+'&utilityid='+utility+'&buildpdf='+buildpdf+'&billid='+billid;
  //alert(url);
  billpdf = window.open(url,'','width=600,height=500,resizable=yes');
}
</script>
<body bgcolor="white" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="servicecode.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Deleted Bills for <%=bldgname%> <%=byear%>, <%=bperiod%></span></td>
</tr>
</table>
<table cellpadding="3" cellspacing="1" align="center" bgcolor="#cccccc">
<tr style="background-color:#dddddd; font-weight: bold;"><td>&nbsp;</td><td>Tenant&nbsp;Name</td><td>Tenant&nbsp;Number</td><td>Utility</td><td>Bill&nbsp;Amt</td><td>Reject&nbsp;Date</td></tr>
<%do until rst1.eof%>
	<tr style="background-color:white">
		<td><a href="javascript:viewBill(<%=rst1("id")%>,'<%=rst1("leaseutilityid")%>', '<%=rst1("ypid")%>','true', <%=rst1("utilityid")%>, true);"><img src="images/pdf_bill.gif" width="21" height="22" border="0"></a><br><!--</td>-->
		<a href="javascript:viewBill(<%=rst1("id")%>,'<%=rst1("leaseutilityid")%>', '<%=rst1("ypid")%>','true', <%=rst1("utilityid")%>, false);">HTML</a></td>
		<td><%=rst1("billingname")%></td>
		<td><%=rst1("TenantNumber")%></td>
		<td><%=rst1("Utilitytype")%></td>
		<td align="right"><%if isnumeric(trim(rst1("totalamt"))) then %><%=formatcurrency(rst1("totalamt"))%><%else%><%=rst1("totalamt")%><%end if%></td>
		<td><%=rst1("reject_date")%></td>
	</tr>
<%rst1.movenext
loop
rst1.close%>
</table>

</form>
</body>
</html>