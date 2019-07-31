<%option explicit

dim meterid, startdate, bldg, interval, utility, billingid,pid
startdate=Request("startdate")
meterid=Request("meterid")
billingid=Request("billingid")
bldg=Request("bldg")
interval=Request("interval")
utility = Request("utility")
pid = Request("pid")

'if instr(luid,"_") then
'    tenantname = right(luid, len(luid)-instr(luid,"_"))
'    luid = left(luid, instr(luid,"_")-1)
'end if

'change default value for interval to 0, so when return from opt_tenantPF.asp page the graph
'will show the same graph for demo purpose - kc(7/14/2008)
if isempty(interval) then 
	interval=0
end if

if not(isdate(startdate)) then startdate=date()
%>
<html>
<head>
<title>LMP Chart</title>
<script>
function propagateVarsUP()
{ var myparent = parent.document.location+""
	if(myparent.indexOf('lmpload.asp?')==-1)
	{	parent.document.forms[0].billingid.value="<%=billingid%>"
    parent.document.forms[0].bldg.value="<%=bldg%>"
    parent.document.forms[0].meterid.value="<%=meterid%>"
    parent.document.forms[0].startdate.value="<%=startdate%>"
    parent.document.forms[0].utility.value="<%=utility%>"
		parent.closeLoadBox('loadFrame1');
		parent.shownav('lmpnav')
	}
}
</script>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" onload="propagateVarsUP();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr>
    <td>
      <div align="center"><img src="<%="makechartlmp2.asp?meterid="& meterid & "&startdate=" & startdate & "&bldg="&bldg&"&interval="&interval&"&billingid="&billingid&"&utility="&utility&"&pid="&pid%>"></div>
    </td>
  </tr>
</table>
</body>
</html>
