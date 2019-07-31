
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<html>
<head>
<script language="javascript" >
 function openpgiMeter() {
	parent.lmp.document.location="http://appserver1.genergy.com/cgi-bin/pgimeter.asp?b=npt&m=24257"
} 

function openImp() {
    parent.document.location="http://appserver1.genergy.com/g1_clients/boma/eri_th/pgi2/lmp.asp?bldg=919"
 }
 
 function openRemoteDisconnect() {
    parent.lmp.document.location="http://appserver1.genergy.com/g1_clients/boma/eri_th/pgi2/remoteDisconnect.asp"
 }
 function opencurrentmeter()
 {
    parent.lmp.document.location="http://appserver1.genergy.com/g1_clients/boma/eri_th/pgi2/currentmeter.asp?b=npt&m=24257"
 }
 function openMeterEdit()
 {
    parent.lmp.document.location="http://appserver1.genergy.com/g1_clients/boma/eri_th/pgi2/meteredit.asp?meterid=24257"
 }
 function openHome()
 {
    parent.document.location="http://appserver1.genergy.com/eri_th/pgi/index.asp?pgi=marketing/test 2.dwf"
 }
</script>
<title></title>
</head> <style type="text/css">
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

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000" onload="parent.closeLoadBox('loadFrame2');">
&nbsp;<br>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
    <td width="51%" valign="top">
<!--      
 //moved to file c:/kamto/original_code for meterInfo.asp 
-->
	<A HREF="javascript:opencurrentmeter()">Current Billing</A><br />
	<A HREF="javascript:openpgiMeter()">Historical Billing</A><br />
	<A HREF="javascript:openImp()">RealTime Meter Data</A><br />
	<A HREF="javascript:openRemoteDisconnect()">Remote Disconnect</A><br />
	<a href="javascript:openMeterEdit()">Meter Detail</a> <br />
	<a href="javascript:openHome()">Back to Power Grid Identifiocation (PGI) </a>
    </td>
  </tr>
</table>

</body>
</html>