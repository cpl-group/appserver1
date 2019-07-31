<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
leaseid = Request("l")
ypid = request("y")

b = request("b")
m = request("m")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Meter Details</title>
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

<script>
function closeme(){
	window.close()
}
function viewbill(ypid,lid) {
			var temp= "invoice.asp?y=" + ypid + "&l=" + lid 
			window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
	}

</script>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
  <tr>
    <td width="46%" bgcolor="#000000"><font face="Arial" size="3" color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Bill Details</font></td>
    <td width="46%" bgcolor="#000000" align="right"><b><font face="Arial" size="3"><a href="javascript:document.location='/eri_th/lmp/opt_accthist.asp?b=<%=b%>&m=<%=m%>&luid=<%=leaseid%>'" style="text-decoration:none;color:white" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><font size="2">Return To Billing History</font></a></font></b></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td> 
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
%>
      <div align="left"></div>
      <table border="1" width="100%" height="1">
        <tr bgcolor="#0099FF" style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
          <td width="7%" align="center"><b>Period</b></td>
          <td width="14%" align="center"><b>Energy Charge</b></td>
          <td width="12%" align="center"><b>Demand Charge</b></td>
          <td width="10%" align="center"><b>Admin Fee</b></td>
          <td width="10%" align="center"><b>Service Fee</b></td>
          <td width="10%" align="center"><b>Sales Tax</b></td>
          <td width="10%" align="center"><b>Total Amt</b></td>
        </tr>
        <tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
          <td width="7%" align="center"><b><%=rst2("billyear")%>/<%=rst2("billperiod")%></b></td>
          <td width="14%" align="center"><b><%=FormatCurrency(rst2("energy"),2)%></b></td>
          <td width="12%" align="center"><b><%=FormatCurrency(rst2("demand"),2)%></b></td>
          <td width="10%" align="center"><b><%=Formatpercent(rst2("Adminfee"),2)%></b></td>
          <td width="10%" align="center"><b><%=FormatCurrency(rst2("Addonfee"),2)%></b></td>
          <td width="10%" align="center"><b><%=FormatCurrency(rst2("tax"),2)%></b></td>
          <td width="10%" align="center"><b><%=FormatCurrency(rst2("totalamt"),2)%></b></td>
        </tr>
      </table>
      <br>
      <table border="0" cellspacing="0" cellpadding="0"><tr><td>
	  <table border="1" width="680" cellpadding="0" cellspacing="0">
        <tr bgcolor="#0099FF" style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
          <td width="20%" align="center"><b>Meter</b></td>
          <td width="20%" align="center"><b>On Peak KWH</b></td>
          <td width="20%" align="center"><b>Off Peak KWH</b></td>
          <td width="20%" align="center"><b>KWH</b></td>
          <td width="20%" align="center"><b>Demand</b></td>
        </tr>
      </table>
	  </td></tr>
	  <tr><td><div style="overflow:auto;height:130">
	  <table border="1" width="680" cellpadding="0" cellspacing="0">
        <%
cmd1.CommandText = "select * from tblmetersbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst1 = cmd1.Execute

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
%>
        <tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
          <td width="20%"><b><%=rst1("Meternum")%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(rst1("onpeak"),0)%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(rst1("offpeak"),0)%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(rst1("kwhused"),0)%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(rst1("demand_P"),0)%></b></td>
        </tr>
        <%
tot_onpeak = tot_onpeak + rst1("onpeak")
tot_offpeak= tot_offpeak+ rst1("offpeak")
tot_kwhused= tot_kwhused + rst1("kwhused")
tot_demand_p= tot_demand_p + rst1("demand_P")

rst1.movenext
wend

else
end if
%>
        <tr bgcolor="#CCCCCC" style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
          <td width="20%"><b>Totals</b></td>
          <td width="20%" align="right"><b><%=formatnumber(tot_onpeak,0)%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(tot_offpeak,0)%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(tot_kwhused,0)%></b></td>
          <td width="20%" align="right"><b><%=formatnumber(tot_demand_P,0)%></b></td>
        </tr>
      <%
set cnn1 = nothing
%>
      </table>
	  </div>
	  </td></tr></table>
	  
	  
    </td>
  </tr>
</table>
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><i><font size="1"> 
  </font></i></b></font></p>
  <table bgcolor="#000000" cellpadding="0" cellspacing="0" width="100%"><tr><td><b><font face="Arial" size="3"><a href="javascript:viewbill('<%=ypid%>','<%=leaseid%>')" style="text-decoration:none;color:white" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><font size="2">View Bill For This Period</font></a></font></b></td></tr></table>
</body>
</html>




