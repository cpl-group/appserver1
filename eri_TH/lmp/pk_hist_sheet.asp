 <!-- #include file="./adovbs.inc" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script>
function loadlmp(m,d,b) {
	var temp = "lmpload.asp?m="+m+"&d="+d+"&s=15&e=2400&i=100&b="+b
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()
	parent.document.forms[0].pd.value = pd
	parent.document.forms[0].nd.value = nd
	

	parent.document.frames.lmp.location = temp
}
function loadtable(meter){
	var b = document.forms.leaseid.bldg.value
	var luid = document.forms.leaseid.luid.value
	var temp = "pk_hist_sheet.asp?m="+meter+"&luid="+luid+"&b="+b
	document.location = temp
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

<%
m = Request.QueryString("m")
b = Request.QueryString("b")
luid= Request.Querystring("luid")

%>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
      <td bgcolor="#000000" height="25"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="360" height="24">
          <param name=movie value="peaks.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#000000">
          <param name="SCALE" value="exactfit">
          <embed src="peaks.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" scale="exactfit" width="360" height="24" bgcolor="#000000">
          </embed> 
        </object></td>
  </tr>
</table></div>
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td width="4%">&nbsp;</td>
      <td width="92%">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr>
      <td width="4%" height="26">&nbsp;</td>
      <td width="92%" height="26"> 
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#000000">
          <%if luid <> "" then 
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		
		strsql = "select meterid, meternum from meters where leaseutilityid = " & luid &" order by meternum"
			
		rst1.Open strsql, cnn1, adOpenStatic
		if not rst1.EOF then 
%>
          <form name="leaseid" method="post" action="">
            <input type="hidden" name="luid" value="<%=luid%>">
            <input type="hidden" name="bldg" value="<%=b%>">
            <tr> 
              <td valign="middle"> 
                <select name="select" onchange="loadtable(this.value)">
                  <option>Select Meter</option>
                  <%while not rst1.eof 
				if cstr(rst1("meterid")) = m then %>
                  <option value="<%=rst1("meterid")%>" selected><%=rst1("meternum") %></option>
                  <% 
				else
			%>
                  <option value="<%=rst1("meterid")%>"><%=rst1("meternum") %></option>
                  <%  end if
				rst1.movenext
				wend 
			%>
                </select>
              </td>
            </tr>
          </form>
          <%
end if
end if
%>
          <tr> 
            <% if m <> "" then 
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		
		strsql = "select top 12 peakdemand.billyear, peakdemand.billperiod, substring(convert(varchar,timefrom,109),13,8) as timepeak, datepeak, demand, kwhused from peakdemand join consumption on peakdemand.meterid=consumption.meterid and peakdemand.billyear=consumption.billyear and peakdemand.billperiod=consumption.billperiod where peakdemand.meterid = " & m &"  order by peakdemand.billyear desc, peakdemand.billperiod desc"
			
		rst1.Open strsql, cnn1, adOpenStatic
		if not rst1.EOF then 
%>
            <td> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="12%"> 
                    <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Bill 
                      Period</font></div>
                  </td>
                  <td width="10%"> 
                    <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Date</font></div>
                  </td>
                  <td width="11%"> 
                    <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Time</font></div>
                  </td>
                  <td width="46%"> 
                    <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">System 
                      Demand</font></div>
                  </td>
                  <td width="21%"> 
                    <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">System 
                      Consumption</font></div>
                  </td>
                </tr>
                <%While not rst1.EOF %>
                <form name="form1" method="post" action="">
                  <tr valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:loadlmp(meter.value, selecteddate.value,bldg.value)"> 
                    <input type="hidden" name="meter" value="<%=m%>">
                    <input type="hidden" name="selecteddate" value="<%=rst1("datepeak")%>">
                    <input type="hidden" name="bldg" value="<%=b%>">
                    <td width="12%"> 
                      <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("billperiod")%>/<%=rst1("billyear")%></font></div>
                    </td>
                    <td width="10%"> 
                      <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("datepeak")%></font></div>
                    </td>
                    <td width="11%"> 
                      <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("timepeak")%></font></div>
                    </td>
                    <td width="46%"> 
                      <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("demand")%></font></div>
                    </td>
                    <td width="21%"> 
                      <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("kwhused")%></font></div>
                    </td>
                  </tr>
                </form>
                <% 	rst1.movenext
					wend %>
              </table>
            </td>
            <% end if 
			end if%>
          </tr>
        </table>
      </td>
      <td width="4%" height="26">&nbsp;</td>
    </tr>
    <tr>
      <td width="4%">&nbsp;</td>
      <td width="92%"><font face="Arial, Helvetica, sans-serif"><a href="<%="options.asp?m="&m&"&b="&b&"&luid="&luid%>" style="text-decoration:none;" onMouseOver="this.style.color = 'gray'" onMouseOut="this.style.color = 'black'"><font size="2"><b>Back 
        To Options</b></font></a></font></td>
      <td width="4%">&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>
