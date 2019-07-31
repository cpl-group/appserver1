<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
leaseid = Request("l")
ypid = request("y")
luid = request("luid")
b = request("b")
m = request("m")
building = request("building")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")

cmd1.ActiveConnection = cnn1
if leaseid<>0 then
	title = "Tenant"
	cmd1.CommandText = "select adminfee*(energy+demand) as afdollar, * from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid &" and posted =1"
else
	title = "Building"
	cmd1.CommandText = "select billyear, billperiod, sum(energy) as energy, sum(demand) as demand, sum(adminfee*(energy+demand)) as afdollar, sum(addonfee) as addonfee, sum(tax) as tax, sum(totalamt) as totalamt from tblbillbyperiod where ypid=" & ypid &" and posted =1 group by billyear, billperiod"
end if
'response.write cmd1.CommandText
'response.end
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
	bperiod = rst2("billperiod")
	byear = rst2("billyear")
end if

%>
<html>

<head>
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
			var temp= "http://pdfmaker.genergyonline.com/pdfMaker/pdfBatchPrint.asp?y=" + ypid + "&l=" + lid 
			window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
	}

</script>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
  <tr>
    <td width="46%" bgcolor="#0099FF"><font face="Arial" size="3" color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Bill Details</font></td>
    <td width="46%" bgcolor="#0099FF" align="right" style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;">
		<%if leaseid<>0 then%>
		<b><font face="Arial" size="3"><a href="javascript:viewbill('<%=ypid%>','<%=leaseid%>')" style="text-decoration:none;color:white" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><font size="2">View Bill For This Period</font></a></font></b>
		<%else%>
		<nobr>
		<a style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;" target="_blank" href="http://pdfmaker.genergyonline.com/pdfmaker/pdfbatchprint.asp?building=<%=server.urlencode(building)%>&bperiod=<%=bperiod%>&byear=<%=byear%>" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><b>Download&nbsp;all&nbsp;bills&nbsp;for&nbsp;this&nbsp;building</b></a>&nbsp;|&nbsp;
		<a style="font-family:arial;font-size:12;text-decoration:none;color:#FFFFFF;" target="_blank" href="http://pdfmaker.genergyonline.com/pdfMaker/pdfBillSummary.asp?building=<%=server.urlencode(building)%>&byear=<%=byear%>&bperiod=<%=bperiod%>" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'white'"><b>Download&nbsp;Bill&nbsp;Summary&nbsp;Report</b></a>
		</nobr>
		<%end if%>
	</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td> 
<%if not rst2.eof then%>
      <div align="left"></div>
      <table border="1" width="100%" height="1">
        <tr bgcolor="#0099FF"> 
          <td width="7%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Period</font></b></td>
          <td width="14%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Energy Charge</font></b></td>
          <td width="12%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Demand Charge</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Admin Fee</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Service Fee</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Sales Tax</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Total Amt</font></b></td>
        </tr>
        <tr> 
          <td width="7%" height="1%" align="center"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></b></td>
			<td width="14%" height="1%" align="center"><p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("energy"),2)%></font></b></td>
          <td width="12%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("demand"),2)%></font></b> 
          </td>
          <td width="10%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatcurrency(rst2("Afdollar"),2)%></font></b> 
          </td>
          <td width="10%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("Addonfee"),2)%></font></b> 
          </td>
          <td width="10%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("tax"),2)%></font></b> 
          </td>
          <td width="10%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("totalamt"),2)%></font></b> 
          </td>
        </tr>
      </table>
      <font face="Arial"> <br>
      </font>
	<%if leaseid<>0 then%>	  
      <table border="1" width="100%" height="1" cellpadding="0" cellspacing="0" align="center">
        <tr bgcolor="#0099FF"> 
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Meter</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">On 
            Peak KWH</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Off 
            Peak KWH</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">KWH</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Demand</font></b></td>
        </tr>
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
        <tr> 
          <td width="20%" height="1%" align="center"> 
            <p align="left"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("Meternum")%> 
              </font></b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("onpeak"),0)%> 
              </font> </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("offpeak"),0)%> 
              </font> </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("kwhused"),0)%> 
              </font> </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("demand_P"),2)%> 
              </font> </b> 
          </td>
        </tr>
        <%
tot_onpeak = tot_onpeak + cDbl(rst1("onpeak"))
tot_offpeak= tot_offpeak+ cDbl(rst1("offpeak"))
tot_kwhused= tot_kwhused + cDbl(rst1("kwhused"))
tot_demand_p= tot_demand_p + cDbl(rst1("demand_P"))

rst1.movenext
wend
end if 'this one is for masking the meter table when building view
end if 'this one is for if has records
if leaseid<>0 then
%>
        <tr bgcolor="#CCCCCC"> 
          <td width="20%" height="1%" align="center"> 
            <div align="center"></div>
            <p align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Totals</font> 
              </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_onpeak,0)%> 
              </font> </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_offpeak,0)%> 
              </font> </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_kwhused,0)%> 
              </font> </b> 
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_demand_P,2)%> 
              </font> </b> 
          </td>
        </tr>
      </table>
      <%
end if
set cnn1 = nothing
%>
    </td>
  </tr>
</table>
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><i><font size="1"> 
  </font></i></b></font></p>
  <table bgcolor="#0099FF" cellpadding="0" cellspacing="0" width="100%"><tr><td>&nbsp;</td></tr></table>
</body>
</html>




