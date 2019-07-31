<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
 .pagebreak {page-break-before: always}
</style>

</head>

<body bgcolor="#FFFFFF" text="#000000" >
<%
'leaseid = Request("l")
ypid = request("ypid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * ,(datediff(day,datestart,dateend)+1)  as difference from tblbillbyperiod where ypid=" & ypid & " order by tenantname"
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
while not rst2.eof 
%>
<table width="100%" border="0">
  <tr>
    <td height="68"><img src="invoice%20logo.jpg" width="202" height="143"></td>
  </tr>
  <tr>
    <td height="257" valign="top"> <font face="Arial, Helvetica, sans-serif" size="1">&nbsp; 
      </font> 
      <table width="80%" border="0" align="center" bordercolor="#FFFFFF">
        <tr> 
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="15%">&nbsp;</td>
          <td width="25%" bgcolor="#CCCCCC"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Invoice 
              Number</font></div>
          </td>
        </tr>
        <tr> 
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="15%">&nbsp;</td>
          <td width="25%" bgcolor="#CCCCCC"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%="EL." & rst2("billperiod") & Right(rst2("billyear"),2)&  "." & rst2("tenantnum") %> 
              </font></div>
          </td>
        </tr>
      </table>
      <font face="Arial, Helvetica, sans-serif" size="1">&nbsp; </font> 
      <table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">
        <tr bgcolor="#CCCCCC"> 
          <td width="9%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Period</font></div>
          </td>
          <td width="9%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">From</font></div>
          </td>
          <td width="11%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">To</font></div>
          </td>
		  <td width="4%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">No.Days</font></div>
          </td>
          <td width="5%"> 
            <div align="center"></div>
          </td>
          <td width="3%"> 
            <div align="center"></div>
          </td>
		  <td width="9%"> 
            <div align="right"></div>
          </td>
		  <td width="8%"><font face="Arial, Helvetica, sans-serif" size="1">Readings</font> 
          </td>
          <td width="12%"> 
            <div align="right"></div>
          </td>
          <td width="9%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">Consumption</font></div>
          </td>
		  <td width="9%"> 
            <div align="right"></div>
          </td>
          <td width="12%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">Demand</font></div>
          </td>
        </tr>
        <tr bgcolor="#CCCCCC" valign="bottom"> 
          <td width="9%"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></div>
          </td>
          <td width="9%"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("datestart")-1%></font></div>
          </td>
          <td width="11%"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("dateend")%></font></div>
          </td>
		  <td width="4%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst2("difference")%></font></div>
          </td>
          <td width="5%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Meter</font></div>
          </td>
		  <td width="3%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Manual 
              Multiplier </font></div>
          </td>
		  <td width="9%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">Previous</font></div>
          </td>
		   
          <td width="8%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">Current</font></div>
          </td>
          <td width="12%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">On 
              Peak</font></div>
          </td>
          <td width="9%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">Off 
              Peak</font></div>
          </td>
          <td width="9%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">KWHR</font></div>
          </td>
          <td width="12%"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">KW</font></div>
          </td>
        </tr>
        <%
cmd1.CommandText = "select * from tblmetersbyperiod where leaseutilityid=" & rst2("leaseutilityid") & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst1 = cmd1.Execute

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0
metercount = 0
while not rst1.eof
%>
        <tr valign="bottom"> 
          <td width="9%" bordercolor="#FFFFFF"> 
            <div align="center"></div>
          </td>
          <td width="9%" height="1%" align="right" bordercolor="#FFFFFF">&nbsp;</td>
          <td width="11%" height="1%" align="right" bordercolor="#FFFFFF">&nbsp;</td>
		  <td width="4%" height="1%" align="right" bordercolor="#FFFFFF">&nbsp;</td>
          <td width="5%" height="1%" align="right" bordercolor="#FFFFFF"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("Meternum")%></font></div>
          </td>
		  <td width="3%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("manualmultiplier"),0)%></font> 
          </td>
          <td width="9%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <%if  rst1("rawprevious")=0 then %>
            &nbsp; 
            <% else%>
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("rawprevious"),0)%></font> <%end if%>
          </td>
		  <td width="8%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <%if  rst1("rawcurrent")=0 then %>
            &nbsp; 
            <% else%>
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("rawcurrent"),0)%> 
              </font> <%end if %>
          </td>
		  <td width="12%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("onpeak"),0)%> 
              </font> 
          </td>
          <td width="9%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("offpeak"),0)%> 
              </font> 
          </td>
          <td width="9%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("kwhused"),0)%></font> 
          </td>
          <td width="12%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("demand_P"))%> 
              </font> 
          </td>
        </tr>
        <%
tot_onpeak = tot_onpeak + cdbl(rst1("onpeak"))
tot_offpeak= tot_offpeak+ cdbl(rst1("offpeak"))
tot_kwhused= tot_kwhused + cdbl(rst1("kwhused"))
tot_demand_p= tot_demand_p + cdbl(rst1("demand_P"))
metercount = metercount + 1
rst1.movenext
wend
%>
        <tr> 
          <td width="9%" bordercolor="#FFFFFF"> </td>
          <td width="9%" height="1%" align="center" bordercolor="#FFFFFF">&nbsp;</td>
          <td width="11%" height="1%" align="center" bordercolor="#FFFFFF">&nbsp;</td>
		  <td width="4%" height="1%" align="center" bordercolor="#FFFFFF">&nbsp;</td>
          <td width="5%" height="1%" align="center" bgcolor="#CCCCCC"> 
            <p align="right">&nbsp; 
          </td>
		  <td width="3%" height="1%" align="center" bgcolor="#CCCCCC"> </td>
		  <td width="9%" height="1%" align="center" bgcolor="#CCCCCC"> </td>
		  <td width="8%" height="1%" align="center" bgcolor="#CCCCCC"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1">Totals 
              </font> </div>
          </td>
          <td width="12%" height="1%" align="center" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(tot_onpeak,0)%> 
              </font> 
          </td>
          <td width="9%" height="1%" align="center" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(tot_offpeak,0)%> 
              </font> 
          </td>
          <td width="9%" height="1%" align="center" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(tot_kwhused,0)%> 
              </font> 
          </td>
          <td width="12%" height="1%" align="center" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatNumber(tot_demand_P)%> 
              </font> 
          </td>
        </tr>
        <%
set cnn1 = nothing
%>
      </table>
      <table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bordercolor="#FFFFFF">&nbsp;</td>
          <td width="13%" bordercolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bordercolor="#FFFFFF"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Admin 
              Fee</b></font></div>
          </td>
          <td width="13%" bordercolor="#FFFFFF"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatPercent(rst2("Adminfee"))%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bordercolor="#FFFFFF"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Service 
              Fee</b></font></div>
          </td>
          <td width="13%" bordercolor="#FFFFFF"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency((rst2("Addonfee")*metercount),2)%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bgcolor="#CCCCCC"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Sub 
              Total</b></font></div>
          </td>
          <td width="13%" bgcolor="#CCCCCC"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("subtotal"),2)%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bgcolor="#CCCCCC"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Sales 
              Tax</b></font></div>
          </td>
          <td width="13%" bgcolor="#CCCCCC"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("tax"),2)%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bgcolor="#CCCCCC"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Total 
              Amt</b></font></div>
          </td>
          <td width="13%" bgcolor="#CCCCCC"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("totalamt"),2)%></font></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td valign="top"> 
      <hr width="80%" align="center">
      <table width="80%" border="0" align="center">
        <tr>
          <td><font face="Arial, Helvetica, sans-serif" size="2">Tenant Name and 
            Address:</font></td>
          <td><font face="Arial, Helvetica, sans-serif" size="2">Make Check Payable 
            To:</font></td>
        </tr>
        <tr> 
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("tenantname")%> 
            (<%=rst2("tenantnum")%>)</font> </b></td>
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btbldgname")%> 
            </font></b></td>
        </tr>
        <tr> 
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("tstrt")%></font> 
            </b></td>
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btstrt")%></font> 
            </b></td>
        </tr>
        <tr> 
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("tcity")%>, 
            <%=rst2("tstate")%> <%=rst2("tzip")%></font></b></td>
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btcity")%>, 
            <%=rst2("btstate")%> <%=rst2("btzip")%></font></b></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td valign="top"> 
    </td>
  </tr>
  <tr>
    <td valign="top"> 
    </td>
  </tr>
  <tr>
    <td valign="top"> 
    </td>
  </tr>
</table>
<br class="pagebreak">
<%

rst2.movenext
wend
%>

</body>
</html>