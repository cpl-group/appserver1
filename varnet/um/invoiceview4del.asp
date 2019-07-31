<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function deletebill(bldg,ypid,luid){
	if(confirm("Delete This Bill?")){
			var temp = "deletebill.asp?bldg=" + bldg + "&ypid=" + ypid + "&luid=" + luid
			document.location=temp	
	} else {
		var msg = "Delete cancelled."
		alert(msg)
	}


}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
luid = Request("luid")
ypid = request("ypid")
bldgnum = Request("bldg")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * from tblbillbyperiod where leaseutilityid=" & luid & " and ypid=" & ypid

cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
%>
<form name="form1" method="post" action="">
  <div align="center">
    <input type="button" name="Button" value="DELETE BILL" onclick="deletebill('<%=bldgnum%>',<%=ypid%>,<%=luid%>)">
  </div>
</form>
<table width="100%" border="0" height="100%">
  <tr>
    <td height="68"><img src="invoice%20logo.jpg" width="202" height="143"></td>
  </tr>
  <tr>
    <td height="485" valign="top">
      <p>&nbsp;</p>
	  <font face="Arial, Helvetica, sans-serif" size="1"> 
      </font>
      <table width="80%" border="0" align="center" bordercolor="#FFFFFF">
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td width="30%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Invoice 
              Number</font></div>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td width="30%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%="EL." & rst2("billperiod") & Right(rst2("billyear"),2)&  "." & rst2("tenantnum") %> 
              </font></div>
          </td>
        </tr>
      </table>
      <font face="Arial, Helvetica, sans-serif" size="1"> </font> 
      <table width="80%" border="0" align="center" bordercolor="#FFFFFF" cellspacing="0">
        <tr bordercolor="#FFFFFF" bgcolor="#CCCCCC"> 
          <td width="13%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Period</font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">From</font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">To</font></div>
          </td>
          <td width="15%"> 
            <div align="center"></div>
          </td>
          <td width="15%"> 
            <div align="center"></div>
          </td>
          <td width="15%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">CONSUMPTION</font></div>
          </td>
          <td width="12%"> 
            <div align="center"></div>
          </td>
          <td width="30%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">DEMAND</font></div>
          </td>
        </tr>
        <tr bordercolor="#FFFFFF" bgcolor="#CCCCCC"> 
          <td width="13%"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("datestart")-1%></font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("dateend")%></font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">METER</font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">On 
              Peak</font></div>
          </td>
          <td width="15%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">Off 
              Peak</font></div>
          </td>
          <td width="12%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">KWHR</font></div>
          </td>
          <td width="30%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="1">KW</font></div>
          </td>
        </tr>
        <%
cmd1.CommandText = "select * from tblmetersbyperiod where leaseutilityid=" & luid & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst1 = cmd1.Execute

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
%>
        <tr bordercolor="#FFFFFF"> 
          <td width="13%"> 
            <div align="center"></div>
          </td>
          <td width="15%" height="1%" align="right">&nbsp;</td>
          <td width="15%" height="1%" align="right">&nbsp;</td>
          <td width="15%" height="1%" align="right" bordercolor="#FFFFFF"> 
            <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("Meternum")%></font></div>
          </td>
          <td width="15%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("onpeak"),0)%> 
              </font> 
          </td>
          <td width="15%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("offpeak"),0)%> 
              </font> 
          </td>
          <td width="12%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("kwhused"),0)%> 
              </font> 
          </td>
          <td width="30%" height="1%" align="center" bordercolor="#FFFFFF"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(rst1("demand_P"),0)%> 
              </font> 
          </td>
        </tr>
        <%
tot_onpeak = tot_onpeak + rst1("onpeak")
tot_offpeak= tot_offpeak+ rst1("offpeak")
tot_kwhused= tot_kwhused + rst1("kwhused")
tot_demand_p= tot_demand_p + rst1("demand_P")

rst1.movenext
wend

%>
        <tr bordercolor="#FFFFFF"> 
          <td width="13%"> </td>
          <td width="15%" height="1%" align="center">&nbsp;</td>
          <td width="15%" height="1%" align="center">&nbsp;</td>
          <td width="15%" height="1%" align="center" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="center"></div>
            <p align="center"><font face="Arial, Helvetica, sans-serif" size="1">Totals 
              </font> 
          </td>
          <td width="15%" height="1%" align="center" bordercolor="#FFFFFF" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(tot_onpeak,0)%> 
              </font> 
          </td>
          <td width="15%" height="1%" align="center" bordercolor="#FFFFFF" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(tot_offpeak,0)%> 
              </font> 
          </td>
          <td width="12%" height="1%" align="center" bordercolor="#FFFFFF" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatnumber(tot_kwhused,0)%> 
              </font> 
          </td>
          <td width="30%" height="1%" align="center" bordercolor="#FFFFFF" bgcolor="#CCCCCC"> 
            <p align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatNumber(tot_demand_P,0)%> 
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
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatPercent(rst2("Adminfee"),2)%></font></div>
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
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("Addonfee"),2)%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Sub 
              Total</b></font></div>
          </td>
          <td width="13%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency((ccur(rst2("energy"))+ccur(rst2("demand"))+ ccur(rst2("Addonfee"))+((ccur(rst2("energy"))+ccur(rst2("demand"))+ ccur(rst2("Addonfee")))* ccur(rst2("Adminfee")))),2)%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Sales 
              Tax</b></font></div>
          </td>
          <td width="13%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("tax"),2)%></font></div>
          </td>
        </tr>
        <tr> 
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
          <td width="69%">&nbsp;</td>
          <td width="12%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="1"><b>Total 
              Amt</b></font></div>
          </td>
          <td width="13%" bgcolor="#CCCCCC" bordercolor="#FFFFFF"> 
            <div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("totalamt"),2)%></font></div>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
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
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btstrt")%></font> 
            </b></td>
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btstrt")%></font> 
            </b></td>
        </tr>
        <tr> 
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btcity")%>, 
            <%=rst2("tstate")%> <%=rst2("btzip")%></font></b></td>
          <td><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("btcity")%>, 
            <%=rst2("btstate")%> <%=rst2("btzip")%></font></b></td>
        </tr>
      </table>
      <p><font size="2"></font></p>
    </td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">BILL 
        NOT FOUND</font> </b></font></div>
    </td>
  </tr>
</table>
<%end if%>
</body>
</html>
