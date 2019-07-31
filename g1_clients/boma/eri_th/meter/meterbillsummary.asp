<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
leaseid = Request("l")
ypid = request("y")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Meter Details</title>
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
</head>
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
    <td width="46%" bgcolor="#6699cc" height="2"><font face="Arial" size="3" color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Meter 
      Details</font></b></font><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"> 
      - Bill Details</font></td>
    <td width="46%" bgcolor="#6699cc" height="2"> 
      <div align="right"><b><font face="Arial" size="3"><a href="javascript:viewbill('<%=ypid%>','<%=leaseid%>')" style="text-decoration:none;" onMouseOver="this.style.color= 'gray'"; onMouseOut="this.style.color = 'Black'"><font size="2">View 
        Bill For This Period</font></a></font></b></div>
    </td>
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
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * from tblbillbyperiod where leaseutilityid=" & leaseid & " and ypid=" & ypid
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
%>
      <div align="left"></div>
      <table border="1" width="100%" height="1">
        <tr bgcolor="#6699cc"> 
          <td width="7%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Period</font></b></td>
          <td width="14%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Energy 
            Charge</font></b></td>
          <td width="12%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Demand 
            Charge</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Admin 
            Fee</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Service 
            Fee</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Sales 
            Tax</font></b></td>
          <td width="10%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Total 
            Amt</font></b></td>
        </tr>
        <tr> 
          <td width="7%" height="1%" align="center"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%></font></b></td>
          <td width="14%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("energy"),2)%></font></b> 
          </td>
          <td width="12%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst2("demand"),2)%></font></b> 
          </td>
          <td width="10%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=Formatpercent(rst2("Adminfee"),2)%></font></b> 
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
      <table border="1" width="100%" height="1" cellpadding="0" cellspacing="0" align="center">
        <tr bgcolor="#6699cc"> 
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
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("demand_P"),0)%> 
              </font> </b> 
          </td>
        </tr>
        <%
tot_onpeak = tot_onpeak + cdbl(rst1("onpeak"))
tot_offpeak= tot_offpeak+ cdbl(rst1("offpeak"))
tot_kwhused= tot_kwhused + cdbl(rst1("kwhused"))
tot_demand_p= tot_demand_p + cdbl(rst1("demand_P"))

rst1.movenext
wend

else
end if
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
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(tot_demand_P,0)%> 
              </font> </b> 
          </td>
        </tr>
      </table>
      <%
set cnn1 = nothing
%>
    </td>
  </tr>
</table>
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><i><font size="1"> 
  </font></i></b></font></p>
</body>
</html>




