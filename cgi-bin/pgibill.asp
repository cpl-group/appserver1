<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<title>Meter Details</title>
</head>
<script>
function closeme(){
	window.close()
}
function loadentry(bldgnum, meterid) {
	var   temp="pgimeter.asp?b=" + bldgnum+' &m='+ meterid
	document.location = temp

}
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
  <tr>
    <td width="46%" bgcolor="#336699" height="2"><font face="Arial" size="3" color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Meter 
      Details</font></b></font><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"> 
      - Bill Details</font></td>
    <td width="46%" bgcolor="#336699" height="2"> 
      <div align="right"><a href="Javascript:history.back()" style="text-decoration:none;color:white"><b>Go 
        Back</b></a></div>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
      <%
leaseid = Request("l")
ypid = request("y")
building = request("b")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,building,"billing")

cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select b.id,  tenantnum, tenantname, billyear, billperiod, isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(Adminfee,0) as Adminfee, isnull(Addonfee,0) as Addonfee, isnull(tax,0) as tax, isnull(totalamt,0) as totalamt from dbBilling.dbo.tblbillbyperiod b where b.reject=0 and leaseutilityid=" & leaseid & " and ypid=" & ypid
cmd1.CommandType = 1
dim billid
billid=0
Set rst2 = cmd1.Execute
if not rst2.eof then
billid = rst2("id")
%>
      <div align="center"> 
         <b><font face="Arial" size="3">&nbsp;<%=rst2("tenantnum")%>
        - <%=rst2("tenantname")%></font></b> 
        <table width="100%" height="1" border="1" cellpadding="0" cellspacing="0">
          <tr bgcolor="#6699CC"> 
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
         </div>
      <font face="Arial">
     <br> 
      </font> 
	  <table border="1" width="100%" height="1" cellpadding="0" cellspacing="0" align="center">
        <tr bgcolor="#6699CC"> 
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Meter</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">On 
            Peak KWH</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Off 
            Peak KWH</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">KWH</font></b></td>
          <td width="20%" height="1%" align="center"><b><font face="Arial, Helvetica, sans-serif" size="1">Demand</font></b></td>
        </tr>
        <%
cmd1.CommandText = "select isnull(demand_p,0) as demand, * from dbBilling.dbo.tblmetersbyperiod m WHERE m.bill_id=" & billid & " and ypid=" & ypid
'response.write cmd1.CommandText
'response.end
cmd1.CommandType = 1
Set rst1 = cmd1.Execute

tot_onpeak = 0
tot_offpeak=0
tot_kwhused=0
tot_demand_p=0

while not rst1.eof
%>
<tr valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:loadentry('<%=rst1("bldgnum")%>','<%=rst1("meterid")%>')"> 
          <td width="20%" height="1%" align="center"> 
            <p align="left"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("Meternum")%> </font></b>
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
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("used"),0)%> 
              </font> </b>
          </td>
          <td width="20%" height="1%" align="center"> 
            <p align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("demand"),0)%> 
              </font> </b>
          </td>
        </tr>
        <%
if not isnull(rst1("onpeak")) 		then tot_onpeak = tot_onpeak + cdbl(rst1("onpeak"))
if not isnull(rst1("offpeak")) 	then tot_offpeak= tot_offpeak+ cdbl(rst1("offpeak"))
if not isnull(rst1("used")) 			then tot_kwhused= tot_kwhused + cdbl(rst1("used"))
if not isnull(rst1("demand_P")) 	then tot_demand_p= tot_demand_p + cdbl(rst1("demand_P"))

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
  <tr>
    <td colspan=5 ><b>Click any meter row for above for usage history information</b></td>
  </tr>
</table>
</body>
</html>




