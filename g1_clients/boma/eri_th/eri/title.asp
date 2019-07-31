<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Genergy ERI Management</title>

<!-- #include file="./adovbs.inc" -->
<link rel="stylesheet" type="text/css" href="../holiday/holiday.css">

<%
bldg1 = Request("qcatnr")
userid= request("userid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getconnect(0,0,"Engineering") 

Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT Buildings.BldgNum, Buildings.BldgName, Buildings.Strt, Count(Tenant_info.Tenant_no) " &_
"AS CountOfTenant_no, Sum(Tenant_info.sqft) AS SumOfsqft, Sum(Tenant_info.ccm) AS SumOfccm, Sum(Tenant_info.ccy) AS SumOfccy " &_
"FROM Buildings INNER JOIN Tenant_info ON Buildings.BldgNum = Tenant_info.Bldg_no " &_
"WHERE (((Tenant_info.Lease_Exp_Date)> { fn NOW() })) " &_
"GROUP BY Buildings.BldgNum, Buildings.BldgName, Buildings.Strt " &_
"HAVING (((Buildings.BldgNum)='" & bldg1 & "'));"
 
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

If rst1.EOF Then
bldgName = "No Building Available"

Else

bldg = rst1(0)
BldgName = rst1(1)
bldgStreet = rst1(2)
tenant= rst1(3)
Sumsqft=rst1(4)
sumCCM=rst1(5)
sumCCY=rst1(6)
avg_sqft=sumCCY/sumsqft
 
end if

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnnl = Nothing
%>
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="none, default">
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


<body bgcolor="#0099FF">
<div align="center">
 <center>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" bordercolor="#333333">
      <tr bgcolor="#0099FF"> 
        <td width="11%"><b><font size="1" face="Arial, Helvetica, sans-serif">123 Madison Ave
          (1065)</font></b></td>
        <td width="26%"><b><font size="1" face="Arial, Helvetica, sans-serif">Acme Property</font></b></td>
        <td width="63%">
          <div align="right"><b><img src="lock.gif" width="16" height="17"></b></div>
        </td>
      </tr>
      <tr>
        <td width="26%" height="2" align="left" valign="bottom" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF">Total 
          no. of ERI Tenants:</font></td>
        <td width="63%" height="2" valign="bottom" align="right" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=tenant%></font></td>
        <td width="63%" height="2" valign="bottom" align="right" bgcolor="#999999">&nbsp;</td>
      </tr>
      <tr>
        <td width="26%" height="2" align="left" valign="bottom" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF">Total 
          of ERI sqft:</font></td>
        <td width="63%" height="2" valign="bottom" align="right" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=formatnumber(sumSqft,0)%></font></td>
        <td width="63%" height="2" valign="bottom" align="right" bgcolor="#999999">&nbsp;</td>
      </tr>
      <tr>
        <td width="30%" height="2" align="left" valign="bottom" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF">Total 
          of ERI Monthly Charge:</font></td>
        <td width="10%" height="2" valign="bottom" align="right" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=FormatCurrency(sumCCM,0)%></font></td>
        <td width="10%" height="2" valign="bottom" align="right" bgcolor="#999999">&nbsp;</td>
      </tr>
      <tr>
        <td width="30%" height="20%" align="left" valign="bottom" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF">Total 
          of ERI Yearly Charge:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=FormatCurrency(sumCCy,0)%></font></td>
        <td width="10%" height="20%" valign="bottom" align="right" bgcolor="#999999">&nbsp;</td>
      </tr>
      <tr>
        <td width="30%" height="20%" align="left" valign="bottom" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF">AVG 
          of ERI $/sqft:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right" bgcolor="#999999"><font size="1" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=FormatCurrency(avg_sqft,2)%></font></td>
        <td width="10%" height="20%" valign="bottom" align="right" bgcolor="#999999">&nbsp;</td>
      </tr>
    </table>
  </center>
</div>


</body>

</html>
