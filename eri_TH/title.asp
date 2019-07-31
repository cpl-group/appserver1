<!-- #include file="secure.inc" --><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Genergy ERI Management</title>

<!-- #include file="./adovbs.inc" -->
<link rel="stylesheet" type="text/css" href="../holiday/holiday.css">
<%
bldg1 = Request("qcatnr")

Set cnn1 = Server.CreateObject("ADODB.Connection")
openStr= "data Source=eri;user Id=web"
cnn1.Open openStr

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


<body>

<div align="center">
 <center>
    <table border="1" width="100%" height="100%">
      <tr> 
        <td width="10%" height="20%"> 
          <p align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Bldg%></font> 
        </td>
        <td width="50%" height "20%"><font size="1" face="Arial, Helvetica, sans-serif"><%=bldgStreet%></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">Total 
          no. of ERI Tenants:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=tenant%></font></td>
      </tr>
      <tr> 
        <td width="10%" height="20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="50%" height "20%"><font size="1" face="Arial, Helvetica, sans-serif"><%=bldgname%></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">Total 
          of ERI sqft:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(sumSqft,0)%></font></td>
      </tr>
      <tr> 
        <td width="10%" height="20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="50%" height "20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">Total 
          of ERI Monthly Charge:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(sumCCM,0)%></font></td>
      </tr>
      <tr> 
        <td width="10%" height="20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="50%" height "20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">Total 
          of ERI Yearly Charge:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(sumCCy,0)%></font></td>
      </tr>
      <tr> 
        <td width="10%" height="20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="50%" height "20%"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">AVG 
          of ERI $/sqft:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=FormatCurrency(avg_sqft,2)%></font></td>
      </tr>
    </table>
 </center>
</div>


</body>

</html>
