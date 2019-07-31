<html>

<head>

<title>Genergy ERI Management</title>

<!-- #include file="./adovbs.inc" -->
<link rel="stylesheet" type="text/css" href="../holiday/holiday.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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
Set cnn1 = Server.CreateObject("ADODB.Connection")
openStr= "data Source=eri;user Id=erimanager;pwd=erimanager"
cnn1.Open openStr

Set rsCat1 = Server.CreateObject("ADODB.Recordset")

sql="SELECT SUM(Tenant_info.CCY) AS Tot_Year, SUM(Tenant_info.ccm) " &_
    "AS Tot_Month, SUM(Tenant_info.sqft) AS Tot_sqft, " &_
    "SUM(Tenant_info.CCY) / SUM(Tenant_info.sqft) " &_ 
    "AS avg_sqft FROM Buildings INNER JOIN Tenant_info ON " &_ 
    "Buildings.BldgNum = Tenant_info.Bldg_no WHERE (Buildings.Owner_id = N'" & Request("portfolioid") & "') AND " &_ 
    "(Tenant_info.Lease_Exp_Date > { fn NOW() })"
    
rsCat1.Open sql, cnn1, adOpenStatic, adLockReadOnly

Portfolio_sqft=rsCat1("tot_sqft")

Set rsCat = Server.CreateObject("ADODB.Recordset")

sql="SELECT Distinct Management FROM Buildings WHERE owner_id = N'" & request("portfolioid") & "'"

rsCat.Open sql, cnn1, adOpenStatic, adLockReadOnly
 %>

<body bgcolor="#FFFFFF">
<table align="center" border="0" width="100%" cellpadding="3" cellspacing="0" height="102">
  <tr> 
    <td align="center" width="503" rowspan="2" height="46"> 
      <h3><font face="Arial, Helvetica, sans-serif"> 
        <%Response.Write rsCat("management")
rscat.close
%>
        </font></h3>
    </td>
    <td align="center" width="84" rowspan="2" height="46"></td>
    <td align="left" width="196" bordercolor="#008000" height="20"><font face="Arial, Helvetica, sans-serif">Yearly 
      ERI Revenue</font></td>
    <td align="right" width="148" height="20"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif"><b><%=FormatCurrency(rsCat1("tot_year"),0)%></b></font>
    </td>
  </tr>
  <tr> 
    <td align="left" width="196" height="20"><font face="Arial, Helvetica, sans-serif">Monthly 
      ERI Revenue</font></td>
    <td align="right" width="148" height="20"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif"><b><%=formatcurrency(rsCat1("tot_month"),0)%></b></font>
    </td>
  </tr>
  <tr> 
    <td align="center" width="503" rowspan="2" height="44"> 
      <h3><font face="Arial, Helvetica, sans-serif">Electric Rent Inclusion Management</font></h3>
    </td>
    <td align="center" width="84" rowspan="2" height="44"></td>
    <td align="left" width="196" height="20"><font face="Arial, Helvetica, sans-serif">ERI 
      sqft</font></td>
    <td align="right" width="148" height="20"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif"><b><%=formatnumber(rsCat1("tot_sqft"),0)%></b></font>
    </td>
  </tr>
  <tr> 
    <td align="left" width="196" height="24"><font face="Arial, Helvetica, sans-serif">Average 
      ERI $/sqft</font></td>
    <td align="right" width="148" height="24"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif"><b><%=formatcurrency(rsCat1("avg_sqft"),2)%></b></font>
    </td>
  </tr>
</table>
<table border="0" width="100%" height="1" cellspacing="1" cellpadding="0">
<tr>
<td width="100%" height="1" valign="bottom" align="left">

<p align="left"><i>Choose a Building to show Tenant
detail&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</i></p>

      <p align="left"><i>Update as <%=date()%></i></p>

<center>

<p>&nbsp;</center></td>
</tr>
</table>
<table border="1" cellpadding="3" cellspacing="4" width="100%">
  <tr bgcolor="#0099FF"> 
    <th align="center" width="10%"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">&nbsp;#</font></p>
    </th>
    <th align="center" width="20%"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Building 
      Name</font></th>
    <th align="center" width="10%"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">ERI 
      Leases</font></th>
    <th align="center" width="15%"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Surveyed 
      ERI SQFT</font></th>
    <th align="center" width="15%"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Yearly 
      Revenue</font></th>
    <th align="center" width="15%"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Monthly 
      Revenue</font></th>
    <th align="center" width="20%" colspan="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Tot 
      ERI Sqft</font></th>
    <th align="center" width="10%"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">AVG 
      $/sqft</font></th>
  </tr>
  <%
' Create and open ADO Connection object.

sql = "SELECT Buildings.BldgNum, Buildings.Strt, COUNT(Tenant_info.Tenant_no) AS Tot_tenant, " &_ 
    "SUM(Tenant_info.ccm) AS tot_ccm, SUM(Tenant_info.CCY) " &_ 
    "AS tot_ccy, SUM(Tenant_info.sqft) AS Tot_sqft, " &_ 
    "SUM(Tenant_info.CCY) / SUM(Tenant_info.sqft) AS avg_sqft,Sur_sqft=isnull((select sum(tenant_info.sqft)as Sqft FROM tenant_info WHERE last_sur_kw <> 0 and  (Tenant_info.Lease_Exp_Date > GETDATE())and bldg_no =bldgnum),0) FROM Buildings INNER JOIN " &_
    "Tenant_info ON Buildings.BldgNum = Tenant_info.Bldg_no WHERE (Buildings.Owner_id = N'" & request("portfolioid") & "') AND " &_ 
    "(Tenant_info.Lease_Exp_Date > { fn NOW() }) GROUP BY Buildings.BldgNum, Buildings.Strt"

rsCat.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Loop thorough recordset object, displaying each record.

Do While Not rsCat.EOF
delta_sqft = (rsCat("tot_sqft") / portfolio_sqft) 
 %>
  <tr> 
    <td valign="top" align="center" width="10%"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
        <% 
   Response.write "<a href=" & chr(34) & "eri.asp?" & _
   "qcatnr=" & rsCat("bldgnum") & chr(34) & ">" & rscat("bldgnum")%>
        </font> 
    </td>
    <td valign="top" align="center" width="20%"> 
      <p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsCat("strt")%></font> 
    </td>
    <td valign="top" align="center" width="10%"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsCat("tot_tenant")%></font> 
    </td>
    <td valign="top" align="center" width="15%"> 
      <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=Formatnumber(rsCat("sur_sqft"),0)%></font></div>
    </td>
    <td valign="top" align="center" width="15%"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=formatcurrency(rsCat("tot_ccy"),0)%></font> 
    </td>
    <td valign="top" align="center" width="15%"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=formatcurrency(rsCat("tot_ccm"),0)%></font> 
    </td>
    <td valign="top" align="center" width="10%"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=formatnumber(rsCat("tot_sqft"),0)%></font> 
    </td>
    <td valign="top" align="center" width="10%"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=formatpercent(delta_sqft,2)%></font> 
    </td>
    <td valign="top" align="center" width="10%" bgcolor="#00FFFF"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=formatcurrency(rsCat("avg_sqft"),2)%></font> 
    </td>
  </tr>
  <%
  rsCat.MoveNext  
Loop
' Close and destroy the recordset and connection objects.
rsCat.Close
Set rsCat = Nothing
cnn1.Close
Set cnn1 = Nothing
%>
</table>


</body>

</html>
