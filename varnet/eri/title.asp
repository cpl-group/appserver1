<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<!-- #include file="./adovbs.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Genergy ERI Management</title>
<%

bldg1 = Request("bldg")

tmpMoveFrame =  "parent.frames.piclist.location = " & Chr(34) & _
                  "piclist.asp?bldg=" & bldg1 & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf

Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

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
<script>
function  open_bldginfo(bldgnum, infotype){
		var temp="buildinginfo.asp?bldgnum='"+bldgnum+"'&infotype="+infotype
       window.open(temp,"", "scrollbars=yes, width=400,height=200" );
}
</script>


</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
 <center>
    <table border="1" width="100%" height="100%" cellpadding="0" cellspacing="0">
      <tr> 
        <td width="50%" height "20%"><font size="4"><%=bldgStreet%></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="3">Total 
          no. of ERI Tenants:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="4"><%=tenant%></font></td>
      </tr>
      <tr> 
        <td width="50%" height "20%"><font size="4"><%=bldgname%></font></td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="3">Total 
          of ERI sqft:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="4"><%=formatnumber(sumSqft,0)%></font></td>
      </tr>
      <tr> 
        <td width="50%" height="0" "20%" bordercolor="#FFFFFF">
		<form name="form1" method="post" action="">
			<input type="hidden" name="bldgnum" value="<%=bldg1 %>">
			<input type="hidden" name="building" value="bld">
			<input type="hidden" name="billing" value="billing">
          <input type="button" name="bldgadd" value="Building Address" onclick="open_bldginfo( bldgnum.value, building.value)"> 
          <input type="button" name="billadd" value="Billing Address"  onclick="open_bldginfo( bldgnum.value, billing.value)">
    	  </form>
        </td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="3">Total 
          of ERI Monthly Charge:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="4"><%=FormatCurrency(sumCCM,0)%></font></td>
      </tr>
      <tr> 
        <td width="50%" height="0" "20%" bordercolor="#FFFFFF">&nbsp;</td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="3">Total 
          of ERI Yearly Charge:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="4"><%=FormatCurrency(sumCCy,0)%></font></td>
      </tr>
      <tr> 
        <td width="50%" height="0" "20%" bordercolor="#FFFFFF">&nbsp;</td>
        <td width="30%" height="20%" align="left" valign="bottom"><font size="3">AVG 
          of ERI $/sqft:</font></td>
        <td width="10%" height="20%" valign="bottom" align="right"><font size="4"><%=FormatCurrency(avg_sqft,2)%></font></td>
      </tr>
    </table>
 </center>
</div>


</body>

</html>
