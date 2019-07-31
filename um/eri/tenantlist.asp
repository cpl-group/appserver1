<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<%
'3/20/2008 N.Ambo amended to gray out offline tenants and move them to the bottom of the lsit of tenants		
%>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Genergy ERI Management</title>
<link rel="Stylesheet" href="styles.css" type="text/css">   
</head>
<body bgcolor="#FFFFFF">

<table border=0 cellpadding="3" cellspacing="1" width="100%">
<tr valign="top" bgcolor="#dddddd" style="font-weight:bold;"> 
  <td width="10%" nowrap>Tenant Num.</td>
  <td width="30%" nowrap>Tenant Name</td>
  <td width="10%" nowrap>Sqft</td>
  <td width="10%" nowrap>Monthly Charge</td>
  <td width="10%" nowrap>Yearly Charge</td>
  <td width="10%" nowrap>$ / Sqft</td>
  <td width="10%" nowrap>Lease Exp. Date</td>
  <td width="10%" nowrap>Move Out Date</td>
</tr>
    <%
bldg = Request("bldg")


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")


Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tenant_info WHERE (bldg_no='" & bldg & "') order by online desc, tenantname"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

If not rst1.EOF then 
rst1.movefirst
end if

Do While Not rst1.EOF
dim fonttag, unfonttag
	if  not rst1("online") then 
		fonttag = "<i><font color='#555555'"
		unfonttag = "</i></font>"
	end if	
%>
    <tr> 
      <td><%=fonttag%><span><%=rst1("Tenant_no")%></span><%=unfonttag%></td>
      <td><%=fonttag%><span><%=rst1("Tenantname")%></span><%=unfonttag%></td>
      
      <td align="right"><% If not IsNull(rst1("sqft")) then %><%=fonttag%><span><%=Formatnumber(rst1("sqft"),0)%></span><%=unfonttag%><% end if %></td>
      <% If IsNull(rst1("ccm")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="right"> 
        <%=fonttag%><span><%=FormatCurrency(rst1("ccm"),2)%>
      </span><%=unfonttag%></td>
      <%end If%>
      <% If IsNull(rst1("ccy")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="right"> 
        <%=fonttag%><span><%=FormatCurrency(rst1("ccy"),2)%>
      </span><%=unfonttag%></td>
      <%end If%>
      <% If IsNull(rst1("cost_sqft")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="right"> 
        <%=fonttag%><span><%=formatcurrency(rst1("cost_sqft"),2)%>
      </span><%=unfonttag%></td>
      <%end If%>
      <td width="10%" height="20" align="center"><%=fonttag%><span><%=rst1("lease_exp_date")%></span><%=unfonttag%></td>
      <td width="10%" height="20" align="right"> 
        <%=fonttag%><span><%=rst1("Move_out_date")%></span><%=unfonttag%>
    </tr>
    <%
    fonttag=""
    unfonttag=""
  rst1.MoveNext  
Loop

'Close and destroy the recordset and connection objects.
rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
%>
  </table>

</body>
</html>

