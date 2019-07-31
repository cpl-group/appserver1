<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="./adovbs.inc" -->
<%  
'3/20/2008 N.Ambo amended to gray out offline tenants and palce them at the bottom of the list
 %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Genergy ERI Management</title>
<%

bldg1 = Request("bldg")

'tmpMoveFrame =  "parent.frames.piclist.location = " & Chr(34) & _
'                  "piclist.asp?bldg=" & bldg1 & chr(34) & vbCrLf 

'Response.Write "<script>" & vbCrLf

'Response.Write tmpMoveFrame
'Response.Write "</script>" & vbCrLf

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"engineering")

sql = "SELECT Buildings.BldgNum, Buildings.BldgName, Buildings.Strt, Count(Tenant_info.Tenant_no) AS CountOfTenant_no,  case  Sum(Tenant_info.sqft) when 0 then 1 else Sum(Tenant_info.sqft) end  AS SumOfsqft, isnull(Sum(Tenant_info.ccm),0) AS SumOfccm, isnull(Sum(Tenant_info.ccy),0) AS SumOfccy FROM Buildings INNER JOIN Tenant_info ON Buildings.BldgNum = Tenant_info.Bldg_no WHERE (((isnull(Tenant_info.Lease_Exp_Date,'1/1/2025'))> { fn NOW() })) and Buildings.BldgNum='" & bldg1 & "' GROUP BY Buildings.BldgNum, Buildings.BldgName, Buildings.Strt "


'response.write sql
'response.end

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
'response.write "SumOfccy is " & SumOfccy
'response.end 
end if

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnnl = Nothing
dim bldgstr
bldgstr = Chr(34) & bldg1 & Chr(34)

%>
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="none, default">
<script>
function addnew(){
  // Load Add new tenant form
  //document.frames.title.location.href="null.htm";
  parent.frames.info.location.href="tenantlist.asp?bldg=" + <%=bldgstr%>;
  location.href="ti_add.asp?bldg=" + <%=bldgstr%>;
}
</script>

<link rel="Stylesheet" href="styles.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" leftmargin="0" topmargin="0">
<form name="form1" method="post" action="">

<table border=0 width="100%" cellpadding="3" cellspacing="0" style="border-bottom:1px solid #cccccc;">
<tr> 
  <td style="font-weight:bold;"><!--[[b]][[%=bldgStreet%]][[/b]] |--> <%=bldgname%>
  <input type="hidden" name="bldgnum" value="<%=bldg1 %>">
  <input type="hidden" name="building" value="bld">
  <input type="hidden" name="billing" value="billing">
  </td>
</tr>
<tr>
  <td>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td align="left">Total no. of ERI Tenants:</td>
    <td align="right"><b><%=tenant%></b></td>
    <td width="15" style="border-right:1px solid #ffffff">&nbsp;</td>
    <td width="5"></td>
    <td align="left">Total of ERI Monthly Charge:</td>
    <td align="right"><b><%=FormatCurrency(sumCCM,0)%></b></td>
    <td width="15" style="border-right:1px solid #ffffff">&nbsp;</td>
    <td width="5"></td>
    <td align="left">AVG of ERI $/sqft:</td>
    <td align="right"><b><%=FormatCurrency(avg_sqft,2)%></b></td>
  </tr>
  <tr> 
    <td align="left">Total of ERI sqft:</td>
    <td align="right"><b><%=formatnumber(sumSqft,0)%></b></td>
    <td width="15" style="border-right:1px solid #ffffff">&nbsp;</td>
    <td width="5"></td>
    <td align="left">Total of ERI Yearly Charge:</td>
    <td align="right"><b><%=FormatCurrency(sumCCy,0)%></b></td>
    <td width="15" style="border-right:1px solid #ffffff">&nbsp;</td>
    <td width="5"></td>
    <td colspan="2"></td>
  </tr>
  </table>    
  </td>
</tr>
</table>

<%	'  if Session("eri") > 2 then %>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #ffffff;">
<tr>  
  <td valign="bottom"><b>Tenants</b></td>
 <td align="right"><%if not(isBuildingOff(request("bldg"))) then%><input type="button" name="Submit2" value="Add New Tenant" onclick="addnew()"><%end if%></td>
 
</tr>
<tr>
  <td colspan="2">
<%	'   end if  	%>
<div id="piclist" style="overflow:auto;height:150px;width:100%;border:1px solid #dddddd;background-color:#ffffff;">
<table border=0 cellpadding="3" cellspacing="1" width="100%">
<tr bgcolor="#dddddd" style="font-weight:normal;"> 
  <td>Tenant&nbsp;Num.</td>
  <td>Tenant Name</td>
  <td>Sqft</td>
  <td>Monthly&nbsp;Charge</td>
  <td>Yearly&nbsp;Charge</td>
  <td>$ / Sqft</td>
  <td>Lease Expiration</td>
  <td>Move Out Date</td>
</tr>
<%


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"engineering")


Set rst1 = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM tenant_info WHERE (bldg_no='" & bldg1 & "') order by online desc, tenantname"  '3/20/2008 N.Ambo amended to take the bldg number from the querystring variable

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.
Response.Write "<script>" & vbCrLf




If rst1.EOF Then
  Response.Write "parent.frames.info.location = " & _
                 Chr(34) & "null.htm" & Chr(34) & vbCrLf
Else
  Response.Write "parent.frames.info.location = " & _
                  Chr(34) & _
                  "info.asp?bldg="&request("bldg")&"&qcatnr=" & rst1("Tenant_no") & _
                  Chr(34) & vbCrLf
End If

Response.Write "</script>" & vbCrLf

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
<tr onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:parent.info.document.location='info.asp?bldg=<%=request("bldg")%>&qcatnr=<%=rst1("Tenant_no")%>'" > 
  <td><%=fonttag%><span><%=rst1("Tenant_no")%></span><%=unfonttag%></td>
  <td><%=fonttag%><span><%=rst1("Tenantname")%></span><%=unfonttag%></td>
  <% If IsNull(rst1("sqft")) then %>
  <td></td>
  <% else %>
  <td><%=fonttag%><span><%=Formatnumber(rst1("sqft"),0)%></span><%=unfonttag%></td>
  <%end If%>
  <% If IsNull(rst1("ccm")) then %>
  <td></td>
  <% else %>
  <td><%=fonttag%><span><%=FormatCurrency(rst1("ccm"),2)%></span><%=unfonttag%></td>
  <%end If%>
  <% If IsNull(rst1("ccy")) then %>
  <td></td>
  <% else %>
  <td><%=fonttag%><span><%=FormatCurrency(rst1("ccy"),2)%></span><%=unfonttag%></td>
  <%end If%>
  <% If IsNull(rst1("cost_sqft")) then %>
  <td></td>
  <% else %>
  <td><%=fonttag%><span><%=formatcurrency(rst1("cost_sqft"),2)%></span><%=unfonttag%></td>
  <%end If%>
  <td><%=fonttag%><span><%if rst1("lease_exp_date")<>"1/1/1900"  then response.write rst1("lease_exp_date")%></span><%=unfonttag%></td>
  <td><%=fonttag%><span><%if rst1("Move_out_date")<>"1/1/1900" then response.write rst1("Move_out_date")%></span><%=unfonttag%></td>
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
</div>
  </td>
</tr>
</table>
</form>


</body>

</html>
