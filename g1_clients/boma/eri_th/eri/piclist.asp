<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
    <%
bldg = Request("qcatnr")

Set cnn1 = Server.CreateObject("ADODB.Connection")
openStr= getconnect(0,0,"Engineering") 
cnn1.Open openStr

Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT *, isnull(last_sur_kwh,0) as lastsurkwh FROM tenant_info WHERE (bldg_no='" & bldg & "') order by tenantname"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<html>
<head>
<title>ERI Tenant list</title>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<link rel="stylesheet" type="text/css" 
      href="../holiday/holiday.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none, default">
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
<%
' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.

Response.Write "<script>" & vbCrLf

If rst1.EOF Then
  Response.Write "parent.frames.info.location = " & _
                 Chr(34) & "blank.htm" & Chr(34) & vbCrLf
Else
  Response.Write "parent.frames.info.location = " & _
                  Chr(34) & _
                  "info.asp" & _
                  "?qcatnr=" & rst1("Tenant_no") & _
                  Chr(34) & vbCrLf
End If

Response.Write "</script>" & vbCrLf
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC" height="8">
  <tr bgcolor="#0099FF"> 
    <td width="10%" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">Tenant 
      #</font></b></td>
    <td width="30%" align="center"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1"><b>Tenant 
      Name</b></font></td>
    <td width="10%" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">Sqft</font></b></td>
    <td width="10%" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">Monthly 
      Charge</font></b></td>
    <td width="10%" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">Yearly 
      Charge</font></b></td>
    <td width="10%" align="center"><b> <font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">$ 
      / Sqft</font></b></td>
    <td width="10%" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">Lease 
      Exp. Date</font></b></td>
    <td width="10%" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="1">Move 
      Out Date</font></b></td>
  </tr>
</table>
<div style="overflow:auto; height:220">
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
    <%Do While Not rst1.EOF
%>
    <tr<% if clng(rst1("lastsurkwh")) > 0 then response.write " bgcolor=""#00CC66"""%><%if isnull(rst1("Move_out_date")) then response.write " bgcolor=""#FFFFCC"""%> bgcolor="#FFFFFF"> 
      <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
        <%
		Response.write "<a href=" & chr(34) & "info.asp?" & "qcatnr=" & rst1("Tenant_no") & chr(34) & " target=info>" & left(rst1("Tenant_no"),4)
%>
        </font></td>
      <td width="30%" height="20" align="center"> 
        <p align="left"><font face="Arial, Helvetica, sans-serif" size="2">Acme Tenant <%=left(rst1("Tenantname"),4)%> 
        </font>
    </td>
    <% If IsNull(rst1("sqft")) then %>
      <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=Formatnumber(rst1("sqft"),0)%> 
        </font>
    </td>
    <%end If%>
    <% If IsNull(rst1("ccm")) then %>
      <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=FormatCurrency(rst1("ccm"),2)%> 
        </font>
    </td>
    <%end If%>
    <% If IsNull(rst1("ccy")) then %>
      <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=FormatCurrency(rst1("ccy"),2)%> 
        </font>
    </td>
    <%end If%>
    <% If IsNull(rst1("cost_sqft")) then %>
      <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=formatcurrency(rst1("cost_sqft"),2)%> 
        </font>
    </td>
    <%end If%>
      <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
        <%if rst1("lease_exp_date")<>"1/1/1900" then response.write rst1("lease_exp_date")%>
        </font></td>
      <td width="10%" height="20" align="center"> 
        <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%if rst1("Move_out_date")<>"1/1/2025"  then response.write rst1("Move_out_date")%> 
        </font>
    </td>
    </tr>
    <%
  rst1.MoveNext  
Loop

'Close and destroy the recordset and connection objects.
rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
%>
  </table></div>

</body>