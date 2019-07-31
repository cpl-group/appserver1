<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<title>Floor Detail</title>
</head>

<body bgcolor="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request("B")
fl = request("F")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)

sql="exec sp_floor_detail '" & bldg & "', '"&fl&"'"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
if not rst1.EOF then 
	riser=rst1("riser")
'response.write riser
'response.end
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#336699" width="23%" height="29" nowrap><span class=standardheader>Riser 
      Detail by Floor as of: <font size="3"><b><%=rst1("date")%></b></span></td>
    <td bgcolor="#336699" width="23%"><div align="right"><a href="javascript:history.back()"><img src="/images/intranet/btn-back.gif" width="68" height="19"></a></div></td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#6699CC"> 
    <td width="16%" height="1%" align="center"><b>Selected 
      Floor</b></td>
    <td width="24%" height="1%" align="center"><b>Riser 
      Serving Floor</b></td>
    <td width="34%" height="1%" align="center"><b>Current 
      Sub-metered Demand (kw)</b></td>
    <td width="26%" height="1%" align="center"><b>Current 
      Sub-metered Power (w/sqft)</b></td>
   
  </tr>
</table>



<div align="left">
  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
    <%
totalpwr = 0
totalDemand = 0
while not rst1.eof
%>
    <form name="form1" method="post" action="">
      <tr> 
        <td width="16%" height="19" align="center"> <div align="right"><b><a href="floorsummary.asp?b=<%=bldg%>&f=<%=fl%>"><font  size="2" face="Arial, Helvetica, sans-serif"><%=rst1("floor")%></font></a><a href="riser.asp?b=<%=bldg%>&r='<%=riser%>'"></a><a href="floorsummary.asp?b=<%=bldg%>&f=<%=fl%>"></a></b></div></td>
        <td width="24%" height="19" align="right"><b><a href="riser.asp?b=<%=bldg%>&r='<%=riser%>'"></a><a href="floorsummary.asp?b=<%=bldg%>&f=<%=fl%>"></a><a href="riser.asp?b=<%=bldg%>&r='<%=rst1("riser")%>'"><font size="2" face="Arial, Helvetica, sans-serif" ><%=rst1("riser")%></font></a></b></td>
        <td width="34%" height="19" align="right"><b><font size="2"face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("current_demand_kw"),2)%></font></b></td>
        <td width="26%" height="19" align="right"><b><font  size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("delivered_wsqft"),2)%></font></b></td>
      </tr>
    </form>
    <% 
		totalpwr = totalpwr +  cdbl(rst1("delivered_wsqft"))
		totalDemand = totaldemand + cdbl(rst1("current_demand_kw"))
		rst1.movenext
		Wend
		
		%>
    <tr> 
       <td style="border-top:1px solid #000000;" align="right" colspan=3><b><font size="2"face="Arial, Helvetica, sans-serif">Total 
        (KW) : <%=formatnumber(totalDemand,2)%></font></b></td>
       <td align="right" style="border-top:1px solid #000000;"><b><font size="2"face="Arial, Helvetica, sans-serif">Total 
        (w/sqft) : <%=formatnumber(totalpwr,2)%></font></b></td>
    </tr>
  </table>
</div>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click any Floor or Riser
    for  detailed information</i></b></font></p>
</div>
<%
else
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b>NO DATA AVAILABLE</b></font></p>
</div>
<%end if 

rst1.close
set cnn1 = nothing
%>
</body>

</html>
