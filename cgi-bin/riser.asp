<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<title>Riser Summary</title>
</head>
<script>
function loadentry(bldg,fl){

	var temp = 'floorsummary.asp?b=' +bldg+'&f='+fl

	
	document.location = temp
	}
</script>
<body bgcolor="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request.querystring("B")
riser= request.querystring("r")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)

sql="exec sp_riser '" & bldg & "',"&riser&""

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

if not rst1.EOF then 
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#336699" width="23%" height="29" nowrap><span class=standardheader>Riser 
      <%=rst1("riser")%>: Floors Served and Usage Summary as of : <font size="3"><b><%=rst1("date") %></b></font></span></td>
    <td bgcolor="#336699" width="23%"><div align="right"><a href="javascript:history.back()"><img src="/images/intranet/btn-back.gif" width="68" height="19"></a></div></td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#6699CC"> 
    <td width="16%" height="28" align="center"><b>Floor</b></td>
    <td width="17%" height="28" align="center"><b>Actual Sub-metered Demand by 
      Floor (KW)</b></td>
    <td width="17%" height="28" align="center"><b>Actual Sub-metered Power by 
      Floor (w/sqft)</b></td>
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
   <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=bldg%>','<%=rst1("floor")%>')"> 
        
        <td width="16%" height="19" align="right"><b><font  size="2" face="Arial, Helvetica, sans-serif"><%=rst1("floor")%></font></b></td>
  
        <td width="17%" height="19" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("usage_by_floor"),2)%></font></b></td>
        <td width="17%" height="19" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("wsqft"),2)%></font></b></td>
    </tr>
 </form>
<% 
		totalDemand = totalDemand + cdbl(rst1("usage_by_floor"))
		totalpwr = totalpwr + cdbl(rst1("wsqft"))
		rst1.movenext
		Wend
		
		%>
    <tr> 
      <td colspan=2 align="right" style="border-top:1px solid #000000;"><b><font size="2"face="Arial, Helvetica, sans-serif">Total 
        (KW) : <%=formatnumber(totalDemand,2)%></font></b></td>
      <td colspan=2 align="right" style="border-top:1px solid #000000;"><b><font size="2"face="Arial, Helvetica, sans-serif">Total 
        (w/sqft) : <%=formatnumber(totalpwr,2)%></font></b></td>
    </tr>
 </table>
</div>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click any Floor 
    for summary information</i></b></font></p>
</div>
<%
else
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b>NO DATA AVAILABLE</b></font></p>
</div>
<%
end if 
rst1.close
set cnn1 = nothing
%>
</body>

</html>
