<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<title>Floor Summary</title>
</head>
<script>
function loadentry(bldg,fl){

	var temp = 'floordetail.asp?b=' +bldg+'&f='+fl

	
	document.location = temp
	}
</script>
<body bgcolor="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request("B")
fl = request("F")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)

sql="exec sp_floor '" & bldg & "', '"&fl&"'"

rst1.Open sql, cnn1

if  rst1.State <> 0 then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#336699" width="23%" height="29" nowrap><span class=standardheader>Floor 
      and Usage Summary as of: <font size="3"><b><%=rst1("date")%></b></font></span></td>
    <td bgcolor="#336699" width="23%"><div align="right"><a href="javascript:history.back()"><img src="/images/intranet/btn-back.gif" width="68" height="19"></a></div></td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#6699CC"> 
    <td width="14%" height="1%" align="center"><b>Selected 
      Floor</b></td>
    <td width="8%" height="1%" align="center"><b>Area 
      Floor (sqft) </b></td>
    <td width="8%" height="1%" align="center"><b>Calculated 
      Deliverable Power (w/sqft)</b></td>
    
    <td width="13%" height="1%" align="center"><b>Current 
      Sub-metered Demand (kw)</b></td>
    <td width="12%" height="1%" align="center"><b>Current 
      Sub-metered Power (w/sqft)</b></td>
  </tr>
</table>

<%
while not rst1.eof
%>

<div align="left">

  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
 <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=bldg%>','<%=fl%>')"> 

      <td width="14%" height="1%" align="center"> 
        <div align="center"><b><font  size="2" face="Arial, Helvetica, sans-serif"><%=rst1("floor")%></font></b></div>
      </td>
      <td width="8%" height="1%" align="right">
        <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" ><%=rst1("sqft")%></font></b></div>
      </td>
      <td width="8%" height="1%" align="right">
        <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("wsqft"),2)%></font></b></div>
      </td>
   
      <td width="13%" height="1%" align="right">
        <div align="center"><b><font  size="2" face="Arial, Helvetica, sans-serif"><%if isNull(rst1("current_demand_kw")) then response.write "unavailable" else response.write formatnumber(rst1("current_demand_kw"),2) end if%></font></b></div>
      </td>
      <td width="12%" height="1%" align="right">
        <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif"><%if isNull(rst1("delivered_wsqft")) then response.write "unavailable" else response.write formatnumber(rst1("delivered_wsqft"),2) end if%></font></b></div>
      </td>
    </tr>
  </table>


</div>


<%
rst1.movenext
wend
rst1.close
set cnn1 = nothing
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click  Floor
    for  detailed  information</i></b></font></p>
</div>
<%

else

	Response.write "<div align='center'>NO DATA FOUND FOR FLOOR " &fl&"</div>"

end if %>
</body>

</html>
