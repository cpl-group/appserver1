<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Riser Summary</title>
</head>
<script>
function loadentry(bldg,fl){

	var temp = 'floorsummary.asp?b=' +bldg+'&f='+fl

	
	document.location = temp
	}
</script>
<body bgcolor="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request.querystring("B")
riser= request.querystring("r")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql="exec sp_riser '" & bldg & "',"&riser&""

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Riser 
      Summary  as of : <%=rst1("date")%>, 
      <%=rst1("time")%></font></b></font></td>
  </tr>
  
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0099FF"> 
    <td width="9%" height="1%" align="center"><b><font size="1" face="Arial">Riser</font></b></td>
    <td width="16%" height="1%" align="center"><b><font face="Arial" size="1">Floor</font></b></td>

    <td width="17%" height="1%" align="center"><b><font size="1" face="Arial">Usage by Floor</font></b></td>
   
  </tr>
</table>



<div align="left">
<table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
<%
while not rst1.eof
%>
<form name="form1" method="post" action=""> 
   <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=bldg%>','<%=rst1("floor")%>')"> 
        <td width="9%" height="19" align="center"> 
          <div align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("riser")%></font></b></div>
      </td>
        <td width="16%" height="19" align="right"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("floor")%></font></b></td>
  
        <td width="17%" height="19" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("usage_by_floor"),2)%></font></b></td>
        
    </tr>
 </form>
<% 
		rst1.movenext
		Wend
		
		%>
 </table>
</div>


<%
rst1.close
set cnn1 = nothing
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click any Floor 
    for summary information</i></b></font></p>
</div>
</body>

</html>
