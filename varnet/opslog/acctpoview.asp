<%@Language="VBScript"%>


<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script>
function po(id1){
	document.location="accpofilter.asp?id1="+id1
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">APPROVED 
        PO's </font></b></div>
    </td>
  </tr>
</table>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")

sqlstr = "select ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber,po.*, employees.[first name]+' '+employees.[last name] as req from po join employees on po.requistioner=substring(employees.username,7,20) where accepted = 1 and closed=0 order by podate desc"

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
%>	
	
<table width="100%" border="0">
  <tr> 
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>NO APPROVED 
        PO's WAITING</b></font></div>
    </td>
  </tr>
</table>
<%
Else
x=0
%>


<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="12%" height="2"><font face="Arial, Helvetica, sans-serif">PO #</font></td>
    <td width="15%" height="2"><font face="Arial, Helvetica, sans-serif">PO Date</font></td>
    <td width="23%" height="2"><font face="Arial, Helvetica, sans-serif">Vendor</font></td>
    <td width="19%" height="2"><font face="Arial, Helvetica, sans-serif">Requistioner</font></td>
    <td width="18%" height="2"><font face="Arial, Helvetica, sans-serif">PO Total</font></td>
    <td width="13%" height="2">&nbsp;</td>
  </tr>
  <%
  While not rst1.EOF
  
  %>
  <form name="form1" method="post" action="">
    <tr> 
      <td width="12%"><a href=<%="poview.asp?po=" & rst1("ponumber") %> ><font face="Arial, Helvetica, sans-serif"><%=rst1("ponumber") %></font></a></td>
      <td width="15%"><font face="Arial, Helvetica, sans-serif"><%=rst1("podate") %></font></td>
      <td width="23%"><font face="Arial, Helvetica, sans-serif"><%=rst1("vendor")%></font></td>
      <td width="19%"><font face="Arial, Helvetica, sans-serif"><%=rst1("req")%> 
        <input type="hidden" name="job" value="<%=rst1("requistioner")%>">
        </font></td>
      <td width="18%"><font face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst1("po_total"))%></font></td>
	  <td width="13%"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif">
		  <input type="hidden" name="id1" value="<%=rst1("id")%>">
          <input type="button" name="closepo" value="Close" onclick="po(id1.value)">
          </font></div>
      </td>
    </tr>
  </form>
  <%
  rst1.movenext
  Wend
  end if
  %>
</table>
<p>&nbsp;</p>
</body>
</html>
