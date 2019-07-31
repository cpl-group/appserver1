<%@Language="VBScript"%>


<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function processpo(poid,action,ponum,podate) {

	if (action=="ACCEPT") {
		var poaction="accept"
	} else 
		if (action=="REJECT"){
		var poaction="reject"
	}else{
		var poaction="question"
		}	
	
	var temp = "processpo1.asp?poid=" + poid + "&poaction=" + poaction + "&ponum=" + ponum + "&podate="+ podate
	
	window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );
}

</script>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">ACCEPT 
        / REJECT SUBMITTED PO's</font></b></div>
    </td>
  </tr>
</table>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")

sqlstr = "select ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber,po.* ,employees.[first name]+' '+employees.[last name] as req from po join employees on po.requistioner=substring(employees.username,7,20) where submitted = 1 and accepted =0 order by podate desc"
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
%>	
	
<table width="100%" border="0">
  <tr> 
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>NO PO's 
        WAITING FOR REVIEW</b></font></div>
    </td>
  </tr>
</table>
<%
Else
x=0
%>


<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="15%"><font face="Arial, Helvetica, sans-serif">PO #</font></td>
    <td width="24%"><font face="Arial, Helvetica, sans-serif">PO Date</font></td>
    <td width="35%"><font face="Arial, Helvetica, sans-serif">Requistioner</font></td>
    <td width="10%"><font face="Arial, Helvetica, sans-serif">PO Total</font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%
  While not rst1.EOF
  
  %><form name="form1" method="post" action="">
  <tr> 
      <td width="15%">
        <%if rst1("question")="True" then %>
        <font color="#0033FF">*</font>
        <%end if%>
        <a href=<%="poview.asp?po=" & rst1("ponumber") %> ><font face="Arial, Helvetica, sans-serif"><%=rst1("ponumber") %></font></a>
<input type="hidden" name="poid" value="<%=rst1("id")%>">
</td>
    <td width="24%"><font face="Arial, Helvetica, sans-serif"><%=rst1("podate") %></font></td>
      <td width="13%"><font face="Arial, Helvetica, sans-serif"><%=rst1("req")%>
		  <input type="hidden" name="job" value="<%=rst1("requistioner")%>">
		  </font></td>
    <td width="10%"><font face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst1("po_total"))%></font></td>
    <td width="16%"> 
      <div align="center"> <font face="Arial, Helvetica, sans-serif"> 
	  <input type="hidden" name="ponum" value="<%=rst1("ponumber")%>">
	  <input type="hidden" name="d" value="<%=rst1("podate") %>">
	  
        <input type="button" name="Button" value="ACCEPT" onclick="processpo(poid.value, this.value,ponum.value,d.value)">
        <input type="button" name="Button" value="REJECT" onclick="processpo(poid.value, this.value,ponum.value,d.value)">
		<font face="Arial, Helvetica, sans-serif" size="1"><a onclick="processpo('<%=rst1("id")%>','Question','<%=rst1("ponumber")%>','<%=rst1("podate")%>')"><img src="question-ccc.gif" border="0"></a>
        </font>
        </font></div>
    </td>
  </tr></form>
  <%
  rst1.movenext
  Wend
  end if
  %>
</table>

</body>
</html>