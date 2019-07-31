<html>
<head>
<title>Utility Meters</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function meter(acctid,utility,bldg,flag){
	var temp="utilitymeter.asp?acctid="+acctid+"&utility="+utility+"&bldg="+bldg+"&flag="+flag
	document.frames.descriptions.location=temp
}
function updatemeter(id1,utility,bldg){
	var temp="utilitymeter.asp?meterid="+id1+"&utility="+utility+"&bldg="+bldg
	//alert(temp)
	document.frames.descriptions.location=temp
}
</script>
<%@Language="VBScript"%>

<%
acctid=Request.querystring("acctid")
bldg=Request.querystring("bldg")
utility=Request.querystring("utility")
id1=Request.querystring("meterid")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")


sqlstr= "select * from meters1 where acctid='" &acctid& "' "

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
if not rst1.EOF  then%>
<body bgcolor="#FFFFFF">
<form name="form2" method="post" action="">

<table width="100%" border="0" height="30">
  <tr> 
    <td bgcolor="#3399CC" height="30" width="89%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Account Number</i></font><font size="4" color="#FFFFFF"> 
        </font><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><%=Request.querystring("acctid")%></font></font></font></font></div>
    </td>
	<td bgcolor="#3399CC" width="11%"> 
	<input type="hidden" name="acct" value="<%=Request.querystring("acctid")%>">
	<input type="hidden" name="utility" value="<%=Request.querystring("utility")%>">
	<input type="hidden" name="bldg" value="<%=Request.querystring("bldg")%>">
	<input type="hidden" name="flag" value="1">
      <input type="button" name="addmeter" value="Add New Meter" onclick="meter(acct.value,utility.value,bldg.value,flag.value)">
    </td>
	</tr>
</table>
</form>
<form name="form3" method="post" action="">
  <table width="813" height="264" border="0">
    <tr > 
      <td width="40%"><IFRAME name="meters" width="100%" height="100%" src="metersearch.asp?acctid=<%=trim(Request.querystring("acctid"))%>&utility=<%=trim(Request.querystring("utility"))%>&bldg=<%=trim(Request.querystring("bldg"))%>" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></td>
      <td width="60%"><IFRAME name="descriptions" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
      </td>
    </tr>
  </table>
  </form>
<%
rst1.close
	 
else%>
<form name="form1" method="post" action="">

<table width="100%" border="0" height="30">
  <tr> 
    <td bgcolor="#3399CC" height="30" width="89%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>No 
        Meters Exist For This Account</i></font></font></div>
    </td>
	
    <td bgcolor="#3399CC" width="11%"> 
      <input type="hidden" name="acct2" value="<%=Request.querystring("acctid")%>">
      <input type="hidden" name="utility2" value="<%=Request.querystring("utility")%>">
      <input type="hidden" name="bldg2" value="<%=Request.querystring("bldg")%>">
      <input type="hidden" name="flag2" value="1">
      <input type="button" name="addmeter2" value="Add New Meter" onClick="meter(acct2.value,utility2.value,bldg2.value,flag2.value)">
    </td>
  </tr>
</table>
</form>
<form name="form3" method="post" action="">
  <table width="813" height="264" border="0">
    <tr > 
      <td width="40%"><iframe name="meters" width="100%" height="100%" src="metersearch.asp?acctid=<%=trim(Request.querystring("acctid"))%>&utility=<%=trim(Request.querystring("utility"))%>&bldg=<%=trim(Request.querystring("bldg"))%>" scrolling="auto" marginwidth="0" marginheight="0" ></iframe></td>
      <td width="60%"><iframe name="descriptions" width="100%" height="100%" src="null.htm"  marginwidth="0" marginheight="0" ></iframe> 
      </td>
    </tr>
  </table>
</form><%
end if%>
<table>
<tr>
<td>
<i><font face="Arial, Helvetica, sans-serif">*Click any highlighted meter to update</font></i>
</td>
</tr>
</table>

<%
set cnn1=nothing
%>

</body>
</html>