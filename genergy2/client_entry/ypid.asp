<%@Language="VBScript"%>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")
acctid=request.querystring("acctid")
bldg=request.querystring("building")
u=request.querystring("utility")
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function findbill(acctid,ypid,by,bp,utility){
var temp
//alert(utility)
 if (utility=="Electricity") {
	temp = "acctdetail.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	}else if (utility=="Gas"){
	 temp = "acctdetailgas.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	}else if (utility=="Steam"){
	 temp = "acctdetailsteam.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	 }else {
	 temp = "acctwdetail.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	 //alert(temp)
	}opener.document.frames.entry.location=temp+"&bldg=<%=bldg%>";
	window.close()
		
}

</script>
<body bgcolor="#FFFFFF" text="#000000">

<%
'response.write u
'response.end
sqlstr="select * from billyrperiod where bldgnum='"&bldg&"' and utility='"&u&"' and dateStart<='"&now()&"' order by billyear desc,billperiod desc"

'response.write sqlstr
'response.end
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1%>

<% if not rst1.eof then%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font size="4" color="#FFFFFF">BILL 
        YEAR PERIODS</font></b></font></div>
    </td>
  </tr>
</table>

<table width="100%" border="0">
  <tr>
   
    <td width="18%"><font face="Arial, Helvetica, sans-serif">Bill Year</font></td>
    <td width="23%"><font face="Arial, Helvetica, sans-serif">Bill Period</font></td>
    <td width="17%"><font face="Arial, Helvetica, sans-serif">Start Date</font></td>
    <td width="28%"><font face="Arial, Helvetica, sans-serif">End Date</font></td>
  </tr>
  <%While not rst1.EOF %>
  <form name="form1" method="post" action="">
     <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:findbill('<%=request.querystring("acctid")%>','<%=rst1("ypid")%>','<%=rst1("billyear")%>','<%=rst1("billperiod")%>','<%=request.querystring("utility")%>')"> 
     
      <td width="18%"><%=rst1("billyear")%></td>
      <td width="23%"> 
        <div align="center"><%=rst1("billperiod")%></div>
      </td>
      <td width="17%"><%=rst1("datestart")%></td>
      <td width="28%"><%=rst1("dateend")%></td>
    </tr>
  </form>
  <%rst1.movenext
  Wend
  %>
</table>
<%else%>
<table width="100%" border="0">
  <tr bgcolor="#3399CC">
    <td> 
      <div align="left"><b><font face="Arial, Helvetica, sans-serif"><i><font color="#FFFFFF" size="2">No 
        bill periods have been defined for this building. Please contact <a href ="mailto:george_nemeth@genergy.com">George 
        Nemeth</a> to add billperiods.</font></i></font></b></div>
    </td>
  </tr>
</table>
<%rst1.close
end if
set cnn1=nothing
%>

</body>
</html>
