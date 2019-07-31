<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim acctid, bldgnum, pid, utility
acctid = request.querystring("acctid")
bldgnum = request.querystring("building")
utility = request.querystring("utility")

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(bldgnum)
rst1.open "SELECT portfolioid FROM buildings WHERE bldgnum='"&bldgnum&"'", cnn1
if not rst1.eof then pid = rst1("portfolioid")
rst1.close
dim DBMainmodIP
DBMainmodIP = "["&getPidIP(pid)&"].Supermod.dbo."
%>
<html>
<head>
<title>Bill Period Selection</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>
<script>

function findbill(acctid,ypid,by,bp,utility){
var temp
//alert(utility)
 if (utility==2) {
	temp = "acctdetail.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	}else if (utility==4){
	 temp = "acctdetailgas.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	}else if (utility==1){
	 temp = "acctdetailsteam.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	 }else {
	 temp = "acctwdetail.asp?acctid=" +acctid+ "&ypid=" +ypid+ "&by=" +by+ "&bp=" +bp+ "&utility=" +utility
	 //alert(temp)
	}
	window.opener.document.all['perioddisplay'].innerText=by+"."+bp
	window.opener.document.all['entryframe'].style.visibility = "visible" 
	window.opener.document.getElementById('entry').src=temp+"&bldg=<%=bldgnum%>";
	window.close()
		
}

</script>
<body bgcolor="#eeeeee">
<%
'response.write u
'response.end
sqlstr="select * from billyrperiod where bldgnum='"&bldgnum&"' and utility='"&utility&"' and dateend<=dateadd(m,6,getdate()) order by billyear desc,billperiod desc"

'response.write sqlstr
'response.end
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1%>

<% if not rst1.eof then%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#3399cc"> 
    <td bgcolor="#3399CC" class="standardheader">Bill Year Periods</td>
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
     <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = '#eeeeee'" onclick="javascript:findbill('<%=request.querystring("acctid")%>','<%=rst1("ypid")%>','<%=rst1("billyear")%>','<%=rst1("billperiod")%>','<%=request.querystring("utility")%>')" style="cursor:hand;"> 
      <td width="18%"><%=rst1("billyear")%></td>
      <td width="23%"><%=rst1("billperiod")%></td>
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
set rst1 = nothing 'TK: 04/28/2006
set cnn1=nothing
%>
</body>
</html>
