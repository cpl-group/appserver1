<html>
<head>
<title>Utility Meters</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function updatemeter(meterid,utility,bldg,acctid){
	var temp="utilitymeter.asp?meterid="+meterid+"&utility="+utility+"&bldg="+bldg+"&acctid="+acctid
	//alert(temp)
	parent.document.frames.descriptions.location=temp
}
function deletemeter(meterid,bldg,acctid,utility){
	var temp="deletemeter.asp?meterid="+meterid+"&bldg="+bldg+"&acctid="+acctid+"&utility="+utility
	var temp3="null.htm"
	//alert(temp)
	if (confirm("Are you sure you want to delete this meter?")){
	parent.document.frames.meters.location=temp
	parent.document.frames.descriptions.location=temp3}
}

</script>
<%@Language="VBScript"%>

<%
acctid=Request.querystring("acctid")
bldg=Request.querystring("bldg")
utility=Request.querystring("utility")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")


sqlstr= "select * from meters1 where acctid=ltrim('" &acctid& "') and bldgnum=ltrim('" &bldg& "') "

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
if not rst1.EOF  then%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Meter 
        List</i></font></font></div>
    </td>
  </tr>
</table>
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC" width="100%"> 
    <td width="25%"> 
        <font face="Arial, Helvetica, sans-serif" color="#000000">Meter:</font><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"> 
        </font></font></td>
    
    <td width="31%"><font face="Arial, Helvetica, sans-serif" color="#000000">Online: 
      </font> </td>
	  
    <td width="44%"><font face="Arial, Helvetica, sans-serif" color="#000000"> 
      </font> </td>
    </tr>
 </table>
 
<table width="100%" border="0">
  <% do until rst1.eof%>
  <form name="detail" method="post" action="">
    <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:updatemeter('<%=rst1("meterid")%>','<%=utility%>','<%=bldg%>','<%=acctid%>')"> 
      <td width="25%"> <font face="Arial, Helvetica, sans-serif" size="1">
	  <input type="hidden" name="meterid" value="<%=rst1("meterid")%>">
	   <input type="hidden" name="acctid" value="<%=Request.querystring("acctid")%>">
	    <input type="hidden" name="bldg" value="<%=Request.querystring("bldg")%>">
<%=rst1("meternum")%></font></td>
      <td width="31%"> <font face="Arial, Helvetica, sans-serif" color="#000000" size="1"> 
        <%if rst1("online") then %>
        <img src="greencheck.gif" width="13" height="15"> 
        <%end if%>
        </font> </td>
	  <td width="44%"><a href="javascript:deletemeter('<%=rst1("meterid")%>','<%=bldg%>','<%=acctid%>','<%=utility%>')"><img src="delete.gif" border="0"></a></td>
    </tr>
    <%rst1.movenext
loop
rst1.close%>
  </form>
</table>
<%end if
set cnn1=nothing
%>


</html>