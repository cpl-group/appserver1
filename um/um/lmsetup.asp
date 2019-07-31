<html>
<head>
<script>
function goback(bldgnum, meternum, srvname, dbnam){
	document.location="http://yahoo.com"
}
function showMeter(){
	var srvname=document.forms[0].srvname.value
	var bldgnum=document.forms[0].bldgnum.value
	var meternum=document.forms[0].meternum.value
	var dbname=document.forms[0].dbname.value
	document.meter.location="lminfo.asp?bldgnum="+bldgnum+"&meternum="+meternum+"&srvname="+srvname+"&dbname="+dbname
	//alert(document.forms[0].srvname.value)
	//alert(meternum)
}

function reload(temp){
	var Ary=temp.split("|")
	bldgnum=Ary[0]
	srvname=Ary[1]
	dbname=Ary[2]
    document.location="lmsetup.asp?bldgnum="+bldgnum+"&srvname="+srvname+"&dbname="+dbname
}
</script>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
sql="select strt, master.dbo.rm.bldgnum as bldgnum,srvname,dbname from buildings, master.dbo.rm where master.dbo.rm.bldgnum=buildings.bldgnum and lm=1"
rst1.Open sql, cnn1, 0, 1, 1
bldgnum=request("bldgnum")
meternum=request("meternum")
srvname=request("srvname")
dbname=request("dbname")
%>
<body bgcolor="#FFFFFF" text="#000000">

<form name="form2">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr >
      <td height="2" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
        <font color="#FFFFFF">Meter LM Setup</font></b></i></font></td>
    </tr>
  </table>
  <%
  if rst1.eof then
  %> 
  <table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i>
		No Building Available</i></font></p>
      </div>
    </td>
  </tr>
  </table>
  <%
  end if
  %>
  <table width="100%" border="0">
    <tr> 
	  <td width="21%"> 
	    
        <select name="temp" onchange='reload(this.value)'>
	  <%
	  if bldgnum="" then
	  %>
	    <option selected>===================</option>
	  <%
	  end if
	  if not rst1.eof then
	  do until rst1.eof
	      temp=trim(rst1("bldgnum"))&"|"&trim(rst1("srvname"))&"|"&trim(rst1("dbname"))
		  'response.write temp
	  	  if trim(rst1("bldgnum"))=bldgnum then
	  %>
	    
	    <option value="<%=temp%>" selected><%=rst1("strt")%></option>
	  <%
	      else
	  %>
	    <option value="<%=temp%>"><%=rst1("strt")%></option>
	  <%
	      end if
	  rst1.movenext
	  loop
	  end if
	  %>
	  </select>
	  </td>		
      
      <%
	  
	  if bldgnum="" then
	  else
	  %>
	    <td width="79%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        Please enter the Meter: 
		
        <input type="text" name="meternum">
		<input type="hidden" name="srvname" value="<%=srvname%>">
		<input type="hidden" name="dbname" value="<%=dbname%>">
		<input type="hidden" name="bldgnum" value="<%=bldgnum%>">
        <input type="button" name="submit3" value="GO" onclick="showMeter()">
		
        <input type="button" name="submit2" value="BACK" onClick='javascript:history.back()'>
        </font>
		</td>
	  <%
	  end if
	  %>	
    </tr>
  </table>
  <p>&nbsp; </p>
</form>  
<IFRAME name="meter" width="100%" height="70%" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></iframe>
<IFRAME name="detail" width="100%" height="30%" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
</body>
</html>
