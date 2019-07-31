<html>
<head>
<script>
function goback(bldgnum, meternum, srvname, dbnam){
	document.location="http://yahoo.com"
}
function process(flag){
	var ary=document.forms[0].bldgnum.value.split("|")
	var bldgnum=ary[0]
	var strt=ary[1]
	var year=document.forms[0].year.value
	var period=document.forms[0].period.value
	ary=document.forms[0].leaseid.value.split("|")
	var leaseid=ary[0]
	if(document.forms[0].bldgnum.value==""){
		alert("Please select a building")
		return
	}
    var msg 
	var temp
	var sql
	if(flag=="v"){
		msg="Are you sure you want to Validate data for\n"
	}else{
		msg="Are you sure you want to Inalidate data for\n"
	}
	if(leaseid==""){
		msg=msg+strt+" Period: "+year+"/"+period
		if(flag=="v"){
			sql="update consumption set validate=1 where billyear="+year+" and billperiod="+period+" and meterid in (select meterid from meters where bldgnum='"+bldgnum+"')"
		}else{
			sql="update consumption set validate=0 where billyear="+year+" and billperiod="+period+" and meterid in (select meterid from meters where bldgnum='"+bldgnum+"')"
		}	
//		temp="validate.asp?bldgnum="+bldgnum+"&year="+year+"&period="+period
	}else{
		msg=msg+ary[1]+" Period: "+year+"/"+period
		if(flag=="v"){
			sql="update consumption set validate=1 where billyear="+year+" and billperiod="+period+" and meterid in (select meterid from meters where bldgnum='"+bldgnum+"' and leaseutilityid='"+leaseid+"')"
		}else{
			sql="update consumption set validate=0 where billyear="+year+" and billperiod="+period+" and meterid in (select meterid from meters where bldgnum='"+bldgnum+"' and leaseutilityid='"+leaseid+"')"
		}
//		temp="validate.asp?bldgnum="+bldgnum+"&year="+year+"&period="+period+"&leaseid="+leaseid
	}
	if(confirm(msg)){
		document.location="validate.asp?sql="+sql		
	}
}

function reload(s){
	var ary=s.split("|")
	var bldgnum=ary[0]
	var strt=ary[1]
	document.location="validate.asp?bldgnum="+bldgnum+"&strt="+strt
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
sql="select bldgnum,strt from buildings order by bldgnum"
rst1.Open sql, cnn1, 0, 1, 1
strsql=request("sql")
if strsql <> "" then
	'response.write strsql
	cnn1.execute strsql
end if
bldgnum=request("bldgnum")
meternum=request("meternum")
strt=request("strt")
dbname=request("dbname")
n=now
%>
<body bgcolor="#FFFFFF" text="#000000">

<form name="form2">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr >
      <td height="2" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
        <font color="#FFFFFF">Validation Page</font></b></i></font></td>
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
  <p><font face="Arial, Helvetica, sans-serif"><b><%=strt%></b></font> </p>
  <div align="right">
  <table width="20%">
  <tr>
  <td>Validate<td><input type="button" name="Submit" value="GO" onClick='process("v")'></td>
  </tr>
  <tr>
  <td>Invalidate<td><input type="button" name="Submit" value="GO" onClick='process("i")'></td>
  </tr>
  </table>
  </div>
  <br>
  <table width="100%" border="0">
    <tr>
	  <td width="14%">Building</td>
	  <td width="86%"> 
        <select name="bldgnum" onchange='reload(this.value)'>
	  <%
	  if bldgnum="" then
	  %>
	    <option value="" selected>Select Buildings</option>
	  <%
	  end if
	  if not rst1.eof then
	  do until rst1.eof
	      if trim(rst1("bldgnum"))=bldgnum then
	  %>
	    
	    <option value="<%=trim(rst1("bldgnum"))&"|"&rst1("strt")%>" selected><%=rst1("strt")%>, <%=rst1("bldgnum")%></option>
	  <%
	      else
	  %>
	    <option value="<%=trim(rst1("bldgnum"))&"|"&rst1("strt")%>"><%=rst1("strt")%>, <%=rst1("bldgnum")%></option>
	  <%
	      end if
	  rst1.movenext
	  loop
	  end if
	  %>
	  </select>
	  </td>
	</tr>
	<tr>
	  <td width="14%">Lease ID</td>
	  <td width="86%"> 
        <select name="leaseid">
		<option value="" selected></option>
	  <%
	  if bldgnum <> "" then
	  	  sql="select billingname,tenantnum,leaseutilityid from tblleases,tblleasesutilityprices where tblleasesutilityprices.billingid=tblleases.billingid and tblleases.bldgnum='"&bldgnum&"' order by billingname"
rst2.Open sql, cnn1, 0, 1, 1
          if not rst2.eof then
	  	  do until rst2.eof
	  %>
	    <option value="<%=rst2("leaseutilityid")&"|"&rst2("tenantnum")&" "&rst2("billingname")%>"> <%=rst2("billingname")%>, <%  %><%=rst2("tenantnum")%>
	  <%
	      rst2.movenext
		  loop
		  end if
	  end if
	  %>
	  </select>
	  </td>
	</tr>
	<tr>
	  <td width="14%">Bill Year</td>
	  <td width="86%"> 
      <select name="year">
	  
		<option value="<%=Year(n)%>"><%=Year(n)%></option>
	    <option value="<%=Year(n)+1%>"><%=Year(n)+1%></option>
		<option value="<%=Year(n)+2%>"><%=Year(n)+2%></option>
		<option value="<%=Year(n)+3%>"><%=Year(n)+3%></option>
		<option value="<%=Year(n)+4%>"><%=Year(n)+4%></option>
		<option value="<%=Year(n)+5%>"><%=Year(n)+5%></option>
      </select>
	 </td>
	</tr>
	<tr>
	  <td width="14%">Bill Period</td>
	  <td width="86%"> 
        <select name="period">
	  <option value="1">1</option>
	  <option value="2">2</option>
	  <option value="3">3</option>
	  <option value="4">4</option>
	  <option value="5">5</option>
	  <option value="6">6</option>
	  <option value="7">7</option>
	  <option value="8">8</option>
	  <option value="9">9</option>
	  <option value="10">10</option>
	  <option value="11">11</option>
	  <option value="12">12</option>
	 </select> 
	 </td> 
	  
	  
    </tr>
  </table>
  <p>&nbsp; </p>
</form>  
</body>
</html>
