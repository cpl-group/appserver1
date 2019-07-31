<%@Language="VBScript"%>
<html>
<head>
<title>Edit Account Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function saveacc(acctnum,vendor,vname,addr,utility,bldg,accounttype, escoRef, locked){
	var temp = "saveacct.asp?acctnum="+acctnum+"&vendor="+vendor+"&vname="+vname+"&addr="+addr+"&utility="+utility+"&bldg="+bldg+"&accounttype="+accounttype+"&escoRef="+escoRef+"&locked="+locked
	parent.document.form1.acctid.value=acctnum
	parent.document.frames.entry.location=temp
	
}

function checkescoBox(option)
{	if(option.value==0)
	{	document.forms['form1'].escoRef.disabled=false;
		document.all['escoRef'].style.backgroundColor='#FFFFFF';
	}else
	{	document.forms['form1'].escoRef.disabled=true;
		document.all['escoRef'].style.backgroundColor='#CCCCCC';
		document.all['escoRef'].selectedIndex=0
	}
}
</script>
<%
utility=request.querystring("utility")
bldg=request.querystring("building")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy2")

sqlstr= "select * from buildings where bldgnum='"&bldg&"'"
	
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1


%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">Account Information - <%=rst1("strt")%>
</font></b></font></div>
    </td>
  </tr>
  </table>
      <form name="form1" method="get" action="">
		<input type="hidden" name="bldg" value="<%=bldg%>">
	   <table width="100%" border="0">
        <tr bgcolor="#CCCCCC">
		    <td width="5%">
           
          </td> 
		    <td bgcolor="#CCCCCC" width="12%"><font face="Arial, Helvetica, sans-serif" color="#000000">AccountID</font></td>   	
            <td bgcolor="#CCCCCC" width="11%"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor</font></td>
            <td width="11%"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor 
              name</font></td>
		    <td width="11%"><font face="Arial, Helvetica, sans-serif" color="#000000">Service 
              Address</font></td>
			
            <td width="11%"><font face="Arial, Helvetica, sans-serif" color="#000000">Utility</font></td>
			
        </tr>
        
        <tr>
		    
      <td width="5%"> 
        <input type="button" name="Button" value="SAVE" onclick="saveacc(acctnum.value,vendor.value,vname.value,addr.value,utility.value,bldg.value,accounttype.value, escoRef.value, locked.value)">
      </td> 
            <td width="12%"><font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="acctnum" >
              </font> </td>  
            <td width="11%"><font face="Arial, Helvetica, sans-serif">
              <input type="text" name="vendor" >
              </font></td>  
            <td width="11%"><font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="vname" >
              </font></td>  
            <td width="11%"><font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="addr">
              </font></td>  
            <td width="11%"><font face="Arial, Helvetica, sans-serif"> 
              <select name="utility">
          <% 
			  Set rst2 = Server.CreateObject("ADODB.recordset")
			  str2="select * from tblutility order by utilitydisplay "
			  rst2.Open str2, cnn1, 0, 1, 1
			  do until rst2.eof
			  if rst2("utility")=utility then
			  %>
			 
          <option value="<%=rst2("utility")%>"selected><%=utility%></option>
		  <%else%>
		   <option value="<%=rst2("utility")%>"><%=rst2("utilitydisplay")%></option>
          <%
			  	end if  
				  rst2.movenext
				  loop
				  rst2.close
			  %>
        </select>
              </font></td>  
            <td width="32%"><font face="Arial, Helvetica, sans-serif">
              </font></td>  
            
        </tr>
		<tr>
			<td></td>
			<td bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif" color="#000000">Account&nbsp;Type</font></td>   	
			<td bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif" color="#000000">Esco&nbsp;Reference</font></td>   	
			<td bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif" color="#000000">Locked&nbsp;Account</font></td>   	
		</tr>
		<tr>
			<td></td>
			<td><select name="accounttype" onchange="checkescoBox(this)"><option value="0">T&nbsp;&amp;&nbsp;D</option><option value="1">Esco</option></select></td>
			<td>
			<select id="escoRef" name="escoRef" style="background-color:#white;">
			<option value="0">none</option>
			<%
			  str2="SELECT * FROM tblacctsetup where Esco=1 and bldgnum='"&bldg&"'"
			  rst2.Open str2, cnn1, 0, 1, 1
			  do until rst2.eof
			  	response.write "<option value="""&rst2("AcctID")&""""
				'if rst2("AcctID")=rst1("Escoref") then response.write " SELECTED"
				response.write ">"&rst2("Vendorname")&"</option>"
				rst2.movenext
			  loop
			  
			  rst2.close
			%>
			</select>
			</td>   	
			<td><input type="checkbox" name="locked" value="1"></td>
		</tr>
      </table>
	  </form>
	  
<%rst1.close

	  set cnn1=nothing%>
	  </body>
</html>
