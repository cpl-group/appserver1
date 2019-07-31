<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>Edit Account Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function updacct(id1,acctid,vendor,name1,addr2,bldg,esco,accounttype,locked){
var temp = "updateacct.asp?id1="+id1+"&acctid="+acctid+"&vendor="+vendor+"&name1="+name1+"&addr2="+addr2+"&bldg="+bldg+"&esco="+esco+"&accounttype="+accounttype+"&locked="+locked
//alert(temp)
	parent.document.frames.entry.location=temp
	parent.document.form1.id1.value=id1
	parent.document.form1.acctid.value=acctid
	parent.document.form1.bldg.value=bldg
	window.close()
}
function meter(acctid,utility,bldg){
	var temp="meterframes.asp?acctid="+acctid+"&utility="+utility+"&bldg="+bldg
	//alert(temp)
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
dim id1
id1=Request.Querystring("acctid")
'response.write id1
'response.end
Dim cnn1, rst1, rst2, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")

if instr(request.servervariables("SCRIPT_NAME"),"/genergy2/")<>0 then cnn1.Open application("cnnstr_genergy2") else cnn1.Open application("cnnstr_genergy1")

sqlstr= "select * from tblacctsetup where acctid='" &id1&"'"
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
%>

<body bgcolor="#FFFFFF">
<form name="form1" method="post" >

      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC">
		  
      <td width="7%" height="20"> 
        <input type="hidden" name="id1" value="<%=rst1("id")%>">
           <input type="hidden" name="acctid" value="<%=rst1("acctid")%>">
          </td> 
		  
      <td bgcolor="#CCCCCC" width="7%" height="20"><font face="Arial, Helvetica, sans-serif" color="#000000">AccountID</font></td>   	
          
      <td bgcolor="#CCCCCC" width="12%" height="20"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor</font></td>
          
      <td width="17%" height="20"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor 
        name</font></td>
		  
      <td width="19%" height="20"><font face="Arial, Helvetica, sans-serif" color="#000000">Service 
        Address</font></td>
			
      <td width="18%" height="20"><font face="Arial, Helvetica, sans-serif" color="#000000">Utility</font></td>
			
      <td width="20%" height="20"><font face="Arial, Helvetica, sans-serif" color="#000000">Building</font></td>
        </tr>
        
        <tr>
		  
      <td width="7%"> 
        <input type="button" name="upd" value="UPDATE" onClick="updacct(id1.value,acctid.value,vendor.value,name1.value,addr2.value,bldg.value,escoRef.value,accounttype.value,locked.value)">
      </td> 
          
        
      <td width="7%"><font face="Arial, Helvetica, sans-serif"><nobr><%=rst1("acctid")%></nobr></font> 
      </td>
        
      <td width="12%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="vendor" value=<%=rst1("vendor")%>>
          </font></td>
        
      <td width="17%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="name1" value=<%=rst1("vendorname")%>>
          </font></td>
        
      <td width="19%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="addr2" value=<%=rst1("serviceaddr")%>>
          </font></td>
        
      <td width="18%"><font face="Arial, Helvetica, sans-serif"> 
	  <input type="hidden" name="utility" value=" <%=rst1("utility")%>">
	  <%=rst1("utility")%> 
        </font></td>
        
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
	  <input type="hidden" name="bldg" value=" <%=rst1("bldgnum")%>">
	  <%=rst1("bldgnum")%> 
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
			<td><select name="accounttype" onchange="checkescoBox(this)"><option value="0">T&nbsp;&amp;&nbsp;D</option><option value="1"<%if rst1("Esco") then response.write " SELECTED"%>>Esco</option></select></td>
			<td>
			<select id="escoRef" name="escoRef" style="background-color:#white;">
			<option value="0">none</option>
			<%
			  dim str2
			  str2="SELECT * FROM tblacctsetup where Esco=1 and bldgnum='"&trim(rst1("bldgnum"))&"'"
			  rst2.Open str2, cnn1, 0, 1, 1
			  do until rst2.eof
			  	response.write "<option value="""&rst2("AcctID")&""""
				if trim(rst2("AcctID"))=trim(rst1("Escoref")) then response.write " SELECTED"
				response.write ">"&rst2("Vendorname")&"</option>"
'				response.write "<BR>"&trim(rst2("AcctID"))&"="&trim(rst1("Escoref"))
				rst2.movenext
			  loop
			  
			  rst2.close
			%>
			</select>
			</td>   	
			<td><input type="checkbox" name="locked" value="1" <%if rst1("locked")="True" then response.write " CHECKED"%>></td>
		</tr>

      </table>
	  
  
    <%
rst1.close
set cnn1=nothing
%></form>
<i><font face="Arial, Helvetica, sans-serif">*If this Account should be removed, 
please contact <a href ="mailto:george_nemeth@genergy.com">George Nemeth</a></font></i> 

<script>
checkescoBox(document.forms['form1'].accounttype)
</script>
</html>  