<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldgnum, utility, pid
bldgnum=Request.Querystring("building")
utility=request.Querystring("utility")
'response.write bldgnum
Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(bldgnum)
rst1.open "SELECT portfolioid FROM buildings WHERE bldgnum='"&bldgnum&"'", cnn1
if not rst1.eof then pid = rst1("portfolioid")
rst1.close
dim DBMainmodIP
DBMainmodIP = "["&getPidIP(pid)&"].Supermod.dbo."

sqlstr= "select * from tblacctsetup where bldgnum='"&bldgnum&"' and utility='"&utility&"'"
'response.write sqlstr
'response.end
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<html>
<head>
<title>Utility Account List</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
</head>
<script>
function editacct(acctid){
	var temp = "editacct.asp?acctid="+acctid+"&building=<%=bldgnum%>"
	window.opener.document.all['entryframe'].style.visibility = "visible"
	window.opener.document.getElementById('entry').src=temp;
	window.close()
}
function selectacct(id1,acctid,bldg,esco){
	window.opener.document.form1.id1.value=id1
	window.opener.document.form1.acctid.value=acctid
	window.opener.document.all['accountdisplay'].innerText=acctid
	window.opener.document.form1.bldg.value=bldg
	if(esco=='True')
	{	window.opener.document.all['enterbillbutton'].style.visibility='hidden'
	}else
	{	window.opener.document.all['enterbillbutton'].style.visibility='visible'
	}
	window.close()
}
</script>
<body bgcolor="#eeeeee">


<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">Account Information </font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td>
<% if rst1.eof then%>
<table width="100%" border="0">
<tr>
          <td> 
            <div align="center"><i><font face="Arial, Helvetica, sans-serif">No 
              accounts are currently setup for this building</font></i> </div>
          </td>
<tr>
</table>
<%else%>
      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC">
		  <td width="5%"></td> 
		  <td width="7%"></td> 
		  <td bgcolor="#CCCCCC" width="15%"><font face="Arial, Helvetica, sans-serif" color="#000000">AccountID</font></td>   	
          <td bgcolor="#CCCCCC" width="27%"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor</font></td>
          <td width="46%"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor&nbsp;name</font></td>
          <td width="46%"><font face="Arial, Helvetica, sans-serif" color="#000000">Other&nbsp;Accounts</font></td>
          <td width="46%"><font face="Arial, Helvetica, sans-serif" color="#000000">Account&nbsp;Type</font></td>
        </tr>
        <% While not rst1.EOF %>
        	 <form name="list1" method="post" action="entry.asp"><tr>
		  <td width="12%">
				<input type="hidden" name="bldg" value="<%=bldgnum%>">
				<input type="hidden" name="utility" value="<%=utility%>">
				<input type="hidden" name="acctid" value="<%=rst1("acctid")%>">
				<input type="hidden" name="id1" value="<%=rst1("id")%>">
				<%if not rst1("esco") then%><input type="button" name="Submit" value="SELECT"  onClick="selectacct(id1.value,acctid.value,bldg.value,'<%=rst1("esco")%>')"><%end if%>
            </td>
			
          <td width="7%"><input type="button" value="<%if isBuildingOff(bldgnum) then%>View<%else%>Edit<%end if%>" name="Edit" onClick="editacct('<%=rst1("acctid")%>');selectacct(id1.value,acctid.value,bldg.value,'<%=rst1("esco")%>')"></td>
          <td width="5%"><font face="Arial, Helvetica, sans-serif"><nobr><%=rst1("acctid")%></nobr></font></td>
          <td width="27%"><font face="Arial, Helvetica, sans-serif"><%=rst1("vendor")%></font></td>
          <td width="46%"><font face="Arial, Helvetica, sans-serif"><%=rst1("vendorname")%></font></td>
          <td width="46%"><font face="Arial, Helvetica, sans-serif"><%
		  if trim(rst1("Escoref"))="0" then
			  response.write "none"
		  else
			  response.write rst1("Escoref")
		  end if
		  %></font></td>
          <td width="46%"><font face="Arial, Helvetica, sans-serif"><%
		  if rst1("Esco")=true then 
		  	response.write "Esco"
		  else 
		  	response.write "T&nbsp;&amp;&nbsp;D"
		  end if
		  %></font></td>
		 
        </tr></form>
		<%
		rst1.movenext
		Wend
		%>
      </table>
    </td>
  </tr>
</table>
<%
end if
rst1.close
set cnn1=nothing
%>


</body>
</html>