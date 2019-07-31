
<html>
<head>
<%@Language="VBScript"%>
<script>
function modify(pid,bldgnum){
	var	temp = "buildingtc.asp?action=edit&pid="+pid+"&bldgnum="+bldgnum;
	parent.document.frames.admin.location=temp
}
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%
pid=Request.QueryString("pid")
editbldgnum=Request.Querystring("bldgnum")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

 
	sql = "select bldgnum, strt,readgroup from buildings where portfolioid = '" & pid& "'"
rst1.Open sql, cnn1, 0, 1, 1

if rst1.eof then
%>
<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><b>Portfolio Not Found</b></font></p>
      </div>
    </td>
  </tr>
</table>
<%
else
if isempty(editbldgnum) then 
count=0
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>BUILDING 
        LIST</b></font></div>
    </td>
  </tr>
</table>
<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="4%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp</font><font size="2"></font></td>
    <td width="13%" height="2"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">ID</font></div>
    </td>
    <td width="75%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">Building</font></td>
    <td width="8%" height="10"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Trip 
        Code </font></div>
    </td>
  </tr>
  <% 
	do until rst1.EOF 
    count=count+1
  %>
  <form name="form1" method="post" action="">
    <tr> 
      <td width=4%> <font size="2"> 
        <input type="hidden" name="pid" value="<%=pid%>">
        <input type="hidden" name="bldgnum" value="<%=rst1("bldgnum")%>">
        <input type="button" name="submit" value="edit" onclick=modify(pid.value,bldgnum.value)>
        </font></td>
      <td width=13%> 
        <div align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rst1("bldgnum")%></font></div>
      </td>
      <td width=75%> <font size="2"> <%=rst1("strt")%> </font></td>
      <td width="8%" height="19"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
          <%=rst1("readgroup")%> </font></div>
      </td>
    </tr>
  </form>
  <%
	rst1.movenext
	loop
%>
</table>
<table width="100%" border="0" bgcolor="#3399CC">
  <tr> 
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=count%> Meter(s) found</font></b></div>
    </td>
  </tr>
</table>

<%
else
count=0
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>BUILDING 
        LIST</b></font></div>
    </td>
  </tr>
</table>
<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="8%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp</font><font size="2"></font></td>
    <td width="4%" height="2"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">ID</font></div>
    </td>
    <td width="84%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">Building</font></td>
    <td width="4%" height="10"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Trip 
        Code </font></div>
    </td>
  </tr>
  <% 
	do until rst1.EOF 
    count=count+1
	if rst1("bldgnum") <> editbldgnum then
  %>
  <form name="form1" method="post" action="">
    <tr> 
      <td width=8%> <font size="2"> 
        <input type="hidden" name="bldgnum" value="<%=rst1("bldgnum")%>">
        <input type="hidden" name="pid" value="<%=pid%>">
        <input type="button" name="submit2" value="edit" onClick=modify(pid.value,bldgnum.value)>
        </font></td>
      <td width=4%> 
        <div align="center"><font size="2"><%=rst1("bldgnum")%></font><font face="Arial, Helvetica, sans-serif"></font></div>
      </td>
      <td width=84%> <font size="2"> <%=rst1("strt")%> </font></td>
      <td width="4%" height="19"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
          <%=rst1("readgroup")%> </font></div>
      </td>
    </tr>
  </form>
  <%else %>
  <form name="form1" method="post" action="savetc.asp">
    <tr> 
      <td width=8%> <font size="2"> 
        <input type="hidden" name="pid" value="<%=pid%>">
        <input type="hidden" name="bldgnum" value="<%=rst1("bldgnum")%>">
        <input type="submit" name="button" value="save">
        </font></td>
      <td width=4%> 
        <div align="center"><font size="2"><%=rst1("bldgnum")%></font><font face="Arial, Helvetica, sans-serif"></font></div>
      </td>
      <td width=84%> <font size="2"> <%=rst1("strt")%> </font></td>
      <td width="4%" height="19"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
          <input type="text" name="readcode" value="<%=rst1("readgroup")%>">
          </font></div>
      </td>
    </tr>
  </form>
  <%
  	end if
	rst1.movenext
	loop
%>
</table>
<table width="100%" border="0" bgcolor="#3399CC">
  <tr> 
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=count%> Meter(s) found</font></b></div>
    </td>
  </tr>
</table>
<%
end if
end if
rst1.close
set cnn1=nothing
%>
</body>
</html>
