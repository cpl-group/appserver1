<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim pid, bldg, count, edit
pid = Request.QueryString("pid")
bldg = Request.Querystring("bldg")

Dim cnn1, rst1, rst2, sql
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(pid,bldg,"billing")
%>
<html>
<head>
<title></title>
<script>
function modify(bldg)
{ var	temp = "buildingtc.asp?pid=<%=pid%>&bldg="+bldg;
  parent.document.frames.admin.location=temp
}
</script>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
<body>
<form name="form1" method="post" action="savetc.asp">
<%
sql = "SELECT bldgnum, strt, readgroup FROM buildings b WHERE portfolioid="&pid&" ORDER BY strt"
rst1.Open sql, cnn1
if rst1.eof then
%><center><b>Portfolio Not Found</b></center><%
else
count=0
%>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#6699CC" class="standardheader" align="center"><b>Buidling List</b></td></tr>
</table>
<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
<tr valign="bottom" bgcolor="#dddddd" style="font-weight:bold;">
  <td width="4%">&nbsp;</td>
  <td width="13%" align="center">ID</td>
  <td width="75%">Building</td>
  <td width="8%" align="center">Trip Code</td>
</tr>
<% 
do until rst1.EOF 
count=count+1
if trim(bldg)=trim(rst1("bldgnum")) then edit=true else edit=false
%>
<tr bgcolor="#ffffff">
  <td>
    <%if edit then%><input type="submit" name="button" value="save"><%else%><input type="button" name="submit" value="edit" onclick="modify('<%=rst1("bldgnum")%>')"><%end if%>
  </td>
  <td align="center"><%=rst1("bldgnum")%></td>
  <td><%=rst1("strt")%></td>
  <td align="center">
    <%if edit then%><input type="text" name="readcode" value="<%=rst1("readgroup")%>"><%else%><%=rst1("readgroup")%><%end if%>
  </td>
</tr>
<%
rst1.movenext
loop
%>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr><td bgcolor="#6699CC" class="standardheader" align="center"><b><%=count%> Meter(s) found</b></td></tr>
</table>
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="pid" value="<%=pid%>">
</form>
<%
end if
rst1.close
%>
</body>
</html>
