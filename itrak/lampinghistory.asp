<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>

<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

fid=request.querystring("fid")

sqlstr= "select l.*,f.* from lamping_sch l join fixtures f on l.fid=f.id where l.fid='"&fid&"' order by datelastchanged" 

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<title>Lamping History</title></head>
<link rel="Stylesheet" href="/genergy2/styles.css">

<body bgcolor="#FFFFFF">

<table width="100%" cellpadding="3" cellspacing="1" border="0">
  <tr bgcolor="336699">
	<td colspan="2"><font face="Arial, Helvetica, sans-serif" color="#ffffff" size="2" class="standard"><b>Lamping History</b></font></td>
</tr>
<% While not rst1.EOF %>
<tr>
	<td width="35%" align="right" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">Fixture Date Changed</font></td>
	<td width="65%" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("datelastchanged")%> </font></td>
</tr>
<tr>
	<td align="right" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">Lamp Quantity</font></td>
	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("lampqty")%></font></td>
</tr>
<tr>
	<td align="right" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">Fixture Electrician</font></td>
	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("electrician")%></font></td>
</tr>
<tr>
	<td align="right" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">Ballast Date Changed</font></td>
	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("bdatelastchanged")%></font></td>
</tr>
<tr>
	<td align="right" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">Ballast Quantity</font></td>
	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("ballastqty")%></font></td>
</tr>
<tr>
	<td align="right" bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">Ballast Electrician</font></td>
	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("belectrician")%></font></td>
</tr>
<tr>
	<td align="right" bgcolor="#eeeeee" style="border-bottom:1px solid #999999;"><font face="Arial, Helvetica, sans-serif" size="2" class="standard">General Comments</font></td>
	<td bgcolor="#eeeeee" style="border-bottom:1px solid #999999;"><font face="Arial, Helvetica, sans-serif" size="2" class="standard"><%=rst1("comments")%> </font></td>
</tr>
  <%
		rst1.movenext
		Wend
		%>
</table>
   
<%

rst1.close
set cnn1=nothing
%>
</html>

