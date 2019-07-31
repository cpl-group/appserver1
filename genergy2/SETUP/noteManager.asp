<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, action, tid, lid, meterid, note, title, id, table, label, sql, sqldisplay
id = request("id")
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
meterid = request("meterid")
action = request("action")
title = request("title")
note = replace(request("note"),"'","''")
if trim(lid)<>"" then
	table = "tenantNotes"
	label = "Lease Utility"
	sql = "INSERT INTO ["&table&"] (bldgid, leaseutilityid, note, [user]) VALUES ('"&bldg&"', '"&lid&"', '"&note&"', '"&getXMLUserName()&"')"
	sqldisplay = "SELECT * FROM "&table&" t WHERE bldgid='"&bldg&"' and leaseutilityid='"&lid&"' ORDER BY date desc"
elseif trim(tid)<>"" then
	table = "accountNotes"
	label = "Tenant Notes"
	sql = "INSERT INTO ["&table&"] (billingid, note, [user], bldgid) VALUES ('"&tid&"', '"&note&"', '"&getXMLUserName()&"','"&bldg&"')"
	sqldisplay = "SELECT * FROM "&table&" t WHERE billingid='"&tid&"' ORDER BY date desc"
elseif trim(bldg)<>"" then
	table = "bldgNotes"
	label = "Building"
	sql = "INSERT INTO ["&table&"] (bldgid, note, [user]) VALUES ('"&bldg&"', '"&note&"', '"&getXMLUserName()&"')"
	sqldisplay = "SELECT * FROM "&table&" t WHERE bldgid='"&bldg&"' ORDER BY date desc"
else
	response.write "Error"
end if

dim cnn1, rst1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim credit, ctype
action = request("action")
credit = request("credit")
ctype = request("ctype")
if not(isnumeric(credit)) then credit = 0
if trim(action)="Add Note" then cnn1.execute sql
  
dim bldgname, portfolioname, rid
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name, region FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
		rid = rst1("region")
	end if
	rst1.close
end if

dim billingname
if trim(bldg)<>"" then
  rst1.open "SELECT billingname FROM tblleases WHERE billingid='"&tid&"'", cnn1
	if not rst1.EOF then
		billingname = rst1("billingname")
	end if
	rst1.close
end if

%>
<html>
<head>
<title>Notes</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="noteManager.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc"><td>
	<table border=0 cellpadding="0" cellspacing="0" width="100%">
	<tr><td><span class="standardheader"><%=label%> Notes | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%'=portfolioname%></a> &gt; <%=bldgname%> &gt; <%=billingname%></span></span></td></tr>
	</table></td>
</tr>
</table>
<%if trim(id)<>"" Then

rst1.open "SELECT * FROM "&table&" t WHERE id="&id, cnn1
if not rst1.eof then
%>
<b>Date:</b> <%=rst1("date")%><br>
<b>Note:</b> <%=rst1("note")%><br>
<div align="right"><a href="javascript:window.close()">Close Window</a></div>
<%
end if
rst1.close%>

<%else%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="30%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Date</td>
    <td width="50%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Note</td>
    <td width="15%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">uid</td>
    <td></td>
</tr>
</table>
<div style="width:100%; overflow:auto; height:100;border-bottom:1px solid #cccccc;">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
'response.write sqldisplay
'response.end
rst1.open sqldisplay, cnn1
do until rst1.eof
if len(rst1("note"))>26 then note = left(rst1("note"), 25)&"..." else note = rst1("note")

%>
<tr bgcolor="#cccccc" valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = '#cccccc'" onClick="javascript:open('noteManager.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid=<%=meterid%>&title=<%=title%>&id=<%=rst1("id")%>','','width=200, height=200, scrollbars=yes')" > 
  <td width="30%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst1("date")%></td>
  <td width="50%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=note%></td>
  <td width="15%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst1("user")%></td>
</tr>
<%
rst1.movenext
loop
%>
</table>
</div>
<table border="0" cellspacing="0" cellpadding="0">
<tr><td><textarea cols="39" rows="3" name="note"></textarea></td><td><input name="action" value="Add Note" type="submit"></td></tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="meterid" value="<%=meterid%>">
<%end if%>
</form>
</body>
</html>