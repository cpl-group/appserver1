<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, groupname, lmp, isportfolio, grouplabel, groupid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
groupname = secureRequest("groupname")
groupid = 0
if trim(bldg)="" or trim(bldg)="0" then isportfolio = 1 else isportfolio = 0

if trim(bldg)="" then lmp = 1 else lmp = 0
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getMainConnect(pid)

dim metersingroup, tenantsingroup, grouptype
grouptype = 0
metersingroup = "|"
tenantsingroup = "|"
if trim(groupname)<>"" then
	rst1.open "SELECT distinct typeid as meterid FROM [group] g, groupitems gi WHERE groupname='"&groupname&"' and g.id=gi.groupid and typecode='m'", cnn1
	do until rst1.EOF
		metersInGroup = metersInGroup & rst1("meterid") & "|"
		rst1.movenext
	loop
	rst1.close
	rst1.open "SELECT gt.id as typeid, g.grouplabel, g.id as groupid FROM [group] g INNER JOIN grouptype gt ON gt.id=g.type WHERE g.groupname='"&groupname&"'"
	if not rst1.eof then
    grouptype = cint(rst1("typeid"))
    grouplabel = rst1("grouplabel")
    groupid = rst1("groupid")
  end if
	rst1.close
end if

'where build
dim where
if trim(bldg)<>"" and trim(bldg)<>"0" then where = where&" and b.bldgnum='"&trim(bldg)&"'"
if trim(tid)<>"" and trim(tid)<>"0" then where = where&" and l.billingid='"&trim(tid)&"'"

dim bldgname, portfolioname, breadcrumbtrail
if trim(bldg)<>"" then 
	rst1.Open "SELECT * FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", getConnect(0,bldg,"billing")
  if not rst1.eof then
  	bldgname = rst1("bldgname")
	  portfolioname = rst1("name")
  end if
	rst1.close
	breadcrumbtrail = "<a href='portfolioedit.asp?pid=" & pid & "' style='color:#ffffff;'>" & portfolioname & "</a> &gt; <a href='buildingedit.asp?pid=" & pid & "&bldg=" & bldg & "' style='color:#ffffff'>" & bldgname & "</a>"
else
  rst1.open "select name from portfolio p where id='" & pid & "'", cnn1
  portfolioname = rst1("name")
	breadcrumbtrail = "<a href=""portfolioedit.asp?pid=" & pid & """ style=""color:#ffffff;"">" & portfolioname & "</a>"
  rst1.close
end if

dim hasInvoice
hasInvoice = true
rst1.open "SELECT * FROM [group] g INNER JOIN grouptype gt ON gt.id=g.type WHERE g.bldgnum='"&bldg&"' and g.type='2'"
if rst1.eof then hasInvoice = false
rst1.close
%>
<html>
<head>
<title>Portfolio View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function flipList(grouptype)
{	document.all['meterlist'].style.display='none';
  document.all['tenantlist'].style.display='none';
  document.all['plantAnalysis'].style.display='none';
  if(grouptype==5)
  { document.all['plantAnalysis'].style.display='inline'
  }else if(grouptype==2)
	{	document.all['tenantlist'].style.display='inline'
	}else
	{	document.all['meterlist'].style.display='inline'
	}
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
<style type="text/css">
li { list-style-type:none;line-height:14pt; }
</style>
</head>
<body>
<FORM action="groupSave.asp" method="get">
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr>
  <td bgcolor="#3399CC">
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td><span class="standardheader">Group Edit | <span style="font-weight:normal;"><%=breadcrumbtrail%></span></span></td>
    <td align="right"><button id="qmark2" onclick="openCustomWin('help.asp?page=groupedit','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
Group type: 
<select name="type" onchange="flipList(this.value)">
<%
rst1.open "SELECT * FROM grouptype WHERE id<>3 AND id<>4 ORDER BY type", getConnect(0,0,"dbCore")
do until rst1.eof
	if ((cint(rst1("id"))<>2 or not(hasInvoice)) and grouptype=0) or (cint(grouptype) = cint(rst1("id"))) then
    if (cint(rst1("id"))=2 and allowGroups("IT Services"))  or cint(rst1("id"))<>2  or allowGroups("gReadingandBilling") then
  		%><option value="<%=rst1("id")%>"<%if cint(grouptype) = cint(rst1("id")) or (grouptype = 0 and cint(rst1("id")) = 1)  then Response.Write " SELECTED"%>><%=rst1("type")%></option><%
    end if
	end if
	rst1.MoveNext
loop
rst1.close
%>
</select>

Name:
<input type="hidden" name="grouptype" value="<%=cint(grouptype)%>">
<input name="name" type="hidden" value="<%=groupname%>">
<input name="groupid" type="hidden" value="<%=groupid%>">
<input name="label" type="text" value="<%=grouplabel%>">
</td>
</tr>
</table>
<br>
<div style="padding:3px"><ul>
<div id="meterlist" style="display: <%if cint(grouptype)=1 or trim(grouptype)=0 or (hasInvoice and trim(grouptype)="") then response.write "inline" else response.write "none"%>;">
<%
dim cbldg, ctenant
cbldg = ""
ctenant = ""
dim sqlthing
sqlthing = "SELECT * FROM meters m INNER JOIN buildings b ON b.bldgnum=m.bldgnum INNER JOIN sysobjects s ON s.name=m.datasource INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=m.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid INNER JOIN super_main sm ON sm.bldgnum=m.bldgnum WHERE m.lmp="&lmp&" and portfolioid='"&pid&"' "&where&" ORDER BY b.bldgnum, tname, meternum"

rst1.open sqlthing, cnn1

if rst1.eof then response.write "No meters with interval data are available."
do until rst1.eof
	if cbldg <> trim(rst1("BldgName")) then
		cbldg = trim(rst1("BldgName"))
		%>
		<br><br><B><%=cbldg%></B>
		<%
	end if
	if ctenant <> trim(rst1("tName")) then
		ctenant = trim(rst1("tName"))
		%>
		<br><%=ctenant%><br>
		<%
	end if
	%><li><input type="Checkbox" value="<%=rst1("meterid")%>" name="meters" <%if instr(metersingroup, "|"&trim(rst1("meterid"))&"|") then Response.Write " CHECKED"%>>&nbsp;<%=rst1("meternum")%><%
	rst1.MoveNext%>
<%
loop
rst1.close
%>
</div>
<div id="tenantlist" style="display: <%if cint(grouptype)=2 then response.write "inline" else response.write "none"%>;">
<%
strsql = "SELECT distinct tenantnum, billingname, strt FROM buildings b INNER JOIN tblleases l ON b.bldgnum=l.bldgnum INNER JOIN tblleasesutilityprices lup ON lup.billingid=l.billingid WHERE "
if trim(tid)<>"" and trim(tid)<>"0" then
  strsql = strsql & "l.billingid='"&tid&"' ORDER BY tenantnum"
else
  strsql = strsql & "l.bldgnum='"&bldg&"' ORDER BY tenantnum"
end if
rst1.open strsql, cnn1
'response.write strsql
'response.end
%>
&nbsp;<br><br><b>Accounts <%if not rst1.eof then response.write " - "&rst1("strt")%></b>
<%
do until rst1.eof%>
	<li><!-- <input type="Checkbox" value="''<%=rst1("tenantnum")%>''" name="tenants" <%if instr(tenantsingroup, "|"&trim(rst1("tenantnum"))&"|") then Response.Write " CHECKED"%>>&nbsp;<%=rst1("tenantnum")%> -->&nbsp;&nbsp;&nbsp;&nbsp;<%=rst1("billingname")%><%
	rst1.MoveNext
loop
rst1.close
%><!--  -->
<input name="tenants" type="hidden" value="''00000''">
</div>
</ul>
<div id="plantAnalysis" style="display: <%if cint(grouptype)=5 then response.write "inline" else response.write "none"%>;">
<%
if isportfolio<>"1" then
  dim inOutList
  rst1.open "SELECT * FROM [group] WHERE clientid="&pid&" AND bldgnum='"&bldg&"' AND type=1 AND id in (SELECT groupid FROM groupitems)", cnn1
  do until rst1.eof
    inOutList = inOutList&";"&rst1("id")&"|"&rst1("grouplabel")
    rst1.movenext
  loop
  inOutList = split(mid(inOutList,2),";")
  rst1.close
  
  dim hasOutGroups, hasInGroups
  hasOutGroups = ""
  hasInGroups = ""
  if cint(grouptype)=5 then
    rst1.open "SELECT * FROM [group] WHERE id in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid and typecode='o' and g.groupname='"&groupname&"')", cnn1
    do until rst1.eof
      hasOutGroups = hasOutGroups &"|"&rst1("groupname")&"|"
      rst1.movenext
    loop
    rst1.close
    rst1.open "SELECT * FROM [group] WHERE id in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid and typecode='i' and g.groupname='"&groupname&"')", cnn1
    do until rst1.eof
      hasInGroups = hasInGroups &"|"&rst1("groupname")&"|"
      rst1.movenext
    loop
    rst1.close
  end if
  %>
  <table>
  <tr><td>Input Groups</td>
      <td>Output Groups</td>
  </tr>
  <tr><td>
        <%
        dim inOutGroup
        for each inOutGroup in inOutList%>
        	<li><input type="Checkbox" value="<%=(split(inOutGroup,"|"))(0)%>" name="inputGroups" <%if instr(hasInGroups, "|"&trim(split(inOutGroup,"|")(1))&"|") then Response.Write " CHECKED"%>>&nbsp;<%=(split(inOutGroup,"|"))(1)%>&nbsp;<%
        next
        %>
      </td>
      <td>
        <%
        for each inOutGroup in inOutList%>
        	<li><input type="Checkbox" value="<%=(split(inOutGroup,"|"))(0)%>" name="outputGroups" <%if instr(hasOutGroups, "|"&trim(split(inOutGroup,"|")(1))&"|") then Response.Write " CHECKED"%>>&nbsp;<%=(split(inOutGroup,"|"))(1)%>&nbsp;<%
        next
        %>
      </td>
  </tr>
  </table>
<%end if%>
</div>

<br>
<input type="hidden" name="isportfolio" value="<%=isportfolio%>">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="edit" value="<%if trim(groupname)<>"" then response.write "1"%>">
<%if grouptype <> 2 and not(isBuildingOff(bldg)) then%><input type="submit" name="action" value="Save"  class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"><%end if%>
<input type="button" value="Cancel" onclick="document.location='groupview.asp?pid=<%=pid%>&bldg=<%=bldg%>';" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
</div>
</FORM>
</body>
</html>
