<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
subnet = request("subnet")
userid = request("userid")

Set cnn1 	= Server.CreateObject("ADODB.Connection")
Set rst1 	= Server.CreateObject("ADODB.recordset")
Set rs 		= Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")

if subnet <> "" and userid <> "" then 
	strsql = "select distinct ip,ipname ,userid, systemid, id from ipindex where ip like '"&subnet&"%' and userid = '"&userid&"'order by ip"
elseif subnet <> "" then 
	strsql = "select distinct ip, ipname ,userid, systemid,id from ipindex where ip like '"&subnet&"%' order by ip"
elseif userid <> "" then 
	strsql = "select distinct ip, ipname ,userid, systemid,id from ipindex where userid = '"&userid&"' order by ip"
else
	strsql = "select distinct ip, ipname ,userid, systemid,id from ipindex order by ip"
end if

rst1.Open strsql, cnn1, 0, 1, 1
%>
<html>
<head>
<title></title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #cccccc; }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><head>
<script>
function updatesystem(key){
	parent.document.frames.sn_entry.location = "systemdetail.asp?key=" + key
	parent.document.all.se.style.visibility="visible"

}
function updateip(key){
	parent.document.frames.sn_entry.location = "ipdetail.asp?key=" + key
	parent.document.all.se.style.visibility="visible"

}
function deleteentry(key, type){
	if(type=="ip"){
	if (confirm("Delete IP Entry?")){
	parent.document.frames.sn_entry.location = "systemmodify.asp?key=" + key + "&modify=delete ip"
	parent.document.all.se.style.visibility="hidden"
	}
	}else{
	if (confirm("Delete System Entry?")){
	parent.document.frames.sn_entry.location = "systemmodify.asp?key=" + key + "&modify=delete system"
	parent.document.all.se.style.visibility="hidden"
	}
	}
}
function showall(){
	var func = eval('document.all.allsystems')
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
	var displaytype = (func.innerHTML != '[-]' ? 'none':'block');
	var tag = document.all//('note162');
	for (i = 0; i < tag.length; i++){
		if (tag[i].name == 'systemset') tag[i].style.display = displaytype
		if (tag[i].name == 'sysfunc') tag[i].innerHTML = func.innerHTML
	} 
}
function trip(id){
	var tag = document.getElementById(id) 
	tag.style.display = (tag.style.display == "block" ? "none" : "block");
	var func = eval('document.all.func'+id)
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
}
</script>
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" class="innerbody">
  <table width="100%" border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff">
    <tr align="left" valign="bottom" bgcolor="#dddddd">
      <td colspan="10" nowrap bgcolor="#FFFFCC" class="tblunderline"> <span id="funcunassigned" name = "unnassignedfunc" style="cursor:hand;text-decoration:none;" onclick="trip('unassigned')">[+]</span> Unassigned Systems</td>
    </tr>
  </table>
<div id="unassigned" name="unassignedset" style="width:100%; height:100;border-bottom:1px solid #cccccc;display:none;">
  <table width="100%" border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff">
    <tr bgcolor="#dddddd" valign="bottom"> 
      <td colspan="2" align="center" valign="middle" nowrap bgcolor="#FFFFCC" class="tblunderline">MANAGE</td>
      <td colspan="9" align="center" bgcolor="#f0f0e0" class="tblunderline" style="border-left:1px solid #e3e3d3;">SYSTEM 
        INFORMATION </td>
    </tr>
    <tr bgcolor="#ffffff" valign="middle"> 
      <td width="4%" align="center" bgcolor="#FFFFCC" class="tblunderline" nowrap>&nbsp;</td>
      <td width="7%" align="center" bgcolor="#FFFFCC" class="tblunderline" nowrap>&nbsp;</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>serial</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>system 
        type</td>
      <td width="7%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>processor</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>memory</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>harddrive</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>nic</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>video</td>
      <td width="9%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>monitor</td>
      <td width="33%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap><div align="right">note</div></td>
    </tr>
    <%
	strsql = "select * from systemsindex where id not in (select systemid from ipindex)"
	rs.Open strsql, cnn1, 0, 1, 1
	if not rs.eof then
		Do until rs.EOF 
	%>
    <form name=form1 method="post" action="">
      <input type="hidden" name="key" value="<%=rs("id")%>">
      <tr bgcolor="#ffffff" valign="middle"> 
        <td align="center" bgcolor="#FFFFCC" class="tblunderline" nowrap><input type="button" name="edit" value="Edit" onClick="updatesystem(key.value)" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"> 
        </td>
        <td align="center" bgcolor="#FFFFCC" class="tblunderline" nowrap><img src="../../opsmanager/joblog/images/delete.gif" onClick="deleteentry(key.value,'system')" width="26" height="22" style="cursor:hand;"></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("serial")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("systemtype")%>&nbsp;</td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("processor")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("memory")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("harddrive")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("nic")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("video")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("monitor")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("note")%></td>
      </tr>
    </form>
    <%  
    rs.movenext
    loop
end if
rs.close
%>
  </table>
</div>

<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <tr bgcolor="#ffffff" valign="middle"> 
    <td width="19%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><span id="allsystems" style="cursor:hand;text-decoration:none;" onclick="javascript:showall()">[+]</span></td>
    <td width="14%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Manage 
      IP </td>
    <td width="10%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">IP</td>
    <td width="27%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">USERNAME</td>
    <td width="40%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap>MACHINE 
      NAME </td>
  </tr>
</table>  
  <%
  
if not rst1.eof then
	Do until rst1.EOF 
%>

<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <tr bgcolor="#eeefff" valign="middle"> 
    <td width="19%" align="center" bgcolor="#eeefff"  class="tblunderline" style="border-left:1px solid #e3e3d3;"><span id="func<%=replace(rst1("ip"),".","_")%>" name = "sysfunc" style="cursor:hand;text-decoration:none;" onclick="trip('<%=replace(rst1("ip"),".","_")%>')">[+]</span> 
    </td>
    <td width="7%" align="center"  class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><input type="button" name="edit2" value="Edit IP" onClick="updateip(<%=rst1("id")%>)" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
    <td width="7%" align="center"  class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><img src="../../opsmanager/joblog/images/delete.gif" onClick="deleteentry(<%=rst1("id")%>,'ip')" width="26" height="22" style="cursor:hand;"></td>
    <td width="10%" align="center"  class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("ip")%></td>
    <td width="27%" align="center"  class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("userid")%></td>
    <td width="40%" align="center"  class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><%=rst1("ipname")%></td>
  </tr>
</table>

<div id="<%=replace(rst1("ip"),".","_")%>" name="systemset" style="width:100%; height:100;border-bottom:1px solid #cccccc;display:none;">
  <table width="100%" border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff">
    <tr bgcolor="#dddddd" valign="bottom"> 
      <td colspan="2" align="center" valign="middle" nowrap bgcolor="#eeefff" class="tblunderline">MANAGE</td>
      <td colspan="9" align="center" bgcolor="#f0f0e0" class="tblunderline" style="border-left:1px solid #e3e3d3;">SYSTEM 
        INFORMATION </td>
    </tr>
    <tr bgcolor="#ffffff" valign="middle"> 
      <td width="4%" align="center" bgcolor="#eeefff" class="tblunderline" nowrap>&nbsp;</td>
      <td width="7%" align="center" bgcolor="#eeefff" class="tblunderline" nowrap>&nbsp;</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap style="border-left:1px solid #e3e3d3;">serial</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>system 
        type</td>
      <td width="7%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>processor</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>memory</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>harddrive</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>nic</td>
      <td width="8%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>video</td>
      <td width="9%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap>monitor</td>
      <td width="33%" align="right" bgcolor="#f0f0e0" class="tblunderline" nowrap><div align="right">note</div></td>
    </tr>
    <%
	strsql = "select * from systemsindex where id = " & rst1("systemid")

	rs.Open strsql, cnn1, 0, 1, 1
	if not rs.eof then
		Do until rs.EOF 
	%>
    <form name=form1 method="post" action="">
      <input type="hidden" name="key" value="<%=rs("id")%>">
      <tr bgcolor="#ffffff" valign="middle"> 
        <td align="center" bgcolor="#eeefff" class="tblunderline" nowrap><input type="button" name="edit" value="Edit" onClick="updatesystem(key.value)" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"> 
        </td>
        <td align="center" bgcolor="#eeefff" class="tblunderline" nowrap><img src="../../opsmanager/joblog/images/delete.gif" onClick="deleteentry(key.value,'system')" width="26" height="22" style="cursor:hand;"></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap style="border-left:1px solid #e3e3d3;"><%=rs("serial")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("systemtype")%>&nbsp;</td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("processor")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("memory")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("harddrive")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("nic")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("video")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("monitor")%></td>
        <td bgcolor="#f0f0e0" align="right" class="tblunderline" nowrap><%=rs("note")%></td>
      </tr>
    </form>
    <%  
    rs.movenext
    loop
end if
rs.close
%>
  </table>
</div>
  <%  
    rst1.movenext
    loop
end if
%>
</body>
</html>
