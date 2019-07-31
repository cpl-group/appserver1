<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!--#INCLUDE file="buildxmlfunctions.asp"-->
<%
dim cid, action, primarymap, newurl, mapid
cid = request("cid")
action = request("action")
primarymap = request("primarymap")
newurl = request("newurl")
mapid = request("mapid")

dim rst1, cnn1, cmd
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
cnn1.open getconnect(0,0,"engineering")
cmd.ActiveConnection = cnn1

if trim(action)="Add" then
	if primarymap=1 then
		cmd.commandText="UPDATE maps SET primarymap=0 WHERE clientid='"&cid&"'"
		cmd.execute
	end if
	cmd.commandText="INSERT maps (url, clientid, primarymap) VALUES ('"&newurl&"', '"&cid&"', '"&primarymap&"')"
	cmd.execute
elseif trim(action)="Delete" then
	if trim(mapid)<>"" then
		cmd.commandText = "DELETE maps WHERE id="&mapid&""
		cmd.execute
		cmd.commandText = "DELETE map_coor WHERE mapid='"&mapid&"'"
		cmd.execute
	end if
end if


rst1.open "SELECT corp_name FROM clients WHERE id='"&cid&"'", cnn1, 0, 1, 1
clientname = rst1("corp_name")
rst1.close


%>
<html>
<head>
<title>Map Setup</title>
<script>
//function selectmap()
//{	var frm = document.forms.['mapform']
//	document.location.href = "mapsetup.asp = frm.mapid.option[1].value)
//}
var picknodesON = 0;

function sendPointInfo(id, x, y, mapid, nodeid, submap, alt)
{	
	if(nodeid=='') nodeid=-1;
	var frm = document.forms['mapform']
	frm.PointX.value = x
	frm.PointY.value = y
	frm.PointId.value = id
	frm.PointAlt.value = alt;
	frm.PointMap.value = submap
	frm.PointNid.value = nodeid
	if(picknodesON) document.frames['picknode'].hilight(nodeid);
	document.all['editfields'].style.visibility='visible';
	document.all['pointadd'].style.visibility='hidden';
}

function newpoint(newX, newY)
{	
	var frm = document.forms['mapform']
	frm.PointX.value=newX;
	frm.PointY.value=newY;
	frm.PointId.value='';
	frm.PointAlt.value='';
	frm.PointMap.value='';
	frm.PointNid.value='';
	if(picknodesON) document.frames['picknode'].hilight('-1');
	document.all['pointadd'].style.visibility='visible';
	document.all['editfields'].style.visibility='hidden';
}

function sendPointAction(action)
{	var frm = document.forms['mapform'];
	var temp='mapPointInterface.asp?mapid=<%=mapid%>'
	temp += '&newX='+frm.PointX.value+'&newY='+frm.PointY.value+'&PointId='+frm.PointId.value+'&PointMap='+frm.PointMap.value+'&PointNid='+frm.PointNid.value+'&action='+action+'&PointAlt='+frm.PointAlt.value
	document.frames['map'].document.location.href = temp;
	frm.PointX.value='';
	frm.PointY.value='';
	frm.PointId.value='';
	frm.PointAlt.value='';
	frm.PointMap.value='';
	frm.PointNid.value='';
}

function confirmDelete(){
  retval = window.confirm("Are you sure you want to delete this item?");
  return retval;
}

</script>
<script src="messages.js" type="text/javascript" language="Javascript1.2"></script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body>
<form method="post" name="mapform">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff;">
<tr bgcolor="#0099ff">
	<td><span class="standardheader"><%=clientname%> | Configure Maps</span></td>
	<td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=cid%>'" class="standard" disabled></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#cccccc">
  <td><span class="standard">&nbsp;</span></td>
	<td><span class="standard"><b>Add New Map</b></span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18%" align="right"><span class="standard">New Map</span></td>
  <td><input name="newurl" type="text"><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('new_map',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a><br><span class="standard" style="font-size:7pt;">e.g., maps/manhattan.jpg</span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td align="right"><span class="standard">Primary Map</span></td>
  <td>
  <span class="standard">
  <input name="primarymap" value="1" type="radio">yes&nbsp;<input name="primarymap" value="0" type="radio" checked>no <a onMouseOut="closeHelpBox()" onMouseOver="helpbox('primary_map',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a><br>
  </span>
  </td>
</tr>
<tr bgcolor="#eeeeee">
	<td><span class="standard">&nbsp;</span></td>
	<td><span class="standard"><input type="submit" name="action" value="Add" class="standard" style="padding-left:10px;padding-right:10px;">&nbsp;<input type="reset" value="Cancel" class="standard"></span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td colspan="2"><span class="standard">&nbsp;</span></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#cccccc">
	<td width="18%" align="right"><span class="standard"><div id="errmsg" class="standard" style="color:#cc3300;"></div></span></td>
  <td>
  
  <table border=0 cellpadding="2" cellspacing="0">
  <tr>
    <td><span class="standard"><b>Edit Map</b></span></td>
    <td>
    <select name="mapid">
    <%
    dim pointoptions
    pointoptions = ""
    rst1.open "SELECT * FROM maps WHERE clientid='"&cid&"' order by primarymap desc, id", cnn1
    if rst1.eof then
      response.write "<option value="""">No maps provided</option>"
    else
      response.write "<option value="""">Select map</option>"
    end if
    do until rst1.eof
      %><option value="<%=rst1("id")%>"<%if trim(mapid)=trim(rst1("id")) then response.write "SELECTED"%>><%=rst1("url")%></option><%
      pointoptions = pointoptions & "<option value="""&rst1("id")&""">"&rst1("url")&"</option>"
      rst1.movenext
    loop
    rst1.close
    %>
    </select>
    </td>
    <td>
    <input type="submit" name="action" value="Edit" class="standard" style="padding-left:4px;padding-right:4px;">
    <input type="submit" name="action" value="Delete" onclick="return confirmDelete();" class="standard" style="padding-left:4px;padding-right:4px;">
    </td>
  </tr>
  </table>
  </td>
</tr>
</table>



</form>
<!--
[[%if action="Edit" and trim(mapid)[[]]"" then%]]
[[script language="JavaScript"]]
	document.location.href="mapsetupedit.asp?cid=[[%=cid%]]&action=[[%=action%]]&primarymap=[[%=primarymap%]]&newurl=[[%=newurl%]]&mapid=[[%=mapid%]]";
[[/script]]
[[%end if%]]
-->
<%
if action="Edit" and trim(mapid)<>"" then
response.redirect "mapsetupedit.asp?cid="&cid&"&action="&action&"&primarymap="&primarymap&"&newurl="&newurl&"&mapid="&mapid&""
elseif (action="Edit" or action="Delete") and trim(mapid)="" then
response.write "<script>document.all.errmsg.innerHTML='Please select a map first&nbsp;'</script>"
end if
%>
<!--#INCLUDE FILE="helpbox.htm"-->
</body>
</html>