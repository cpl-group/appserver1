<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
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
cnn1.open application("cnnstr_lighting")
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
{	document.all['points'].style.visibility = 'visible'
	document.all['add'].style.visibility='hidden';
	if(nodeid=='') nodeid=-1;
	var frm = document.forms['mapform']
//	alert(id+", "+x+", "+y+", "+mapid+", "+nodeid+", "+submap);
	frm.PointX.value = x
	frm.PointY.value = y
	frm.PointId.value = id
	frm.PointAlt.value = alt;
	frm.PointMap.value = submap
	frm.PointNid.value = nodeid
	if(picknodesON) document.frames['picknode'].hilight(nodeid);
	document.all['pointedit'].style.visibility='visible';
	document.all['pointadd'].style.visibility='hidden';
}

function newpoint(newX, newY)
{	document.all['points'].style.visibility = 'visible'
	document.all['add'].style.visibility='hidden';
	var frm = document.forms['mapform']
	frm.PointX.value=newX;
	frm.PointY.value=newY;
	frm.PointId.value='';
	frm.PointAlt.value='';
	frm.PointMap.value='';
	frm.PointNid.value='';
	if(picknodesON) document.frames['picknode'].hilight('-1');
	document.all['pointadd'].style.visibility='visible';
	document.all['pointedit'].style.visibility='hidden';
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
	document.all['points'].style.visibility = 'hidden';
}
</script>
<style type="text/css">
.standard { font-family:Arial,Helvetica,sans-serif;font-size:8pt; }
.bottomline { border-bottom:1px solid #eeeeee; }
.floorlink { font-family:Arial,Helvetica,sans-serif;font-size:8pt; color:#0099ff; }
a.floorlink:hover { color:lightgreen; }
.shrunkenheader { font-family:Arial,Helvetica,sans-serif;font-size:7pt;font-weight:bold; }
</style>
</head>

<body>
<form method="post" name="mapform">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff;">
<tr bgcolor="#0099ff">
	<td><font face="Arial, Helvetica, sans-serif" color="#ffffff"><span class="standard"><b>Map Editor</b></span></font></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr valign="top" bgcolor="#eeeeee">
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

<input type="submit" name="action" value="Edit" class="standard">
<input type="submit" name="action" value="Delete" class="standard">
</td>
</tr>
<tr bgcolor="#cccccc">
<td><input type="button" name="action" value="Add New Map" onclick="document.all['add'].style.visibility='visible';document.all['points'].style.visibility='hidden';	document.all['pointadd'].style.visibility='hidden';document.all['pointedit'].style.visibility='hidden'; document.all['addcell'].style.backgroundColor='#eeeeee';" class="standard"></td>
</tr>
<tr>
<td id="addcell">
<div id="add" class="standard" style="visibility:hidden;background-color:#eeeeee;padding:3px;">
New URL: <input name="newurl" type="text"><br>
Primary Map: <input name="primarymap" value="1" type="radio">yes&nbsp;<input name="primarymap" value="0" type="radio" checked>no<br>
<input type="submit" name="action" value="Add" class="standard">&nbsp;<input type="button" value="Cancel Add" class="standard">
</div>
</td></tr></table>
<div id="points" style="visibility:hidden;">
Point Editor:<br>
X <input type="text" name="PointX">
Y <input type="text" name="PointY">
Alt Text <input type="text" name="PointAlt">
<input type="hidden" name="PointId">
<select name="PointMap">
	<%=pointoptions%>
</select>
<select name="PointNid" disabled style="font-color:black">
<option value=""></option>
<%
	rst1.open "SELECT * FROM nodes INNER JOIN label on label.id=nodes.labelid WHERE nodes.clientid='"&cid&"' ORDER BY name, nodeid", cnn1, adOpenStatic
	do until rst1.eof
		%><option value="<%=rst1("nodeid")%>"><%=rst1("name")%>(<%=rst1("nodeid")%>)</option><%
		rst1.movenext
	loop
	rst1.close
%>
</select>
<input type="button" name="action" value="Select Node..." onclick="frames['picknode'].location.href='picknode.asp?cid=<%=cid%>'; document.all['picknodes'].style.visibility='visible'; document.all['picknodes'].style.position='relative';">
<br>
<div id="pointadd" style="visibility:hidden"><input type="Button" name="action" value="Add Point" onclick="sendPointAction('Add')"></div>
<div id="pointedit" style="visibility:hidden"><input type="Button" name="action" value="Delete Point" onclick="sendPointAction('Delete')"> <input type="button" name="action" value="Edit Point" onclick="sendPointAction('Edit')"></div>
</div>
</form>
<%if action="Edit" and trim(mapid)<>"" then%>
	<table width="100%" border=0><tr><td width="100%"><iframe name="map" src="mapPointInterface.asp?mapid=<%=mapid%>" width="100%" height="300"></iframe></td>
	<td><div id="picknodes" style="visibility:hidden;position:absolute;width:200;"><iframe name="picknode" src="null.htm" width="100%" height="300"></iframe></div></td>
	</tr></table>
<%end if%>
</body>
</html>