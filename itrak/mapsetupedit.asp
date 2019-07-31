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
{	//document.all['points'].style.visibility = 'visible'
//	document.all['add'].style.visibility='hidden';
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
//	document.all['pointedit'].style.visibility='visible';
//	document.all['pointadd'].style.visibility='hidden';
}

function newpoint(newX, newY)
{	//document.all['points'].style.visibility = 'visible'
//	document.all['add'].style.visibility='hidden';
	var frm = document.forms['mapform']
	frm.PointX.value=newX;
	frm.PointY.value=newY;
	frm.PointId.value='';
	frm.PointAlt.value='';
	frm.PointMap.value='';
	frm.PointNid.value='';
	if(picknodesON) document.frames['picknode'].hilight('-1');
//	document.all['pointadd'].style.visibility='visible';
//	document.all['pointedit'].style.visibility='hidden';
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
//	document.all['points'].style.visibility = 'hidden';
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
	<td><span class="standard" style="color:#ffffff"><b>Configure Maps for <%=clientname%> | Edit Map</b></span></td>
	<td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=cid%>'" class="standard"></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#cccccc">
	<td><span class="standard">&nbsp;</span></td>
	<td><span class="standard"><b>Set Points</b></span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18%"><span class="standard">&nbsp;</span></td>
	<td>
	<span class="standard">
  <table border=0 cellpadding="5" cellspacing="0">
  <tr>
    <td width="18"><span class="standard"><img src="images/num_one.gif" alt="1" width="13" height="13" border="0"></span></td>
    <td>
    <table border=0 cellpadding="0" cellspacing="0">
    <tr valign="middle">
      <td><span class="standard">Select map:&nbsp;</span></td>
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
      <td><span class="standard">&nbsp;</span></td>
      <td>
      <input type="submit" name="action" value="Edit" class="standard" style="padding-left:4px;padding-right:4px;">
      <input type="submit" name="action" value="Delete" onclick="return confirmDelete();" class="standard" style="padding-left:4px;padding-right:4px;">
<!--
      [[a onMouseOut="closeHelpBox()" onMouseOver="helpbox('pickmap',event.x,event.y)"]][[img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"]][[/a]]
-->
      </td>
    </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td><span class="standard"><img src="images/num_two.gif" alt="2" width="13" height="13" border="0"></span></td>
    <td colspan="2"><span class="standard">Click on the map image below to set a point at the location of your mouse click, or click an existing point (red pylon) to edit it</span></td>
  </tr>
  </table>
	</span>
	</td>
</tr>
<tr bgcolor="#cccccc">
  <td><span class="standard">&nbsp;</span></td>
	<td>
  <table border=0 cellpadding="5" cellspacing="0">
  <tr valign="top">
    <td width="18"><span class="standard"><img src="images/num_three.gif" alt="3" width="13" height="13" border="0"></span></td>
    <td><span class="standard"><b>Edit Point</b></span></td>
  </tr>
  </table>
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td width="18%" align="right"><span class="standard">Coordinates</span></td>
	<td>
	<span class="standard">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr valign="middle">
    <td width="30" align="right"><span class="standard">X:&nbsp;</span></td>
    <td><input type="text" name="PointX" size="4"></td>
    <td><span class="standard">&nbsp;&nbsp;</span></td>
    <td><span class="standard">Y:&nbsp;</span></td>
    <td><input type="text" name="PointY" size="4"></td>
    <td><span class="standard">&nbsp;</span></td>
    <td><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('xy',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></td>
  </tr>
  </table>
	</span>
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Alt Text</span></td>
	<td style="padding-left:33px;"><input type="text" name="PointAlt"><input type="hidden" name="PointId"></td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
	<td align="right" style="padding-top:7px;"><span class="standard">Link Point To</span></td>
	<td>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td width="30" align="right"><span class="standard">Map:&nbsp;</span></td>
    <td>
    <span class="standard"> 
    <select name="PointMap">
    <option>Select map</option>
      <%=pointoptions%>
    </select>
    </span>
    </td>
    <td width="14"><span class="standard">&nbsp;</span></td>
    <td><span class="standard">or node:&nbsp;</span></td>
    <td>
    <span class="standard"> 
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
    </select> &nbsp;<input type="button" name="action" value="Select Node..." onclick="frames['picknode'].location.href='picknode.asp?cid=<%=cid%>'; document.all['picknodes'].style.visibility='visible'; document.all['picknodes'].style.position='relative';" class="standard">
    <a onMouseOut="helpup('help_link');" onMouseOver="helpdrop('help_link','link_map_pt_to');"><img name="help_link_img" src="images/question-rt.gif" alt="?" title="" width="22" height="13" hspace="4" border="0"></a>
    </span>
    </td>
  </tr>
  </table>
  <div id="help_link" class="standard" style="display:'none';margin-left:30px; margin-right:40px;padding-top:6px;padding-bottom:6px;"></div>    
	</td>
</tr>
<tr bgcolor="#cccccc">
	<td align="right"><span class="standard">&nbsp;</span></td>
	<td>
  <table border=0 cellpadding="5" cellspacing="0">
  <tr valign="middle">
    <td width="18"><img src="images/num_four.gif" alt="4" width="13" height="13" border="0"></td>
    <td><div id="pointadd" style="visibility:visible;"><input type="Button" id="addbutton" name="action" value="Add Point" onclick="sendPointAction('Add');this.style.fontWeight='normal';" class="standard"> <input type="button" name="action" id="editbutton" value="Edit Point" onclick="sendPointAction('Edit');this.style.fontWeight='normal';" class="standard"> <input type="Button" name="action" value="Delete Point" onclick="sendPointAction('Delete');document.forms['mapform'].addbutton.style.fontWeight='normal';document.forms['mapform'].editbutton.style.fontWeight='normal';" class="standard"> </div></td>
    <td><a onMouseOut="closeHelpBox()" onMouseOver="helpbox('save_point',event.x,event.y)"><img src="images/question-ccc.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a></td>
  </tr>
  </table>
	</td>
</tr>
</table>

</form>
<%if action="Edit" and trim(mapid)<>"" then%>
	<table width="100%" border=0 cellpadding="2" cellspacing="0">
	<tr>
	  <td width="80%"><iframe name="map" src="mapPointInterface.asp?mapid=<%=mapid%>" width="100%" height="300"></iframe></td>
  	<td width="200"><div id="picknodes" style="visibility:visible;width:200;background:#999999;"><iframe name="picknode" src="null.htm" width="100%" height="300"></iframe></div></td>
	</tr>
	</table>
<%end if%>
<!--#INCLUDE FILE="helpbox.htm"-->
</body>
</html>