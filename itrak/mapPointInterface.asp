<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim mapid, cid, mapref, newX, newY, PointId, PointMap, PointNid, PointAlt, action
mapid = request.querystring("mapid")
newX = request.querystring("newX")
newY = request.querystring("newy")
PointId = request.querystring("PointId")
PointMap = request.querystring("PointMap")
PointNid = request.querystring("PointNid")
PointAlt = request.querystring("PointAlt")
action = request.querystring("action")
dim cnn1, rst1, sqlstr, cmd

set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getconnect(0,0,"engineering")
cmd.activeConnection=cnn1

if action="Add" then
	cmd.commandText = "INSERT map_coor (x,y,mapid,nodeid,submap,alt) VALUES ('"&newX&"', '"&newY&"', '"&mapid&"', '"&PointNid&"', '"&PointMap&"', '"&PointAlt&"')"
	cmd.execute
elseif action="Edit" then
	cmd.commandText = "UPDATE map_coor SET x='"&newX&"', y='"&newY&"', mapid='"&mapid&"', nodeid='"&PointNid&"', submap='"&PointMap&"', alt='"&PointAlt&"' WHERE id='"&PointId&"'"
	cmd.execute
elseif action="Delete" then
	cmd.commandText = "DELETE map_coor WHERE id='"&PointId&"'"
	cmd.execute
end if


sqlstr = "SELECT * FROM maps m LEFT JOIN map_coor mc on m.id=mc.mapid WHERE m.id="&mapid
rst1.open sqlstr, cnn1

if not(rst1.eof) then mapref = rst1("url")
%>

<html>
<head>
<title>Maps</title>
<script>
var workingpoint = -1;
function setPoint(point)
{	if(point!=workingpoint)
	{	document.all['newpoint'].style.visibility='hidden'
		if(workingpoint!=-1)
		{	workingpoint.style.borderColor='#006699';
			workingpoint.style.borderWidth=0;
			workingpoint.style.left=parseInt(workingpoint.style.left)+5;
			workingpoint.style.top=parseInt(workingpoint.style.top)+5
		}
		workingpoint = point
		workingpoint.style.borderColor='#996600'
	}
}
function newPoint(x, y)
{	if(workingpoint!=-1)
	{	workingpoint.style.borderColor='#006699';
		workingpoint.style.borderWidth=0;
		workingpoint.style.left=parseInt(workingpoint.style.left)+5;
		workingpoint.style.top=parseInt(workingpoint.style.top)+5
		workingpoint = -1;
	}
	document.all['newpoint'].style.left=parseInt(x)-5;
	document.all['newpoint'].style.top=parseInt(y)-5;
	document.all['newpoint'].style.visibility='visible'
}

function overPoint(point)
{	if(point!=workingpoint)
	{	point.style.borderWidth=5;
		point.style.left=parseInt(point.style.left)-5;
		point.style.top=parseInt(point.style.top)-5;
	}
}
function outPoint(point)
{	if(point!=workingpoint)
	{	point.style.borderWidth=0;
		point.style.left=parseInt(point.style.left)+5;
		point.style.top=parseInt(point.style.top)+5
	}
}
</script>
</head>
<body bgcolor="#999999">
<img onclick="newPoint(document.body.scrollLeft+event.x-2, document.body.scrollTop+event.y-2);parent.newpoint(document.body.scrollLeft+event.x-2, document.body.scrollTop+event.y-2);parent.mapform.addbutton.style.fontWeight='bold';parent.mapform.editbutton.style.fontWeight='normal';" style="position:absolute; left:0; top:0" id="map" src="<%=mapref%>" border="0">
	<div id="newpoint" style="position:absolute; visibility:hidden; border-width:5; border-style:solid; border-color:#996600; cursor:hand; left:0; top:0"><a><img src="bullet-add1.gif" alt="" width="10" height="12" border="0"></a></div>
<%
if trim(rst1("x"))<>"" then
	do until rst1.eof%>
		<div style="position:absolute; border-width:0; border-style:solid; border-color:#006699; cursor:hand; left:<%=rst1("x")%>; top:<%=rst1("y")%>" onmouseover="overPoint(this)" onmouseout="outPoint(this)" onclick="setPoint(this)"><a onclick="parent.sendPointInfo(<%=rst1("id")%>, <%=rst1("x")%>, <%=rst1("y")%>, <%=rst1("mapid")%>, '<%=rst1("nodeid")%>', '<%=rst1("submap")%>', '<%=rst1("alt")%>'); parent.mapform.editbutton.style.fontWeight='bold';parent.mapform.addbutton.style.fontWeight='normal';"><img src="bullet-add1.gif" width="10" height="12" border="0" alt="<%=rst1("alt")%>"></a></div>
	<%rst1.movenext
	loop
end if
%>
</body>
</html>
