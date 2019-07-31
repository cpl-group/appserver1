<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim mapid, cid, mapref
mapid = request.querystring("mapid")
cid = request.querystring("cid")

dim cnn1, rst1, sqlstr

if trim(mapid)<>"" then
	sqlstr = "SELECT * FROM maps m LEFT JOIN map_coor mc on m.id=mc.mapid WHERE m.clientid='"&cid&"' and m.id="&mapid
else
	sqlstr = "SELECT * FROM maps m LEFT JOIN map_coor mc on m.id=mc.mapid WHERE m.clientid="&cid&" and m.primarymap=1"
end if
'response.write sqlstr
'response.end
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getconnect(0,0,"engineering")
rst1.open sqlstr, cnn1

if not(rst1.eof) then mapref = rst1("url")
%>

<html>
<head>
<title>Maps</title>
<script>
var Xoffset = 0;
function init()
{	var framewidth = window.document.body.clientWidth
	var imagewidth = document.all['map'].width
	if(framewidth-imagewidth>0)
	{	Xoffset = Math.round((framewidth-imagewidth)/2)
	}else
	{	Xoffset = 0;
	}
	document.all['map'].style.left = Xoffset
	setdotsX()
}

function setdotsX()
{	var dots = document.getElementsByTagName('div')
	for(i=0;i<dots.length;i++)
	{	dots[i].style.left = (parseInt(dots[i].Xoffset)+parseInt(Xoffset));
	}
}

function nullspace()
{
}

function overPoint(point)
{	point.style.borderWidth=5;
	point.style.left=parseInt(point.style.left)-5;
	point.style.top=parseInt(point.style.top)-5;
}
function outPoint(point)
{	point.style.borderWidth=0;
	point.style.left=parseInt(point.style.left)+5;
	point.style.top=parseInt(point.style.top)+5
}


</script>
</head>
<body onload="init();" onresize="init()">
<img style="position:absolute; left:0; top:40" id="map" src="<%=mapref%>" border="0">
<%do until rst1.eof%>
<% 
	dim tmpY 
	tmpY = rst1("y") + 40
	if trim(rst1("y"))<>"" and trim(rst1("x")) then%>
		<div Xoffset="<%=rst1("x")%>" style="position:absolute; border-width:0; border-style:solid; border-color:#006699; cursor:hand; left:0; top:<%=tmpY%>" onmouseover="overPoint(this)" onmouseout="outPoint(this)"><a <%if trim(rst1("submap"))<>"" then%>href="maps.asp?cid=<%=cid%>&mapid=<%=rst1("submap")%>"<%end if%> onclick="<%if trim(rst1("nodeid"))<>"" then%>parent.frames.menu.colapsetree();parent.frames.menu.opennid(<%=rst1("nodeid")%>);parent.frames['menu'].document.location.hash='<%=rst1("nodeid")%>'<%end if%>"><img src="bullet-add1.gif" alt="<%=rst1("alt")%>" width="16" height="16" border="0"></a></div>
<%	end if
	rst1.movenext
loop
%>
</body>
</html>
