<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!--#INCLUDE file="buildxmlfunctions.asp"-->
<%
dim cid
cid = request("cid")

dim cnn1, rst1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getconnect(0,0,"engineering")
rst1.open "SELECT * FROM nodes INNER JOIN label on label.id=nodes.labelid WHERE nodes.clientid="&cid&" ORDER BY fatherref, position, name", cnn1, adOpenStatic
copyresults()
rst1.close

dim xmlobj, xslobj
set xmlobj = server.createobject("MSXML2.DomDocument")
set xslobj = server.createobject("MSXML2.DomDocument")
xslobj.async = False
xslobj.load Server.MapPath("setuptreemenu.xsl")

buildtree xmlobj, treerecord
cnn1.close
%>
<HTML xmlns:genergy>
<HEAD>
<TITLE>New Menus</TITLE>
<STYLE>
genergy\:root
{	behavior: url(expand_js_gStatic.htc);
	font: 0pt arial;
	white-space : nowrap;
	mv--indent : 0;
}
genergy\:branch
{	behavior: url(expand_js_gStatic.htc);
	font: 9pt arial;
	white-space : nowrap;
	mv--indent : 20;
	cursor : hand;
}
genergy\:leaf
{	behavior: url(expand_js_gStatic.htc);
	font: 9pt arial;
	white-space : nowrap;
	cursor : hand;
}
</STYLE>
<SCRIPT>
var moveselecting = 0;
function sendNodeInfo(nid, fid, labelid, position, target, nlink, labelname)
{	//alert("nid"+nid+"\nfid"+fid+"\nlabel"+nlabel+"\nlink"+nlink);
	var frm = parent.document.forms['mapform'];
	frm.PointNid.value = nid;
	hilight(nid)
	document.all['genergymenu'].expandAll();
}

function hilight(nodeid)
{	var nodes = document.getElementsByName('node');
	var colorstring
	colorstring = "#ccccff"
	for(i=0;i<nodes.length;i++)
	{	if(nodes[i].nid==nodeid)
		{	nodes[i].style.backgroundColor=colorstring;
		}else if(nodes[i].style.backgroundColor==colorstring)
		{	nodes[i].style.backgroundColor='#FFFFFF';
		}
	}
}

function makeActiveSelect(a)
{
}
</SCRIPT>
</HEAD>
<BODY bgcolor="#999999" onload="document.all['genergymenu'].expandAll();parent.picknodesON=1;">
	<%Response.Write xmlobj.transformNode (xslobj)%>
</body>
</html>