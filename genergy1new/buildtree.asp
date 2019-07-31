<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE file="sessioncheck.asp"-->
<!--#INCLUDE file="buildxmlfunctions.asp"-->
<%
dim cid
cid = request("cid")

dim cnn1, rst1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_lighting")

dim xmlobj, xslobj
set xslobj = server.createobject("MSXML2.DomDocument")
xslobj.async = False
xslobj.load Server.MapPath("treemenu.xsl")

if not(isobject(application("xmltree"&cid))) then 
	response.write "built from database"
	rst1.open "SELECT * FROM nodes INNER JOIN label on label.id=nodes.labelid WHERE nodes.clientid="&cid&" ORDER BY fatherref, position, name", cnn1, adOpenStatic
	copyresults()
	rst1.close

	set xmlobj = server.createobject("MSXML2.FreeThreadedDOMDocument")
	buildtree xmlobj, treerecord
	set application("xmltree"&cid) = xmlobj
else
	rst1.open "SELECT distinct logo FROM clients WHERE id="&cid, cnn1
	if not rst1.eof then logo = rst1("logo")
	rst1.close
	set xmlobj = application("xmltree"&cid)
'	application("xmltree"&cid) = null ' to force clear cache
end if
cnn1.close
set cnn1 = nothing
'response.write xmlobj.xml
%>
<HTML xmlns:genergy>
<HEAD>
<TITLE>New Menus</TITLE>
<STYLE>
a
{	text-decoration:none
}
a:hover
{	color: #99DDFF
}
genergy\:root
{	behavior: url(expand_js_g1.htc);
	font: 0pt arial;
	white-space : nowrap;
	mv--indent : 0;
}
genergy\:branch
{	behavior: url(expand_js_g1.htc);
	font: 9pt arial;
	white-space : nowrap;
	mv--indent : 20;
}
genergy\:leaf
{	behavior: url(expand_js_g1.htc);
	font: 9pt arial;
	white-space : nowrap;
	cursor : hand;
}
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
</STYLE>
<script>
function expandtree()
{	document.all['root'].expandAll()
}
function colapsetree()
{	document.all['root'].closeAll()
	document.all['root'].expandNode();
}
function opennid(nid)
{	var nodes = Array()
	var branch = document.getElementsByTagName('branch')
	var leaf = document.getElementsByTagName('leaf')
	for(i=0;i<leaf.length;i++)
	{	nodes[nodes.length] = leaf[i];
	}
	for(i=0;i<branch.length;i++)
	{	nodes[nodes.length] = branch[i];
	}
	
	for(i=0;i<nodes.length;i++)
	{	if(nodes[i].nid==nid)
		{	nodes[i].expandNode();
			if(nodes[i].fid!='0')
			{	opennid(nodes[i].fid);
			}
		}
	}
}
</script>
</HEAD>
<BODY onload="document.all['root'].expandNode();" bgcolor="#999999" LINK="#FFFFFF" marginwidth="1" marginheight="1" text="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%Response.Write xmlobj.transformNode (xslobj)%>
</BODY>
</HTML>