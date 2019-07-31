<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE file="buildxmlfunctions.asp"-->
<%
dim cid, action, nid, fid, link, position, target, labelid, addlabel, types, mnid, mlink, mlabel, cnid, clink, clabel
cid = request("cid")
action = request("action")
nid = request("nid")
fid = request("fid")
link = request("link")
position = request("position")
target = request("target")
labelid = request("labelid")
addlabel = trim(request("addlabel"))
types = request("type")
mnid = request("mnid")
mlink = request("mlink")
mlabel = request("mlabel")
cnid = request("cnid")
clink = request("clink")
clabel = request("clabel")


'response.contentType = "text/xml"
'on error goto printerror
dim cnn1, rst1, cmd, strsql, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_lighting")
cmd.activeconnection = cnn1
if addlabel<>"" then
	dim newlabel
	newlabel = false
	rst1.open "SELECT name, id FROM label WHERE clientid="&cid&" and lower(name)='"&lcase(addlabel)&"'", cnn1
	newlabel = rst1.eof
	if not newlabel then labelid = rst1("id")
	rst1.close
	if newlabel then 'if label does not already exist
		strsql = "INSERT into Label (name, type, clientid) VALUES ('"&addlabel&"', '"&types&"', '"&cid&"')"
		cmd.commandText = strsql
		cmd.execute
		rst1.open "SELECT id FROM label WHERE clientid="&cid&" and lower(name)='"&lcase(addlabel)&"'", cnn1
		labelid = rst1("id")
		rst1.close
	end if
end if

if trim(nid)<>"" then
	if action="Edit" then
		strsql = "UPDATE nodes SET labelid='"&labelid&"', link='"&link&"', position="&position&", target='"&target&"' WHERE clientid='"&cid&"' and nodeid="&nid&""
		cmd.CommandText = strsql
		cmd.Execute
	elseif action="Add Child" then
		if trim(nid)<>"" then
			strsql = "INSERT into nodes (labelid, link, fatherref, clientid, position, target) VALUES (21, '', "&nid&", '"&cid&"', 1, 'main')"
			cmd.CommandText = strsql
			cmd.Execute
			strsql = "UPDATE nodes SET relative=1 WHERE nodeid="&nid&""
			cmd.CommandText = strsql
			cmd.Execute
		end if
	elseif action="Move" then
		cmd.CommandText = "sp_getnodes"
		cmd.CommandType = adCmdStoredProc
		'input params
		Set prm = cmd.CreateParameter("nodeid", adVarChar, adParamInput, 10)
		cmd.Parameters.Append prm
		cmd.Parameters("nodeid") = nid
		set rst1 = cmd.execute
		dim isowngrandpa
		isowngrandpa = false
		do until rst1.eof
			if trim(rst1("nodeid"))=trim(mnid) then
				isowngrandpa = true
			end if
			rst1.movenext
		loop
		rst1.close
		if not isowngrandpa then
			cmd.CommandType = adCmdText
			strsql = "UPDATE nodes SET fatherref="&mnid&" WHERE nodeid="&nid
			cmd.CommandText = strsql
			cmd.execute
			strsql = "UPDATE nodes SET relative=1 WHERE nodeid="&mnid
			cmd.CommandText = strsql
			cmd.execute
			cmd.CommandText = "sp_getnodes"
			cmd.CommandType = adCmdStoredProc
			'input params
			cmd.Parameters("nodeid") = fid
			set rst1 = cmd.execute
		'	if not rst1.eof  then response.write rst1.recordcount
			dim recordcount
			recordcount = 0
			do until rst1.eof
				recordcount = recordcount+1
				rst1.movenext
			loop
			rst1.close
			if recordcount=1 then
				cmd.CommandType = adCmdText
				strsql = "UPDATE nodes SET relative=0 WHERE nodeid="&fid
				cmd.CommandText = strsql
				cmd.execute
			end if
		else
			response.write "<span style=""color:red"">Can not move a node into one of its children.</span>"
		end if
	elseif action="Copy" then
		cmd.CommandText = "sp_copy_node"
		cmd.CommandType = adCmdStoredProc
		'input params
		Set prm = cmd.CreateParameter("nodeid", adVarChar, adParamInput, 10)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("did", adVarChar, adParamInput, 10)
		cmd.Parameters.Append prm
		cmd.Parameters("nodeid") = nid
		cmd.Parameters("did") = cnid
		cmd.execute
	elseif action="Delete" then
		cmd.CommandText = "sp_deletenode"
		cmd.CommandType = adCmdStoredProc
		'input params
		Set prm = cmd.CreateParameter("nodeid", adVarChar, adParamInput, 10)
		cmd.Parameters.Append prm
		Set cmd.ActiveConnection = cnn1
		cmd.Parameters("nodeid") = nid
		cmd.execute
	end if
elseif trim(action)<>"" then
	response.write "<span style=""color:red"">Please select a node before add, edit or deleting a node.</span>"
end if











'###########display###########'
rst1.open "SELECT * FROM nodes INNER JOIN label on label.id=nodes.labelid WHERE nodes.clientid='"&cid&"' ORDER BY fatherref, position, name", cnn1, adOpenStatic
copyresults()
rst1.close

dim xmlobj, xslobj
set xmlobj = server.createobject("MSXML2.FreeThreadedDomDocument")
set xslobj = server.createobject("MSXML2.DomDocument")
xslobj.async = False
xslobj.load Server.MapPath("setuptreemenu.xsl")

buildtree xmlobj, treerecord

if trim(action)<>"" then set application("xmltree"&cid) = xmlobj
'response.write xmlobj.xml
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
	mv--indent : 20px;
	cursor : hand;
}
genergy\:leaf
{	behavior: url(expand_js_gStatic.htc);
	font: 9pt arial;
	white-space : nowrap;
	mv--indent : 20px;
	cursor : hand;
}
</STYLE>
<SCRIPT>
var moveselecting = 0;
var copyselecting = 0;
function sendNodeInfo(nid, fid, labelid, position, target, nlink, labelname)
{	//alert("nid"+nid+"\nfid"+fid+"\nlabel"+nlabel+"\nlink"+nlink);
	var frm = document.forms['node'];
	frm.addLabel.value="";
	if(moveselecting==1)
	{	if(nid!=frm.nid.value)
		{	frm.mnid.value = nid;
			document.all['mnid'].innerText = nid;
			document.all['mlabel'].innerText = labelname;
			document.all['mlink'].innerText = nlink;
			hilight(nid);
		}else
		{	alert("Cannot move node to itself");
		}
	}else if(copyselecting==1)
	{	if(nid!=frm.nid.value)
		{	frm.cnid.value = nid;
			document.all['cnnid'].innerText = nid;
			document.all['clabel'].innerText = labelname;
			document.all['clink'].innerText = nlink;
			hilight(nid);
		}else
		{	alert("Cannot copy node to itself");
		}
	}else
	{	frm.nid.value = nid;
		frm.fid.value = fid;
		frm.labelid.value = labelid;
		frm.link.value = nlink;
		frm.position.value = position;
		frm.target.value = target;
		synchSelects();
		checkaddlabel();
		hilight(nid);
	}
	document.all['genergymenu'].expandAll();
}

function checkaddlabel()
{	var str = document.forms['node'].addLabel.value;
	str = str.replace(/\s*/,'');
	if(str!='')
	{	document.all['addlabel'].style.visibility='visible';
		document.forms['node'].region.disabled=1;
		document.forms['node'].city.disabled=1;
		document.forms['node'].address.disabled=1;
		document.forms['node'].service.disabled=1;
	}else
	{	document.all['addlabel'].style.visibility='hidden';
		document.forms['node'].region.disabled=0;
		document.forms['node'].city.disabled=0;
		document.forms['node'].address.disabled=0;
		document.forms['node'].service.disabled=0;
	}
}

function synchSelects()
{	var frm = document.forms['node'];
	frm.region.value = frm.labelid.value;
	frm.city.value = frm.labelid.value;
	frm.address.value = frm.labelid.value;
	frm.service.value = frm.labelid.value;
}

function makeActiveSelect(str)
{	if(moveselecting==0)
	{	document.all['region'].style.visibility="hidden";
		document.all['city'].style.visibility="hidden";
		document.all['address'].style.visibility="hidden";
		document.all['service'].style.visibility="hidden";
		
		document.all['region'].style.position="absolute";
		document.all['city'].style.position="absolute";
		document.all['address'].style.position="absolute";
		document.all['service'].style.position="absolute";
		if(str!='default')
		{	document.all[str].style.visibility='visible';
			document.all[str].style.position="relative";
			document.forms['node'].type.value=str;
		}else
		{	makeActiveSelect('service');
		}
	}
}


function hilight(nodeid)
{	var nodes = document.getElementsByName('node');
	var colorstring
	if((moveselecting==0)&&(copyselecting==0))
	{	colorstring = "#ccccff"
	}else
	{	colorstring = "#ffcccc"
	}
	for(i=0;i<nodes.length;i++)
	{	//alert(nodes[i].style.backgroundColor+'|'+colorstring+(nodes[i].style.backgroundColor==colorstring));
		if(nodes[i].nid==nodeid)
		{	nodes[i].style.backgroundColor=colorstring;
		}else if(nodes[i].style.backgroundColor==colorstring)
		{	nodes[i].style.backgroundColor='#FFFFFF';
		}
	}
}
function setscroll()
{	document.all['treewindow'].scrollTop=document.forms['node'].scroll.value;
}

function cancelCopy()
{	hilight(-1);
	copyselecting=0;
	document.all['copy'].style.visibility='hidden';
	document.all['cnnid'].innerText = '';
//	document.all['cfid'].innerText = '';
	document.all['clabel'].innerText = '';
	document.all['clink'].innerText = '';
}

function cancelMove()
{	hilight(-1);
	moveselecting=0;
	document.all['move'].style.visibility='hidden';
	document.all['mnid'].innerText = '';
//	document.all['mfid'].innerText = '';
	document.all['mlabel'].innerText = '';
	document.all['mlink'].innerText = '';
}

</SCRIPT>
<style type="text/css">
.standard { font-family:Arial,Helvetica,sans-serif;font-size:8pt; }
.bottomline { border-bottom:1px solid #eeeeee; }
.floorlink { font-family:Arial,Helvetica,sans-serif;font-size:8pt; color:#0099ff; }
a.floorlink:hover { color:lightgreen; }
.shrunkenheader { font-family:Arial,Helvetica,sans-serif;font-size:7pt;font-weight:bold; }
</style>
</HEAD><!--  -->
<BODY bgcolor="#FFFFFF" onload="document.all['genergymenu'].expandAll(document.all['treewindow'],document.forms['node'].scroll.value);setTimeout('setscroll()',200)">
<table width="100%" border="0" height="33">
<tr bgcolor="#0099ff"><td align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><span class="standard">Client : <%=clientname%></span></font></b></td></tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0"><tr>
<td valign="top">
	<div id="treewindow" style="overflow:auto;width:250;height:340;">
	<%Response.Write xmlobj.transformNode (xslobj)%>
	</div>
</td>
<td valign="top" width="90%"><form name="node" method="post" onsubmit="scroll.value=document.all['treewindow'].scrollTop;">
<input type="hidden" name="scroll" value="<%=request("scroll")%>">
<span style="background-color:#CCCCFF">Label:</span><br>
<%
dim labelregion, labelcity, labeladdress, labelservice

rst1.open "SELECT * FROM Label WHERE clientid='"&cid&"' ORDER BY name ", cnn1
do until rst1.eof
	dim tempstr
	tempstr = "<option link value="&rst1("id")&">"&rst1("name")&"</option>"
	if rst1("type")="region" then
		labelregion = labelregion & tempstr
	elseif rst1("type")="city" then
		labelcity = labelcity & tempstr
	elseif rst1("type")="address" then
		labeladdress = labeladdress & tempstr
	elseif rst1("type")="service" then
		labelservice = labelservice & tempstr
	end if
	rst1.movenext
loop
rst1.close
%>
<nobr>
<select name="type" id="type" onchange="makeActiveSelect(this.value)">
<option value="region">Region</option>
<option value="city">City</option>
<option value="address">Address</option>
<option value="service">Service</option>
</select>
<select name="region" id="region" style="visibility:hidden;position:absolute;" onchange="labelid.value=this.value; synchSelects()">
<%=labelregion%>
</select>
<select name="city" id="city" style="visibility:hidden;position:absolute;" onchange="labelid.value=this.value; synchSelects()">
<%=labelcity%>
</select>
<select name="address" id="address" style="visibility:hidden;position:absolute;" onchange="labelid.value=this.value; synchSelects()">
<%=labeladdress%>
</select>
<select name="service" id="service" style="visibility:hidden;position:absolute;" onchange="labelid.value=this.value; synchSelects()">
<%=labelservice%>
</select>
&nbsp;<input type="text" name="addLabel" onKeyUp="checkaddlabel()">&nbsp;<span id="addlabel" style="visibility:hidden; color:red">Add Label</span><br>
<input type="Hidden" name="labelid" value=""></nobr>
<table><tr><td>Node&nbsp;id&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>Parent&nbsp;id&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>Link</td>
<td>Target&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>Position</td></tr>
<tr valign="top"><td><input type="text" name="nid" size="5" readonly></td>
<td><input type="text" name="fid" size="5" readonly></td>
<td><input type="text" name="link">
	<select onchange="link.value=this.value; target.value='main'">
		<option>Select Map...</option>
		<%
		rst1.open "SELECT * FROM maps WHERE clientid='"&cid&"'", cnn1
		do until rst1.eof
			response.write "<option value=""maps.asp?cid="&cid&"&mapid="&rst1("id")&""">"&rst1("url")&"</option>"
			rst1.movenext
		loop
		%>
	</select></td>
<td><input type="text" name="target" size="10" value="main"></td>
<td><input type="text" name="position" size="3"></td></tr></table>
<nobr>
<input type="submit" name="action" value="Edit">
<input type="submit" name="action" value="Add Child">
<input type="Button" value="Move To..." onclick="if(nid.value!=''){if(copyselecting==1)cancelCopy();moveselecting=1;document.all['move'].style.visibility='visible'}">
<input type="Button" value="Copy To..." onclick="if(nid.value!=''){if(moveselecting==1)cancelMove();copyselecting=1;document.all['copy'].style.visibility='visible'}">
<input type="submit" name="action" value="Delete">
</nobr>
<div id="move" style="visibility:hidden">
<span style="background-color:#FFCCCC">Move Node</span><br>
<table border="0"><tr><td><input type="hidden" name="mnid" readonly>
<span id="mnid"></span></td>
<td><span id="mlabel"></span>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><span id="mlink"></span></td></tr></table>
<input type="submit" name="action" value="Move">
<input type="Button" value="Cancel Move" onclick="cancelMove()">
</div>
<div id="copy" style="visibility:hidden">
<span style="background-color:#FFCCCC">Copy Node</span><br>
<table border="0"><tr><td><input type="hidden" name="cnid" readonly>
<span id="cnnid"></span></td>
<td><span id="clabel"></span>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><span id="clink"></span></td></tr></table>
<input type="submit" name="action" value="Copy">
<input type="Button" value="Cancel Copy" onclick="cancelCopy();">
</div>
<input type="hidden" name="cid" value="<%=cid%>">
</form>
</td></tr></table>

</BODY>
</HTML>
