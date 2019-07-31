<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!--#INCLUDE file="buildxmlfunctions.asp"-->
<!--#INCLUDE file="treenode_functions.asp"-->
<%
sub screenfor(byref source, lookup, ParamValue)
	dim screenforStart, screenforEnd
	screenforStart = instr(source,lookup)
	if screenforStart<>0 then 
		screenforEnd = instr(mid(source,screenforStart),"&")
		if screenforEnd = 0 then
			source = left(source,screenforStart-1)&lookup&ParamValue
		else
			source = left(source,screenforStart-1)&lookup&ParamValue&mid(source,screenforEnd+screenforStart-1)
		end if
	end if
end sub

dim cid, action, nid, fid, link, position, target, labelid, newLabel, types, mnid, mlink, mlabel, cnid, clink, clabel, errstr, movestr, reaction, nlink, labelname, buildingconvert, bldglist, addbldglist
cid = request("cid")
action = request("action")
nid = request("nid")
fid = request("fid")
link = request("link")
position = request("position")
target = request("target")
labelid = request("labelid")
newLabel = trim(request("newLabel"))
types = request("type")
mnid = request("mnid")
mlink = request("mlink")
mlabel = request("mlabel")
cnid = request("cnid")
clink = request("clink")
clabel = request("clabel")
reaction = request("reaction")
nlink = request("nlink")
labelname = request("labelname")
buildingconvert = request("buildingconvert")
bldglist = request("bldglist")
addbldglist = request("addbldglist")

'response.write action
'response.contentType = "text/xml"
'on error goto printerror
dim cnn1, rst1, cmd, strsql, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getconnect(0,0,"engineering")
cmd.activeconnection = cnn1
if newLabel<>"" then
	labelid = addLabel(cid, newLabel, types, cnn1)
'	dim newlabel
'	newlabel = false
'	rst1.open "SELECT name, id FROM label WHERE clientid="&cid&" and lower(name)='"&lcase(newLabel)&"'", cnn1
'	newlabel = rst1.eof
'	if not newlabel then labelid = rst1("id")
'	rst1.close
'	if newlabel then 'if label does not already exist
'		strsql = "INSERT into Label (name, type, clientid) VALUES ('"&newLabel&"', '"&types&"', '"&cid&"')"
'		cmd.commandText = strsql
'		cmd.execute
'		rst1.open "SELECT id FROM label WHERE clientid="&cid&" and lower(name)='"&lcase(newLabel)&"'", cnn1
'		labelid = rst1("id")
'		rst1.close
'	end if
end if
if trim(nid)<>"" then
	if action = "Add Building To Tree" then
		strsql = "select * from facilityinfo WHERE id="&addbldglist
		rst1.Open strsql, cnn1
		dim bnid, facInfo, maintNid, lightingNid, bid, addr
		bid = rst1("id")
		addr = rst1("address")
		bnid = addBuildingNode(cid, nid, addbldglist, addr, "address", 1, 1, "", cnn1)
			facInfo = addBuildingNode(cid, bnid, bid, "Facility Info", "service", 1, 1, "editbldg.asp?id="&bid, cnn1)
				addBuildingNode cid, facInfo, bid, "Primary Contact", "service", 1, 0, "contactinfo.asp?bldg="&bid, cnn1
			addBuildingNode cid, bnid, bid, "Floor Plans", "service", 2, 0, "", cnn1
			addBuildingNode cid, bnid, bid, "Furniture Plans", "service", 3, 0, "", cnn1
			addBuildingNode cid, bnid, bid, "Reflected Ceiling", "service", 4, 0, "", cnn1
			addBuildingNode cid, bnid, bid, "Voice/Data Plan", "service", 5, 0, "", cnn1
			maintNid = addBuildingNode(cid, bnid, bid, "Maintenance", "service", 6, 1, "", cnn1)
				lightingNid = addBuildingNode(cid, maintNid, bid, "Lighting", "service", 1, 1, "", cnn1)
					addBuildingNode cid, lightingNid, bid, "Fixture Management", "service", 1, 0, "fixtypesearch.asp?cid="&cid&"&bldg="&bid, cnn1
					addBuildingNode cid, lightingNid, bid, "Lighting Management", "service", 2, 0, "floorsearch.asp?bldg="&bid, cnn1
					addBuildingNode cid, lightingNid, bid, "Lighting Reports", "service", 3, 0, "reportingindex.asp?bldg="&bid, cnn1
				addBuildingNode cid, maintNid, bid, "Mechanical PGI", "service", 2, 0, "", cnn1
				addBuildingNode cid, maintNid, bid, "Other", "service", 3, 0, "", cnn1
		cnn1.execute = "UPDATE facilityinfo SET nodeid="&bnid&" WHERE id="&addbldglist
		rst1.close
	elseif action="Edit" then
		strsql = "UPDATE nodes SET labelid='"&labelid&"', link='"&link&"', position="&position&", target='"&target&"' WHERE clientid='"&cid&"' and nodeid="&nid&""
		cmd.CommandText = strsql
		cmd.Execute
		if trim(buildingconvert) = "1" then
			rst1.open "SELECT address FROM facilityinfo WHERE id="&bldglist, cnn1
			if not rst1.eof then
				labelid = addLabel(cid, rst1("address"), types, cnn1)
				strsql = "UPDATE nodes SET labelid='"&labelid&"' WHERE clientid='"&cid&"' and nodeid="&nid&""
				cmd.CommandText = strsql
				cmd.Execute
			end if
			rst1.close
			cmd.CommandText = "sp_getnodes"
			cmd.CommandType = adCmdStoredProc
			'input params
			Set prm = cmd.CreateParameter("nodeid", adVarChar, adParamInput, 10)
			cmd.Parameters.Append prm
			cmd.Parameters("nodeid") = nid
			set rst1 = cmd.execute
			
			cmd.CommandType = adCmdText
			dim templink, rsBuildingAssign
			set rsBuildingAssign = server.createobject("ADODB.recordset")
			do until rst1.eof
				rsBuildingAssign.open "SELECT * FROM nodes WHERE nodeid="&rst1("nodeid"), cnn1, adOpenStatic, adLockOptimistic
				templink = trim(rsBuildingAssign("link"))
				screenfor templink,"bldg=",bldglist
				screenfor templink,"b=",bldglist
				screenfor templink,"building=",bldglist
				screenfor templink,"id=",bldglist
				screenfor templink,"clientid=",cid
				screenfor templink,"cid=",cid
				'response.write templink &"<br>"
				rsBuildingAssign("link").value = templink
				rsBuildingAssign.update
				rsBuildingAssign.close
				rst1.movenext
			loop
			rst1.close
	'		response.end
			
		end if
	elseif action="Add Child" or reaction = "Add Child" then
		if trim(nid)<>"" then
			strsql = "INSERT into nodes (labelid, link, fatherref, clientid, position, target) VALUES (21, '', "&nid&", '"&cid&"', 1, 'main')"
			cmd.CommandText = strsql
			cmd.Execute
			strsql = "UPDATE nodes SET relative=1 WHERE nodeid="&nid&""
			cmd.CommandText = strsql
			cmd.Execute
			rst1.open "SELECT * FROM nodes INNER JOIN label on nodes.labelid=label.id WHERE nodes.clientid='"&cid&"' ORDER BY nodeid desc", cnn1
			dim childnid, childfid, childlabelid, childposition, childtarget, childnlink, childlabelname
			childnid = rst1("nodeid")
			childfid = rst1("fatherref")
			childlabelid = rst1("labelid")
			childposition = rst1("position")
			childtarget = rst1("target")
			childnlink = rst1("link")
			childlabelname = rst1("name")
			rst1.close
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
			movestr = "Cannot move a node into one of its children."
		end if
	elseif action="Copy" then
		if cint(nid)<>0 then
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
		else
			errstr = "The root node can not be copy."
		end if
	elseif action="Delete" then
		if cint(nid)<>0 then
			cmd.CommandText = "sp_deletenode"
			cmd.CommandType = adCmdStoredProc
			'input params
			Set prm = cmd.CreateParameter("nodeid", adVarChar, adParamInput, 10)
			cmd.Parameters.Append prm
			Set cmd.ActiveConnection = cnn1
			cmd.Parameters("nodeid") = nid
			cmd.execute
		else
			errstr = "The root node can not be deleted."
		end if
	end if
elseif trim(action)<>"" then
	errstr = "Please select a node before adding, editing or deleting a node."
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
var serviceLinks = new Array();
<%
rst1.open "SELECT l.name, serviceLink FROM label l LEFT JOIN serviceLinks sl ON l.name=sl.serviceName where l.clientid="&cid&" and l.type='service' ORDER BY l.name", cnn1
dim tempservicelink, serviceIndex
serviceIndex = 0
do until rst1.eof
	tempservicelink = rst1("serviceLink")
	if not(isnull(tempservicelink)) then tempservicelink = replace(tempservicelink,"cid=","cid="&cid)
	response.write "serviceLinks["&serviceIndex&"] = """&tempservicelink&""";"&vbCrLf
	rst1.movenext
	serviceIndex = serviceIndex+1
loop
rst1.close
%>

function sendNodeInfo(nid, fid, labelid, position, target, nlink, labelname)
{	//alert("nid"+nid+"\nfid"+fid+"\mlabel"+mlabel+"\nlink"+nlink);
  movefrom = labelname;
	var frm = document.forms['node'];
	frm.newLabel.value="";
	if(moveselecting==1)
	{	if(nid!=frm.nid.value)
		{	frm.mnid.value = nid;
			document.all['mlabel'].innerText = labelname;
			hilight(nid);
		}else
		{	alert("Cannot move node to itself");
		}
	}else if(copyselecting==1)
	{	if(nid!=frm.nid.value)
		{	frm.cnid.value = nid;
			document.all['clabel'].innerText = labelname;
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
		frm.bldglist.value = '';
		frm.addbldglist.value = '';
		synchSelects();
		checkaddlabel();
		hilight(nid);
	}
	document.all['genergymenu'].expandAll();
	document.forms['node'].newLabel.value=labelname;
  document.forms["node"].nlink.value = nlink;
  document.forms["node"].labelname.value = labelname;
  document.all["err1"].style.visibility = "hidden";
  document.all["err2"].style.visibility = "hidden";
}

function checkaddlabel()
{	var str = document.forms['node'].newLabel.value;
	str = str.replace(/\s*/,'');
	if(str!='')
	{	//document.all['newLabel'].style.visibility='visible';
		document.forms['node'].region.disabled=1;
		document.forms['node'].city.disabled=1;
//		document.forms['node'].address.style.display='none';
		document.forms['node'].address.disabled=1;
		document.forms['node'].service.disabled=1;
		document.forms['node'].newLabel.value = str.replace(/'/,'');//'
	}else
	{	//document.all['newLabel'].style.visibility='hidden';
		document.forms['node'].region.disabled=0;
		document.forms['node'].city.disabled=0;
//		document.forms['node'].address.style.display='none';
		document.forms['node'].address.disabled=1;
		document.forms['node'].service.disabled=0;
	}
}

function synchSelects()
{	var frm = document.forms['node'];
//  document.all["addressmsg"].innerHTML = "";
	frm.region.value = frm.labelid.value;
	frm.city.value = frm.labelid.value;
	frm.address.value = frm.labelid.value;
//	if (frm.address.selectedIndex >= 0) { document.all["addressmsg"].innerHTML = "Change " + frm.address[frm.address.selectedIndex].innerHTML + " to: "; }
	frm.service.value = frm.labelid.value;
}

function makeActiveSelect(str)
{  
  if(moveselecting==0)
	{	
	  document.all['region'].style.visibility="hidden";
		document.all['city'].style.visibility="hidden";
		document.all['address'].style.visibility="hidden";
		document.all['service'].style.visibility="hidden";
		
		document.all['region'].style.position="absolute";
		document.all['city'].style.position="absolute";
		document.all['address'].style.position="absolute";
		document.all['service'].style.position="absolute";
//    document.all["addressmsg"].innerHTML = "";

		if (str!='default') {
      document.forms['node'].type.value=str;
      document.forms["node"].bldglist.style.display = "none";
      document.forms["node"].bldglist2.style.display = "none";
      document.forms["node"].link.value = "";
      
      switch(str){
        case "region":
          document.forms["node"].newLabel.style.display = "inline";
          break;
          
        case "city":
          document.forms["node"].newLabel.style.display = "inline";
          break;
          
        case "address":
          document.forms["node"].newLabel.style.display = "none";
          document.forms["node"].bldglist.style.display = "inline";
//          document.forms["node"].address.style.display = "none";
//	if (document.forms['node'].address.selectedIndex >= 0) { document.all["addressmsg"].innerHTML = "Change " + document.forms['node'].address[document.forms['node'].address.selectedIndex].innerHTML + " to: "; } else { document.all["addressmsg"].innerHTML = "Uh oh"; }
          
        case "service":
          document.forms["node"].newLabel.style.display = "none";
          document.all[str].style.visibility='visible';
          document.all[str].style.position="relative";
          break;
      }
    } else {	
		  makeActiveSelect('region');
		}
	}
}

function toggleTasks(str){
  switch (str){
    case "":
      hideAll();
      break;
    
    case "addchild":
      hideAll();
      document.forms["node"].reaction.value="Add Child";
      document.forms["node"].scroll.value=document.all['treewindow'].scrollTop;
      document.forms["node"].submit();
      break;
    
    case "addbldg":
      hideAll();
      document.all['bldgdiv'].style.display = "block";
      break;
    
    case "editnode":
      hideAll();
      document.all['editfields'].style.display = "block";
      document.forms["node"].buildingconvert.value=0
      document.forms["node"].buildingadd.value=0
      break;
    
    case "movenode":
      hideAll();
      if(document.forms["node"].nid.value!=''){if(copyselecting==1)cancelCopy();moveselecting=1;} 
      document.all['move'].style.display = "block";
      document.all['fromnode'].innerHTML = "&quot;" + movefrom + "&quot;";
      break;
    
    case "copynode":
      hideAll();
      if(document.forms["node"].nid.value!=''){if(moveselecting==1)cancelMove();copyselecting=1;}
      document.all['copy'].style.display = "block";
      break;
  }
}

function changeServices(idx){
  document.forms["node"].labelid.value=document.forms["node"].service.value;
  synchSelects();
  if (document.forms["node"].service[idx].value == "New") { document.forms["node"].newLabel.style.display = "inline"; } else { document.forms["node"].newLabel.style.display = "none" }
  document.forms["node"].newLabel.value="";
  re = /bldg=/;
  if (idx > 0) { idx--; } //compensate for an extra item in the services pulldown versus the serviceLinks array
  if (re.test(serviceLinks[idx]) > 0) {
    document.all["bldglist2"].style.display="inline";
  } else { 
    document.all["bldglist2"].style.display="none";  
  }
  document.forms["node"].link.value = serviceLinks[idx];
}

function hideAll(){
  document.all['editfields'].style.display = "none";
  document.all['move'].style.display = "none";
  document.all['copy'].style.display = "none";
//  document.forms["node"].addbldglist.style.display = "none"
//  document.all['addbldgbutton'].style.display = "none"
    document.all['bldgdiv'].style.display = "none";
}

function checkforNodeInfo(){
  <%if reaction = "Add Child" then%>
    sendNodeInfo('<%=childnid%>', '<%=childfid%>', '<%=childlabelid%>', '<%=childposition%>', '<%=childtarget%>', '<%=childnlink%>', '<%=childlabelname%>');
    document.all['editfields'].style.display = "block";
  <%end if%>
}

function checkForBldg(){
  retval = false;
  if (document.forms["node"].addbldglist.selectedIndex > 0) {
    retval = true;
  } else {
    alert("Please select a building to add to the tree.");
  }
  return retval;
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
		{	nodes[i].style.background='#999999';
		}
	}
}
function setscroll()
{	document.all['treewindow'].scrollTop=document.forms['node'].scroll.value;
}

function cancelCopy()
{	hilight(-1);
	copyselecting=0;
    document.all['copy'].style.display = "none";
//	document.all['copy'].style.visibility='hidden';
//	document.all['cnnid'].innerText = '';
//	document.all['cfid'].innerText = '';
	document.all['clabel'].innerText = '(Select a second node)';
//	document.all['clink'].innerText = '';
	document.node.tasks.selectedIndex = 0;
	copyselecting = 0;
}

function cancelMove()
{	hilight(-1);
	moveselecting=0;
    document.all['move'].style.display = "none";
//	document.all['move'].style.visibility='hidden';
//	document.all['mnid'].innerText = '';
//	document.all['mfid'].innerText = '';
	document.all['mlabel'].innerText = '(Select a second node)';
//	document.all['mlink'].innerText = '';
	document.node.tasks.selectedIndex = 0;
	moveselecting = 0;
}

function cancelEdit()
{	hilight(-1);
	moveselecting=0;
	document.all['editfields'].style.display='none';
	document.node.tasks.selectedIndex = 0;
}

function confirmDelete(){
  if(document.forms["node"].nid.value!="0") {
    retval = window.confirm("Are you sure you want to delete this item?");
  } else { 
    retval = false;
    alert("Sorry, you can't delete this node.");
  }
  return retval;
}

</SCRIPT>
<link rel="Stylesheet" href="/genergy2/styles.css">
<script src="messages.js" type="text/javascript" language="Javascript1.2"></script>
</HEAD><!--  -->
<BODY bgcolor="#FFFFFF" onload="document.all['genergymenu'].expandAll(document.all['treewindow'],document.forms['node'].scroll.value);setTimeout('setscroll()',200);checkforNodeInfo();">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff;">
<tr bgcolor="#0099ff">
	<td><span class="standardheader"><%=clientname%> | Configure Menu</span></td>
	<td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=cid%>'" class="standard">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=cid%>'" class="standard" disabled>&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=cid%>'" class="standard"></td>
</tr>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="0">
<tr>
<td valign="top" bgcolor="#999999">
	<div id="treewindow" style="overflow:auto;width:250;height:462;padding:3px;">
	<%Response.Write xmlobj.transformNode (xslobj)%>
	</div>
</td>
<td width="16"><span class="standard">&nbsp;</span></td>
<td valign="top" width="90%"><form name="node" method="post" onsubmit="scroll.value=document.all['treewindow'].scrollTop;">

<div id="err1" class="standard" style="color:#cc0000;visibility:visible;"><%=errstr%></div>
<table border=0 cellpadding="5" cellspacing="0" width="100%">
<tr>
	<td colspan="2" bgcolor="#cccccc"><span class="standard"><b>Organize and Edit Your Facilities Portfolio</b></span></td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
  <td width="18"><img src="images/num_one.gif" alt="1" align="left" width="13" height="13" hspace="2" border="0"></td>
	<td><span class="standard">Click a node in the tree structure at left <a onMouseOut="closeHelpBox()" onMouseOver="helpbox('click_node_first',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a><br></td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
  <td width="18"><img src="images/num_two.gif" alt="2" align="left" width="13" height="13" hspace="2" border="0"></td>
	<td><span class="standard">
	<select name="tasks" onchange="toggleTasks(this.value);">
	<option value="">Select an action
	<option value="addbldg">Add building
	<option value="addchild">Add new node
	<option value="editnode">Edit/delete node
	<option value="movenode">Move node
	<option value="copynode">Copy node
	</select>
	<input type="hidden" name="reaction" value="">
	</span>
  <a onMouseOut="helpup('help_select_tree_action');" onMouseOver="helpdrop('help_select_tree_action','select_tree_action');"><img name="help_select_tree_action_img" src="images/question-rt.gif" alt="?" title="" width="22" height="13" hspace="4" border="0"></a>
  <div id="help_select_tree_action" class="standard" style="display:'none';margin-right:40px;line-height:10pt;padding-top:6px;padding-bottom:6px;"></div>
	
	</td>
</tr>
</table>
<div id="bldgdiv" style="display:none;">	
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr bgcolor="#eeeeee">
  <td width="18"><img src="images/num_three.gif" alt="3" align="left" width="13" height="13" hspace="2" border="0"></td>
	<td><span class="standard">
  <!-- begin building list -->
  <select name="addbldglist" id="addbldglist" onchange="buildingadd.value=1;buildingconvert.value=0;synchSelects();document.forms['node'].newLabel.value='';">
  <option value="" selected>Select a building</option>
    <%
    rst1.open "SELECT address, id FROM facilityinfo WHERE clientid='"&cid&"'", cnn1
    do until rst1.eof
      response.write "<option value="""&rst1("id")&""">"&rst1("address")&"</option>"
      rst1.movenext
    loop
    rst1.close
    %>
  </select>
  <a onMouseOut="closeHelpBox()" onMouseOver="helpbox('add_bldg_to_tree',event.x,event.y)"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a>
  </td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18"><img src="images/num_four.gif" alt="4" align="left" width="13" height="13" hspace="2" border="0"></td>
	<td><span class="standard">
  <input type="submit" name="action" value="Add Building To Tree" onclick="return checkForBldg();" class="standard">
  <input type="hidden" name="buildingadd" value="0">
  <!-- end building list -->
	
	<!--input type="submit" name="action" value="Add Child" class="standard" style="margin-top:8px;"-->
	<!--input type="button" id="taskbutton" value="Submit" onclick="toggleTasks(document.forms['node'].tasks[document.forms['node'].tasks.selectedIndex].value)" class="standard"-->
	</span>
	</td>
</tr>
</table>
</div>

<span id="err2" class="standard" style="color:#cc0000;visibility:visible;"><%=movestr%></span>

<!-- begin move fields -->
<div id="move" style="display:none">
<table border=0 cellpadding="5" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee">
  <td colspan="2"><span class="standard">Move node <span id="fromnode">NO NODE SELECTED</span> below:</span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18"><img src="images/num_three.gif" alt="3" align="left" width="13" height="13" hspace="2" border="0"></td>
  <td width="100%"><span id="mlabel" class="standard" style="font-weight:bold;">(Select a second node)</span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18"><img src="images/num_four.gif" alt="4" align="left" width="13" height="13" hspace="2" border="0"></td>
  <td>
  <span id="mlink" class="standard"></span>
  <input type="hidden" name="mnid" readonly>
  <span id="mnid"></span>
  <input type="submit" name="action" value="Move" class="standard" style="padding-left:4px;padding-right:4px;">&nbsp;
  <input type="Button" value="Cancel Move" onclick="cancelMove()" class="standard">
  </td>
</tr>
</table>
</div>
<!-- end move fields -->

<!-- begin copy fields -->
<div id="copy" style="display:none">
<table border=0 cellpadding="5" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee">
  <td colspan="2"><span class="standard">Duplicate this node below:</span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18"><img src="images/num_three.gif" alt="3" align="left" width="13" height="13" hspace="2" border="0"></td>
  <td><span id="clabel" class="standard" style="font-weight:bold;">(Select a second node)</span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td width="18"><img src="images/num_four.gif" alt="4" align="left" width="13" height="13" hspace="2" border="0"></td>
  <td>
  <input type="submit" name="action" value="Copy" class="standard" style="padding-left:4px;padding-right:4px;">&nbsp;&nbsp;<input type="Button" value="Cancel Copy" onclick="cancelCopy();" class="standard">&nbsp;
  <input type="hidden" name="cnid" readonly><span id="cnnid" class="standard"></span>
  </td>
</tr>
</table>
</div>
<!-- end copy fields -->

<!-- begin Edit/delete node fields -->
<div id="editfields" style="display:none">
<table border=0 cellpadding="5" cellspacing="1" width="100%">
<tr bgcolor="#cccccc">
	<td colspan="2">
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="18"><img src="images/num_three.gif" alt="3" align="left" width="13" height="13" hspace="2" border="0"></td>
    <td width="10"></td>
    <td width="200"><span class="standard"><b>Edit Node</b></span></td>
    <td align="right"><span class="standard" style="color:#666666">Node&nbsp;id: <input type="text" name="nid" size="5" readonly> &nbsp;Parent&nbsp;id <input type="text" name="fid" size="5" readonly></span></td>
  </tr>
  </table>
	</td>
</tr>

<tr valign="top" bgcolor="#eeeeee">
  <td width="18%" align="right"><span class="standard">Node Type</span></td>
	<td bgcolor="#eeeeee">
  <!-- begin node type -->
	<span class="standard">	
  <input type="hidden" name="scroll" value="<%=request("scroll")%>">
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
  <select name="type" id="type" onchange="makeActiveSelect(this.value);">
  <option value="region">Region</option>
  <option value="city">City</option>
  <option value="address">Building</option>
  <option value="service">Service</option>
  </select><a onMouseOut="helpup('help_node_type');" onMouseOver="helpdrop('help_node_type','node_type');"><img name="help_node_type_img" src="images/question-rt.gif" alt="?" title="" width="22" height="13" hspace="4" border="0"></a>
  <div id="help_node_type" class="standard" style="display:'none';margin-right:40px;padding-top:6px;padding-bottom:6px;"></div>
  <!-- end node type -->
  </td>
</tr>

<tr bgcolor="#eeeeee">
  <td width="18%" align="right"><span class="standard">Label</span></td>
	<td bgcolor="#eeeeee">
  <!-- begin label pulldowns -->
	<span class="standard">
  <select name="region" id="region" style="visibility:hidden;position:absolute;" onchange="labelid.value=this.value; synchSelects()">
  <%=labelregion%>
  </select>
  <select name="city" id="city" style="visibility:hidden;position:absolute;" onchange="labelid.value=this.value; synchSelects()">
  <%=labelcity%>
  </select>
  <select name="address" id="address" disabled style="visibility:hidden;position:absolute;display:'inline'" onChange="labelid.value=this.value; synchSelects()">
  <%=labeladdress%>
  </select>
  <select name="service" id="service" style="visibility:hidden;position:absolute;margin-right:6px;" onchange="changeServices(this.selectedIndex);">
  <OPTGROUP label='Service Options'>
  <option value="New">Add new service
  </OPTGROUP>
  <OPTGROUP label='Existing Services'>
  <%=labelservice%>
  </OPTGROUP>
  </select>
  <!-- end label pulldowns -->
  
  
  <!-- begin building list -->

  <select name="bldglist" id="bldglist" style="display:none;margin-right:6px;" onchange="buildingconvert.value=1;buildingadd.value=0;synchSelects();document.forms['node'].newLabel.value='';">
    <option>Select Building...</option>
    <%
    rst1.open "SELECT address, id FROM facilityinfo WHERE clientid='"&cid&"'", cnn1
    do until rst1.eof
      response.write "<option value="""&rst1("id")&""">"&rst1("address")&"</option>"
      rst1.movenext
    loop
    rst1.close
    %>
  </select>

  <!-- end building list -->
  
  <input type="hidden" name="buildingconvert" value="0">
  <!-- begin label text field (newLabel) -->
  <input type="text" name="newLabel" onKeyUp="checkaddlabel()" style="display:'inline';">&nbsp;<!--input type="button" name="bldgbutton" value="Edit Facility Information..." onclick="location='updatebldg.asp?id=&labelname=<%=labelname%>'" class="standard" style="display:none;"-->
  <!-- end label text field (newLabel) -->
  <input type="Hidden" name="labelid" value=""><a onMouseOut="closeHelpBox();document.forms['node'].selectmap.style.display='inline';" onMouseOver="helpbox('whats_a_label',event.x,event.y);document.forms['node'].selectmap.style.display='none';"><img src="images/question.gif" alt="?" title="" width="13" height="13" hspace="4" border="0"></a>
  </td>
</tr>

<tr valign="top" bgcolor="#eeeeee">
  <td align="right"><span class="standard">Link</span></td>
  <td><input type="text" name="link">
  <!-- begin building list 2 -->
  <select name="bldglist2" id="bldglist2" style="display:none;">
  <option>Select Building...</option>
  <%
  rst1.open "SELECT address, id FROM facilityinfo WHERE clientid='"&cid&"'", cnn1
  do until rst1.eof
    response.write "<option value="""&rst1("id")&""">"&rst1("address")&"</option>"
    rst1.movenext
  loop
  rst1.close
  %>
  </select>
  <!-- end building list 2 -->
  
  <!-- begin map select -->
  <span class="standard">or&nbsp;</span>
	<select name="selectmap" onchange="link.value=this.value; target.value='main'" style="display:'inline';">
  <option>Select Map...</option>
  <%
  rst1.open "SELECT * FROM maps WHERE clientid='"&cid&"'", cnn1
  do until rst1.eof
    response.write "<option value=""maps.asp?cid="&cid&"&mapid="&rst1("id")&""">"&rst1("url")&"</option>"
    rst1.movenext
  loop
  rst1.close
  %>
  </select><a onMouseOut="helpup('help_link');" onMouseOver="helpdrop('help_link','link');"><img name="help_link_img" src="images/question-rt.gif" alt="?" title="" width="22" height="13" hspace="4" border="0"></a>
  <!-- end map select -->
  <div id="help_link" class="standard" style="display:'none';margin-right:40px;padding-top:6px;padding-bottom:6px;"></div>
  </td>	
</tr>
<tr bgcolor="#eeeeee">
  <td align="right"><span class="standard">Target</span></td>
  <td><input type="text" name="target" size="10" value="main"></td>
</tr>
<tr bgcolor="#eeeeee">
  <td align="right"><span class="standard">Position</span></td>
  <td><input type="text" name="position" size="3"></td>
</tr>
<tr bgcolor="#cccccc">
  <td width="18%"><span class="standard">&nbsp;</span></td>
	<td>
	<input type="submit" name="action" value="Edit" class="standard" style="padding-left:6px;padding-right:6px;"> 
	<input type="submit" name="action" value="Delete" onclick="return confirmDelete();" class="standard" style="padding-left:6px;padding-right:6px;">
  <input type="button" value="Cancel Edit" onclick="cancelEdit()" class="standard">
	</td>
</tr>
</table>
</div>
<!-- end Edit/delete node fields -->

<input type="hidden" name="nlink" value="">
<input type="hidden" name="labelname" value="">
</form>
</td></tr></table>

<!--#INCLUDE FILE="helpbox.htm"-->
</BODY>
</HTML>
