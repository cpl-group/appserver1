<%
'################################################'
'# These functions are for adding labels and    #'
'# nodes in the tree structure                  #'
'################################################'

'############## addBuildingNode ##############'
'bid = building id-->taken from last update
'cid = client id -->passed from node setup page
'nid = node id -->passed from node setup page
'label = node label -->passed from node setup page
'labeltype = address|service|region|city
'position = node position
'relative = boolean --> whether ndoe has children
'link = node hyperlink if has one
'cnn1 = connection object
'
'adds a hard coded default tree 
'to the node specified
function addBuildingNode(cid, nid, bid, label, labeltype, position, relative, link, byref cnn1)
	dim cmNodes, addlabelid, rsNodes
	set cmNodes = server.createobject("ADODB.command")
	addlabelid = addLabel(cid, label, labeltype, cnn1)
	cmNodes.activeConnection = cnn1
	cmNodes.commandText = "INSERT into nodes (labelid, link, fatherref, clientid, position, target, relative) VALUES ("&addlabelid&", '"&link&"', "&nid&", '"&cid&"', "&position&", 'main', "&relative&")"
	cmNodes.execute
	
	set rsNodes = server.createobject("ADODB.recordset")
	rsNodes.open "SELECT max(nodeid) as nodeid FROM nodes WHERE clientid="&cid, cnn1
	addBuildingNode = cInt(trim(rsNodes("nodeid")))
	rsNodes.close
end function
'################################################'

'################### addLabel ###################'
'cid = client id -->passed from node setup page
'label = node id -->the label string
'labeltype = address|service|region|city
'cnn1 = connection object
'
'checks to see if specified label already
'exsists, and adds if it does not
'returns label id
function addLabel(cid, label, labeltype, byref cnn1)
	dim rsLabel, cmLabel, ALstrsql
	set rsLabel = server.createobject("ADODB.recordset")
	rsLabel.open "SELECT name, id FROM label WHERE clientid="&cid&" and lower(name)='"&lcase(label)&"'", cnn1
	if rsLabel.eof then
		set cmLabel = server.createobject("ADODB.command")
		cmLabel.activeconnection = cnn1
		ALstrsql = "INSERT into Label (name, type, clientid) VALUES ('"&label&"', '"&labeltype&"', '"&cid&"')"
		cmLabel.commandText = ALstrsql
		cmLabel.execute
		rsLabel.close
		rsLabel.open "SELECT id FROM label WHERE clientid="&cid&" and lower(name)='"&lcase(label)&"'", cnn1
		addLabel = rsLabel("id")
		rsLabel.close
	else
		addLabel = trim(rsLabel("id"))
		rsLabel.close
	end if
end function
'################################################'
%>