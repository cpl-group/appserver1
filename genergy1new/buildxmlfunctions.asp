<%
dim logo, rstClient, clientname, level, treepath(20,3), branch, leaf, treerecord, nodelength

sub copyresults()
	nodelength = rst1.RecordCount
	redim treerecord(nodelength,12)
	dim treeindex
	treeindex=0
	do until rst1.eof
		treerecord(treeindex,0) = rst1("labelid")
		treerecord(treeindex,1) = rst1("link")
		treerecord(treeindex,2) = rst1("fatherref")
		treerecord(treeindex,3) = rst1("relative")
		treerecord(treeindex,4) = rst1("clientid")
		treerecord(treeindex,5) = rst1("nodeid")
		treerecord(treeindex,6) = rst1("target")
		treerecord(treeindex,7) = rst1("position")
		treerecord(treeindex,8) = rst1("id")
		treerecord(treeindex,9) = rst1("Name")
		treerecord(treeindex,10) = rst1("type")
		treerecord(treeindex,11) = rst1("clientid")
		
		treeindex = treeindex+1
		rst1.movenext
	loop
end sub

sub buildtree(byref xmlobj, byref treerecord)
	dim genergymenu
	set genergymenu = xmlobj.createNode("element", "genergymenu", "")
	set rstClient = server.createobject("ADODB.Recordset")
	rstClient.open "SELECT * FROM clients WHERE id="&cid, cnn1
	if not rstClient.eof then
		genergymenu.setAttribute "Corp_name", rstClient("Corp_name")
		genergymenu.setAttribute "address", rstClient("address")
		genergymenu.setAttribute "city", rstClient("city")
		genergymenu.setAttribute "state", rstClient("state")
		genergymenu.setAttribute "zip", rstClient("zip")
		genergymenu.setAttribute "logo", rstClient("logo")
		genergymenu.setAttribute "contact", rstClient("contact")
		genergymenu.setAttribute "contactPhone", rstClient("contactPhone")
	end if
	logo = rstClient("logo")
	clientname = rstClient("Corp_name")
	rstClient.close
	
	level = 0
	treepath(level,0) = 0
	treepath(level,1) = 1
	set treepath(level,2) = genergymenu
	xmlobj.appendChild(treepath(level,2))
	if nodelength<>0 then
		findchild()
	end if
end sub

sub findchild()
	do until level=-1
		dim inc, count
		inc = 0
		count=0
		do while not(inc=nodelength or count = treepath(level,1))
			if treerecord(inc,2)=treepath(level,0) then
				count = count + 1
			end if
			if count < treepath(level,1) then inc = inc + 1
		loop
	'	response.write "abouttocheck:"
		if not inc=nodelength then 'make node of specified sibling number in treepath(level,1)
			makenode(inc)
		else 'doesn't have this sibling number so must go back a level to check if there is a next sibling there
			treepath(level,0) = ""
			treepath(level,1) = ""
			set treepath(level,2) = nothing
			level = level-1
			if level>-1 then 'if has -1 level it means the hierarchy has ended, xml should be then complete
				treepath(level,1) = treepath(level,1)+1
			end if
		end if
	loop
end sub

sub makenode(inc)
	'response.write inc&"<br>"
'	dim i
'	for i = 0 to level
'		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
'	next
'	response.write treerecord(inc,9)&"|"&treerecord(inc,1)&"("&treepath(level,0)&"|"&treepath(level,1)&")<br>"
	if treerecord(inc,3)="True" then 'has a child or two go get them
		set branch = xmlobj.createNode("element", "branch", "")
		branch.setAttribute "label", treerecord(inc,9)
		branch.setAttribute "fid", treerecord(inc,2)
		branch.setAttribute "nid", treerecord(inc,5)
		branch.setAttribute "lid", treerecord(inc,0)
		branch.setAttribute "type", treerecord(inc,10)
		branch.setAttribute "target", treerecord(inc,6)&""
		branch.setAttribute "position", treerecord(inc,7)
		if trim(treerecord(inc,1))<>"" then branch.setAttribute "link", treerecord(inc,1)
		treepath(level,2).appendChild(branch)
		level = level+1
		'response.write "level("&level&")"
		treepath(level,0) = cLNG(treerecord(inc,5))
		treepath(level,1) = 1
		set treepath(level,2) = branch
'		response.write treepath(level,0)&"|"&treepath(level,1)&"|level:"&level
	else 'is a leaf with no children
		set branch = xmlobj.createNode("element", "leaf", "")
		branch.setAttribute "label", treerecord(inc,9)
		branch.setAttribute "fid", treerecord(inc,2)
		branch.setAttribute "nid", treerecord(inc,5)
		branch.setAttribute "lid", treerecord(inc,0)
		branch.setAttribute "type", treerecord(inc,10)
		branch.setAttribute "target", treerecord(inc,6)&""
		branch.setAttribute "position", treerecord(inc,7)
		if trim(treerecord(inc,1))<>"" then branch.setAttribute "link", treerecord(inc,1)
		treepath(level,2).appendChild(branch)
		treepath(level,1) = cINT(treepath(level,1))+1
	end if
end sub
%>