<%'option explicit

Sub loadNewXML(username)
  Dim xml
  Set xml = server.createobject("MSXML2.FreeThreadedDOMDocument")
  Dim root
  Dim node
  Set root = xml.createNode("element", "root", "")
  Set node = xml.createNode("element", "userinfo", "")
  node.setAttribute "username", username
  root.appendChild node
  Set node = xml.createNode("element", "dataservers", "")
  root.appendChild node
  Set node = xml.createNode("element", "keys", "")
  root.appendChild node
  Set node = xml.createNode("element", "groups", "")
  root.appendChild node
  Set node = xml.createNode("element", "views", "")
  root.appendChild node
  xml.appendChild root
  set session("xmlUserObj") = xml
End Sub

Function getXMLUserName()
  Dim node, xml
  Set xml = getXmlSession()
  Set node = xml.selectSingleNode("//userinfo")
  getXMLUserName = node.Attributes.getNamedItem("username").Text
End Function

Sub setBuilding(bldgnum, ip,pid,nodename,sqlport)
On Error Resume Next
  'Proper handling of NULL values in pid
  if IsNull(pid) or  pid = "" then pid = 0
  Dim node, dataservers, xml
  Set xml = getXmlSession()
  Set dataservers = xml.selectSingleNode("//dataservers")
  Set node = xml.createNode("element", "building", "")
  node.setAttribute "bldgnum", bldgnum
  node.setAttribute "name", nodename
  node.setAttribute "ip", ip
  node.setAttribute "sqlport", sqlport
  node.setAttribute "offline", 0
  node.setAttribute "pid", pid
  dataservers.appendChild node
  If Err.number <> 0 then
   'response.Write Err.description & "--> pid = " & pid & " bldgnum= " & bldgnum & " ip = " & ip & " Port=" & sqlport
   'response.End()
  End If

End Sub

Function getAllBuildings()
  Dim nodes, i, xml
  Set xml = getXmlSession()
  Set nodes = xml.selectNodes("//building")
  Dim outarray()
  ReDim outarray(nodes.length, 2)
  For i = 0 To nodes.length - 1
    outarray(i, 0) = nodes.Item(i).Attributes.getNamedItem("bldgnum").Text
    outarray(i, 1) = nodes.Item(i).Attributes.getNamedItem("ip").Text
    outarray(i, 3) = nodes.Item(i).Attributes.getNamedItem("offline").Text
  Next
  getAllBuildings = outarray
End Function

Function isBuildingOff(bldgnum)
  isBuildingOff = false
  Dim node, xml
  set xml = getXmlSession()
  set node = xml.selectSingleNode("//building[@bldgnum='"&bldgnum&"']")
  If TypeName(node) <> "Nothing" Then 
  	if TypeName(node.Attributes.getNamedItem("offline")) <> "Nothing" then
	    if node.Attributes.getNamedItem("offline").Text = "1" then isBuildingOff = true
	end if
  end if
End Function

Function getBuildingIP(bldgnum)
	getBuildingIP =  getBuildingIPCom(bldgnum)
End Function

Function getBuildingIPCom(bldgnum)
  getBuildingIPCom = 0
  Dim node, xml, rstB
  set rstB = server.createobject("ADODB.recordset")
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//building[translate(@bldgnum,'"&ucase(bldgnum)&"','"&lcase(bldgnum)&"')='"&lcase(bldgnum)&"']")
  If TypeName(node) = "Nothing" Then 
'    if hasGroup("Genergy Users") then
      rstB.open "SELECT ip as ipout,pid,strt,isnull(sqlport,'1433') as sqlport from super_main m inner join portfolio p on p.id = m.pid left join buildings b on b.bldgnum = m.bldgnum WHERE m.bldgnum='"&bldgnum&"'", getConnect(0,0,"dbCore")
      if not rstB.eof then
        if trim(rstB("ipout"))<>"" then
          getBuildingIPCom = rstB("ipout")
          setBuilding bldgnum, getBuildingIPCom,rstB("pid"),rstb("strt"),rstb("sqlport")
        end if
      end if
      rstB.close
'    end if
    if trim(getBuildingIPCom)="" or trim(getBuildingIPCom)="0" then
      response.write "Error getting building info for '"&bldgnum&"'. Your Session is not configured for this building."
      response.end
    end if
  Else 
    getBuildingIPCom = node.Attributes.getNamedItem("ip").Text
  end if
End Function

Function getBuildingPortMain(bldgnum)
  getBuildingPortMain = 0
  Dim node, xml, rstB
  set rstB = server.createobject("ADODB.recordset")
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//building[translate(@bldgnum,'"&ucase(bldgnum)&"','"&lcase(bldgnum)&"')='"&lcase(bldgnum)&"']")
  If TypeName(node) = "Nothing" Then 
'    if hasGroup("Genergy Users") then
      rstB.open "SELECT ip as ipout,pid,strt,isnull(sqlport,'1433') as sqlport from super_main m inner join portfolio p on p.id = m.pid left join buildings b on b.bldgnum = m.bldgnum WHERE m.bldgnum='"&bldgnum&"'", getConnect(0,0,"dbCore")
      if not rstB.eof then
        if trim(rstB("ipout"))<>"" then
          getBuildingPortMain = rstB("sqlport")
          setBuilding bldgnum, getBuildingIPCom,rstB("pid"),rstb("strt"),rstb("sqlport")
        end if
      end if
      rstB.close
'    end if
    if trim(getBuildingPortMain)="" or trim(getBuildingPortMain)="0" then
      response.write "Error getting building info for '"&bldgnum&"'. Your Session is not configured for this building."
      response.end
    end if
  Else 
    getBuildingPortMain = node.Attributes.getNamedItem("sqlport").Text
  end if
End Function


function getAllBuildingIP()
	dim outarray(1)
	outarray(0) = application("superIP")
	getAllBuildingIP = outarray
end function

function getAllBuildingIPCom()
	Dim nodes, i, xml, temp
	Set xml = getXmlSession()
	temp = ""
	Set nodes = xml.selectNodes("//building")'[not(following::building/@bldgnum=@bldgnum)]/@ip")
	Dim outarray
	For i = 0 To nodes.length - 1
		if instr(temp, nodes.Item(i).Attributes.getNamedItem("ip").Text)=0 then temp = temp & trim(nodes.Item(i).Attributes.getNamedItem("ip").Text) & ","
	Next
	if temp="" then 
		response.write "There are no buildings loaded in your session."
		response.end
	end if
	outarray = split(left(temp, len(temp)-1),",")
	getAllBuildingIPCom = outarray
end function

Sub setPortfolio(pid, ip,nodename,sqlport)
	Dim node, dataservers, xml
	Set xml = getXmlSession()
	Set dataservers = xml.selectSingleNode("//dataservers")
	Set node = xml.createNode("element", "portfolio", "")
	node.setAttribute "pid", pid
	node.setAttribute "ip", ip
	node.setAttribute "sqlport", sqlport
	node.setAttribute "name", nodename	
	dataservers.appendChild node
End Sub
Function getPIDPortMain(pid)
  getPIDPortMain = 0
  Dim node, xml, rstB
  set rstB = server.createobject("ADODB.recordset")
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  Set node = xml.selectSingleNode("//portfolio[translate(@pid,'"&ucase(pid)&"','"&lcase(pid)&"')='"&lcase(pid)&"']")
  If TypeName(node) = "Nothing" Then 
	getPIDPortMain = "1433"
  Else 
    getPIDPortMain = node.Attributes.getNamedItem("sqlport").Text
  end if
  
    if trim(getPIDPortMain)="" or trim(getPIDPortMain)="0" then
      response.write "Error getting building info for '"&bldgnum&"'. Your Session is not configured for this building."
      response.end
    end if
End Function

Function getPortfolio()
	getPortfolio=getKeyValue("pid")
end function

function getPID(bldgnum)
  getPID = 0
  Dim node, xml, rstB
  set rstB = server.createobject("ADODB.recordset")
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//building[translate(@bldgnum,'"&ucase(bldgnum)&"','"&lcase(bldgnum)&"')='"&lcase(bldgnum)&"']")
  If TypeName(node) = "Nothing" Then 

	rstB.open "SELECT ip as ipout,pid,name,isnull(sqlport,'1433') as sqlport from super_main m inner join portfolio p on p.id = m.pid left join buildings b on b.bldgnum = m.bldgnum WHERE m.bldgnum='"&bldgnum&"'", Application("dbdefault") & "dbCore;"

      if not rstB.eof then
        if trim(rstB("ipout"))<>"" then
	      'Proper Handling of NULL
		   getPID = rstB("pid")
           setBuilding bldgnum, rstB("ipout"),getPID, rstB("name"),rstB("sqlport")
        end if
      end if
      rstB.close
  Else 
    getPID = node.Attributes.getNamedItem("pid").Text
  end if

end function

Function getPidIP(pid)
	getPidIP=getPidIPMain(pid)
End Function

function getIP(pid, bldg)
	if pid <> "" then 
		getIP=getPidIP(pid)
	elseif bldg <> "" then 
		getIP=getBuildingIP(bldg)
	end if 
end function
function getPort(pid,bldg)
	if pid <> "" then 
		getPort=getPIDPortMain(pid)
	elseif bldg <> "" then 
		getPort=getBuildingPortMain(bldg)
	end if 
end function 
Function getPidIPMain(pid)
  Dim node, xml
  set xml = getXmlSession()
   
  set node = xml.selectSingleNode("//portfolio[@pid='" & pid & "']")

   If TypeName(node) = "Nothing" Then 
    getPidIPMain = Application("CoreIP")
   Else 
    getPidIPMain = node.Attributes.getNamedItem("ip").Text
  end if
End Function

Function printXML()
  Dim xml, node
  Set xml = getXmlSession()
  Set node = xml.selectSingleNode("/")
  printXML = node.xml
End Function

Function getXmlSession()
  Checkxml()
  set getXmlSession = session("xmlUserObj")
End Function
Function CheckEmptyxml()
  if isempty(session("xmlUserObj")) then
	CheckEmptyxml = false
  else 
    CheckEmptyxml = true	
  end if
end Function

Sub Checkxml()
  if isempty(session("xmlUserObj")) then
    response.write "There is no user session established, please login again."
	'dim name
	'For Each name In Request.ServerVariables
	'	response.write name&" = "&Request.ServerVariables(name)&"</br>"
	'next
	'dim k
	'dim l
	'l = Session.Contents.Count
	'Response.Write("Session variables: " & l & "<br>") 
	'For k=1 to l
	'  Response.Write(Session.Contents(k) & "<br>") 
	'Next
    'response.end
  end if
end Sub

Sub setKeyValue(key, Value)
  Dim node, keys, xml
  Set xml = getXmlSession()
  Set keys = xml.selectSingleNode("//keys")
  Set node = xml.selectSingleNode("//keys/key[@name='"&key&"']")
  if isnull(value) then value = ""
  if typename(node) = "Nothing" then
    Set node = xml.createNode("element", "key", "")
    node.setAttribute "name", key
    node.setAttribute "value", value
    keys.appendChild node
  else
    node.setAttribute "name", key
    node.setAttribute "value", value
  end if
End Sub

Sub setBuildingOffline(bldgnum, value)
  Dim node, keys, xml
  Set xml = getXmlSession()
  Set keys = xml.selectSingleNode("//building")
  Set node = xml.selectSingleNode("//building[@bldgnum='"&bldgnum&"']")
  if typename(node) <> "Nothing" then
    node.setAttribute "offline", value
  end if
End Sub

Function getKeyValue(keyname)
  Dim node, xml
  set xml = getXmlSession()
  set node = xml.selectSingleNode("//key[@name='" & keyname & "']")
  If TypeName(node) = "Nothing" Then getKeyValue = "" Else getKeyValue = node.Attributes.getNamedItem("value").Text
End Function

Sub setGroup(groupname)
  Dim node, groups, xml
  Set xml = getXmlSession()
  Set groups = xml.selectSingleNode("//groups")
  Set node = xml.selectSingleNode("//groups/group[@name='"&groupname&"']")
  if typename(node) = "Nothing" then
    Set node = xml.createNode("element", "group", "")
    node.setAttribute "name", groupname
    groups.appendChild node
  else
    node.setAttribute "name", groupname
  end if
End Sub

Function hasGroup(groupname)
  Dim node, xml
  set xml = getXmlSession()
  set node = xml.selectSingleNode("//group[@name='"&groupname&"']")
  If TypeName(node) = "Nothing" Then hasGroup = false Else hasGroup = true
End Function

Sub setView(viewname, ip, pid)
  if isnull(viewname) then viewname = "ERR" end if
  if isnull(ip) then ip = "ERR" end if
  if isnull(pid) then pid = "ERR" end if

  Dim node, views, xml
  Set xml = getXmlSession()
  Set views = xml.selectSingleNode("//views")
  Set node = xml.createNode("element", "view", "")
  node.setAttribute "viewname", viewname
  node.setAttribute "ip", ip
  node.setAttribute "pid", pid
  views.appendChild node
End Sub

Function getViewIP(viewname)
	getViewIP = application("superIP")
End Function

Function getViewIPCom(viewname)
  Dim node, xml
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  if viewname = "" then 
  	getViewIPCom = 0 
  else 
	  set node = xml.selectSingleNode("//view[translate(@viewname,'"&ucase(viewname)&"','"&lcase(viewname)&"')='"&lcase(viewname)&"']")
	  If TypeName(node) = "Nothing" Then 
		getViewIPCom = 0 
		response.write "Your account was not configured to view '"&viewname&"'. Try logging in again or contact Genergy's IT department."
		response.end
	  Else 
		getViewIPCom = node.Attributes.getNamedItem("ip").Text
	  end if
  end if
End Function

Function getViewPortfolio(viewname)
	getViewPortfolio = application("superIP")
End Function

Function getViewPortfolio(viewname)
  Dim node, xml
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//view[translate(@viewname,'"&ucase(viewname)&"','"&lcase(viewname)&"')='"&lcase(viewname)&"']")
  If TypeName(node) = "Nothing" Then 
    getViewPortfolio = 0
  Else 
    getViewPortfolio = node.Attributes.getNamedItem("pid").Text
  end if
End Function

sub deleteView(viewname)
  Dim node, xml, parent
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set parent = xml.selectSingleNode("//views")
  set node = xml.selectSingleNode("//view[translate(@viewname,'"&ucase(viewname)&"','"&lcase(viewname)&"')='"&lcase(viewname)&"']")
  If TypeName(node) <> "Nothing" Then 
	parent.removeChild node
  end if
End sub
%>
