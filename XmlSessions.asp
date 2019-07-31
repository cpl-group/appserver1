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

Sub setBuilding(bldgnum, ip)
  Dim node, dataservers, xml
  Set xml = getXmlSession()
  Set dataservers = xml.selectSingleNode("//dataservers")
  Set node = xml.createNode("element", "building", "")
  node.setAttribute "bldgnum", bldgnum
  node.setAttribute "ip", ip
  dataservers.appendChild node
End Sub

Function getAllBuildings()
  Dim nodes, i, xml
  Set xml = getXmlSession()
  Set nodes = xml.selectNodes("//building")
  Dim outarray()
  ReDim outarray(nodes.length, 1)
  For i = 0 To nodes.length - 1
    outarray(i, 0) = nodes.Item(i).Attributes.getNamedItem("bldgnum").Text
    outarray(i, 1) = nodes.Item(i).Attributes.getNamedItem("ip").Text
  Next
  getAllBuildings = outarray
End Function

Function getBuildingIP(bldgnum)
  Dim node, xml
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//building[translate(@bldgnum,'"&ucase(bldgnum)&"','"&lcase(bldgnum)&"')='"&lcase(bldgnum)&"']")
  If TypeName(node) = "Nothing" Then 
    getBuildingIP = 0 
'    response.write "Error getting building info for '"&bldgnum&"'. Your Session is not configured for this building."
'    response.end
  Else 
    getBuildingIP = node.Attributes.getNamedItem("ip").Text
  end if
End Function

function getAllBuildingIP()
  Dim nodes, i, xml, temp
  Set xml = getXmlSession()
  temp = ""
  Set nodes = xml.selectNodes("//building")'[not(following::building/@bldgnum=@bldgnum)]/@ip")
  Dim outarray
  For i = 0 To nodes.length - 1
    if instr(temp, nodes.Item(i).Attributes.getNamedItem("ip").Text)=0 then temp = temp & trim(nodes.Item(i).Attributes.getNamedItem("ip").Text) & ","
  Next
  outarray = split(left(temp, len(temp)-1),",")
  getAllBuildingIP = outarray
end function

Sub setPortfolio(pid, ip)
  Dim node, dataservers, xml
  Set xml = getXmlSession()
  Set dataservers = xml.selectSingleNode("//dataservers")
  Set node = xml.createNode("element", "portfolio", "")
  node.setAttribute "pid", pid
  node.setAttribute "ip", ip
  dataservers.appendChild node
End Sub

Function getPidIP(pid)
  Dim node, xml
  set xml = getXmlSession()
  set node = xml.selectSingleNode("//portfolio[@pid='" & pid & "']")
  If TypeName(node) = "Nothing" Then 
    getPidIP = 0 
    response.write "Error getting portfolio info for '"&pid&"'. Your Session is not configured for this portfolio."
    response.end
  Else 
    getPidIP = node.Attributes.getNamedItem("ip").Text
  end if
End Function

Sub printXML()
  Dim xml, node
  Set xml = getXmlSession()
  Set node = xml.selectSingleNode("/")
  response.write node.xml
End Sub

Function getXmlSession()
  Checkxml()
  set getXmlSession = session("xmlUserObj")
End Function

Sub Checkxml()
  if isempty(session("xmlUserObj")) then
    response.write "There is no user session established, please login again."
    response.end
  end if
end Sub

Sub setKeyValue(key, Value)
  Dim node, keys, xml
  Set xml = getXmlSession()
  Set keys = xml.selectSingleNode("//keys")
  Set node = xml.selectSingleNode("//keys/key[@name='"&key&"']")
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

Sub setView(viewname, ip, portfolio)
  Dim node, views, xml
  Set xml = getXmlSession()
  Set views = xml.selectSingleNode("//views")
  Set node = xml.createNode("element", "view", "")
  node.setAttribute "viewname", viewname
  node.setAttribute "ip", ip
  node.setAttribute "portfolio", portfolio
  views.appendChild node
End Sub

Function getViewIP(viewname)
  Dim node, xml
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//view[translate(@viewname,'"&ucase(viewname)&"','"&lcase(viewname)&"')='"&lcase(viewname)&"']")
  If TypeName(node) = "Nothing" Then 
    getViewIP = 0 
  Else 
    getViewIP = node.Attributes.getNamedItem("ip").Text
  end if
End Function

Function getViewPortfolio(viewname)
  Dim node, xml
  set xml = getXmlSession()
  xml.SetProperty "SelectionLanguage","XPath"
  set node = xml.selectSingleNode("//view[translate(@viewname,'"&ucase(viewname)&"','"&lcase(viewname)&"')='"&lcase(viewname)&"']")
  If TypeName(node) = "Nothing" Then 
    getViewPortfolio = -1
  Else 
    getViewPortfolio = node.Attributes.getNamedItem("portfolio").Text
  end if
End Function
%>
