<!--#INCLUDE FILE="XmlSessions.asp"-->
<%
'securetrace
Dim UserGroups, username, domainUnit, domainServer, password, domainName, userGrps, x, cnnLog, logstr, rstMakeIP
function allowGroups(groups)
  dim group
  allowGroups = false
  groups = split(groups,",")
  for each group in groups
    if hasGroup(group) then allowGroups = true
  next
end function

function denyGroups(groups)
  dim group
  denyGroups = true
  groups = split(groups,",")
  for each group in groups
    if hasGroup(group) then denyGroups = false
  next
end function

function checklogin(uname, pword, org)
  Dim UserGroups , username, domainUnit, domainServer, password, Domainname
	username = trim(uname)
	password = trim(pword)		
	domainName ="ghnet"
	domainUnit = org
	domainServer = "ghnetpdc1"
  err.Clear  
  Dim adsLDAP,adsUser
 on error resume next 
  Set adsLDAP = GetObject("LDAP:") 'LDAP object to authenticate with
	strADsPath = "LDAP://" & domainServer &"/DC="& domainName &" ,DC=" & domainUnit & ",DC=com"
	strLogin = ",CN=" & domainUnit & ", DC=" & domainName & ", DC=com"
	strPassword = password
  Set adsUser = adsLDAP.OpenDSObject(strADsPath, username & strLogin, strPassword,1)
  adsUser.GetInfo
		
  If ( Err ) Then
    'handle failed login
  	checklogin = "Login failed, please try again."
	  response.cookies ("genergyone i:2").expires = date() - 365		
  Else
    'handle successful login
  	loadNewXML(username)
  	If ( Err ) Then
		checklogin = "Error in LoadNewXML"
	else 
		checklogin = loadADIinfo(username)
			
		'setBuildingOffline
		if session("autologin") = 1 then 
			response.cookies ("genergyone i:2")("password") = password
			response.cookies ("genergyone i:2")("username") = username
			response.cookies ("genergyone i:2").expires = date() + 5
		end if 
		tempsessions(username)
	end if
  End if
  loadIPs(0)
  if checklogin = "" then checklogin = "UNKNOWN SYSTEM ERROR<br>PLEASE TRY AGAIN LATER" 
end function

function loadADIinfo(username)
	Dim usrGrps, uGroups
	Dim cnn, rs,sql
	set cnn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.Recordset")
	cnn.open getConnect(0,0,"dbCore")
	sql = "select * from adusers_genergyusers where username = '"&username&"'"
	rs.open sql, cnn
		
	if not rs.eof then 
	
		setKeyValue "fullname", rs("FullName")
		setKeyValue "user", username
		groups = split(rs("pgroup"),"|")
		for Each uGroups in groups
			if usrgrps = "" then
			  usrgrps = trim(uGroups)
			else
			  usrgrps = usrgrps & "," & trim(uGroups)
			end if
		  setGroup trim(uGroups)
		next
		if instr(rs("pgroup"),"MultiAccess") then 
			loadADIinfo = 1
		else
			rs.close 
			sql = "select * from logintracking where username = '"&username&"' and ip <> '"&Request.ServerVariables("Remote_Addr")&"'"
			rs.open sql, cnn
			if rs.eof then 
				loadADIinfo = 1
				sql = "delete from logintracking where username = '"&username&"';insert into logintracking (username, lastpageview,ip) values ('"&username&"','initial login','"&Request.ServerVariables("Remote_Addr")&"')"
				cnn.execute sql
			else
				loadADIinfo = "User is already logged into the system. Please try again in 15 minutes."
			end if
		end if 
	else
				loadADIinfo = "SYSTEM ERROR: PROFILE SETTINGS NOT FOUND"
	end if
	  	rs.close
	set uGroup = nothing
  	set adsUsers = nothing
    
	set cnn = nothing
	set rs  = nothing	
end function

function tempsessions(username)
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  Set rst1 = Server.CreateObject("ADODB.recordset")
  cnn1.Open getConnect(0,0,"dbCore")
  
  strsql = "SELECT * from employees where login='" & username & "'" 
  rst1.Open strsql, cnn1, 0, 1, 1
  	if not rst1.EOF then 
	  Session("login") = rst1("login")
	  Session("name") = rst1("name")
	  Session("roleid")=4
	  Session("um") = rst1("um")
	  Session("eri") = rst1("eri")
	  Session("opslog") = rst1("opslog")
	  Session("ts") = rst1("ts")
	  Session("corp") = rst1("corp")
	  Session("it") = rst1("it")
	  Session("admin") = rst1("admin")
 
	  setKeyValue "login", trim(rst1("login"))
	  setKeyValue "name", trim(rst1("name"))
	  setKeyValue "roleid", 4
	  setKeyValue "um", rst1("um")
	  setKeyValue "eri", rst1("eri")
	  setKeyValue "opslog", rst1("opslog")
	  setKeyValue "ts", rst1("ts")
	  setKeyValue "corp", rst1("corp")
	  setKeyValue "it", rst1("it")
	  setKeyValue "admin", rst1("admin")
	  setKeyValue "fMessage", "You are currently logged on."
	end if 
  rst1.close
  Session("fMessage") = "You are currently logged on."
  
  strsql = "UPDATE employees SET status=1 where login = '" & Request.Form("login") & "'"
  cnn1.execute strsql
  set cnn1 = nothing
  set cnn2 = nothing
end function

function logger(string)
  'Logging Update
  set cnnLog = server.createobject("ADODB.connection")
  cnnLog.open getConnect(0,0,"dbCore")
  dim username
  username = getXMLUserName()
  logstr = "Insert into logtable (uid, log) values ('" & username & "','" & replace(string,"'","''")& "')"
  cnnlog.Execute logstr
  set cnnlog = nothing
  'end Log
end function

function loadIPs(pid)
  dim rst1, cnn1, cnn2,rst2,sql, where
  if pid = 0 then pid = "" else where = "WHERE pid="&pid 
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  Set rst1 = Server.CreateObject("ADODB.recordset")
  
  Set cnn2 = Server.CreateObject("ADODB.Connection")
  Set rst2 = Server.CreateObject("ADODB.recordset")

  cnn1.Open getConnect(0,0,"dbCore")
  
  rst1.open "select * from super_main m inner join portfolio p on p.id = m.pid left join buildings b on b.bldgnum = m.bldgnum "&where&" ORDER BY m.bldgnum, m.pid", cnn1
  do until rst1.eof
    if isnull(rst1("bldgnum")) then setPortfolio rst1("pid"), rst1("ip"),rst1("name") else setBuilding rst1("bldgnum"), rst1("ip"),rst1("pid"),rst1("strt")
	
		cnn2.open getConnect(rst1("pid"),rst1("bldgnum"),"Billing")
	
		rst2.open "SELECT * FROM [group] WHERE (type=1 or type=5) and  bldgnum in (select bldgnum from buildings where portfolioid=" & rst1("pid") & ")", cnn2
			do until rst2.eof
				setView rst2("id"), rst2("ip"), PID
				rst2.movenext
			loop
		rst2.close
	
		'find offline buildings
		sql = "SELECT bldgnum, strt  FROM buildings WHERE offline=1 and portfolioid=" & rst1("pid")
		rst2.open sql, cnn2
		do until rst2.eof
			setBuildingOffline rst2("bldgnum"), 1
			rst2.movenext
		loop
		rst2.close
		cnn2.close	
    rst1.movenext
  loop
  rst1.close
  
  if pid <> "" then where = "and clientid="&pid 
  
  'find external sqlserver port mappings
  rst1.open "SELECT * FROM EXTportMappings", cnn1
  do until rst1.eof
    setKeyValue rst1("ip"), rst1("port")
    rst1.movenext
  loop
  rst1.close
  set rst1 = nothing
  set rst2 = nothing
  set cnn1 = nothing
  set cnn2 = nothing
end function

function loadServices(userid)
	dim rst1, cnn1, servicelist
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	Set rst1 = Server.CreateObject("ADODB.recordset")
	cnn1.Open getConnect(0,0,"dbCore")
	
	sql = "select * from g2_userservices where userid =" & userid
	rst1.open sql, cnn1
	
	do until rst1.eof
		setKeyValue "Service|"&trim(rst1("serviceid")), trim(rst1("servicelevel"))& "|"& trim(rst1("serviceindex"))
		rst1.movenext
	loop
  
	rst1.close
	set rst1 = nothing
 	set cnn1 = nothing
end function

function getLocalConnect(bldg)
	getLocalConnect=getConnect(0,bldg,"billing")
end function

function getConnect(pid, bldg, class_req)
	Dim PIDSTR, CLASSSTRING,PIDVAL
	
	if (trim(pid) = "" or pid="0") and (bldg<>"0" and bldg<>"") then 
		PIDVAL=getPID(bldg)	
	else
		PIDVAL=pid
	end if 
	
	Select case lcase(trim(class_req))
		case "billing" 
			CLASSSTRING="dbBilling;"
			PIDSTR = "PID" & PIDVAL
		case "intervaldata" 
		    CLASSSTRING="dbIntervalData;"
			PIDSTR = "PID" & PIDVAL
		case "archive" 
			CLASSSTRING="dbArchiveIntervalData;"
			PIDSTR = "PID" & PIDVAL
		case "dbcore" 
			CLASSSTRING="dbCore;"
			PIDSTR = "dbDefault"
		case "intranet" 
			CLASSSTRING="main;"
			PIDSTR = "dbIntranet"
		case "engineering" 
			CLASSSTRING="dbEngineering;"
			PIDSTR = "dbDefault"
		case else 
			getConnect = "ERROR - PROPER CONNECTION CLASS NOT DEFINED"
			exit function
	end select 
		
	if Application(PIDSTR)="" then 
		PIDSTR="dbDefault"
	end if

	getConnect = Application(PIDSTR) & CLASSSTRING	
end function

function getLocalConnectCom(bldg)
	getLocalConnectCom=getConnect(0,bldg,"intervaldata")
end function

function getMainConnect(pid)
    getMainConnect = getConnect(pid,0,"billing")
end function

function checkClientLogin(uname, pword, org)
	Dim UserGroups , username, domainUnit, domainServer, password, Domainname, sql, pid
	username = trim(uname)
	domainName ="ghnet"
	domainUnit = org
	domainServer = "ghnetpdc1"
	password = trim(pword)
	dim cnn1, rsVis, rsRole
	set rsRole = server.createobject("ADODB.Recordset")
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	Set rsVis = Server.CreateObject("ADODB.Recordset")
	cnn1.Open getConnect(0,0,"dbCore")
  err.Clear  
  Dim adsLDAP,adsUser
  
on error resume next 
  Set adsLDAP = GetObject("LDAP:") 'LDAP object to authenticate with
	strADsPath = "LDAP://" & domainServer &"/DC="& domainName &" ,DC=" & domainUnit & ",DC=com"
	strLogin = ",CN=" & domainUnit & ", DC=" & domainName & ", DC=com"
	strPassword = password
  Set adsUser = adsLDAP.OpenDSObject(strADsPath, username & strLogin, strPassword,1)
  adsUser.GetInfo
  If ( Err ) Then 'handle failed login
 	  sql = "SELECT * FROM clients WHERE (username = '" & username & "') ; "
	  rsVis.Open sql, cnn1, adOpenDynamic, adLockOptimistic
	  If rsVis.EOF Then 'user ID address not in database, returning visitor
  	  fMsg = "USER ID not found."
	  Else  
	    If password = Trim(rsVis("paswd")) Then
        loadNewXML(username)
  			setKeyValue "initialpath", Trim(rsVis("initial_page"))
	  		setKeyValue "roleid", rsVis("roleid")
  			rsRole.open "SELECT Label FROM tblrole WHERE roleid=" & getKeyValue("roleid"), cnn1
  			setKeyValue "rolename", rsRole("Label")
  			rsRole.close
  			setKeyValue "loggedin", True
  			setKeyValue "loginemail", username
        
        'old sessions should be fazed out
  			session("initialpath") = getKeyValue("initialpath")
  			session("userid") = getXMLUserName()
  			session("RoleID") = getKeyValue("roleid")
  			session("RoleName") = getKeyValue("rolename")
  			session("loggedin") = True
  			session("loginemail") = getKeyValue("loginemail")
        pid = rsVis("portfolio_id")
  			checkClientLogin = "1"
			else
  			checkClientLogin = "0"
	    end if
	End If
	  rsVis.Close
  Else 'handle successful login
    sql = "SELECT * FROM clients WHERE (username = '" & username & "') ; "
    rsVis.Open sql, cnn1, adOpenDynamic, adLockOptimistic
    if not rsVis.EOF then
      loadNewXML(username)
 			setKeyValue "initialpath", Trim(rsVis("initial_page"))
  			setKeyValue "roleid", rsVis("roleid")
 			rsRole.open "SELECT Label FROM tblrole WHERE roleid=" & getKeyValue("roleid"), cnn1
 			setKeyValue "RoleName", rsRole("Label")
 			rsRole.close
 			setKeyValue "loggedin", True
 			setKeyValue "loginemail", username
			session("initialpath") = getKeyValue("initialpath")
			session("userid") = getXMLUserName()
			session("RoleID") = getKeyValue("roleid")
			session("RoleName") = getKeyValue("rolename")
			session("loggedin") = True
			session("loginemail") = getKeyValue("loginemail")
      		pid = rsVis("portfolio_id")
 			checkClientLogin = "1"
    else
   		response.write "USER IS NOT CONFIGURED CORRECTLY FOR THE GENERGYONE SYSTEM. PLEASE CONTACT SUPPORT@GENERGY.COM"
  		response.end
    end if
    rsVis.close
  	Dim usrGrps
  	Set User = GetObject("WinNT://"& domainName & "/" & username)
  	for Each uGroups in User.groups
  		if usrgrps = "" then
  		  usrgrps = uGroups.Name
  		else
  		  usrgrps = usrgrps & "," & uGroups.Name
  		end if
      setGroup uGroups.Name
  	next
  	checkClientLogin = "1"
  End if
  if checkClientLogin = "1" and isnumeric(pid) then loadIPs(pid)
end function

function checkClientLoginG2(appid,uname, pword, org, aLogin)
	Dim UserGroups , username, domainUnit, domainServer, password, Domainname, sql, pid
	username = trim(uname)
	domainName ="ghnet"
	domainUnit = org
	domainServer = "ghnetpdc1"
	password = trim(pword)
	Dim adsLDAP,adsUser
	on error resume next 
	Err.clear
	Set adsLDAP = GetObject("LDAP:") 'LDAP object to authenticate with
	strADsPath = "LDAP://" & domainServer &"/DC="& domainName &" ,DC=" & domainUnit & ",DC=com"
	strLogin = ",CN=" & domainUnit & ", DC=" & domainName & ", DC=com"
	strPassword = password
	Set adsUser = adsLDAP.OpenDSObject(strADsPath, username & strLogin, strPassword,1)
	adsUser.GetInfo
  If ( Err ) Then 'handle failed login
			checkClientLoginG2 = "0"
  Else 'handle successful login
		if aLogin = 1 then 
			response.cookies ("genergyone App:"&Session("appid"))("password") = password
			response.cookies ("genergyone App:"&Session("appid"))("username") = username
			response.cookies ("genergyone App:"&Session("appid")).expires = date() + 5
			response.cookies ("genergyone App:"&Session("appid"))("org") = org
		end if 
		loadNewXML(username)
		select case appid
			case 1,2,3,5
				checkClientLoginG2 = loadADG2info(username)
			case 4
				checkClientLoginG2 = loadADIinfo(username)
				
		end select
		Dim usrGrps
		Set User = GetObject("WinNT://"& domainName & "/" & username)
		for Each uGroups in User.groups
			if usrgrps = "" then
			  usrgrps = uGroups.Name
			else
			  usrgrps = usrgrps & "," & uGroups.Name
			end if
		  setGroup uGroups.Name
		next
  End if
  if checkClientLoginG2 = "1" then if isnumeric(getKeyValue("pid")) then loadIPs(getKeyValue("pid"))
  
end function

function loadADG2info(username)
	Dim sql,rsVis, cnn1
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	Set rsVis = Server.CreateObject("ADODB.Recordset")
	cnn1.Open getConnect(0,0,"dbCore")

	sql = "SELECT * FROM g2_accounts WHERE (userid = '" & username & "') ; "
    rsVis.Open sql, cnn1, adOpenDynamic, adLockOptimistic
    if not rsVis.EOF then
 			setKeyValue "usertype", Trim(rsVis("usertype"))
 			setKeyValue "namelabel", Trim(rsVis("name"))
			setKeyValue "bldgid", Trim(rsVis("bldgid"))			
			setKeyValue "pid", Trim(rsVis("pid"))
			setKeyValue "userid", Trim(rsVis("id"))
			setKeyValue "username", Trim(rsVis("userid"))
			session("appid") = Trim(rsVis("appid"))
			setKeyValue "appid", session("appid")
 			loadServices(rsVis("id"))			 			
			loadADG2info = "1"
    else
   		response.write "USER CONFIGURATION IS NOT YET COMPLETE. PLEASE CONTACT SUPPORT@GENERGY.COM"
  		response.end
    end if
    rsVis.close
end function

function checkGroup(h)
  checkgroup=1
end function

function secureRequest(requestString)
	dim nastyPunc, nastyPuncFound, dirtyInput, punc, badWord
	dirtyInput = Request(requestString)
	dirtyInput = replace(dirtyInput,"'","''")
	nastyPunc = array("--",";")
	nastyPuncFound = false
	
	For Each punc in nastyPunc
		if inStr(dirtyInput, punc) <> 0 then
			'response.write("found nasty punctuation")
			nastyPuncFound = true
			dirtyInput = replace(dirtyInput, punc, "")
		end if
	next
	
	dim badWords
	badWords = array( "select", "insert", "update", "delete", "drop", "create", "exec", "alter","truncate")
	For Each badWord in badWords
		if ((inStr(lcase(dirtyInput), badWord) <> 0) AND (nastyPuncFound = true)) then
			Response.write("<title>Invalid Input</title>Invalid input.")
			'flagIP( request.ServerVariables(Remote_addr) )
			Response.end()
			secureRequest = "invalid input"
		end if
	next
	
	secureRequest = dirtyInput
end function

function activeservice(serviceid)
dim activeservices

activeservices = getKeyValue("Service|"&serviceid)

	if activeservices <> "" then 
		activeservices 	= split(activeservices, "|")
		entrytype 		= activeservices(0)
		entryvalue		= activeservices(1)
		Select Case entrytype
			case "p"
				pid  		= entryvalue				
			case "b"
				bldg 		= entryvalue
				pid		= getKeyValue("pid")
			case "t" 
				tenantid = entryvalue
				bldg		= getKeyValue("bldgid")
				pid		= getKeyValue("pid")
		end select
	else
		response.redirect "/genergyonev2/demo/notavailable.html"
	end if
end function
function sendmail(toadd, fromadd, subject, message)
	dim email, body
	set email = server.createObject("CDONTS.NewMail")
	email.To= toadd
	email.From= fromadd
	email.Subject = Subject
	email.Body = message
	email.Bodyformat=0
	email.Mailformat=1
	email.Send 
end function
function securetrace() 

  if CheckEmptyXML then
		Dim cnn, rs,sql
		set cnn = server.createobject("ADODB.Connection")
		set rs = server.createobject("ADODB.Recordset")
		cnn.open getConnect(0,0,"dbCore")
		
		sql = "delete from logintracking where username='"&trim(getXMLUsername())&"';insert into logintracking (username,lastpageview,ip) values ('"&trim(getXMLUsername())&"','"&Request.ServerVariables("Script_Name")&"','"&Request.ServerVariables("Remote_Addr")&"')"
		cnn.execute sql		
	end if 
end function 

function GenerateUserList(name,UsersRS,multiple,sizeo,user)
dim tracktype,company
if (name <> "userlogin") then
Response.Write("<select name=""" & name & """ " & multiple & " size = """ & sizeo & """>") 
          
							
							tracktype = ""
							company = ""
							Do while Not(UsersRS.EOF)
							
			             if tracktype = "" then       
					       	 Response.Write("<OPTGROUP Label=""" & trim(UsersRS("Company")) & """>")   
   					       	 Response.Write("<OPTGROUP Label=""&nbsp;&nbsp;" & trim(UsersRS("department")) & """>")   
						 elseif (trim(tracktype) <> trim(UsersRS("department"))) and (trim(company) = trim(UsersRS("Company")) ) then 
   					      	 Response.Write("</OPTGROUP><OPTGROUP Label=""&nbsp;&nbsp;" & trim(UsersRS("department")) & """>")   
                         elseif (trim(company) <> trim(UsersRS("Company"))) then
					      	 Response.Write("</OPTGROUP></OPTGROUP>")
							 Response.Write("<OPTGROUP Label=""" & trim(UsersRS("Company")) & """>")   
							 Response.Write("<OPTGROUP Label=""&nbsp;&nbsp;" & trim(UsersRS("department")) & """>")   
		    			  end if 
						  tracktype = trim(UsersRS("department"))
						  company = trim(UsersRS("Company"))
						if len(UsersRS("username")) <> 0 then 
									
          					Response.Write("<option value=""" & UsersRS("username") & """ ")
							
							if trim(UsersRS("username")) = user then Response.Write("selected") end if
							
				            Response.Write(">")
				            Response.Write(UsersRS("fullname"))
					        Response.Write("</option>") 
          
						end if 
								UsersRS.MoveNext
							Loop
							UsersRS.Movefirst
					
        Response.Write("</select>")
		
else


Response.Write("<select name=""" & name & """ " & multiple & " size = """ & sizeo & """>") 
          
							
							tracktype = ""
							company = ""
							Do while Not(UsersRS.EOF)
							
			             if tracktype = "" then       
					       	 Response.Write("<OPTGROUP Label=""" & trim(UsersRS("Company")) & """>")   
   					      elseif (trim(tracktype) <> trim(UsersRS("Company"))) then 
   					      	 Response.Write("</OPTGROUP><OPTGROUP Label=""" & trim(UsersRS("Company")) & """>")   
                          end if 
						  tracktype = trim(UsersRS("Company"))
						if len(UsersRS("username")) <> 0 then 
									
          					Response.Write("<option value=""" & UsersRS("username") & """ ")
							
							if trim(UsersRS("username")) = user then Response.Write("selected") end if
							
				            Response.Write(">")
				            Response.Write(UsersRS("fullname") & " (" & UsersRS("username") & ")")
					        Response.Write("</option>") 
          
						end if 
								UsersRS.MoveNext
							Loop
							UsersRS.Movefirst
					
        Response.Write("</select>")

end if

end function

%>
