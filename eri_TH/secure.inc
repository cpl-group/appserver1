<%
function checkgroup(groupname)
Dim UserGroups , username, domainUnit, domainServer, password, domainName, userGrps, x, y
Set UserGroups = CreateObject("Scripting.Dictionary")
userGrps = Session("uGroups")
if vartype(userGrps)=0 then
	%>
	<script>
	window.top.document.location='/genergy2/login.asp'
	</script>
	<%
	response.end
end if
x = cint(ubound(userGrps))
	for y=0 to x
		UserGroups.Add userGrps(y),"Group"
	next
	if UserGroups.Exists(groupname) then
		checkgroup = "1"
	else
	
		checkgroup = "0"
	end if
end function 
function checkClientLogin(uname, pword, org)
Dim UserGroups , username, domainUnit, domainServer, password, Domainname
	username = trim(uname)
	domainName ="ghnet"
	domainUnit = org
	domainServer = "2012dc"
	password = trim(pword)

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
	
	'Check old security system
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	cnn1.Open application("cnnstr_security")

	' Create and open an ADO recordset object.
	  sql = "SELECT * " & _
			 "FROM clients " & _
			 "WHERE (username = '" & username & "') ; "
	  Set rsVis = Server.CreateObject("ADODB.Recordset")
	  
	  rsVis.Open sql, cnn1, adOpenDynamic, adLockOptimistic
	
	 If rsVis.EOF Then
			
	'     user ID address not in database, returning visitor
		  fMsg = "USER ID not found."
	  Else  
		   If password = Trim(rsVis("paswd")) Then
			session("initialpath")= Trim(rsVis("initial_page"))'&"?fc="&Oct(rnd*1000000)
			session("userid") = username
			session("RoleID") = rsVis("roleid")
			Dim rsRole
			set rsRole = server.createobject("ADODB.Recordset")
			rsRole.open "SELECT Label FROM tblrole WHERE roleid=" & session("RoleID"), cnn1
			session("RoleName") = rsRole("Label")
			rsRole.close
			session("loggedin") = True
			session("loginemail") = username
			checkclientlogin = "1"
			else
			checkclientlogin = "0"
		   end if
		   
	End If
	  
	  rsVis.Close
  Else
    'handle successful login
	'Check old security system
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	cnn1.Open application("cnnstr_security")

	' Create and open an ADO recordset object.
	  sql = "SELECT * " & _
			 "FROM clients " & _
			 "WHERE (username = '" & username & "') ; "
	  Set rsVis = Server.CreateObject("ADODB.Recordset")
	  
	  rsVis.Open sql, cnn1, adOpenDynamic, adLockOptimistic
	  
	  if not rsVis.EOF then
			session("initialpath")= Trim(rsVis("initial_page"))'&"?fc="&Oct(rnd*1000000)
			session("userid") = username
			session("RoleID") = rsVis("roleid")
			set rsRole = server.createobject("ADODB.Recordset")
			rsRole.open "SELECT Label FROM tblrole WHERE roleid=" & session("RoleID"), cnn1
			session("RoleName") = rsRole("Label")
			session("pid") = rsVis("portfolioid")
			rsRole.close
			session("loggedin") = True
			session("loginemail") = username
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
	next
	grparray = split(usrgrps,",")
	Session("uGroups") = grparray
	Session("user") = username
	uGroup = nothing
	adsUsers = nothing
	checklogin = "1"
 End if
end function
%>
