<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>Client Setup/Update</title>
<script language="JavaScript" type="text/javascript">
function fillup(name){
  //document.location="usrdetail.asp?username="+name
  //document.site.src="usrsite.asp?username="+name
  alert(document.site.src);
}
function submitform(choice){
	document.form1.choice.value=choice
	document.form1.submit()
}
function viewconsole(mode){
  document.site.location="./g1nav/g1nav.asp?mode="  + mode;
}
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
function cloneid(userid){
	var url
	url = "cloneid.asp?uid=" + userid
	openwin(url, 400,150)

}
function launchoptions(uid){

	var selection = document.form1.glboptList.value
	var cid = selection.split('_')[0];
	var label = selection.split('_')[1];
	var url = '/um/security/optionsList.asp?username='+uid+'&csid='+cid+'&label=' + label;
	openwin(url, 250,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
<%
userid = Request("userid")
Session("editemail")=userid

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
strsql = "SELECT * FROM clients where username ='" & userid & "'"
rst1.Open strsql, cnn1, 0, 1, 1
if not rst1.EOF then 
	uName 		= trim(rst1("name"))
	uLogin 		= trim(rst1("username"))
	uPassword	= trim(rst1("paswd"))
	uEmail		= trim(rst1("email"))
	uCompany 	= trim(rst1("company"))
	uCustompage = trim(rst1("custompage"))
	uStartpage 	= trim(rst1("initial_page"))	 
	uKey 		= trim(rst1("clientkey"))
	pid 		= trim(rst1("portfolio_id"))
end if
rst1.close
%>

<body bgcolor="#eeeeee" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<form name="form1" method="post" action="usrmodify.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr> 
      <td colspan="2" align="right"  bgcolor="#666699"><div align="left"><span class="standardheader"><a name="Top"></a>Account 
          Details </span></div></td>
      <td  bgcolor="#666699"><div align="left"></div></td>
    </tr>
    <tr> 
      <td colspan="2" align="right"><div align="left">Name</div></td>
      <td><div align="left">Login Username </div></td>
    </tr>
    <tr> 
      <td colspan="2" align="right"><div align="left"> 
          <input type="text" name="name" size="50%" value="<%=uName%>">
        </div></td>
      <td>	  <%
		  dim userlist, usernamelist
		
		  'strDomain = "ghnet"
		  'strGroup = "clientOperations"
		  'set oGroup = GetObject("WinNT://" & strDomain & "/" & strGroup)
		  'For each olUser in oGroup.Members
		  'usernamelist 	= usernamelist & ","& olUser.Name
		  'userlist		= userlist & "," & olUser.FullName
		  'next
		  'usernamelist 	= split(usernamelist,",")

		  'set oGroup = nothing
	  %>
		<div align="left"> 
		<%
		Dim cnn,UsersRS
		set cnn 	= server.createobject("ADODB.connection")
		cnn.open getConnect(0,0,"dbCore")	  
		set UsersRS = server.createobject("ADODB.recordset")
		UsersRS.open "select CASE WHEN Company IS  NULL THEN 'Other' ELSE Company END As 'Company',fullname,username from dbo.ADusers_genergyone order by company,fullname", cnn
		UsersRS.MoveFirst
		GenerateUserList "userlogin",UsersRS,"","" ,trim(uLogin)
		UsersRS.Close
		cnn.Close
		set UsersRS = nothing
		set cnn = nothing
		 %>
       </div></td>
    </tr>
    <tr> 
      <td colspan="2" align="right"><div align="left">Email Address </div></td>
      <td><div align="left">Login Password </div></td>
    </tr>
    <tr> 
      <td colspan="2" align="right"><div align="left"> 
          <input type="text" name="email" size="50%" value="<%=uEmail%>">
        </div></td>
      <td><div align="left"> 
          <input type="text" name="passwd" size="50%" value="<%=uPassword%>">
        </div></td>
    </tr>
    <tr> 
      <td colspan="2" align="right"><div align="left">Company Name (PID)</div></td>
      <td> <div align="left"> 
          <input name="defaultpage" type="radio" value="0" <%if uCustompage = "False" then%>checked<% end if %> onClick="document.form1.page.value='/g1_clients/index2.asp'">
          Default Console 
          <input type="radio" name="defaultpage" value="1" <%if uCustompage = "True" then%>checked<% end if %> onClick="document.form1.page.value=''">
          Custom Console </div></td>
    </tr>
    <tr> 
      <td colspan="2" align="right"><div align="left"> 
	  <%
		strsql = "SELECT portfolio,ID, name FROM dbo.portfolio order by name"
		rst1.Open strsql, cnn1, 0, 1, 1
		if not rst1.EOF then 
	  %>
		<select name="company_portfolio">
        	<%while not rst1.EOF %>
		    <option value="<%=rst1("name") & "_" & rst1("id")%>"  <% if lcase(trim(rst1("id"))) = lcase(pid) then %> selected <%end if%>><%=rst1("name") & " (" & rst1("portfolio") & ")"%></option>
        	<%
			rst1.movenext
			wend
			%>
		  </select>
          <%
		else
	  		response.write "No Companies available in the system. Please contact Data Services"
		end if
		rst1.close
	  %>
        </div></td>
      <td> <div align="left"> 
          <input name="page" type="text" value="<%=uStartpage%>" size="50%" >
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <input name="choice" type="hidden"> <input name="key" type="hidden" value="<%=uKey%>"> <input type="button" name="close"  style="border:1px outset #ddffdd;background-color:ccf3cc;"  <%if ukey ="" then %> value="Save" onclick="submitform('Save')" <% else %> value="Update" onclick="submitform('Update')" <%end if%>> 
        <input type="button" name="delete" value="Delete" style="border:1px outset #ddffdd;background-color:ccf3cc;" onclick="submitform('Delete')"> 
        <input type="button" name="clone" value="Clone User To This Account" style="border:1px outset #ddffdd;background-color:ccf3cc;" onclick="cloneid('<%=uLogin%>')"> 
      </td>
      <td><div align="right"> </div></td>
    </tr>
    <tr> 
      <td colspan="3"><hr></td>
    </tr>
    <%if ukey <> "" then %>
    <tr bgcolor="#666699"> 
      <td><span class="standardheader">CONSOLE VIEWS</SPAN></td>
      <td><span class="standardheader">TREE TOOLS</SPAN></td>
      <td><span class="standardheader">GLOBAL OPTIONS</SPAN></td>
    </tr>
    <tr> 
      <td><input type="button" name="loadnav" value="Refresh Edit Console View" style="border:1px outset #ddffdd;background-color:ccf3cc;width:200px;height:20px;height:20px" onClick="viewconsole('edit');"></td>
      <td><input type="button" name="Add Entry" value="Add Entry" style="border:1px outset #ddffdd;background-color:ccf3cc;width:200px;height:20px" onClick="document.site.location='addentries.asp?mode=initentry&userid=<%=userid%>'"></td>
      <td><select name="glboptList" style="border:1px outset #ddffdd;background-color:ccf3cc;width:200px;height:20px">
          <%
  rst1.open "SELECT * FROM tblCoreServices WHERE CSID in (SELECT csid FROM tbladdons)", cnn1
  do until rst1.eof
    %>
          <option  value="<%=rst1("csid")%>_<%=rst1("Label")%>"><%=rst1("Label")%></option>
          <%
    rst1.movenext
  loop
  rst1.close
  set cnn1 = nothing
%>
        </select></td>
    </tr>
    <tr> 
      <td><input type="button" name="loadnav2" value="Refresh Live Console View" style="border:1px outset #ddffdd;background-color:ccf3cc;width:200px;height:20px" onClick="viewconsole('live');"></td>
      <td><input type="button" name="autoorder" value="Auto Order Navigation Tree" style="border:1px outset #ddffdd;background-color:ccf3cc;width:200px;height:20px" onClick="if (confirm('This will re-order your entire user tree in building number order. Are you sure you want to proceed? THIS CANNOT BE UNDONE')) {openwin('autoorder.asp', 400,150)}"></td>
      <td> <input type="button" name="action" value="View Options" onClick="launchoptions('<%=uLogin%>')" style="border:1px outset #ddffdd;background-color:ccf3cc;width:200px;height:20px"> 
      </td>
    </tr>
    <%end if%>
  </table>
      </table>
  </form>
<IFRAME name="site" width="100%" height="300" src="./g1nav/<% if uKey = "" then %>null.htm <%else%>g1nav.asp?mode=edit&userid=<%=userid%><%end if%>" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
</body>

</html>
