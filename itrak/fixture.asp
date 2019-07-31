<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>

<%
		if isempty(Session("name")) then
'			Response.Redirect "../index.asp"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
		
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getconnect(0,0,"engineering")
		
id=request.querystring("id")

sqlstr = "select * from fixtures where id='"& id &"'"
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then

%>
<title>Fixtures</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="updatefix.asp">
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="12%" height="2"> <font size="2"> 
        <input type="hidden" name="id" value="<%=rst1("id")%>">
        <input type="hidden" name="bldg" value="<%=rst1("bldgnum")%>">
        <font face="Arial, Helvetica, sans-serif">Fixture Type </font></font></td>
      <td width="15%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Floor</font></td>
      <td width="16%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Room</font></td>
      <td width="16%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Fixture 
        Quantity</font></td>
      <td width="16%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Lamp 
        Quantity</font></td>
      <td width="20%" height="10"><font size="2" face="Arial, Helvetica, sans-serif">Ballast 
        Quantity</font></td>
      <td width="21%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Comments</font></td>
    </tr>
    <tr> 
      <td  valign="top" width=12%><font face="Arial, Helvetica, sans-serif" size="2"> 
        <select name="type">
          <%set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select  fix_catalog+' '+lamp_catalog as type ,id from fixture_types "
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
				do until rst2.eof	
		%>
          <option value="<%=rst2("id")%>"><%=rst2("type")%></option>
          <%
					rst2.movenext
					loop
					end if
					%>
        </select>
        </font> </td>
      <td  valign="top" width="15%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="floor" value="<%=rst1("floor")%>">
        </font></td>
      <td valign="top" width="16%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="room" value="<%=rst1("room")%>">
        </font></td>
      <td valign="top"  width="16%" height="19"><font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="fixqty" value="<%=rst1("fixtureqty")%>">
        </font></td>
      <td valign="top"  width="16%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("lampqty")%> 
        </font></td>
      <td valign="top"  width="16%" height="19"><font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="balqty" value="<%=rst1("ballast_qty")%>">
        </font></td>
      <td width="21%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <textarea name="comments"><%=rst1("comments")%></textarea>
        </font></td>
    </tr>
    <tr> 
      <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="submit" name="choice22"  value="Update">
        </font></td>
      <td width="15%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
      <td width="16%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
    </tr>
  </table>
</form>
</body>
</html>
<%end if%>
