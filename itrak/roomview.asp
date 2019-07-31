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


Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic
id=request.querystring("id")
'response.write id
'response.end
		sqlstr= "select * from room where id='"&id&"'"
	
'response.write sqlstr
'response.end	
rst1.Open sqlstr, cnn1, 0, 1, 1
%>


<title>New Room</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="saveroom.asp">
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="16%" height="2"><font face="Arial, Helvetica, sans-serif">Floor</font></td>
      <td width="21%" height="10"><font face="Arial, Helvetica, sans-serif">Room 
        Name</font></td>
      <td width="18%" height="10"><font face="Aral, Helvetica, sans-serif">SQFT</font></td>
    </tr>
    <tr> 
      <td width=16%> <font face="Arial, Helvetica, sans-serif"> 
	   <input type="hidden" name="id" value=<%rst1("id")%>">
        <input type="hidden" name="floor" value="<%rst1("floor")%>">
		<input type="hidden" name="bldg" value="<%rst1("bldg")%>">
        <%=fl%></font></td>
      <td width="21%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="room" value="<%rst1("room")%>">
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="sqft" value="<%rst1("sqft")%>">
        </font></td>
    </tr>
    <tr> 
      <td width=16%><font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="Submit" value="Update">
        </font> </td>
      <td width="21%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
    </tr>
  </table>
	
  
</form>
</body>
</html>
<%rst1.close
set cnn1=nothing%>