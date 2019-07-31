<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
bldgnum=Request("bldgnum")
address=Request("address")
sqft=Request("sqft")
rev=Request("rev")	
%>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="capaddbldg.asp">
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Building 
        No. </font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Address</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">SQFT</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Revision</font></td>
    </tr>
    <tr> 
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="bldgnum" value="<%=bldgnum%>">
		<%=bldgnum%>
        </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="address" value="<%=address%>">
		<%=address%>
        </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="sqft" value="<%=sqft%>">
		<%=sqft%>
        </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="rev" value="<%=rev%>">
		<%=rev%>
        </font></td>
    </tr>
  </table>
  <br>
  <input type="button" name="submit" value="Update">
  <input type="button" name="submit" value="Back">
  <input type="button" name="submit" value="Add Floor">
</form>
<IFRAME name="floor" width="100%" height="150" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<IFRAME name="riser" width="100%" height="150" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
</body>
</html>
