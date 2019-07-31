<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
server.ScriptTimeout=300
tripcode 	= request("tripcode")
billperiod 	= request("billperiod")
extended 	= request("extended")

if extended = "" then extended = false

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rs	 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")

strsql = "select pid, bldgnum, tripcode from super_tripcodes where tripcode = "& tripcode &" GROUP BY pid, bldgnum, tripcode ORDER BY bldgnum"
rst1.Open strsql, cnn1, 0, 1, 1
%>
<html>
<head>
<title>Trip Sheet</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #cccccc; }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><head>
<script>
function updateentry(key){
	parent.document.frames.tripedit.location = "tripdetail.asp?key=" + key
	parent.document.all.te.style.visibility="visible"

}
function deletetrip(key, bldgname){

	if (confirm("Delete trip for" + bldgname+"?")){
	parent.document.frames.tripedit.location = "tripmodify.asp?key=" + key +"&modify=Delete"
	parent.document.all.te.style.visibility="hidden"
	}
}
</script>
<%if not rst1.eof then%>
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" class="innerbody" onload="window.print()">
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <tr bgcolor="#dddddd" valign="bottom"> 
    <td colspan="6" valign="top" align="left" bgcolor="#eeefff" class="tblunderline" style="solid #e3e3d3;"><font size="4">Trip 
      Sheet for Trip <%=tripcode%>, Bill Period <%=billperiod%></font></td>
    <td valign="bottom" align="right" bgcolor="#eeefff" class="tblunderline" colspan=<%if extended then%>8<%else%>4<%end if%> style="solid #e3e3d3;"><font size="2">Meter 
      Reader: ___________________________</font></td>
  </tr>
  <% 
	dim rst2, floor, tenantName, tenantNo,uid
	set rst2 = server.createobject("adodb.recordset")
	Do until rst1.EOF 
		currentbldg = rst1("bldgnum")
		if not(isBuildingOff(currentbldg)) then
			strsql = "exec sp_Trip_MeterInfo '"&currentbldg&"', "&Year(now())&", "&billperiod&", "&tripcode
		
			rst2.open strsql, getConnect(0,currentbldg,"billing")
			if not rst2.eof then
			%>
		  <tr bgcolor="#ffffff" valign="middle"> 
		    <td align="left" bgcolor="#f0f0e0" class="tblunderline" style="border-left:1px solid #e3e3d3;" colspan=<%if extended then%>14<%else%>10<%end if%>><strong><%=rst2("bldgname")%> 
		      (<%=rst2("bldgnum")%>) - [<%=billperiod%>/<%=rst2("billyear")%>]</strong></td>
		  </tr>
			
		  <tr align="center" valign="top" bgcolor="#ffffff"> 
		    <td align="left" nowrap bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Utility</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter 
		      Number</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Floor</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Tenant</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Location</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous 
		      <%if extended then%>
		      On Peak 
		      <%end if%>
		      Reading</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">Current 
		      <%if extended then%>
		      On Peak 
		      <%end if%>
		      Reading</td>
				<% if extended then %>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous 
		      Off Peak Reading</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">Current 
		      Off Peak Reading</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous 
		      Mid Peak Reading</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">Current 
		      Mid Peak Reading</td>
				<% end if %>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous 
		      Demand</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">Current 
		      Demand</td>
				
		    <td width="100" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Notes</td>
			</tr>
		<% do until rst2.eof %>
		    <tr bgcolor="#ffffff" valign="top" > 
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><%=rst2("utility")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("meternum")%>&nbsp;</td>
				<td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("floor")%>&nbsp;</td>
				<%
				if len(rst2("tenenat")) > 22 then 
					tenantName = "" & left(rst2("tenenat"),20) & "..."
				else
					tenantName =  rst2("tenenat")  
				end if
				%>
				<td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=tenantName%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("location")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("rawprevious")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">&nbsp;</td>
		      <% if extended then %>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("rawpreviousoff")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("rawpreviousint")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">&nbsp;</td>
			  <% end if %> 
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst2("rawprev")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border-left:1px solid #e3e3d3;">&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">&nbsp;</td>
		    </tr>
		  <%  
		    rst2.movenext
			loop
			else 
			%>
			<tr><td colspan=<%if extended then%>14<%else%>10<%end if%>>NO METERS FOUND FOR BUILDING <%=currentbldg%></td></tr>
			<%
			end if
			rst2.close
		
		end if
		rst1.movenext
    loop
%>
</table>
</body>
<%else %>
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" class="innerbody">
	<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td class="standardheader" bgcolor="#999999" align="center">NO TRIP SHEET FOUND FOR TRIPCODE <%=tripcode%></td>
  </tr>
</table>

</body>
<% end if %>
</html>
