<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
server.ScriptTimeout=300
tripcode 	= request("tripcode")
billperiod 	= request("billperiod")
extended 	= request("extended")
bldgnum = request("bldgnum")
billyear = request("billyear")

if extended = "" then extended = false

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rs	 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")

if bldgnum = "All" then
    strsql = "select pid, bldgnum, tripcode from super_tripcodes where tripcode = "& tripcode &" GROUP BY pid, bldgnum, tripcode ORDER BY bldgnum"
else
    strsql = "select distinct(pid), bldgnum, tripcode from super_tripcodes where tripcode = "& tripcode &" and bldgnum = '" + bldgnum + "'"
end if
rst1.Open strsql, cnn1, 0, 1, 1
%>
<html>
<head>

<title>Trip Sheet</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css" media="print">
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
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" onload="window.print()">
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <% 
	dim rst2, floor, tenantName, tenantNo,uid
	set rst2 = server.createobject("adodb.recordset")
	Do until rst1.EOF 
		currentbldg = rst1("bldgnum")
		if not(isBuildingOff(currentbldg)) then
		    strsql = "sp_Trip_MeterInfo_2months_test '"&currentbldg&"', "&billyear&", "&billperiod&", "&tripcode
		   	rst2.open strsql, getConnect(0,currentbldg,"billing")
			if not rst2.eof then
			%>
  
  <tr bgcolor="#dddddd" valign="bottom"> 
    <td colspan="<%if extended then%>17<%else%>10<%end if%>" valign="top" align="left" bgcolor="#eeefff" class="tblunderline" style="solid #000000;"><font size="4">Trip 
      Sheet for Trip <%=tripcode%>, Bill Period <%=billperiod%>/<%=rst2("billyear")%></font></td>
  </tr>
  <tr>
    <td height="40" valign="bottom" align="left" bgcolor="#eeefff" class="tblunderline" colspan="<%if extended then%>17<%else%>10<%end if%>" style="solid #000000;"><font size="2">
    Meter Reader: _________________ Date:______________ Time In:______________  Time Out:______________</font></td>
  </tr>
  		  <tr bgcolor="#ffffff" valign="middle"> 
		    <td align="left" bgcolor="#f0f0e0" class="tblunderline" style="border:1px solid #000000;" colspan=<%if extended then%>17<%else%>10<%end if%>><strong>
		   <font size=3> <%=rst2("bldgname")%> (<%=rst2("bldgnum")%>)</font></strong></td>
		  </tr>
		  <% if( rst2("bldgnum") = "200PARK") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****ALL METERS IN THIS BUILDING NEED LADDERS****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "1370S") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****ALL METERS IN THIS BUILDING NEED LADDERS****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "1001-6") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****ALL METERS IN THIS BUILDING NEED LADDERS****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "260Mad") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****Please stop by the building office to meet with Benny Bash.  His number is 212-971-0111 x 155 or 646-387-5816.****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "261Mad") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****Please stop by the building office to meet with Benny Bash.  His number is 212-971-0111 x 155 or 646-387-5816.****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "945") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****Please contact Don Capp at 516-250-4671/DCapp@steelequities.com or Althea Hager at 516-576-3165/Ahager@steelequities.com for help in locating any meter.****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "044SW") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****PC&S meters may require code 8453 to reset max demand.****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("bldgnum") = "123W") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****This building requires a ladder to read the meters. Ask building engineers for a ladder.****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <% if( rst2("portfolioid") = "73") then  %>
		        <tr bgcolor="#ffffff" valign="middle">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>****Please take full Model and Serial numbers****</b>  
		            </td>
		        </tr>
		  <%end if %>
		  <tr align="center" valign="top" bgcolor="#ffffff"> 
		    <td align="left" nowrap bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Utility</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Meter 
		      Number</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Floor</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Location</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Tenant</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      <%if extended then%>
		      2 Month On Peak 
		      <%end if%>
		      2 Month Reading</td>
			
            <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      <%if extended then%>
		      On Peak 
		      <%end if%>
		      Reading</td>	

		    <td bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">Current 
		      <%if extended then%>
		      On Peak 
		      <%end if%>
		      Reading</td>
				<% if extended then %>

            <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      2 Month Mid Peak Reading</td>
				
			<td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      Mid Peak Reading</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">Current 
		      Mid Peak Reading</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      2 Month Off Peak Reading</td>

            <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      Off Peak Reading</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">Current 
		      Off Peak Reading</td>
				
		    
				<% end if %>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      2 Month Demand</td>

            <td bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">Previous 
		      Demand</td>
				
		    <td bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">Current 
		      Demand</td>
				
		    
		    </tr>
		    
		<% 
		    dim rst3, meterid, mySql, ticket
	        set rst3 = server.createobject("adodb.recordset")
		    mySql = "SELECT * from dbo.tickets where ticketfortype = 'meterid' AND [closed] = 0 AND Billyear ="&billyear&" AND billperiod = '" & billperiod & "'"
		    
		    rst3.open mySql, cnn1
		    do until rst3.eof
		        ticket = rst3("ticketfor")
		        ticket = split(ticket,"-")(1)
		 
		        response.write(ticketfor)
		        meterid = meterid + ", " + ticketfor
		        rst3.movenext
			loop
			
			response.write("meterid : " + meterid)
		    
		    do until rst2.eof %>
		    <tr <% %> bgcolor="#ffffff" valign="top" > 
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" nowrap><%=rst2("utility")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("meternum")%>&nbsp;</td>
				<td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("floor")%>&nbsp;</td>
				<td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("location")%>&nbsp;</td>
				<%
				if len(rst2("tenenat")) > 22 then 
					tenantName = "" & left(rst2("tenenat"),20) & "..."
				else
					tenantName =  rst2("tenenat")  
				end if
				%>
				<td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=tenantName%>&nbsp;</td>
              <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawprevious2")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawprevious")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">&nbsp;</td>
		      <% if extended then %>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawpreviousint2")%>&nbsp;</td>
              <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawpreviousint")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawpreviousoff2")%>&nbsp;</td>
              <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawpreviousoff")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">&nbsp;</td>
		      
			  <% end if %> 
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawprev2")%>&nbsp;</td>
              <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;"><%=rst2("rawprev")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeeccc" class="tblunderline" style="border:1px solid #000000;">&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;">&nbsp;</td>
		    </tr>
		    <tr bgcolor="#ffffff" valign="top">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>Notes: </b>  
		            </td>
		    </tr>
		    <% if( rst2("reader_notes") <> null OR rst2("reader_notes")<>"") then  %>
		        <tr bgcolor="#ffffff" valign="top">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>Reader Notes: <%=rst2("reader_notes")%></b>  
		            </td>
		        </tr>
		    <%end if %>
		    <% if( rst2("locationnotes") <> null OR rst2("locationnotes")<>"") then  %>
		        <tr bgcolor="#ffffff" valign="top">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>Location Notes: <%=rst2("locationnotes")%></b>  
		            </td>
		        </tr>
		    <%end if %>
            <% if( rst2("cnoteprev") <> null OR rst2("cnoteprev")<>"") then  %>
		        <tr bgcolor="#ffffff" valign="top">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>Consumption Notes: <%=rst2("cnoteprev")%></b>  
		            </td>
		        </tr>
		    <%end if %>
            <% if( rst2("dnoteprev") <> null OR rst2("dnoteprev")<>"") then  %>
		        <tr bgcolor="#ffffff" valign="top">
		            <td colspan=<%if extended then%>17<%else%>10<%end if%> bgcolor="#eeefff" class="tblunderline" style="border:1px solid #000000;" >
		              <b>Demand Notes: <%=rst2("dnoteprev")%></b>  
		            </td>
		        </tr>
		    <%end if %>
          <%  
		    rst2.movenext
			loop
		    %>
		    <tr>
		        <td height="40" valign="bottom" align="right" bgcolor="#eeefff" class="tblunderline" colspan=<%if extended then%>17<%else%>10<%end if%> ><b>Print Name _________________ Signature _________________ Date ___________</b></td>
		    </tr>
		    <%
			else 
			%>
			<tr><td colspan=<%if extended then%>17<%else%>10<%end if%>>NO METERS FOUND FOR BUILDING <%=currentbldg%></td></tr>
			<%
			end if
			rst2.close
			%>	    
        <div style="page-break-before:always" />
        <%
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
