<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--#include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
dim notoolbar, ccolor
if not(allowGroups("Genergy Users,clientOperations")) then
notoolbar = 1
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
'*****************************************
'12/28/2007 N.Ambo added function checkValues for validity check for portfolio #
'*****************************************

dim cnn1, rst1,rst2, strsql, ticketcount, masterticketid,opentickets, criticalopentickets, totaltickets
dim pid, edit, strsql2

pid = secureRequest("pid")
edit = secureRequest("edit")
if edit="" then edit = false

set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,0,"billing")


dim portfolio, portfolioname, billtemplate, paymenttext, offline
if trim(pid)<>"" then
	rst1.Open "SELECT * FROM portfolio WHERE id='"&pid&"'", cnn1
	if not rst1.EOF then
		portfolio = rst1("portfolio")
		portfolioname = rst1("name")
		paymenttext = rst1("paymentterm")
		billtemplate = rst1("templateid")
		offline = rst1("offline")
	end if
	rst1.close
end if
if not isnumeric(billtemplate) then billtemplate = 0

dim ticket
set ticket = New tickets
'ticket.Label="Portfolio"
'ticket.Note = "Master Ticket for Portfolio ID " & pid
'ticket.ccuid  = "rbdept"
'ticket.client = 1
if pid<>"" then ticket.findtickets "portfolioid", pid
%>
<html>
<head>
<title>Portfolio Edit</title>
<script language="JavaScript">
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function buildingEdit(bldg)
{	
<% if notoolbar then %>
//  if (parent.frames.length > 2) {
  	parent.toolbarfrm.location = 'toolbar.asp?pid=<%=pid%>&bldg='+bldg;
    parent.contentfrm.location = 'buildingedit.asp?pid=<%=pid%>&bldg='+bldg;
//  } else {
<% else %>
    document.location = 'buildingedit.asp?pid=<%=pid%>&bldg='+bldg;
<% end if %>
//  }
}
//12/28/2007 N.Ambo added this function to make the portfolio # field mandatory
function checkValues()
{	
	if (document.form2.portfolio.value == "")
	{		
		document.form2.action.value = "Nosave"
		alert("A value must be entered for Portfolio #.")		
	}	
}

</script>
<script src="/quickhelp/quickhelp.js" type="text/javascript" language="Javascript1.2"></script>
<style type="text/css">
.mgmtlink:hover { color:#3399cc; }
.custlink:hover { color:#339999; }
a.custlink { color:#006666; }
</style>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="portfoliosave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td>
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td>
    <span class="standardheader">
		<%if trim(pid)<>"" then%>
			<%if (edit) then%>
        Update
      <% else %>
        View
      <%end if%>
      Portfolio | <span style="font-weight:normal;"><%=portfolioname%></span>
    <%else%>
      Add New Portfolio
    <%end if%>
    </span>
    </td>
		<td align="right">
			<%if pid<>"" then ticket.MakeButton%>
			<button id="qmark2" onclick="openCustomWin('help.asp?page=portfolioedit','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button>
		</td>
  </tr>
  </table>
	</td>
</tr>
<tr bgcolor="#eeeeee">
  <td style="border-bottom:1px solid #999999">
  <% if (edit) or trim(pid)="" then %>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td align="right"><span class="standard">Portfolio Name</span></td> 
    <td>
    <input type="text" name="portfolioname" value="<%=portfolioname%>">
    </td>
    <td align="right"><span class="standard">Bill Template</span></td> 
    <td>
    <select name="billtemplate">
 
 
  <!--****************added line of code and rst2 to accomodate to drop down selection to default to genergy basic. 8/11/2008 - Michelle T. ***************-->
   
 <%  rst2.open "select * from billtemplates where name = 'genergy basic'",cnn1 %>
  <option value="<%=rst2("id")%>"<%if cint(billtemplate)="genergy basic" then response.write " SELECTED"%>><%="genergy basic"%></option>  
      rst2.close
 
  <%
      rst1.open "SELECT  * FROM billtemplates where name not in ('genergy basic') order by name", cnn1
    
      do until rst1.eof
     
        %>
        
     <option value="<%=rst1("id")%>"<%if cint(billtemplate)=cint(rst1("id")) then response.write " SELECTED"%>><%=rst1("name")%></option><%
        rst1.movenext
    loop
      rst1.close
    %>

    </select>
    </td>
  </tr>
  <tr>
    <td align="right"><span class="standard">Portfolio&nbsp;#</span></td>
    <td>
    <% if trim(pid)<>"" then%>
    <%=portfolio%>
    <% else %>
    <input type="text" name="portfolio" value="<%=portfolio%>">
    <% end if %>
    </td>  
	<td align="right">Payment Term</td>
	<td><input type="text" name="paymenttext" value="<%=trim(paymenttext)%>" size="50" maxlength="250"></td>
	<td>&nbsp;Offline&nbsp;<input type="checkbox" name="offline" value="1" <%if offline="True" then response.write "CHECKED"%>></td>
  </tr>
  <%if trim(pid)="" then%>
  <tr><td colspan="5">Number of Buildings in Portfolio</td></tr>
  <tr><td></td><td colspan="4">
      <table>
      <tr><td>15 minute buildings</td><td><input name="buildings15" size="4"></td></tr>
      <tr><td>1 minute building</td><td><input name="buildings1" size="4"></td></tr>
      </table></td></tr>
  <%end if%>
  <tr> 	
    <td>&nbsp;</td>
    <td>
      <%if trim(pid)<>"" and (edit) then%>
        <input type="submit" name="action" value="Update" class="standard"  style="border:1px outset #ddffdd;background-color:ccf3cc;">
        <input type="button" name="cancel" value="Cancel" onclick="location='portfolioedit.asp?pid=<%=pid%>&edit=0';" class="standard"  style="border:1px outset #ddffdd;background-color:ccf3cc;">
      <%else%>
        <input type="submit" name="action" value="Save" onclick="checkValues()" class="standard"  style="border:1px outset #ddffdd;background-color:ccf3cc;">
        <input type="button" name="cancel" value="Cancel" onclick="location='portfolioview.asp';" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
      <%end if%>
    </td>
  </tr>
  </table>
  
  <% else %>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
            <td><b><%=portfolioname%> Portfolio&nbsp; (<%=portfolio%>)</b></td>
  </tr>
  </table>  
  <input type="button" name="action" value="Edit Portfolio" onclick="location='portfolioedit.asp?pid=<%=pid%>&edit=1';" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;margin:3px;">

  <% end if%>
	</td>
</tr>
<%
if trim(pid)<>"" then
%>
<tr bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td>
    <!-- begin core component links -->
    <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:document.location = 'groupView.asp?pid=<%=pid%>'" class="mgmtlink">Manage portfolio groups</a><br>
    <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:document.location = 'contactView.asp?pid=<%=pid%>'" class="mgmtlink">Manage portfolio contacts</a><br>
    <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:document.location= 'buildingTransfer.asp?pid=<%=pid%>'" class="mgmtlink">Transfer building</a>
    <!-- end core component links -->
    <%
    rst1.open "SELECT * FROM custom_links WHERE code=0 and unitid='"&pid&"'", cnn1
    do while not rst1.eof
      response.write "<img src=""images/aro-rt.gif"" align=""absmiddle""  hspace=""2"" border=""0""><a href=""javascript:openCustomWin('"&rst1("link")&"?pid="&pid&"','customlink', 'width="&rst1("width")&",height="&rst1("height")&"');"" class=""custlink"">"&rst1("label")&"</a><br>"
      rst1.movenext
    loop
    rst1.close
    %>
    </td>
  </tr>
  </table>
	</td>
</tr>
</table>
<% if trim(pid)<>"" then %>
<table width="100%" border=0 cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee">
      <td valign="middle" align="left" style="border-bottom:1px solid #999999;border-top:1px solid #999999;">
	  	<%ticket.Display pid,true, true, false%>
	  </td>
    </tr>
</table>
	<%end if %>
<input type="hidden" name="pid" value="<%=pid%>">
<%
  strsql = "SELECT * FROM buildings b left join regions on b.region=regions.id where portfolioid='" & pid & "'"
  'response.write strsql
  'response.end
  rst1.open strsql, cnn1'"select *, regions.city as rcity from buildings left join regions on buildings.region=regions.id where portfolioid='" & pid & "'", cnn1
  if not rst1.eof then %>
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee">
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Buildings In The <%=portfolio%> Portfolio</b></span></td>
      <td colspan="2" align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><input type="button" value="Add Building" onclick="buildingEdit('');" id=1 name=1 class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
    </tr>
		<tr valign="top" bgcolor="#dddddd">
			<td width="25%"><span class="standard"><b>Building Name</b></span></td>
			<td width="45%"><span class="standard"><b>Address</b></span></td>
			<td width="15%"><span class="standard"><b>Region</b></span></td>
			<td width="15%"><span class="standard"><b>Building Number</b></span></td>
		</tr>
		<% do until rst1.eof 
			ccolor = ""
			if isBuildingOff(rst1("bldgnum")) then ccolor="class=""grayout"""
			%>
		<tr valign="top" bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="buildingEdit('<%=rst1("bldgnum")%>');">
			<td <%=ccolor%>><span class="standard"><%=rst1("bldgname")%></span></td>
			<td <%=ccolor%>>
			<span class="standard">
			<% if not rst1("strt")="" then response.write rst1("strt") & "<br>"%>
			<%=rst1("city")%><%if not (rst1("state")="" or rst1("zip")="") then response.write ",&nbsp;" end if%><%=rst1("state")%> &nbsp;<%=rst1("zip")%>
			</span>
			</td>
			<td <%=ccolor%>><span class="standard"><%=rst1("city")%></span></td>
			<td <%=ccolor%>><span class="standard"><%=rst1("bldgnum")%></span></td>
		</tr>
		<% rst1.movenext
		loop %>
    </table>
    <% else %>
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr bgcolor="#eeeeee">
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Buildings In The <%=portfolio%> Portfolio</b></span></td>
      <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><input type="button" value="Add Building" onclick="buildingEdit('');" id=1 name=1 class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
    </tr>
    <tr>
      <td colspan="2">No buildings have been set up yet in this portfolio.</td>
    </tr>
    </table>
  <% end if
  rst1.close%>
  <%
  strsql = "SELECT DISTINCT b.*,regions.city FROM dbo.view_buildings b left join regions on b.region=regions.id where oldPid='" & pid & "' AND endDate IS NOT NULL"
  rst1.open strsql, cnn1 '"select *, regions.city as rcity from buildings left join regions on buildings.region=regions.id where portfolioid='" & pid & "'", cnn1
  if not rst1.eof then %>
    <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee">
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Transfered Buildings In The <%=portfolio%> Portfolio</b></span></td>
      <td colspan="4" align="right" style="border-top:1px solid #ffffff; border-bottom:1px solid #999999;"></td>
    </tr>
		<tr valign="top" bgcolor="#dddddd">
			<td width="15%"><span class="standard"><b>Building Name</b></span></td>
			<td width="35%"><span class="standard"><b>Address</b></span></td>
			<td width="15%"><span class="standard"><b>Region</b></span></td>
			<td width="10%"><span class="standard"><b>Building Number</b></span></td>
			<td width="15%"><span class="standard"><b>Transfered To</b></span></td>
		    <td width="10%"><span class="standard"><b>Transfered Date</b></span></td>
		</tr>
		<% do until rst1.eof 
			ccolor = ""
			if isBuildingOff(rst1("bldgnum")) then ccolor="class=""grayout"""
			%>
		<tr valign="top" bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="buildingEdit('<%=rst1("bldgnum")%>');">
			<td <%=ccolor%>><span class="standard"><%=rst1("bldgname")%></span></td>
			<td <%=ccolor%>>
			<span class="standard">
			<% if not rst1("strt")="" then response.write rst1("strt") & "<br>"%>
			<%=rst1("city")%><%if not (rst1("state")="" or rst1("zip")="") then response.write ",&nbsp;" end if%><%=rst1("state")%> &nbsp;<%=rst1("zip")%>
			</span>
			</td>
			<td <%=ccolor%>><span class="standard"><%=rst1("city")%></span></td>
			<td <%=ccolor%>><span class="standard"><%=rst1("bldgnum")%></span></td>
			<%   strsql2 = "SELECT p.name FROM dbo.view_buildings b LEFT JOIN portfolio p ON p.id = newPid WHERE endDate IS NULL AND oldPid='" & pid & "' AND bldgNum ='" + rst1("bldgNum") + "'" 
			     rst2.open strsql2, cnn1
             %>
            <td <%=ccolor%>><span class="standard"><%=rst2("name")%></span></td>
            <% rst2.close %>
			<td <%=ccolor%>><span class="standard"><%=rst1("endDate")%></span></td>
		</tr>
		<% rst1.movenext
		loop 
		end if
		rst1.close
		%>
    </table>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr><td height="10"></td></tr>
<tr>
  <td bgcolor="#dddddd"><input type="button" value="Add Building" onclick="buildingEdit('');" id=1 name=1 class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
</tr>
</table>
<%
end if
%>
</form>
</body>
</html>

