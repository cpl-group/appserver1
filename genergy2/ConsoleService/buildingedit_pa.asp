<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim notoolbar

dim pid, bldg, edit,masterticketid,opentickets,criticalopentickets, totaltickets,bldgid, ticketcount,sqlstr

pid = request("pid")
bldg = request("bldgnum")
edit = request("edit")

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
if trim(bldg)<>"" then cnn1.open getLocalConnect(bldg)

dim bldgname, address, city, state, zip, sqft, action, region, portfolioname, customsrc, facilityType, btbldgname, btstrt, btcity, btstate, btzip, offline, ContactName, ContactPhone

if trim(bldg)<>"" then
	sqlstr="SELECT b.id as bldgid, * FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'"
	rst1.Open sqlstr, cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		address = rst1("strt")
		city = rst1("city")
		state = rst1("state")
		zip = rst1("zip")
		sqft = rst1("sqft")
		pid = rst1("portfolioid")
		region = rst1("region")
		portfolioname = rst1("name")
		customsrc = rst1("customsrc")
		facilityType = rst1("facilityType")
		btbldgname = rst1("btbldgname")
		btstrt = rst1("btstrt")
		btcity = rst1("btcity")
		btstate = rst1("btstate")
		btzip = rst1("btzip")
		bldgid = rst1("bldgid")
		offline = rst1("offline")
		ContactName = rst1("ContactName")
		ContactPhone = rst1("ContactPhone")
	end if
	rst1.close
end if


if portfolioname="" then
rst1.Open "select name from portfolio where id='" & pid & "'", getConnect(pid,bldg,"billing")
  if not rst1.EOF then
    portfolioname = rst1("name")
  end if
rst1.close
end if

dim ticket
set ticket = New tickets
ticket.Label="Building"
ticket.Note = "Master Ticket for Building " & bldg
ticket.ccuid  = "rbdept"
ticket.client = 1
if bldg<>"0" then ticket.findtickets "bldgnum", bldg
%>
<html>
<head>
<title>Building View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}

function tenantEdit(tid)
{	
<% if notoolbar then %>
//  if (parent.frames.length > 2) {
    parent.toolbarfrm.location = 'toolbar.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid='+tid;
<% end if %>
//  }
  document.location.href = 'tenantedit_pa.asp?pid=<%=pid%>&edit=0&bldg=<%=bldg%>&tid='+tid;
}

function billperiodView()
{	document.location.href = 'billPeriodView.asp?pid=<%=pid%>&bldg=<%=bldg%>';
}

function toggleInfoDisplay(util){
  utildiv = util + "info";
  utilbutton = util + "button"
  if (document.all[utildiv].style.display == "none") {
    document.all[utildiv].style.display = "inline";
    document.all[utilbutton].value = "Hide Info";
    //document.all[utilbutton].style.border = "1px inset #ffffff"
  } else {
    document.all[utildiv].style.display = "none";
    document.all[utilbutton].value = "Show Info";
    //document.all[utilbutton].style.border = "1px outset #ffffff"
  }
}
function JumpTo(url){
	var frm = document.forms['form1'];
	var url = url + "?pid=<%=pid%>&bldg=<%=bldg%>&building=<%=bldg%>&utilityid=2";
	window.document.location=url;
}
</script>
<style type="text/css">
.mgmtlink:hover { color:#3399cc; }
.custlink:hover { color:#339999; }
a.custlink { color:#006666; }
</style>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="buildingsave.asp">
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr> 
      <td bgcolor="#000000"> </td>
    </tr>
    <tr bgcolor="#3399cc"> 
      <td> <table border=0 cellpadding="0" cellspacing="0" width="100%">
          <tr> 
            <td width="53%"> <span class="standardheader"> 
              <%if trim(bldg)<>"" then%>
              Update Building | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> 
              &gt; <%=bldgname%></span> 
              <%else%>
              Add New Building | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> 
              <%end if%>
              </span> 
			 </td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #999999"> 
        <% if (edit) or trim(bldg)="" then %>
        <!-- Here exists the Code For the New Building And Update Building-- By Rahul Aggarwal--> 
        <% else %>
        <table border=0 cellpadding="3" cellspacing="0" width="100%">
          <tr valign="top"> 
            <td><b <%if isBuildingOff(bldg) then%>class="grayout"<%end if%>>Building # <%=bldg%> (<%=bldgname%>)</b></td>
            <td align="right"> <input type="button" id="buildingbutton" value="Hide Info" onClick="toggleInfoDisplay('building');" class="standard" style="cursor:hand;text-decoration:underline;background-color:#eeeeee;border:0;color:336699;">
              &nbsp; </td>
          </tr>
        </table>
        <div id="buildinginfo" style="display:inline;border:1px solid #cccccc;width:100%;padding:3px;"> 
          <table border=0 cellpadding="3" cellspacing="0">
            <tr valign="top" bgcolor="#eeeeee"> 
              <td>Address:</td>
              <td> <%=address%><br> <%=city%>
                <% if city<>"" and (state<>"" or zip<>"") then response.write ", " end if%>
                <%=state%>&nbsp;<%=zip%> </td>
            </tr>
            <tr valign="top" bgcolor="#eeeeee"> 
              <td>Region:</td>
              <td> 
                <% 
    rst1.open "SELECT city FROM regions WHERE id='" & trim(region) & "'", getConnect(pid,0,"dbCore")
    if not rst1.EOF then response.write rst1("city")
    rst1.close
    %>
              </td>
            </tr>
            <tr valign="top" bgcolor="#eeeeee"> 
              <td><span class="standard">Size:</span></td>
              <td><%=sqft%>&nbsp;sqft</td>
            </tr>
            <tr valign="top" bgcolor="#eeeeee">
              <td>Status:</td>
              <td><%if isBuildingOff(bldg) then%>Offline<%else%>Online<%end if%></td>
            </tr>
          </table>
        </div>
	<% end if %>	
      </td>
    </tr>
  </table>
<input type="hidden" name="pid" value="<%=pid%>">
<%
if trim(bldg)<>"" then

	sqlstr="SELECT * FROM tblleases  left join  (select count(*) as tickets, ticketfor as bid from ["&Application("CoreIP")&"].dbCore.dbo.tickets where ticketfortype='tid' and ticketfor in (select '"&split(getBuildingIP(bldg),"\")(1)&"-' + ltrim(convert(varchar(10),billingid)) from tblleases) and closed=0 and (billyear<>'' and billperiod <> '') group by ticketfor) a on a.bid = '"&split(getBuildingIP(bldg),"\")(1)&"-' + ltrim(convert(varchar(10),tblleases.billingid)) WHERE bldgnum='"&bldg&"' order by billingname"
	
	rst1.Open sqlstr, cnn1
%>
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr bgcolor="#eeeeee">
    <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Accounts</b></span></td>
  </tr>
  </table>
  <% if not rst1.EOF then %>
    <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#dddddd">
      <td width="40%"><span class="standard"><b>Billing Name</b></span></td>
      <td width="20%"><span class="standard"><b>Account Number</b></span></td>
      <td width="20%"><span class="standard"><b>Floor</b></span></td>
      <td width="20%"><span class="standard"><b>sqft</b></span></td>
    </tr>
    </table>
    <!--[[div style="overflow:auto;width:100%;height:200px;border:1px solid #cccccc;"]]-->
    <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <%do until rst1.EOF
		dim fonttag, unfonttag
		if lcase(trim(rst1("leaseexpired"))) = "true" then
			fonttag = "<i><font color='#555555'"
			unfonttag = "</i></font>"
		end if%>
    	<tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" onClick="tenantEdit(<%=rst1("billingid")%>);">
      		<td width="40%"><%=fonttag%><span class="standard"><%=rst1("billingname")%></span><%=unfonttag%></td>
      		<td width="20%"><%=fonttag%><span class="standard"><%=rst1("tenantnum")%></span><%=unfonttag%></td>
      		<td width="20%"><%=fonttag%><span class="standard"><%=rst1("Flr")%></span><%=unfonttag%></td>
      		<td width="20%"><%=fonttag%><span class="standard"><%=rst1("sqft")%></span><%=unfonttag%></td>
    	</tr>
    	<%
		fonttag=""
		unfonttag=""
		rst1.movenext
	loop%>
    </table>
    <!--[[/div]]-->
  <% else %>
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr>
    <td colspan="2"><span class="standard">There are no accounts set up for this building.</span></td>
  </tr>
  </table>
  <% end if
  rst1.close
else %>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#eeeeee">
    <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Accounts</b></span></td>
  </tr>
<tr>
  <td><span class="standard">None</span></td>
</tr>
</table>
<% end if
%>
<%if trim(bldg)<>"" then%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr><td>&nbsp;</td></tr>
<tr>
	<td bgcolor="#dddddd"><font color="#555555"><i>Offline Tenants</i></font></td></tr>
</table>
<% end if%>
</form>
</body>
</html>
