<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim notoolbar
if not(allowGroups("Genergy Users,clientOperations")) then
notoolbar = 1
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, edit,masterticketid,opentickets,criticalopentickets, totaltickets,bldgid, ticketcount,sqlstr

pid = request("pid")
bldg = request("bldg")
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
		'pid = rst1("portfolioid")
		region = rst1("region")
		'portfolioname = rst1("name")
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


'if portfolioname="" then
rst1.Open "select name from portfolio where id='" & pid & "'", getConnect(0,0,"billing")
  if not rst1.EOF then
    portfolioname = rst1("name")
  end if
rst1.close
'end if

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
  document.location.href = 'tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid='+tid;
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
              </span> </td>
            <td width="50%" align="right" nowrap><select name="select" onChange="JumpTo(this.value)">
                <option value="#" selected>Jump to...</option>
                <option value="/genergy2/billing/processor_select.asp">Bill Processor</option>
                <% if (isBuildingTransfered(pid, bldg) = 0) then %>
                <option value="../validation/re_index.asp">Review Edit</option>
                <option value="/genergy2/manualentry/entry_select.asp">Manual Entry</option>
                <option value="/genergy2/billentry/entry.asp">Utility Bill Entry</option>
                <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem Report</option>
                <option value="/genergy2/accounting_files/historic_acctFile.asp">Accounting Transactions</option>
                <% end if %>
              </select> <%if (not isBuildingOff(bldg) AND (isBuildingTransfered(pid, bldg) = 0)) then ticket.MakeButton%>
              <button id="qmark2" onClick="openCustomWin('help.asp?page=buildingedit','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) 
              Quick Help</button></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #999999"> 
        <% if (edit) or trim(bldg)="" then %>
        <table border=0 cellpadding="3" cellspacing="0">
          <%if trim(bldg)="" then%>
          <tr>
            <td colspan="4"><b>Enter number of interval meters required for new 
              building</b></td>
          </tr>
          <tr>
            <td align="right"></td>
            <td colspan="4">15 minute meters&nbsp;
              <input name="Submeters15" type="text" value="" size="4">
              &nbsp;&nbsp;1 minute meters&nbsp;
              <input name="Submeters1" type="text" value="" size="4"></td>
          </tr>
          <!--   <tr><td align="right">ERI</td>
      <td><input type="Checkbox" name="submetered" value="1"></td>
  </tr> -->
          <tr>
            <td colspan="8"><hr size="1" noshade></td>
          </tr>
          <%end if%>
          <tr> 
            <td align="right"><span class="standard">Building&nbsp;Name</span></td>
            <td><input type="text" name="bldgname" value="<%=bldgname%>" />&nbsp;Offline&nbsp;
            <input type="checkbox" name="offline" <% if isBuildingOff(bldg) then %> checked <% end if %> /></td>
            <!-- <%if isBuildingOff(bldg) then%>onclick="if(this.checked){document.all['updatebutton'].style.display='none';}else{document.all['updatebutton'].style.display='inline';}"<%end if%> value="1" <%if "True"=offline then response.write "CHECKED"%> -->
            <td width="30">&nbsp;</td>
            <td  align="right"><span class="standard">Building&nbsp;#</span></td>
            <td>
              <% if trim(bldg)="" then %>
              <input type="text" name="bldg" value="<%=bldg%>">
              <%else%>
              <%=bldg%>
              <input type="hidden" name="bldg" value="<%=bldg%>">
              <% end if %>
            </td>
            <td width="30">&nbsp;</td>
            <td align="right" class="standard">Billing&nbsp;Name</td>
            <td><input type="text" name="btbldgname" value="<%=btbldgname%>"></td>
            <td align="right" class="standard">Contact Name</td>
            <td><input type="text" name="bldgcontactname" value="<%=ContactName%>"/></td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td align="right"><span class="standard">Address</span></td>
            <td><input type="text" name="address" value="<%=address%>"></td>
            <td width="30">&nbsp;</td>
            <td align="right"><span class="standard">Region</span></td>
            <td> 
              <% 
    rst1.open "SELECT * FROM regions ORDER BY city", getConnect(pid,bldg,"dbCore")
    %>
              <select name="region">
                <% do until rst1.eof %>
                <option value="<%=rst1("id")%>"<%if trim(region)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("city")%></option>
                <%
    rst1.movenext
    loop %>
              </select> 
              <% rst1.close %>
            </td>
            <td width="30">&nbsp;</td>
            <td align="right"><span class="standard">Billing&nbsp;Address</span></td>
            <td><textarea cols="25" rows="2" name="btstrt" wrap="off" style="overflow-x: scroll;"><%=btstrt%></textarea></td>
            <td align="right" class="standard" valign="top">Contact Phone</td>
            <td valign="top"><input type="text" name="bldgcontactphone" value="<%=ContactPhone%>"/></td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td align="right"><span class="standard">City</span></td>
            <td><input type="text" name="city" value="<%=city%>"></td>
            <td width="30">&nbsp;</td>
            <td align="right"><span class="standard">Square&nbsp;Footage</span></td>
            <td> <table border=0 cellpadding="0" cellspacing="0">
                <tr valign="middle"> 
                  <td><input type="text" name="sqft" size="10" maxlength="10" value="<%=sqft%>"></td>
                  <td width="4"><span class="standard">&nbsp;</span></td>
                  <td><span class="standard">SQFT</span></td>
                </tr>
              </table></td>
            <td width="30">&nbsp;</td>
            <td align="right"><span class="standard">Billing&nbsp;City</span></td>
            <td><input type="text" name="btcity" value="<%=btcity%>"></td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td align="right"><span class="standard">State</span></td>
            <td> <table border=0 cellpadding="0" cellspacing="0">
                <tr valign="middle"> 
                  <td><input type="text" name="state" size="4" value="<%=state%>"></td>
                  <td width="12"><span class="standard">&nbsp;</span></td>
                  <td><span class="standard">Zip Code&nbsp;</span></td>
                  <td><input type="text" name="zip" size="10" value="<%=zip%>"></td>
                </tr>
              </table></td>
            <td width="30">&nbsp;</td>
            <td width="30">Facility&nbsp;Type</td>
            <td width="30"> 
              <%rst1.open "SELECT * FROM facilityType ORDER BY description", getConnect(pid,bldg,"billing")%>
              <select name="facilityType">
                <% do until rst1.eof %>
                <option value="<%=rst1("id")%>"<%if trim(facilityType)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("Description")%></option>
                <%
    rst1.movenext
    loop %>
              </select> 
              <% rst1.close %>
            </td>
            <td width="30">&nbsp;</td>
            <td align="right"><span class="standard">Billing&nbsp;State</span></td>
            <td> <table border=0 cellpadding="0" cellspacing="0">
                <tr valign="middle"> 
                  <td><input type="text" name="btstate" size="4" value="<%=btstate%>"></td>
                  <td width="12"><span class="standard">&nbsp;</span></td>
                  <td><span class="standard">Zip Code&nbsp;</span></td>
                  <td><input type="text" name="btzip" size="10" value="<%=btzip%>"></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td>&nbsp;</td>
            <td colspan="4"> 
			
              <%= trim(bldg)%>
			  <% 'response.end%>
			  <% if (not isBuildingOff(bldg) OR allowGroups("IT Services") ) then
			      if trim(bldg)<>"" and (edit) then%>
                  <input id="updatebutton" type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;<%if (isBuildingOff(bldg) AND not allowGroups("IT Services")) then%>display:none<%end if%>"> 
                  <!--[[input type="submit" name="action" value="Delete" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"]]-->
                  <input type="button" name="action" value="Cancel" onClick="history.back();" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
                  <%else %>
                  <input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
                  <input type="button" name="action" value="Cancel" onClick="history.back();" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
                  <%end if
               end if%>
            </td>
          </tr>
        </table>
        <% else %>
        <table border=0 cellpadding="3" cellspacing="0" width="100%">
          <% if criticalopentickets <> 0 then %>
          <tr align="center"> 
            <td colspan =2 bgcolor="#FF0000"><b><a href="#" onClick="window.open('/genergy2_intranet/itservices/ttracker/troublesearch.asp?searchstring=<%=bldg%>&searchbox=True&action=Search&buildings=True','SearchNotes','width=800,height=400, scrollbars=no')"><%=criticalopentickets%> 
              Open Ticket(s) for this building that will effect Reading and Billing.</a></b> 
            </td>
          </tr>
          <% end if %>
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
		<%'if not(isBuildingOff(bldg)) then%>
			<input type="button" name="action" value="Edit Building Information" onClick="document.location='buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&edit=1';" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;margin:3px;">
		<%'end if%>
      </td>
    </tr>
    <% if (isBuildingTransfered(pid, bldg) = 0) then %>
    <tr bgcolor="#eeeeee"> 
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"> 
        <table border="0" cellpadding="3" cellspacing="0">
          <tr valign="top"> 
            <td> 
              <!-- begin core component links -->
              <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:billperiodView();" class="mgmtlink">Manage bill periods for this building</a><br>
			  <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:document.location='groupView.asp?pid=<%=pid%>&bldg=<%=bldg%>'" class="mgmtlink">Manage building groups</a><br>
			  <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="javascript:document.location='contactView.asp?pid=<%=pid%>&bldg=<%=bldg%>'" class="mgmtlink">Manage building contacts</a><br>
              <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('bsRatePicker.asp?pid=<%=pid%>&buildingNum=<%=bldg%>&view1970=false','NoteManager','width=500,height=375, scrollbars=yes');" class="mgmtlink">Building Specific Invoice Amount</a><br>
			</td>
            <td width="8">&nbsp;</td>
			<td>
			<%if not(isBuildingOff(bldg)) then%>
			  <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('AddonFeeSetup.asp?pid=<%=pid%>&bldg=<%=bldg%>','AddonFeeSetup','width=460,height=260, scrollbars=no');" class="mgmtlink">Setup Global Addon Fees</a><br>
			  <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('AutomationSetup.asp?pid=<%=pid%>&bldg=<%=bldg%>','BillProcessSetup','width=265,height=375, scrollbars=no');" class="mgmtlink">Bill Processor Setup</a><br> 
			<%end if%>
			  <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('/genergy2/UMreports/reportUM.asp?pid=<%=pid%>&bldgid=<%=bldgid%>','Reporter','width=800,height=600, scrollbars=yes,resizable=yes');" class="mgmtlink">Utility Manager Reporter</a><br> 
              <%
	'need to find out if anyone is on the 1970 rate and display the link if they are.
	dim sql_1470, rst_1470
	sql_1470 = "select distinct tl.bldgnum as building from tblleasesutilityprices tlup inner join tblleases tl on tl.billingId = tlup.billingId where tlup.rateTenant = '16' and tl.bldgnum = '"&bldg&"'"
	set rst_1470 = server.createobject("ADODB.RECORDSET")
	rst_1470.open sql_1470, cnn1
	if not rst_1470.eof then 'a record exists, that means there is a 1470 tenant out there... somewhere
		%>
              <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('bsRatePicker.asp?pid=<%=pid%>&buildingNum=<%=bldg%>&view1970=true','NoteManager','width=500,height=300, scrollbars=yes');" class="mgmtlink">Edit 1970 rate increase</a><br> 
              <%
	end if
	%>
      <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('gatewaysetup.asp?bldgnum=<%=bldg%>&select=Yes','getawaysetup','width=400,height=210, scrollbars=no');" class="mgmtlink">Gateway Setup</a><br>
	          <!-- end core component links -->
              <%
    if trim(customsrc)<>"" then
      response.write "*Contains Custom fields"
    end if
    
    rst1.open "SELECT * FROM custom_links WHERE code=1 and unitid='"&pid&"'", cnn1
    if not(rst1.eof) and not(isBuildingOff(bldg)) then %>
            </td>
            <td width="8">&nbsp;</td>
            <td> 
              <% 
    do while not rst1.eof
      response.write "<img src=""images/aro-rt.gif"" align=""absmiddle""  hspace=""2"" border=""0""><a href=""javascript:openCustomWin('"&rst1("link")&"?pid="&pid&"&bldg="&bldg&"','customlink', 'width="&rst1("width")&",height="&rst1("height")&"');"" class=""custlink"">"&rst1("label")&"</a><br>"
      rst1.movenext
      loop
      end if
    %>
            </td>
             <td width="8">&nbsp;</td>
            <td><img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('/genergy2/eri_th/meterservices/MeterMaintenanceLetter.asp?pid=<%=pid%>&bldgNum=<%=bldg%>','MaintenanceLetter','width=450,height=375, scrollbars=no');" class="mgmtlink">Maintenance Letter</a><br>
            <!--****-->
            <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0"><a href="#" onClick="open('/genergy2/eri_th/meterservices/AddNewDataSource.asp?pid=<%=pid%>&bldgNum=<%=bldg%>&bldgname=<%=bldgname%>','AddDataSource','width=450,height=375, scrollbars=no');" class="mgmtlink">Add New DataSource</a><br> 
           <!--****-->
             <img src="images/aro-rt.gif" align="absmiddle" hspace="2" border="0" ><a href="#" onClick="open('/genergy2/setup/accountingFileSetup.asp?pid=<%=pid%>&bldgNum=<%=bldg%>&bldgname=<%=bldgname%>', 'AccoutingFileSetup', 'width=450,height=450, scrollbars=no');" class="mgmtlink">Accounting File Setup</a><br />
            </td>
          </tr>
        </table>
        <% 
rst1.close
%>
        <% end if %>
      </td>
    </tr>
 <% end if %>
    <tr bgcolor="#eeeeee">
	<td valign="middle" align="left">
		<% if not(edit) and trim(bldg)<>"" then ticket.Display pid,true, true, true%>
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
    <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;" align="right"><%if (not isBuildingOff(bldg) AND (isBuildingTransfered(pid, bldg) = 0))  then%><input type="button" value="Add Account" onClick="tenantEdit('');" id=1 name=1 class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%>&nbsp;</td>
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
      		<td width="40%"><%=fonttag%><span class="standard"><%=rst1("billingname")%><%if rst1("tickets") <> "" then %> <font color="#FF0000">[<%=rst1("tickets")%> critical ticket(s)]</font><%end if%></span><%=unfonttag%></td>
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
<td bgcolor="#dddddd" width="12%"><%if (not(isBuildingOff(bldg)) AND (isBuildingTransfered(pid, bldg) = 0)) then%><input type="button" value="Add Account" onClick="tenantEdit('');" id=1 name=1 class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><% end if%>&nbsp;</td><td bgcolor="#dddddd"><font color="#555555"><i>Offline Tenants</i></font></td></tr>
</table>
<% end if%>
</form>
</body>
</html>
